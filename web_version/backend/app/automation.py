import threading
import time
from collections.abc import Callable

from selenium import webdriver
from selenium.common.exceptions import ElementClickInterceptedException, NoAlertPresentException, TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


class RegistrationCancelled(Exception):
    pass


class NoAvailableSession(Exception):
    pass


class RegistrationBot:
    URL = "https://www.mvdis.gov.tw/m3-emv-trn/exm/locations#"

    def __init__(self, stop_event: threading.Event, status: Callable[[str], None], keep_browser: bool):
        self.stop_event = stop_event
        self.status = status
        self.keep_browser = keep_browser
        self.driver = None

    def _cancel_check(self):
        if self.stop_event.is_set():
            raise RegistrationCancelled("使用者已停止作業")

    def _click_first(self, locators, timeout=10):
        last_error = None
        for locator in locators:
            try:
                element = WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable(locator))
                try:
                    self._click_element(element)
                except ElementClickInterceptedException:
                    self.driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                    WebDriverWait(self.driver, 2).until(
                        lambda _driver: element.is_displayed() and element.is_enabled()
                    )
                    self.driver.execute_script("arguments[0].click();", element)
                return
            except (TimeoutException, ElementClickInterceptedException) as exc:
                last_error = exc
        raise TimeoutException("找不到可點擊的頁面按鈕") from last_error

    def _wait_until_unblocked(self, timeout=20):
        def page_is_unblocked(driver):
            overlays = driver.find_elements(By.CSS_SELECTOR, ".blockUI.blockOverlay")
            return not any(overlay.is_displayed() for overlay in overlays)

        try:
            WebDriverWait(self.driver, timeout).until(page_is_unblocked)
        except TimeoutException as exc:
            raise TimeoutException("網站載入遮罩超過 20 秒仍未消失") from exc

    def _click_element(self, element):
        self._wait_until_unblocked()
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        for attempt in range(3):
            try:
                WebDriverWait(self.driver, 5).until(lambda _driver: element.is_displayed() and element.is_enabled())
                element.click()
                return
            except ElementClickInterceptedException:
                if attempt == 2:
                    break
                self._wait_until_unblocked()
                time.sleep(0.25)
        self._wait_until_unblocked()
        self.driver.execute_script("arguments[0].click();", element)

    def _wait_for_query_response(self, timeout=15):
        dialog_locators = (
            (By.XPATH, "//div[contains(@class,'ui-dialog') and not(contains(@style,'display: none'))]//a[contains(normalize-space(.),'確定') or contains(normalize-space(.),'繼續') or contains(normalize-space(.),'同意')]"),
            (By.XPATH, "//body/div[not(contains(@style,'display: none'))]//center/a[3]"),
        )

        def response_ready(driver):
            try:
                return ("alert", driver.switch_to.alert)
            except NoAlertPresentException:
                pass
            rows = driver.find_elements(By.CSS_SELECTOR, "#trnTable tbody tr")
            if any(row.is_displayed() for row in rows):
                return ("results", None)
            for locator in dialog_locators:
                for element in driver.find_elements(*locator):
                    if element.is_displayed() and element.is_enabled():
                        return ("dialog", element)
            return False

        try:
            response_type, target = WebDriverWait(self.driver, timeout).until(response_ready)
        except TimeoutException as exc:
            raise TimeoutException("查詢送出後 15 秒內未收到確認視窗或場次結果") from exc
        if response_type == "alert":
            target.accept()
        elif response_type == "dialog":
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
            try:
                target.click()
            except ElementClickInterceptedException:
                self.driver.execute_script("arguments[0].click();", target)

    def _wait_for_signup_confirmation(self, timeout=15):
        dialog_locators = (
            (By.XPATH, "//div[contains(@class,'ui-dialog') and not(contains(@style,'display: none'))]//a[contains(normalize-space(.),'確定') or contains(normalize-space(.),'同意') or contains(normalize-space(.),'繼續') or contains(normalize-space(.),'報名')]"),
            (By.XPATH, "/html/body/div[11]/div[2]/a"),
            (By.XPATH, "//body/div[not(contains(@style,'display: none'))]//div[2]/a[contains(@class,'btn') or contains(@class,'button') or contains(@class,'std_btn')]"),
        )

        def signup_ready(driver):
            try:
                return ("alert", driver.switch_to.alert)
            except NoAlertPresentException:
                pass
            form = driver.find_elements(By.ID, "idNo")
            if any(element.is_displayed() for element in form):
                return ("form", None)
            for locator in dialog_locators:
                for element in driver.find_elements(*locator):
                    if element.is_displayed() and element.is_enabled():
                        return ("dialog", element)
            return False

        try:
            response_type, target = WebDriverWait(self.driver, timeout).until(signup_ready)
        except TimeoutException as exc:
            raise TimeoutException("點擊報名後 15 秒內未找到確認彈窗或個資表單") from exc
        if response_type == "alert":
            target.accept()
        elif response_type == "dialog":
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
            try:
                target.click()
            except ElementClickInterceptedException:
                self.driver.execute_script("arguments[0].click();", target)

    def run(self, data):
        options = Options()
        if self.keep_browser:
            options.add_experimental_option("detach", True)
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        try:
            self.status("正在啟動 Chrome…")
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            self.driver.get(self.URL)
            self._cancel_check()
            self.status("正在設定查詢條件…")
            wait = WebDriverWait(self.driver, 15)
            Select(wait.until(EC.presence_of_element_located((By.ID, "licenseTypeCode")))).select_by_visible_text(data["駕照類型"])
            Select(wait.until(EC.presence_of_element_located((By.ID, "dmvNoLv1")))).select_by_visible_text(data["目的地區"])
            Select(wait.until(EC.presence_of_element_located((By.ID, "dmvNo")))).select_by_visible_text(data["目的監理所"])
            date_input = wait.until(EC.element_to_be_clickable((By.ID, "expectExamDateStr")))
            date_input.clear()
            date_input.send_keys(data["考試日期"])
            date_input.send_keys(Keys.TAB)
            self.driver.execute_script("arguments[0].blur();", date_input)
            try:
                WebDriverWait(self.driver, 3).until(
                    EC.invisibility_of_element_located((By.ID, "ui-datepicker-div"))
                )
            except TimeoutException:
                self.driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
            self._click_first(((By.CSS_SELECTOR, "#form1 a.std_btn[onclick*='query']"), (By.XPATH, "//*[@id='form1']//a[contains(@onclick,'query')]")))
            self.status("查詢已送出，正在等待網站回應…")
            self._wait_for_query_response()
            self.status("正在查詢可報名場次…")
            self._wait_until_unblocked()
            rows = WebDriverWait(self.driver, 20).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "#trnTable tbody tr")))
            for row in rows:
                self._cancel_check()
                if "額滿" in row.text or "重考" in row.text:
                    continue
                links = row.find_elements(By.CSS_SELECTOR, "td:nth-child(4) a")
                if links:
                    self._click_element(links[0])
                    self.status("已選擇場次，正在處理報名確認…")
                    self._fill(data)
                    self.status("資料已送出，請在瀏覽器確認結果")
                    return
            raise NoAvailableSession("目前查詢到的場次皆已額滿或不可報名")
        except (RegistrationCancelled, NoAvailableSession, TimeoutException):
            raise
        except WebDriverException as exc:
            raise RuntimeError(f"瀏覽器操作失敗：{exc.msg}") from exc
        finally:
            if self.driver and (not self.keep_browser or self.stop_event.is_set()):
                try:
                    self.driver.quit()
                except WebDriverException:
                    pass

    def _fill(self, data):
        self._wait_for_signup_confirmation()
        self.status("正在填寫報名資料…")
        wait = WebDriverWait(self.driver, 15)
        for element_id, key in {"idNo": "身分證字號", "birthdayStr": "生日", "name": "姓名", "contactTel": "電話", "email": "電子郵件"}.items():
            element = wait.until(EC.element_to_be_clickable((By.ID, element_id)))
            element.clear()
            element.send_keys(data[key])
        self._cancel_check()
        self._click_first(((By.XPATH, "//*[@id='form1']//a[contains(normalize-space(.),'送出') or contains(normalize-space(.),'確定')]"), (By.CSS_SELECTOR, "#form1 table a.btn:first-of-type"), (By.XPATH, "//*[@id='form1']/table/tbody/tr[6]/td/a[1]")))
