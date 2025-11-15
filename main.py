#coding = utf-8
import os
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

class LicenseRegistrationForm:
    def __init__(self, root):
        self.root = root
        self.root.title("駕照報名表單")
        self.root.geometry("500x650")
        self.root.resizable(False, False)
        self.data_file = os.path.join("driver_license", "driver_data.json")
        
        self.stations_data = {
            '臺北市區監理所（含金門馬祖）': ['士林監理站(臺北市士林區承德路5段80號)', '基隆監理站(基隆市七堵區實踐路296號)', '金門監理站(金門縣金湖鎮黃海路六之一號)', '連江監理站(連江縣南竿鄉津沙村155號)'],
            '臺北區監理所（北宜花）': ['臺北區監理所(新北市樹林區中正路248巷7號)', '板橋監理站(新北市中和區中山路三段116號)', '宜蘭監理站(宜蘭縣五結鄉中正路二段9號)', '花蓮監理站(花蓮縣吉安鄉中正路二段152號)', '玉里監理分站(花蓮縣玉里鎮中華路427號)', '蘆洲監理站(新北市蘆洲區中山二路163號)'],
            '新竹區監理所（桃竹苗）': ['新竹區監理所(新竹縣新埔鎮文德路三段58號)', '新竹市監理站(新竹市自由路10號)', '桃園監理站(桃園市介壽路416號)', '中壢監理站(桃園縣中壢市延平路394號)', '苗栗監理站(苗栗市福麗里福麗98號)'],
            '臺中區監理所（中彰投）': ['臺中區監理所(臺中市大肚區瑞井里遊園路一段2號)', '臺中市監理站(臺中市北屯路77號)', '埔里監理分站(南投縣埔里鎮水頭里水頭路68號)', '豐原監理站(臺中市豐原區豐東路120號)', '彰化監理站(彰化縣花壇鄉南口村中山路二段457號)', '南投監理站(南投縣南投市光明一路301號)'],
            '嘉義區監理所（雲嘉南）': ['嘉義區監理所(嘉義縣朴子市朴子七路29號)', '東勢監理分站(雲林縣東勢鄉新坤村新坤路333號)', '雲林監理站(雲林縣斗六市雲林路二段411號)', '新營監理站(臺南市新營區大同路55號)', '臺南監理站(臺南市崇德路1號)', '麻豆監理站(臺南市麻豆區北勢里新生北路551號)', '嘉義市監理站(嘉義市東區保建街89號)'],
            '高雄市區監理所': ['高雄市區監理所(高雄市楠梓區德民路71號)', '苓雅監理站(高雄市三民區建國一路454號)', '旗山監理站(高雄市旗山區旗文路123-1號)'],
            '高雄區監理所（高屏澎東）': ['高雄區監理所(高雄市鳳山區武營路361號)', '臺東監理站(臺東市正氣北路441號)', '屏東監理站(屏東市忠孝路222號)', '恆春監理分站(屏東縣恒春鎮草埔路11號)', '澎湖監理站(澎湖縣馬公市光華里121號)']
        }
        
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        title_label = ttk.Label(main_frame, text="駕照報名表單", font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        ttk.Label(main_frame, text="駕照類型:", font=('Arial', 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.license_type = ttk.Combobox(main_frame, width=30, state='readonly')
        self.license_type['values'] = ('普通重型機車', '普通輕型機車 (50cc 以下)', '普通小型車', '職業小型車', '普通大貨車', '職業大貨車', '普通大客車', '職業大客車', '普通聯結車', '職業聯結車')
        self.license_type.current(0)
        self.license_type.grid(row=1, column=1, pady=5, padx=(10, 0))
        
        ttk.Label(main_frame, text="姓名:", font=('Arial', 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.name_entry = ttk.Entry(main_frame, width=32)
        self.name_entry.grid(row=2, column=1, pady=5, padx=(10, 0))
        
        ttk.Label(main_frame, text="生日:", font=('Arial', 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
        birthday_frame = ttk.Frame(main_frame)
        birthday_frame.grid(row=3, column=1, pady=5, padx=(10, 0), sticky=tk.W)
        
        self.birth_year = ttk.Combobox(birthday_frame, width=8, state='readonly')
        self.birth_year['values'] = list(range(2010, 1919, -1))
        self.birth_year.current(10)
        self.birth_year.grid(row=0, column=0, padx=(0, 5))
        
        ttk.Label(birthday_frame, text="年").grid(row=0, column=1, padx=(0, 5))
        
        self.birth_month = ttk.Combobox(birthday_frame, width=5, state='readonly')
        self.birth_month['values'] = list(range(1, 13))
        self.birth_month.current(0)
        self.birth_month.grid(row=0, column=2, padx=(0, 5))
        
        ttk.Label(birthday_frame, text="月").grid(row=0, column=3, padx=(0, 5))
        
        self.birth_day = ttk.Combobox(birthday_frame, width=5, state='readonly')
        self.birth_day['values'] = list(range(1, 32))
        self.birth_day.current(0)
        self.birth_day.grid(row=0, column=4, padx=(0, 5))
        
        ttk.Label(birthday_frame, text="日").grid(row=0, column=5)
        
        ttk.Label(main_frame, text="電話:", font=('Arial', 10)).grid(row=4, column=0, sticky=tk.W, pady=5)
        self.phone_entry = ttk.Entry(main_frame, width=32)
        self.phone_entry.grid(row=4, column=1, pady=5, padx=(10, 0))
        
        ttk.Label(main_frame, text="電子郵件:", font=('Arial', 10)).grid(row=5, column=0, sticky=tk.W, pady=5)
        self.email_entry = ttk.Entry(main_frame, width=32)
        self.email_entry.grid(row=5, column=1, pady=5, padx=(10, 0))
        
        ttk.Label(main_frame, text="身分證字號:", font=('Arial', 10)).grid(row=6, column=0, sticky=tk.W, pady=5)
        self.id_entry = ttk.Entry(main_frame, width=32)
        self.id_entry.grid(row=6, column=1, pady=5, padx=(10, 0))
        
        ttk.Label(main_frame, text="考試日期:", font=('Arial', 10)).grid(row=7, column=0, sticky=tk.W, pady=5)
        exam_frame = ttk.Frame(main_frame)
        exam_frame.grid(row=7, column=1, pady=5, padx=(10, 0), sticky=tk.W)
        
        self.exam_year = ttk.Combobox(exam_frame, width=8, state='readonly')
        self.exam_year['values'] = list(range(2025, 2027))
        self.exam_year.current(0)
        self.exam_year.grid(row=0, column=0, padx=(0, 5))
        
        ttk.Label(exam_frame, text="年").grid(row=0, column=1, padx=(0, 5))
        
        self.exam_month = ttk.Combobox(exam_frame, width=5, state='readonly')
        self.exam_month['values'] = list(range(1, 13))
        self.exam_month.current(0)
        self.exam_month.grid(row=0, column=2, padx=(0, 5))
        
        ttk.Label(exam_frame, text="月").grid(row=0, column=3, padx=(0, 5))
        
        self.exam_day = ttk.Combobox(exam_frame, width=5, state='readonly')
        self.exam_day['values'] = list(range(1, 32))
        self.exam_day.current(0)
        self.exam_day.grid(row=0, column=4, padx=(0, 5))
        
        ttk.Label(exam_frame, text="日").grid(row=0, column=5)
        
        ttk.Label(main_frame, text="監理所區域:", font=('Arial', 10)).grid(row=10, column=0, sticky=tk.W, pady=5)
        self.destination_region = ttk.Combobox(main_frame, width=30, state='readonly')
        self.destination_region['values'] = list(self.stations_data.keys())
        self.destination_region.current(0)
        self.destination_region.bind('<<ComboboxSelected>>', self.update_destination_station)
        self.destination_region.grid(row=10, column=1, pady=5, padx=(10, 0))
        
        ttk.Label(main_frame, text="監理所:", font=('Arial', 10)).grid(row=11, column=0, sticky=tk.W, pady=5)
        self.destination_station = ttk.Combobox(main_frame, width=30, state='readonly')
        self.destination_station['values'] = self.stations_data['臺北市區監理所（含金門馬祖）']
        self.destination_station.current(0)
        self.destination_station.grid(row=11, column=1, pady=5, padx=(10, 0))
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=12, column=0, columnspan=2, pady=(30, 0))
        
        submit_btn = ttk.Button(button_frame, text="提交", command=self.save_data)
        submit_btn.grid(row=0, column=0, padx=5)

        start_btn = ttk.Button(button_frame, text="開始", command=self.spider_submit)
        start_btn.grid(row=0, column=2, padx=5)
        
        clear_btn = ttk.Button(button_frame, text="清除", command=self.clear_form)
        clear_btn.grid(row=0, column=1, padx=5)

        self.load_data()
    
    def update_destination_station(self, event):
        region = self.destination_region.get()
        stations = self.stations_data.get(region, [])
        self.destination_station['values'] = stations
        if stations:
            self.destination_station.current(0)
        
    def save_data(self):
        birthday = f"{int(self.birth_year.get()) - 1911}{int(self.birth_month.get()):02d}{int(self.birth_day.get()):02d}"
        exam_date = f"{int(self.exam_year.get()) - 1911}{int(self.exam_month.get()):02d}{int(self.exam_day.get()):02d}"
        
        data = {
            '駕照類型': self.license_type.get(),
            '姓名': self.name_entry.get(),
            '生日': birthday,
            '電話': self.phone_entry.get(),
            '電子郵件': self.email_entry.get(),
            '身分證字號': self.id_entry.get(),
            '考試日期': exam_date,
            '目的地區': self.destination_region.get(),
            '目的監理所': self.destination_station.get()
        }
        
        if not data['姓名']:
            messagebox.showwarning("警告", "請輸入姓名")
            return
        if not data['電話']:
            messagebox.showwarning("警告", "請輸入電話")
            return
        if not data['電子郵件']:
            messagebox.showwarning("警告", "請輸入電子郵件")
            return
        if not data['身分證字號']:
            messagebox.showwarning("警告", "請輸入身分證字號")
            return
            
        message = "\n".join([f"{key}: {value}" for key, value in data.items()])
        messagebox.showinfo("表單提交成功", f"已提交以下資料:\n\n{message}")

        save_data = {
            'form_data': data,
            'form_state': {
                'license_type_index': self.license_type.current(),
                'birth_year': self.birth_year.get(),
                'birth_month': self.birth_month.get(),
                'birth_day': self.birth_day.get(),
                'exam_year': self.exam_year.get(),
                'exam_month': self.exam_month.get(),
                'exam_day': self.exam_day.get(),
                'destination_region_index': self.destination_region.current(),
                'destination_station_index': self.destination_station.current()
            }
        }
        
        with open(self.data_file, 'w', encoding='utf-8') as f:
            json.dump(save_data, f, ensure_ascii=False, indent=2)

    def clear_form(self):
        self.name_entry.delete(0, tk.END)
        self.phone_entry.delete(0, tk.END)
        self.email_entry.delete(0, tk.END)
        self.id_entry.delete(0, tk.END)
        self.license_type.current(0)
        self.destination_region.current(0)
        self.destination_station['values'] = self.stations_data['臺北市區監理所（含金門馬祖）']
        self.destination_station.current(0)
        self.birth_year.current(10)
        self.birth_month.current(0)
        self.birth_day.current(0)
        self.exam_year.current(0)
        self.exam_month.current(0)
        self.exam_day.current(0)

    def spider_submit(self):
        self.save_data()
        with open(self.data_file, 'r', encoding='utf-8') as f:
            save_data = json.load(f)

        form_data = save_data.get('form_data', {})
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.get('https://www.mvdis.gov.tw/m3-emv-trn/exm/locations#')

        license_type = Select(driver.find_element(By.ID, 'licenseTypeCode'))
        license_type.select_by_visible_text(form_data['駕照類型'])

        place1 = Select(driver.find_element(By.ID, 'dmvNoLv1'))
        place1.select_by_visible_text(form_data['目的地區'])

        place2 = Select(driver.find_element(By.ID, 'dmvNo'))
        place2.select_by_visible_text(form_data['目的監理所'])
        
        date = driver.find_element(By.ID, 'expectExamDateStr')
        date.send_keys(form_data['考試日期'] + '\n')

        search = driver.find_element(By.XPATH, '//*[@id="form1"]/div/a')
        search.click()

        move = driver.find_element(By.XPATH, '/html/body/div[11]/div/center/a[3]')
        move.click()

        texts = driver.find_element(By.XPATH, '//*[@id="trnTable"]/tbody')
        texts = texts.find_elements(By.TAG_NAME, 'tr')
        for i in range(len(texts)):
            text = texts[i].text
            if '額滿' not in text and '重考' not in text:
                sighup = driver.find_element(By.XPATH, f'//*[@id="trnTable"]/tbody/tr[{i + 1}]/td[4]/a')
                sighup.click()

                signup2 = driver.find_element(By.XPATH, '/html/body/div[11]/div[2]/a')
                signup2.click()

                id = driver.find_element(By.ID, 'idNo')
                id.send_keys(form_data['身分證字號'])

                birday = driver.find_element(By.ID, 'birthdayStr')
                birday.send_keys(form_data['生日'] + '\n')

                name = driver.find_element(By.ID, 'name')
                name.send_keys(form_data['姓名'])

                phone_num = driver.find_element(By.ID, 'contactTel')
                phone_num.send_keys(form_data['電話'])
                
                email = driver.find_element(By.ID, 'email')
                email.send_keys(form_data['電子郵件'])

                over = driver.find_element(By.XPATH, '//*[@id="form1"]/table/tbody/tr[6]/td/a[1]')
                over.click()

    def load_data(self):
        if not os.path.exists(self.data_file):
            return
            
        with open(self.data_file, 'r', encoding='utf-8') as f:
            save_data = json.load(f)

        form_data = save_data.get('form_data', {})
        form_state = save_data.get('form_state', {})
            
        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, form_data.get('姓名', ''))
            
        self.phone_entry.delete(0, tk.END)
        self.phone_entry.insert(0, form_data.get('電話', ''))
            
        self.email_entry.delete(0, tk.END)
        self.email_entry.insert(0, form_data.get('電子郵件', ''))
            
        self.id_entry.delete(0, tk.END)
        self.id_entry.insert(0, form_data.get('身分證字號', ''))
            
        if 'license_type_index' in form_state:
            self.license_type.current(form_state['license_type_index'])
            
        if 'birth_year' in form_state:
            self.birth_year.set(form_state['birth_year'])
            self.birth_month.set(form_state['birth_month'])
            self.birth_day.set(form_state['birth_day'])
            
        if 'exam_year' in form_state:
            self.exam_year.set(form_state['exam_year'])
            self.exam_month.set(form_state['exam_month'])
            self.exam_day.set(form_state['exam_day'])
            
        if 'destination_region_index' in form_state:
            self.destination_region.current(form_state['destination_region_index'])
            self.update_destination_station(None)
            self.destination_station.current(form_state['destination_station_index'])
            
        messagebox.showinfo("成功", "已載入上次儲存的資料")

def main():
    root = tk.Tk()
    LicenseRegistrationForm(root)
    root.mainloop()

if __name__ == "__main__":
    main()