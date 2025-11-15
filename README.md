🚗 自動化駕照報名工具（Auto License Registration Bot）

一個以 Python + Selenium + UV 建立的自動化駕照報名程式
可自動填寫學員資料，自動選取考試日期，自動送出報名資訊。
只需提前輸入個人資料，即可一鍵完成流程。

✨ 功能特色（Features）

🔄 全自動化流程：自動進入系統 → 填資料 → 選日期 → 提交。

🧾 支援事先設定資料（姓名、身分證、生日、電話、班別等）。

⚡ 使用 uv 管理環境，速度快、簡潔、依賴乾淨。

🌐 以 Selenium 驅動 Chrome，實際模擬使用者操作。

📁 可記錄日誌與錯誤資訊，方便排查。

🧩 模組化程式架構，方便擴充與維護。

📦 安裝（Installation）

Clone 專案：

git clone <your-repo-url>
cd auto-license-registration


使用 uv 建立虛擬環境並安裝依賴：

uv sync


uv 會依照 pyproject.toml 自動安裝所需套件（如 selenium、webdriver-manager 等）。

❓ 常見問題（FAQ）
1. ChromeDriver 版本不符？

你使用了 WebDriver Manager，因此會自動下載對應版本。

若仍出錯：

uv pip install --upgrade webdriver-manager

2. 報名網站改版導致失敗？

請更新：

selector

XPaths

填寫欄位名稱

我也能協助你更新程式。

3. 主程式找不到 config 檔？

請確認：

config/user_info.json


存在且格式正確。