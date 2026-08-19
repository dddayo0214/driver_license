# 自動化駕照報名工具

使用 Python、Tkinter 與 Selenium 建立的桌面工具，可保存報名資料、查詢監理站考試場次，並自動填寫報名表單。

## 功能

- 表單資料格式及日期驗證
- 身分證字號檢查碼驗證與欄位遮蔽
- 自動保存與載入本機資料
- 背景執行瀏覽器流程，GUI 不會因等待而凍結
- 執行狀態、停止按鈕、120 秒搜尋超時
- 執行期間防止重複啟動
- 可選擇完成後是否保留 Chrome
- 不含個資的本機執行紀錄

## 專案結構

```text
driver_license/
├── main.py                    # Tkinter 介面與流程協調
├── automation.py              # Selenium 報名流程
├── validation.py              # 表單及身分證驗證
├── storage.py                 # JSON 讀寫
├── stations.py                # 監理站設定
├── user_info.example.json     # 無真實個資的資料範例
├── test_validation.py         # 核心驗證測試
├── pyproject.toml
└── uv.lock
```

`user_info.json` 會在儲存表單時建立，其中含有個人資料，已由 `.gitignore` 排除。請勿將它傳給他人或提交至版本控制。

## 環境需求

- Python 3.13 以上
- Chrome
- [uv](https://docs.astral.sh/uv/)

## 安裝與執行

```powershell
uv sync
uv run main.py
```

第一次執行瀏覽器自動化時，webdriver-manager 可能需要連線下載相容的 ChromeDriver。

## 操作方式

1. 填寫資料並按「儲存資料」。
2. 選擇考試日期、區域及監理站。
3. 按「開始報名」。
4. 可在畫面下方查看執行狀態；需要中止時按「停止」。
5. 自動送出後，請在瀏覽器確認網站結果。

「清除」只清除畫面欄位，不會刪除磁碟中的 `user_info.json`。

## 測試

```powershell
uv run python -m unittest -v
```

測試不會開啟瀏覽器，也不會送出報名。

## 維護提醒

監理服務網若改版，可能需要更新 `automation.py` 中的 selector。程式優先使用語意或 ID 定位，並保留必要的備援 XPath。
