# 本機網頁版駕照報名工具

網頁版由 React 前端與 FastAPI 後端組成，只監聽本機位址，不提供公開網路服務。

## 功能

- 響應式網頁報名表單
- 電話、Email、身分證檢查碼、日期與監理站驗證
- 使用 Fernet 對稱加密保存個資
- 後端背景啟動 Selenium
- 即時狀態更新、停止作業、防止重複執行
- 沒有可報名場次時立即結束
- 可選擇完成後保留 Chrome

## 一鍵啟動（Windows）

在 PowerShell 進入此目錄後執行：

```powershell
.\start.ps1
```

第一次執行會下載依賴。看到瀏覽器開啟後即可使用；回到 PowerShell 按 Enter 可停止本機服務。

## 手動啟動

終端機一：

```powershell
cd backend
$env:UV_CACHE_DIR = ".uv-cache"
uv sync
uv run uvicorn app.main:app --host 127.0.0.1 --port 8000
```

終端機二：

```powershell
cd frontend
npm install
npm run dev
```

開啟 `http://127.0.0.1:5173`。

## 個資安全

報名資料加密後保存在 `backend/.data/user_info.enc`，金鑰保存在同一台電腦的 `backend/.data/secret.key`，兩者皆被 Git 忽略。請勿分享 `.data` 目錄。

這種設計可避免個資直接以明文留在磁碟，但無法防止已能登入並完整存取該電腦帳號的攻擊者。若刪除 `secret.key`，原有加密資料將無法復原。

後端只綁定 `127.0.0.1`，請勿自行改成 `0.0.0.0` 對外開放。

## 測試與建置

```powershell
cd backend
uv run pytest

cd ..\frontend
npm run build
```
