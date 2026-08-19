# 自動化駕照報名工具

專案已分為兩個獨立版本：

- `python_version/`：原本的 Tkinter 桌面版。
- `web_version/`：FastAPI + React 本機網頁版。

兩個版本都在使用者自己的電腦執行，由 Selenium 開啟 Chrome 操作監理服務網。

駕照報名最後須手動填寫驗證碼。

## 桌面版

```powershell
cd python_version
uv sync
uv run main.py
```

詳細說明請見 [`python_version/README.md`](python_version/README.md)。

## 網頁版

```powershell
cd web_version
.\start.ps1
```

第一次啟動會安裝前後端依賴，完成後自動開啟 `http://127.0.0.1:5173`。
詳細說明請見 [`web_version/README.md`](web_version/README.md)。
