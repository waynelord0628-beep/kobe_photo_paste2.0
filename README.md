# Kobe強強照片黏貼

本專案是離線本機版照片黏貼紀錄表工具，可匯入照片、排序、標註、填寫地點與說明，並列印成 A4 紀錄表。

## 功能

- 多張照片匯入
- 匯入前預覽、刪除、拖曳排序
- 主畫面照片排序、上下移動
- 個別填寫地點與說明
- 統一地點 / 統一說明，只填入空白欄位
- 照片標註：畫筆、箭頭、圓圈、文字、編號
- 箭頭、圓圈、文字、編號可選取、移動、調整大小
- 選取物件後可按物件上的 X 或鍵盤 Delete 刪除
- 同一張照片重新開啟編輯時，編號會接續該照片上次的下一號
- A4 列印，每頁 2 張照片
- 專案儲存 / 載入 JSON，照片與標註狀態會一起保存
- 透過 GitHub Releases 檢查並安裝新版

## 直接使用

已打包版本放在：

```text
dist/Kobe強強照片黏貼.exe
```

執行後會開啟本機桌面視窗，不會跳出瀏覽器分頁。

## 更新方式

程式內可按「檢查更新」。如果 GitHub Releases 有新版，程式會下載新版 exe，關閉目前版本後自動替換並重新啟動。

## 開發方式

主要程式在：

```text
src/photo35.html
```

可直接用瀏覽器打開 `src/photo35.html` 測試功能。

## 打包方式

需要 Python 與 PyInstaller：

```powershell
pip install -r requirements.txt
pyinstaller --noconsole --onefile --name "Kobe強強照片黏貼" --collect-all webview --add-data "src/photo35.html;." launcher.py
```

打包完成後，exe 會在：

```text
dist/Kobe強強照片黏貼.exe
```

## 檔案說明

- `src/photo35.html`：主程式，離線單頁 HTML app
- `launcher.py`：本機桌面視窗啟動器
- `requirements.txt`：打包所需 Python 套件
- `dist/Kobe強強照片黏貼.exe`：目前打包好的 Windows 執行檔
