# Kobe強強照片黏貼

本機桌面版照片黏貼工具。可匯入照片、排序、填寫地點與說明、編輯標註，並輸出列印或 Word 模板。

## 功能

- 多張照片匯入與拖曳排序
- 匯入前預覽排序與刪除
- 地點、說明欄位與批次填入
- 編輯照片標註：筆、箭頭、圓圈、文字、編號
- 箭頭、圓圈、文字、編號可作為物件移動、調整大小、刪除
- `Delete` 鍵可刪除目前選取的標註物件
- 專案儲存與載入 JSON
- A4 列印模式
- Word 模板匯出
- GitHub Releases 檢查更新與自動替換新版 exe

## Word 模板

程式沿用舊版 Word 模板設定，並新增較密集的直式照片版型。

目前支援：

- `1頁2張橫式`：一頁上下 2 張，適合橫式照片
- `1頁2張直式`：一頁左右 2 張，適合直式照片
- `1頁3張橫式`：橫向 A4，一頁 3 張
- `1頁4張直式`：直式 A4，2 欄 x 2 排
- `1頁6張直式`：橫向 A4，3 欄 x 2 排

匯出時會依照目前畫面上的照片順序輸出，並使用照片卡片內的「說明」作為 Word 說明文字；若沒有說明，會改用地點；兩者都沒有時使用照片編號。

## 開發執行

```powershell
pip install -r requirements.txt
python launcher.py
```

## 打包 exe

```powershell
pip install -r requirements.txt
pyinstaller --noconsole --onefile --name "Kobe強強照片黏貼" --collect-all webview --add-data "src/photo35.html;." --add-data "templates;templates" launcher.py
```

輸出檔會在：

```text
dist/Kobe強強照片黏貼.exe
```

## 更新

程式內可按「檢查更新」。如果 GitHub Releases 有新版，程式會下載新版 exe，關閉目前版本後自動替換並重新啟動。

## 檔案說明

- `src/photo35.html`：主要前端介面
- `launcher.py`：桌面版啟動器與 pywebview API
- `template_export.py`：Word 模板匯出功能
- `templates/`：Word 模板檔
- `requirements.txt`：Python 依賴
