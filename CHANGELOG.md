# Changelog

## 1.1.0

- 多選文件或連續啟動轉檔時，共用一個進度視窗。每份文件都有狀態與結果。
- 由本機 Microsoft Word、PowerPoint 或 Excel 唯讀開檔並匯出 PDF。Word 與 PowerPoint 支援修復開啟及替代的 Office PDF 匯出方法。
- 單檔失敗會顯示原因並繼續處理後續文件。已取消第三方備援轉檔。
- 匯出檔通過 PDF 標頭與結尾檢查後才放到原檔旁，並避免同時轉換時覆寫同名檔案。
- 新增 Windows 11 x64 和 ARM64 原生右鍵命令、簽章發行封裝及對應的移除流程。

## 1.0.0

Initial stable release for `right-click-office-to-pdf`.

- Added Windows right-click conversion for DOC, DOCX, PPT, PPTX, XLS, XLSX, ODT, ODS, and ODP files.
- Added SendTo shortcut for reliable multi-file conversion.
- Added Microsoft Office first conversion strategy.
- Added LibreOffice / OpenOffice / OxOffice fallback through `soffice.exe`.
- Added one-line online install and uninstall scripts.
- Added user-level installation under `%LOCALAPPDATA%\RightClickOfficeToPDF`.
- Added user-level registry integration under HKCU only.
- Added safer PowerShell launcher path to preserve Unicode file paths more reliably.
- Added Chinese Windows menu name: `轉換為 PDF 檔案`.
