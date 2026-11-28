# Microsoft Store submission draft

## 繁體中文

在 Windows 11 檔案總管選取一份或多份文件，按右鍵選擇「轉換為 PDF 檔案」。程式會呼叫電腦上的 Word、PowerPoint 或 Excel，將 PDF 儲存在原檔旁。同名檔案會自動加上序號。

批次轉檔共用一個進度視窗，可查看每份文件的結果與失敗原因。支援 DOC、DOCX、ODT、PPT、PPTX、ODP、XLS、XLSX 和 ODS。需要另外安裝並啟用相應的 Microsoft Office 桌面應用程式。

## English

Select one or more documents in Windows 11 File Explorer and choose **Convert to PDF** from the context menu. The app uses your installed copy of Word, PowerPoint, or Excel and saves each PDF beside its source document. If that name already exists, the app adds a number.

Follow the batch in one progress window and select a file to see its result or error. Supported formats include DOC, DOCX, ODT, PPT, PPTX, ODP, XLS, XLSX, and ODS. The corresponding Microsoft Office desktop apps must be installed and activated separately.

## Notes for certification

On a Windows 11 PC with Microsoft Word and PowerPoint installed and activated, select a local DOCX and PPTX in File Explorer. Choose **Convert to PDF** in the context menu, including under **App extensions** when Windows groups extensions there. Verify that the progress window shows both files and that each PDF appears beside its source. Repeat with an existing output filename to check numbered output.

The runFullTrust capability enables the desktop conversion process to automate locally installed Microsoft Office through COM. The package uses default Windows file-system virtualization.

Privacy policy: https://kaoshou.github.io/right-click-office-to-pdf/privacy/

Developer website: https://github.com/kaoshou/right-click-office-to-pdf

Support: https://github.com/kaoshou/right-click-office-to-pdf/issues
