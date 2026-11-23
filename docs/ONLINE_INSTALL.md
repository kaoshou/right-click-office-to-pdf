# 線上安裝與移除

在要使用右鍵轉檔的帳號開啟 Windows PowerShell。請先確認本機安裝了要使用的 Microsoft Office 桌面應用程式，例如轉 Word 文件需要 Word。

## 安裝

下載安裝腳本後即可檢視內容，再執行第二行。

```powershell
Invoke-WebRequest -UseBasicParsing 'https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main/install.ps1' -OutFile "$env:TEMP\right-click-office-to-pdf-install.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\right-click-office-to-pdf-install.ps1"
```

安裝程式會讀取儲存庫的 `VERSION`，下載同版號的 GitHub Release ZIP，將腳本安裝到 `%LOCALAPPDATA%\RightClickOfficeToPDF`。在 Windows 11 上，安裝程式會依 x64 或 ARM64 選擇對應的已簽章 MSIX，向目前使用者註冊新版右鍵選單。

若該版本的發行 ZIP 尚未上架，安裝程式會顯示警告，再從原始碼安裝傳統選單。若 Windows 無法信任 MSIX 簽章或組織政策阻止部署，安裝輸出會列出原因。此時 Windows 11 可從「顯示其他選項」使用轉檔。要使用新版選單，請採用已發布且受電腦信任的正式發行包。

## 移除

```powershell
Invoke-WebRequest -UseBasicParsing 'https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main/uninstall.ps1' -OutFile "$env:TEMP\right-click-office-to-pdf-uninstall.ps1"
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\right-click-office-to-pdf-uninstall.ps1"
```

也可從發行 ZIP 的解壓目錄執行 `.\Uninstall-RightClickOfficeToPDF.ps1`。如需離線安裝、建立發行包或在測試機驗收選單，請參閱[部署與發行](DEPLOYMENT.md)。
