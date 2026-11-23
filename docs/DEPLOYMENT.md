# 部署與發行

使用者需要 Windows 和可正常開啟文件的 Microsoft Office 桌面版。Word、Excel、PowerPoint 各自處理對應格式。Windows 11 的新版右鍵選單支援 x64、ARM64，Windows 10 使用傳統選單。

## 建立正式發行 ZIP

建置電腦需要 CMake、Visual Studio C++ 建置工具（含 x64 與 ARM64 工具組）以及 Windows SDK 的 MakeAppx 和 SignTool。將含私鑰的程式碼簽署憑證放在目前使用者的個人憑證存放區。憑證必須受目標電腦信任。

在專案根目錄執行以下指令。將指紋換成實際值。套件的 Publisher 會沿用憑證主體名稱。

```powershell
$cert = Get-Item 'Cert:\CurrentUser\My\<憑證指紋>'
.\tools\Build-ReleaseZip.ps1 -CertificateThumbprint $cert.Thumbprint -Publisher $cert.Subject
```

若有 RFC 3161 時間戳記服務，可加上 `-TimestampUrl` 和服務網址。發行腳本依 `VERSION` 建立兩個架構的 MSIX，用 SignTool 簽署及驗證，並將安裝腳本、選單啟動器和 MSIX 放入同一個 ZIP。例如 `VERSION` 為 `1.1.0` 時，在 GitHub Release `v1.1.0` 上傳 `release/right-click-office-to-pdf-v1.1.0.zip`，線上安裝程式便能取得該版本。後續版本維持相同的 Publisher，並遞增版本號碼。

## 開發建置與驗證

在 x64 開發機可用下面的指令建立未簽章測試版，並直接從套件測試選單。這個流程只建置、解壓和測試，適合不註冊右鍵選單的工作站。

```powershell
.\tools\Build-ReleaseZip.ps1 -Development -Architecture x64
$version = (Get-Content .\VERSION -Raw).Trim()
$testDir = Join-Path $env:TEMP ([guid]::NewGuid().ToString())
Expand-Archive ".\release\right-click-office-to-pdf-v$version-dev.zip" -DestinationPath $testDir
powershell -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Test-ModernMenu.ps1 "$testDir\ModernMenu\RightClickOfficeToPDF-x64.msix"
```

測試直接載入套件中的 COM 命令，檢查顯示條件、多選和啟動器傳參。可用 `-Architecture arm64` 建置 ARM64 套件，並在同架構測試機執行原生測試。開發者若要另外測試檔案總管實際註冊，應在已啟用開發人員模式的測試機，於測試版 ZIP 解壓目錄執行 `.\Install-RightClickOfficeToPDF.ps1 -DevelopmentMenu`。

## 部署驗收

1. 在測試機解壓正式發行 ZIP，從解壓目錄執行 `powershell -NoProfile -ExecutionPolicy Bypass -File .\Install-RightClickOfficeToPDF.ps1`。受信任的正式簽章套件可使用一般使用者帳號安裝。
2. 在 Windows 11 選取一份 DOCX 或 PPTX，確認新版選單可見「轉換為 PDF 檔案」。若尚未更新，請重新登入後再檢查。Windows 10 可直接在傳統選單檢查。
3. 混選 Word、PowerPoint 與 Excel 文件啟動轉檔，確認進度視窗列出結果，並打開同資料夾內產生的 PDF 核對頁面。再次轉檔時，原先的 PDF 應保留。
4. 轉檔期間分別從新版與傳統選單加入另一份文件，確認兩個入口共用進度視窗。若組織常從郵件附件資料夾轉檔，可另外驗證 AppData 內的文件也能在原目錄取得 PDF。套件已關閉檔案系統寫入虛擬化，讓 Office 與兩種選單使用相同的實體路徑。[Microsoft 封裝虛擬化說明](https://learn.microsoft.com/windows/msix/desktop/flexible-virtualization)
5. 在測試機執行 `.\Uninstall-RightClickOfficeToPDF.ps1`，確認新版及傳統右鍵選單已移除。

如果只想測試轉檔與進度視窗，可直接從原始碼或發行 ZIP 執行 `.\Show-ConversionProgress.ps1 -Paths @('C:\測試\講義.docx', 'C:\測試\簡報.pptx')`，檔案路徑換成自己的測試文件。
