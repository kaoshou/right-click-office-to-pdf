# Online Install

本文件說明 **Right Click Office to PDF** 的一行指令安裝與移除方式。

## Repository

Repository：

```text
https://github.com/kaoshou/right-click-office-to-pdf
```

## 單行安裝

```powershell
iwr -UseBasicParsing https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main/install.ps1 -OutFile "$env:TEMP\right-click-office-to-pdf-install.ps1"; powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\right-click-office-to-pdf-install.ps1"
```

此指令會下載 `install.ps1` 到暫存資料夾，再執行安裝流程。安裝程式會再從 GitHub 下載必要檔案並安裝到目前使用者的：

```text
%LOCALAPPDATA%\RightClickOfficeToPDF
```

## 單行移除

```powershell
iwr -UseBasicParsing https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main/uninstall.ps1 -OutFile "$env:TEMP\right-click-office-to-pdf-uninstall.ps1"; powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\right-click-office-to-pdf-uninstall.ps1"
```

## 為什麼不用 `irm ... | iex`

本專案刻意使用：

```powershell
iwr ... -OutFile ...; powershell -File ...
```

而不是：

```powershell
irm ... | iex
```

主要原因是下載到暫存檔後，使用者或管理者可以先檢查腳本內容，較適合對一般使用者公開。
