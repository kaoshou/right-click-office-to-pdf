# Right Click Office to PDF

**Right Click Office to PDF** 是一個 Windows 右鍵快速轉 PDF 工具。使用者可以在檔案總管中選取 Office 文件，透過右鍵選單直接轉成 PDF。此工具純粹使用 PowerShell 呼叫 Office 工具來將檔案轉換為 PDF，所以使用者電腦必須先安裝相關工具。

本工具採用 **Microsoft Office 優先、LibreOffice / OpenOffice / OxOffice 備援** 的混合式轉檔策略：

```text
DOC / DOCX      → 優先 Microsoft Word        → 失敗再嘗試 LibreOffice / OpenOffice / OxOffice
XLS / XLSX      → 優先 Microsoft Excel       → 失敗再嘗試 LibreOffice / OpenOffice / OxOffice
PPT / PPTX      → 優先 Microsoft PowerPoint  → 失敗再嘗試 LibreOffice / OpenOffice / OxOffice
ODT / ODS / ODP → 優先 Microsoft Office      → 失敗再嘗試 LibreOffice / OpenOffice / OxOffice
```

---

## 主要功能

- 支援 DOC、DOCX、PPT、PPTX、XLS、XLSX、ODT、ODS、ODP。
- 支援單一檔案右鍵轉 PDF，也可多選檔案後透過右鍵批次轉 PDF。
- 優先使用本機 Microsoft Office，提高 Word、Excel、PowerPoint 文件版面相容性。
- Microsoft Office 不可用或單檔轉換失敗時，自動嘗試 LibreOffice、OpenOffice、OxOffice 的 `soffice.exe`。

---

## 單行指令快速安裝 / 移除

使用者可免手動下載 ZIP，只要開啟 **Windows PowerShell**，執行下列指令即可。

### 執行單行指令快速安裝

```powershell
iwr -UseBasicParsing https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main/install.ps1 -OutFile "$env:TEMP\right-click-office-to-pdf-install.ps1"; powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\right-click-office-to-pdf-install.ps1"
```

### 執行單行指令快速移除

```powershell
iwr -UseBasicParsing https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main/uninstall.ps1 -OutFile "$env:TEMP\right-click-office-to-pdf-uninstall.ps1"; powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\right-click-office-to-pdf-uninstall.ps1"
```

### 線上安裝安全說明

上述指令會先將 GitHub 上的安裝腳本下載到 Windows 暫存資料夾，再執行該腳本。這比 `irm ... | iex` 更適合對一般使用者公開，因為使用者可以先檢查下載的 `.ps1` 內容。完整說明請見 [`docs/ONLINE_INSTALL.md`](docs/ONLINE_INSTALL.md)。

---

## 手動下載與安裝

若不使用單行指令安裝，也可以下載此專案後解壓縮，例如：

```text
C:\Users\你的帳號\Downloads\right-click-office-to-pdf
```

在該資料夾開啟 PowerShell，執行：

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-RightClickOfficeToPDF.ps1
```

安裝程式會將核心檔案複製到：

```text
%LOCALAPPDATA%\RightClickOfficeToPDF
```

並建立：

```text
右鍵傳統選單項目
支援格式的右鍵選單登錄設定
```

---

## 使用方式

### 單一或多個檔案轉 PDF

1. 在 Windows 檔案總管選取一個或多個支援的檔案，例如 `.docx`、`.pptx`、`.xlsx`。
2. 按右鍵。
3. 點選「轉換為 PDF 檔案」。
4. 工具會在文字介面中列出收到的檔案，並逐一轉換。

```text
右鍵 → 轉換為 PDF 檔案
```

在 Windows 11 中，傳統 shell 右鍵選單項目可能會出現在「顯示其他選項」內。  
本工具只會針對支援的 Office / OpenDocument 格式註冊右鍵選單；不支援的檔案類型不會出現此選項。

---

## 輸出規則

PDF 預設輸出到原檔案相同資料夾：

```text
講義.docx   → 講義.pdf
簡報.pptx  → 簡報.pdf
成績表.xlsx → 成績表.pdf
```

若同名 PDF 已存在，工具會自動增加序號，避免覆蓋舊檔：

```text
簡報.pdf
簡報 (1).pdf
簡報 (2).pdf
```

---

## 解除安裝

### 使用一行指令移除

```powershell
iwr -UseBasicParsing https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main/uninstall.ps1 -OutFile "$env:TEMP\right-click-office-to-pdf-uninstall.ps1"; powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\right-click-office-to-pdf-uninstall.ps1"
```

### 使用本機腳本移除

在專案資料夾中執行：

```powershell
powershell -ExecutionPolicy Bypass -File .\Uninstall-RightClickOfficeToPDF.ps1
```

解除安裝會移除：

```text
右鍵選單登錄設定
支援格式的右鍵選單登錄設定
%LOCALAPPDATA%\RightClickOfficeToPDF
```

---

## 支援的轉檔引擎

### Microsoft Office

若本機已安裝 Microsoft Office，工具會優先使用：

```text
Word.Application       → DOC / DOCX / ODT
Excel.Application      → XLS / XLSX / ODS
PowerPoint.Application → PPT / PPTX / ODP
```

這通常能取得最接近原始文件的 PDF 排版。

### LibreOffice / OpenOffice / OxOffice

若 Microsoft Office 不可用或該檔案轉換失敗，工具會尋找：

```text
soffice.exe
```

常見搜尋位置包含：

```text
C:\Program Files\LibreOffice\program\soffice.exe
C:\Program Files\OpenOffice 4\program\soffice.exe
C:\Program Files\Apache OpenOffice 4\program\soffice.exe
C:\Program Files\OxOffice\program\soffice.exe
C:\Program Files\OSSII\OxOffice\program\soffice.exe
```

也可透過環境變數指定：

```text
SOFFICE_PATH
```

---

## 安全說明摘要

- 本工具只在本機執行，不會上傳文件，不會呼叫雲端 API。
- 本工具不內建 Microsoft Office、LibreOffice、OpenOffice 或 OxOffice，也不繞過任何軟體授權。
- 使用 Microsoft Office COM 自動化轉檔時，工具會盡量以唯讀模式開啟檔案，並嘗試停用巨集自動化執行。
- 不建議轉換來源不明、可能含惡意巨集或惡意內容的 Office 文件。
- 若檔案處於保護檢視、密碼保護、開啟需互動確認、或含外部連結提示，背景轉檔可能失敗。
- 一行安裝指令會下載並執行 GitHub 上的 PowerShell 腳本；正式對外發布時，請提醒使用者確認 Repository 來源與腳本內容。

完整安全說明請見 [`SECURITY.md`](SECURITY.md) 與 [`docs/SAFETY.md`](docs/SAFETY.md)。

---

## Windows 11 右鍵選單說明

本工具採用穩定的傳統 Windows Shell 整合方式。於 Windows 11 中，功能通常會顯示在：

```text
右鍵 → 顯示其他選項 → 轉換為 PDF 檔案
```

或：

```text
右鍵 → 顯示其他選項 → 傳送到 → 轉換為 PDF 檔案
```

本專案不修改 Windows 11 全域右鍵選單樣式，也不將 Windows 11 強制改回 Windows 10 傳統右鍵選單。

---

## 已知限制

- Excel 轉 PDF 會受工作表列印範圍、紙張方向、縮放比例影響。
- Office 文件若需要密碼、手動確認、信任中心互動、外部連結確認，可能無法背景轉換。
- PowerPoint 若包含特殊字型、影音、外掛內容，PDF 可能與原始播放效果不同。
- LibreOffice / OpenOffice / OxOffice 備援轉檔的版面可能與 Microsoft Office 有差異。

---

## 專案資訊

- Project: `right-click-office-to-pdf`
- Version: `1.0.0`
- Platform: Windows 10 / Windows 11
- Runtime: Windows PowerShell 5.1 或 PowerShell 7+
- Developer: Yu-Han Cheng（鄭郁翰）
- GitHub: `https://github.com/kaoshou/right-click-office-to-pdf`

---

## 免責聲明

本專案最初僅為個人使用需求所開發，現以「現狀（AS IS）」方式提供，不附任何明示或暗示之保證，包含但不限於商業適售性、特定用途適用性及不侵權保證。
使用者應自行承擔使用本軟體之所有風險。對於因使用本軟體所造成之任何直接、間接、附帶、衍生性或特殊損害（包含但不限於資料遺失、文件損毀、轉換錯誤、商業中斷或工作損失等），作者與貢獻者概不負責。
本工具可能涉及 Office 文件、PDF 與其他敏感資料之處理，使用者應自行確認其使用方式符合所在地之法律、組織資訊安全政策及個人資料保護相關規範。
本專案與 Microsoft、Adobe、LibreOffice、The Document Foundation、OSSII 或其他第三方公司或組織無任何附屬、授權、贊助或背書關係。
文件轉換結果仍可能因 Office 版本、字型、版面設定或第三方軟體環境差異而有所不同。對於重要文件，請務必自行再次確認轉換結果之正確性。
繼續使用本軟體即視為您已閱讀、理解並同意上述內容。
