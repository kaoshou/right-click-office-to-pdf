# Security Policy

## 專案定位

**Right Click Office to PDF** 的設計目標是提供本機、低干擾、可解除安裝的 Windows 轉 PDF 捷徑。

本工具只在使用者的 Windows 電腦本機執行，不會將文件上傳到網路服務，也不會呼叫雲端 API 進行轉換。

## 安全設計

- 安裝於目前使用者的 `%LOCALAPPDATA%\RightClickOfficeToPDF`。
- 只寫入目前使用者層級的 `HKCU` 右鍵選單設定。
- 不需要系統管理員權限。
- 不修改 Windows 11 全域右鍵選單模式。
- 不包含將 Windows 11 右鍵選單強制改回傳統模式的腳本。
- 同名 PDF 已存在時會自動加序號，不直接覆蓋既有 PDF。
- Microsoft Office 轉換時會盡量以唯讀模式開啟文件，並嘗試停用巨集自動化執行。

## 使用者應注意的風險

Office 文件本身可能包含巨集、外部連結、嵌入物件或其他安全風險。請不要轉換來源不明、未經信任或可能含惡意內容的文件。

以下情況可能導致轉換失敗或需要人工處理：

- 檔案需要密碼。
- 檔案處於保護檢視。
- 檔案開啟時需要信任中心確認。
- 檔案含外部連結更新提示。
- 檔案使用罕見字型或特殊外掛內容。

## 線上安裝注意事項

README 提供的一行安裝指令會從 GitHub 下載 `install.ps1`，再執行安裝流程。這類指令很方便，但也代表使用者必須信任該 GitHub Repository 的內容。

建議對外發布時提醒使用者：

1. 確認網址是否為正式 Repository。
2. 可先下載 `install.ps1` 檢查內容。
3. 不要執行來路不明或被第三方修改過的安裝指令。

## 回報安全問題

若發現安全疑慮，請透過 GitHub Issues 或直接聯絡專案維護者。

Developer: Yu-Han Cheng（鄭郁翰）
GitHub: https://github.com/kaoshou/right-click-office-to-pdf
