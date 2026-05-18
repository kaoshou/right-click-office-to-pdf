# 部署建議

## 適合情境

本工具適合：

- 個人電腦快速轉 PDF。
- 校內行政人員桌面環境。
- 教師批次轉換教材。
- 不需要伺服器化的大量文件轉檔情境。

## 不建議情境

不建議將本工具直接用於：

- 長時間背景服務。
- Web Server 伺服器端轉檔服務。
- 多使用者同時呼叫同一套 Office COM 的自動化服務。
- 處理大量來源不明文件的非隔離環境。

## 校內部署流程建議

1. 由 IT 人員確認 Office 版本與授權狀態。
2. 於測試電腦安裝本工具。
3. 測試常見文件格式：DOCX、PPTX、XLSX。
4. 測試常用 Office 與 OpenDocument 格式轉檔。
5. 確認防毒軟體與 PowerShell 執行政策不會阻擋。
6. 再提供給使用者安裝。

## 解除安裝

使用：

```powershell
powershell -ExecutionPolicy Bypass -File .\Uninstall-RightClickOfficeToPDF.ps1
```
