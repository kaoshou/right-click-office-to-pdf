# Usage

## 安裝

### 單行快速安裝

```powershell
iwr -UseBasicParsing https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main/install.ps1 -OutFile "$env:TEMP\right-click-office-to-pdf-install.ps1"; powershell -NoProfile -ExecutionPolicy Bypass -File "$env:TEMP\right-click-office-to-pdf-install.ps1"
```

### 手動安裝

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-RightClickOfficeToPDF.ps1
```

## 單檔轉換

```text
選取檔案 → 右鍵 → 顯示其他選項 → 轉換為 PDF 檔案
```

適合單一 `.docx`、`.pptx`、`.xlsx` 等檔案。

## 多檔轉換

```text
選取多個檔案 → 右鍵 → 顯示其他選項 → 傳送到 → 轉換為 PDF 檔案
```

多檔批次轉換最推薦使用「傳送到」方式。

## 拖曳到 CMD 啟動器

也可以將檔案拖曳到：

```text
RightClickOfficeToPDF.cmd
```

## 輸出位置

PDF 會輸出到原始檔案所在資料夾。同名 PDF 存在時，會自動加序號避免覆蓋。
