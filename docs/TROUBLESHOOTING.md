# Troubleshooting / 疑難排解

## 1. 右鍵沒有看到「轉換為 PDF 檔案」

Windows 11 通常會把傳統右鍵功能放在：

```text
右鍵 → 顯示其他選項
```

多檔批次轉換請使用：

```text
右鍵 → 顯示其他選項 → 傳送到 → 轉換為 PDF 檔案
```

## 2. 出現 File not found

請確認檔案仍存在，且路徑沒有被移動。新版已處理常見的空白與括號檔名問題。

若仍發生，請嘗試將檔案放到較短路徑，例如：

```text
C:\Temp\test.pptx
```

再測試一次。

## 3. Microsoft Office 轉換失敗

可能原因：

- Office 尚未完成首次啟動設定。
- Office 授權未啟用。
- 文件處於保護檢視。
- 文件需要密碼。
- 文件開啟時需要手動確認外部連結或修復提示。
- Office 背景已有未關閉的錯誤視窗。

處理方式：

1. 手動開啟該文件，確認可以正常開啟。
2. 手動另存 PDF 測試。
3. 關閉所有 Word、Excel、PowerPoint 視窗。
4. 再重新使用右鍵轉換。

## 4. LibreOffice / OpenOffice / OxOffice 備援失敗

請確認已安裝相容軟體，且可找到 `soffice.exe`。

也可設定環境變數：

```text
SOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe
```

## 5. Excel 輸出 PDF 版面不正確

Excel 轉 PDF 取決於工作表的列印設定。

請在 Excel 中先設定：

- 列印範圍
- 紙張方向
- 紙張大小
- 縮放比例
- 分頁線

再執行轉換。

## 6. PowerPoint 轉 PDF 缺字或版面跑掉

常見原因：

- 缺少簡報使用的字型。
- 簡報包含特殊外掛、影音或動畫。
- 備援引擎與 Microsoft Office 的排版相容性不同。

建議優先使用安裝 Microsoft Office 的電腦轉換正式簡報。
