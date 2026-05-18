# Right Click Office to PDF uninstaller
# Version: 1.0.0

$extensions = @(".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".odt", ".ods", ".odp")

foreach ($ext in $extensions) {
    $base = "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell\RightClickOfficeToPDF"
    Remove-Item -LiteralPath $base -Recurse -Force -ErrorAction SilentlyContinue
}

$legacyPsPath = "HKCU:\Software\Classes\*\shell\RightClickOfficeToPDF"
Remove-Item -LiteralPath $legacyPsPath -Recurse -Force -ErrorAction SilentlyContinue

$sendTo = [Environment]::GetFolderPath("SendTo")

@(
    "Convert to PDF.lnk",
    "Convert to PDF.cmd",
    "Convert to PDF (Office First).lnk",
    "Convert to PDF (Office First).cmd",
    "轉換成 PDF.lnk",
    "轉換成 PDF.cmd",
    "轉換為 PDF 檔案.lnk",
    "轉換為 PDF 檔案.cmd"
) | ForEach-Object {
    Remove-Item -LiteralPath (Join-Path $sendTo $_) -Force -ErrorAction SilentlyContinue
}

$installDir = Join-Path $env:LOCALAPPDATA "RightClickOfficeToPDF"
Remove-Item -LiteralPath $installDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Right Click Office to PDF has been uninstalled."
