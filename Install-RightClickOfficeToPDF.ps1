# Right Click Office to PDF installer
# Version: 1.0.0
# Direct right-click context menu edition.
# Registers only supported Office/OpenDocument file formats.

$ErrorActionPreference = "Stop"

$installDir = Join-Path $env:LOCALAPPDATA "RightClickOfficeToPDF"
New-Item -ItemType Directory -Path $installDir -Force | Out-Null

$sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$files = @(
    "Convert-OfficeToPdf.ps1",
    "Invoke-RightClickOfficeToPDF.ps1",
    "Uninstall-RightClickOfficeToPDF.ps1"
)

foreach ($file in $files) {
    $src = Join-Path $sourceDir $file

    if (-not (Test-Path -LiteralPath $src)) {
        throw "Required file not found: $file"
    }

    Copy-Item -LiteralPath $src -Destination (Join-Path $installDir $file) -Force
}

# Remove obsolete launchers from older builds.
Remove-Item -LiteralPath (Join-Path $installDir "RightClickOfficeToPDF.cmd") -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $installDir "Launch-RightClickOfficeToPDF.vbs") -Force -ErrorAction SilentlyContinue

$culture = [System.Globalization.CultureInfo]::CurrentUICulture.Name

if ($culture -like "zh-TW*" -or $culture -like "zh-Hant*" -or $culture -like "zh-HK*" -or $culture -like "zh-MO*") {
    $menuText = "轉換為 PDF 檔案"
} else {
    $menuText = "Convert to PDF"
}

$powershellExe = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
$launcher = Join-Path $installDir "Invoke-RightClickOfficeToPDF.ps1"

$extensions = @(".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".odt", ".ods", ".odp")

# Remove old SendTo entries from previous builds.
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

# Remove legacy generic all-file entry if any.
$legacyPsPath = "HKCU:\Software\Classes\*\shell\RightClickOfficeToPDF"
Remove-Item -LiteralPath $legacyPsPath -Recurse -Force -ErrorAction SilentlyContinue

foreach ($ext in $extensions) {
    $base = "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell\RightClickOfficeToPDF"
    $commandKey = Join-Path $base "command"

    if (Test-Path -LiteralPath $base) {
        Remove-Item -LiteralPath $base -Recurse -Force -ErrorAction SilentlyContinue
    }

    New-Item -Path $commandKey -Force | Out-Null

    New-ItemProperty -Path $base -Name "MUIVerb" -Value $menuText -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $base -Name "Icon" -Value "shell32.dll,70" -PropertyType String -Force | Out-Null

    # Player enables multi-select support for supported file types.
    New-ItemProperty -Path $base -Name "MultiSelectModel" -Value "Player" -PropertyType String -Force | Out-Null

    # "%1" is the primary selected file.
    # %* may contain the complete multi-selected file list on supported Explorer shell paths.
    # We pass both and de-duplicate in Invoke-RightClickOfficeToPDF.ps1.
    # %1 MUST be quoted to handle spaces and parentheses in file names.
    $cmd = '"' + $powershellExe + '" -NoProfile -STA -ExecutionPolicy Bypass -File "' + $launcher + '" "%1" %*'
    Set-ItemProperty -Path $commandKey -Name "(default)" -Value $cmd
}

Write-Host ""
Write-Host "Installed: $menuText"
Write-Host "Install directory: $installDir"
Write-Host ""
Write-Host "Usage:"
Write-Host "Right click one or multiple supported files -> $menuText"
Write-Host ""
Write-Host "Supported extensions:"
Write-Host ($extensions -join ", ")
