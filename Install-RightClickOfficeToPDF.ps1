# Install for the current user. -DevelopmentMenu is for testing an unsigned build.
param([switch]$DevelopmentMenu)
$ErrorActionPreference = "Stop"

# Use the native Windows PowerShell Appx module even when launched from pwsh or WOW64.
if ($PSEdition -eq 'Core' -or ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess)) {
    $system = if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) { 'Sysnative' } else { 'System32' }
    $windowsPS = Join-Path $env:WINDIR "$system\WindowsPowerShell\v1.0\powershell.exe"
    $arguments = @()
    if ($DevelopmentMenu) { $arguments += '-DevelopmentMenu' }
    & $windowsPS -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath @arguments
    if ($LASTEXITCODE -ne 0) { throw "Installer failed (exit $LASTEXITCODE)." }
    return
}
$installDir = Join-Path $env:LOCALAPPDATA "RightClickOfficeToPDF"
New-Item -ItemType Directory -Path $installDir -Force | Out-Null

$sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$files = @(
    "Convert-OfficeToPdf.ps1",
    "Invoke-RightClickOfficeToPDF.ps1",
    "Show-ConversionProgress.ps1",
    "ProgressControls.cs",
    "Uninstall-RightClickOfficeToPDF.ps1"
)

foreach ($file in ($files + 'ModernMenu.ps1')) {
    if (-not (Test-Path -LiteralPath (Join-Path $sourceDir $file) -PathType Leaf)) {
        throw "Required file not found: $file"
    }
}

foreach ($file in $files) {
    $src = Join-Path $sourceDir $file

    Copy-Item -LiteralPath $src -Destination (Join-Path $installDir $file) -Force
}

# Remove obsolete launchers from older builds.
Remove-Item -LiteralPath (Join-Path $installDir "RightClickOfficeToPDF.cmd") -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $installDir "Launch-RightClickOfficeToPDF.vbs") -Force -ErrorAction SilentlyContinue

$menuText = "轉換為 PDF 檔案"

$powershellExe = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
$launcher = Join-Path $installDir 'Show-ConversionProgress.ps1'
$cmd = '"' + $powershellExe + '" -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "' + $launcher + '" "%1" %*'
. (Join-Path $sourceDir 'ModernMenu.ps1')
$arch = Get-ExplorerArchitecture
if ($arch) {
    $nativeLauncher = Join-Path $sourceDir "Launchers\$arch\OfficeToPDF.exe"
    if (Test-Path -LiteralPath $nativeLauncher -PathType Leaf) {
        $targetExe = Join-Path $installDir 'OfficeToPDF.exe'
        Copy-Item -LiteralPath (Join-Path $sourceDir 'Invoke-Selection.ps1') -Destination $installDir -Force
        Copy-Item -LiteralPath $nativeLauncher -Destination $targetExe -Force
        $cmd = '"' + $targetExe + '" "%1" %*'
    }
}

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

    Set-ItemProperty -Path $commandKey -Name "(default)" -Value $cmd
}

Write-Host ""
try {
    Install-ModernMenu -SourceDir $sourceDir -InstallDir $installDir -DevelopmentMenu:$DevelopmentMenu
} catch {
    Write-Warning ("Windows 11 新版選單安裝失敗。可在「顯示其他選項」使用轉檔。原因：{0}" -f $_.Exception.Message)
}
Write-Host "Installed: $menuText"
Write-Host "Install directory: $installDir"
Write-Host ""
Write-Host "Usage:"
Write-Host "Right click one or multiple supported files -> $menuText"
Write-Host ""
Write-Host "Supported extensions:"
Write-Host ($extensions -join ", ")
