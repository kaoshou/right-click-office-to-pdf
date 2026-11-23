# Remove this tool from the current user account.
$ErrorActionPreference = 'Stop'
if ($PSEdition -eq 'Core' -or ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess)) {
    $system = if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) { 'Sysnative' } else { 'System32' }
    $windowsPS = Join-Path $env:WINDIR "$system\WindowsPowerShell\v1.0\powershell.exe"
    & $windowsPS -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
    if ($LASTEXITCODE -ne 0) { throw "Uninstaller failed (exit $LASTEXITCODE)." }
    return
}

if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { throw 'LOCALAPPDATA is unavailable.' }
$installDir = Join-Path ([IO.Path]::GetFullPath($env:LOCALAPPDATA)) 'RightClickOfficeToPDF'
$publisherFile = Join-Path $installDir 'modern-menu.publisher'
if (Test-Path -LiteralPath $publisherFile -PathType Leaf) {
    $publisher = (Get-Content -LiteralPath $publisherFile -Raw).Trim()
    # Unregister before removing loose development files. A failed removal can be retried.
    Get-AppxPackage -Name 'Kaoshou.RightClickOfficeToPDF.Menu' -ErrorAction Stop |
        Where-Object { $_.Publisher -eq $publisher -and -not $_.IsResourcePackage } |
        ForEach-Object { Remove-AppxPackage -Package $_.PackageFullName -ErrorAction Stop }
}
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

if (Test-Path -LiteralPath $installDir) {
    $installed = Get-Item -LiteralPath $installDir -Force
    if ($installed.Attributes -band [IO.FileAttributes]::ReparsePoint) {
        throw "Refusing to recursively remove a reparse point: $installDir"
    }
    Remove-Item -LiteralPath $installed.FullName -Recurse -Force -ErrorAction Stop
}
Write-Host "Right Click Office to PDF has been uninstalled."
