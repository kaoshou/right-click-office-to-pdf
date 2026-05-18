# Online installer for right-click-office-to-pdf
# Version: 1.0.0

$ErrorActionPreference = "Stop"

$repo = "https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main"
$tempDir = Join-Path $env:TEMP ("RightClickOfficeToPDF-Install-" + [guid]::NewGuid().ToString("N"))

New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$files = @(
    "Convert-OfficeToPdf.ps1",
    "Invoke-RightClickOfficeToPDF.ps1",
    "Install-RightClickOfficeToPDF.ps1",
    "Uninstall-RightClickOfficeToPDF.ps1"
)

foreach ($file in $files) {
    Write-Host "Downloading: $file"
    Invoke-WebRequest -UseBasicParsing -Uri "$repo/$file" -OutFile (Join-Path $tempDir $file)
}

Write-Host ""
Write-Host "Starting installer..."
powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $tempDir "Install-RightClickOfficeToPDF.ps1")
