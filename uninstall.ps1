# Online uninstaller for right-click-office-to-pdf
# Version: 1.0.0

$ErrorActionPreference = "Stop"

$repo = "https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main"
$tempDir = Join-Path $env:TEMP ("RightClickOfficeToPDF-Uninstall-" + [guid]::NewGuid().ToString("N"))

New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

Invoke-WebRequest -UseBasicParsing -Uri "$repo/Uninstall-RightClickOfficeToPDF.ps1" -OutFile (Join-Path $tempDir "Uninstall-RightClickOfficeToPDF.ps1")

powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $tempDir "Uninstall-RightClickOfficeToPDF.ps1")
