$ErrorActionPreference = 'Stop'
$uri = 'https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main/Uninstall-RightClickOfficeToPDF.ps1'
$scriptPath = Join-Path ([IO.Path]::GetTempPath()) ('RightClickOfficeToPDF-Uninstall-' + [guid]::NewGuid().ToString('N') + '.ps1')
try {
    Invoke-WebRequest -UseBasicParsing -Uri $uri -OutFile $scriptPath
    & $scriptPath
} finally {
    if (Test-Path -LiteralPath $scriptPath) { Remove-Item -LiteralPath $scriptPath -Force }
}
