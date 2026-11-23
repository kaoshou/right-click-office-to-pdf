# Fetch a single versioned release for both the classic and Windows 11 menus.
$ErrorActionPreference = 'Stop'
$github = 'https://github.com/kaoshou/right-click-office-to-pdf'
$raw = 'https://raw.githubusercontent.com/kaoshou/right-click-office-to-pdf/main'
$tempRoot = [IO.Path]::GetTempPath()
$tempDir = Join-Path $tempRoot ('RightClickOfficeToPDF-Install-' + [guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($tempDir) | Out-Null
try {
    $versionFile = Join-Path $tempDir 'VERSION'
    Invoke-WebRequest -UseBasicParsing -Uri "$raw/VERSION" -OutFile $versionFile
    $version = (Get-Content -LiteralPath $versionFile -Raw).Trim()
    if ($version -notmatch '^\d+\.\d+\.\d+$') { throw "Unexpected release version: $version" }
    $zip = Join-Path $tempDir 'package.zip'
    $bundle = Join-Path $tempDir 'bundle'
    try {
        Invoke-WebRequest -UseBasicParsing -Uri "$github/releases/download/v$version/right-click-office-to-pdf-v$version.zip" -OutFile $zip
        $sourceDir = $bundle
    } catch {
        if ([int]$_.Exception.Response.StatusCode -ne 404) { throw }
        Write-Warning "The release ZIP for v$version has not been published. Installing the classic context menu from source."
        Invoke-WebRequest -UseBasicParsing -Uri "$github/archive/refs/heads/main.zip" -OutFile $zip
        $sourceDir = Join-Path $bundle 'right-click-office-to-pdf-main'
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Compression.ZipFile]::ExtractToDirectory($zip, $bundle)
    & (Join-Path $sourceDir 'Install-RightClickOfficeToPDF.ps1')
} finally {
    $stage = Get-Item -LiteralPath $tempDir
    if ($stage.Parent.FullName -eq ([IO.DirectoryInfo]$tempRoot).FullName.TrimEnd('\') -and $stage.Name -like 'RightClickOfficeToPDF-Install-*') {
        Remove-Item -LiteralPath $stage.FullName -Recurse -Force
    }
}
