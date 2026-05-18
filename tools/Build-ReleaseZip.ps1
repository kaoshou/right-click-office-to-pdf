$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$version = Get-Content -Path (Join-Path $projectRoot "VERSION") -Raw
$version = $version.Trim()
$releaseDir = Join-Path $projectRoot "release"
$zipPath = Join-Path $releaseDir ("right-click-office-to-pdf-v{0}.zip" -f $version)

if (Test-Path -LiteralPath $releaseDir) {
    Remove-Item -LiteralPath $releaseDir -Recurse -Force
}
New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null

$exclude = @(".git", "release")
$tempDir = Join-Path $env:TEMP ("right-click-office-to-pdf-release-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

Get-ChildItem -Path $projectRoot -Force | Where-Object { $exclude -notcontains $_.Name } | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $tempDir -Recurse -Force
}

Compress-Archive -Path (Join-Path $tempDir "*") -DestinationPath $zipPath -Force
Remove-Item -LiteralPath $tempDir -Recurse -Force

Write-Host "Release ZIP created: $zipPath"
