[CmdletBinding(DefaultParameterSetName = 'Signed')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Signed')]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory, ParameterSetName = 'Development')]
    [switch]$Development,

    [string]$Publisher = 'CN=kaoshou',
    [Parameter(ParameterSetName = 'Signed')]
    [string]$TimestampUrl,

    [Parameter(ParameterSetName = 'Development')]
    [ValidateSet('x64', 'arm64')]
    [string[]]$Architecture = @('x64', 'arm64')
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$version = (Get-Content -LiteralPath (Join-Path $projectRoot 'VERSION') -Raw).Trim()
if ($version -notmatch '^\d+\.\d+\.\d+$') { throw "Invalid VERSION: $version" }
$releaseDir = Join-Path $projectRoot 'release'
[IO.Directory]::CreateDirectory($releaseDir) | Out-Null
$suffix = if ($Development) { '-dev' } else { '' }
$zipName = "right-click-office-to-pdf-v$version$suffix.zip"
$zipPath = Join-Path $releaseDir $zipName

$tempRoot = [IO.Path]::GetTempPath()
$tempDir = Join-Path $tempRoot ("right-click-office-to-pdf-release-" + [guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($tempDir) | Out-Null
try {
    $payload = Join-Path $tempDir 'payload'
    $packages = Join-Path $payload 'ModernMenu'
    [IO.Directory]::CreateDirectory($packages) | Out-Null
    $buildArgs = @{ OutputDirectory = $packages; Publisher = $Publisher }
    if (-not $Development) {
        $buildArgs.CertificateThumbprint = $CertificateThumbprint
        if ($TimestampUrl) { $buildArgs.TimestampUrl = $TimestampUrl }
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    foreach ($arch in $Architecture) {
        & (Join-Path $PSScriptRoot 'Build-ModernMenu.ps1') -Architecture $arch @buildArgs
        $msix = Join-Path $packages "RightClickOfficeToPDF-$arch.msix"
        if (-not (Test-Path -LiteralPath $msix -PathType Leaf)) {
            throw "Missing $arch menu package."
        }
        # The classic menu can use the same GUI launcher on Windows 10.
        $launcherDir = Join-Path $payload "Launchers\$arch"
        [IO.Directory]::CreateDirectory($launcherDir) | Out-Null
        $package = [IO.Compression.ZipFile]::OpenRead($msix)
        try {
            [IO.Compression.ZipFileExtensions]::ExtractToFile($package.GetEntry('OfficeToPDF.exe'), (Join-Path $launcherDir 'OfficeToPDF.exe'))
        } finally { $package.Dispose() }
    }
    foreach ($name in @(
        'Convert-OfficeToPdf.ps1', 'Invoke-RightClickOfficeToPDF.ps1', 'Install-RightClickOfficeToPDF.ps1',
        'ModernMenu.ps1', 'OfficePdfText.ps1', 'Show-ConversionProgress.ps1', 'ProgressControls.cs', 'Uninstall-RightClickOfficeToPDF.ps1', 'install.ps1', 'uninstall.ps1',
        'README.md', 'SECURITY.md', 'LICENSE', 'CHANGELOG.md', 'VERSION', 'docs'
    )) {
        Copy-Item -LiteralPath (Join-Path $projectRoot $name) -Destination $payload -Recurse
    }
    Copy-Item -LiteralPath (Join-Path $projectRoot 'ModernMenu\Invoke-Selection.ps1') -Destination $payload
    $archive = Join-Path $tempDir $zipName
    [IO.Compression.ZipFile]::CreateFromDirectory($payload, $archive)
    # Publish only a complete archive, even if a build is interrupted or retried.
    $next = Join-Path $releaseDir ([guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [IO.File]::Copy($archive, $next)
        if ([IO.File]::Exists($zipPath)) {
            [IO.File]::Replace($next, $zipPath, [Management.Automation.Language.NullString]::Value)
        } else {
            [IO.File]::Move($next, $zipPath)
        }
    } finally {
        [IO.File]::Delete($next)
    }
} finally {
    $stage = Get-Item -LiteralPath $tempDir
    if ($stage.Parent.FullName -eq ([IO.DirectoryInfo]$tempRoot).FullName.TrimEnd('\') -and $stage.Name -like 'right-click-office-to-pdf-release-*') {
        Remove-Item -LiteralPath $stage.FullName -Recurse -Force
    }
}
Write-Host "Release ZIP created: $zipPath"
