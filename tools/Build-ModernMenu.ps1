[CmdletBinding()]
param(
    [ValidateSet('x64','arm64')][string]$Architecture = 'x64',
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [string]$Publisher = 'CN=kaoshou',
    [string]$CertificateThumbprint,
    [uri]$TimestampUrl
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$source = Join-Path $root 'ModernMenu'
$version = [Version]((Get-Content -LiteralPath (Join-Path $root 'VERSION') -Raw).Trim())
$output = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($output) | Out-Null
$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
# MSBuild still creates internal paths subject to MAX_PATH.
$work = Join-Path $tempRoot ('pdf-' + [IO.Path]::GetRandomFileName())
New-Item -ItemType Directory -Path $work -ErrorAction Stop | Out-Null

function Run-Native([string]$Executable, [string[]]$Arguments) {
    & $Executable @Arguments
    if ($LASTEXITCODE -ne 0) { throw "$Executable exited with code $LASTEXITCODE." }
}

try {
    $finder = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    $component = if ($Architecture -eq 'arm64') { 'Microsoft.VisualStudio.Component.VC.Tools.ARM64' } else { 'Microsoft.VisualStudio.Component.VC.Tools.x86.x64' }
    $query = @('-latest','-products','*','-requires',$component,'-property')
    $vs = & $finder @query installationPath
    if (!$vs) { throw "Install Visual Studio C++ build tools for $Architecture." }
    $major = ([Version](& $finder @query installationVersion)).Major
    $generator = "Visual Studio $major " + (& $finder @query catalog_productLineVersion)
    $sdkRoot = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots').KitsRoot10
    $hostArch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'x64' }
    $sdk = Get-ChildItem (Join-Path $sdkRoot 'bin') -Directory | Where-Object {
        $_.Name -match '^10\.0\.\d+\.\d+$' -and (Test-Path (Join-Path $_.FullName "$hostArch\makeappx.exe"))
    } | Sort-Object { [Version]$_.Name } -Descending | Select-Object -First 1
    if (!$sdk) { throw 'Install the Windows 10/11 SDK packaging tools.' }
    $sdkBin = Join-Path $sdk.FullName $hostArch
    $build = Join-Path $work 'build'
    $stage = Join-Path $work 'package'
    $targetArch = if ($Architecture -eq 'arm64') { 'ARM64' } else { 'x64' }
    Run-Native cmake @('-S', $source, '-B', $build, '-G', $generator, '-A', $targetArch, "-DCMAKE_GENERATOR_INSTANCE=$vs")
    Run-Native cmake @('--build', $build, '--config', 'Release')
    Run-Native cmake @('--install', $build, '--config', 'Release', '--prefix', $stage)
    Copy-Item -LiteralPath (Join-Path $source 'Invoke-Selection.ps1') -Destination $stage
    Copy-Item -Path (Join-Path $source 'Logo*.png') -Destination $stage
    foreach ($script in 'OfficePdfText.ps1', 'Show-ConversionProgress.ps1','ProgressControls.cs','Invoke-RightClickOfficeToPDF.ps1','Convert-OfficeToPdf.ps1') {
        Copy-Item -LiteralPath (Join-Path $root $script) -Destination $stage
    }
    $manifest = [xml](Get-Content -LiteralPath (Join-Path $source 'AppxManifest.xml') -Encoding UTF8 -Raw)
    $identity = $manifest.Package.Identity
    $identity.SetAttribute('Publisher', $Publisher)
    $identity.SetAttribute('ProcessorArchitecture', $Architecture)
    $identity.SetAttribute('Version', ('{0}.{1}.{2}.{3}' -f $version.Major,$version.Minor,[Math]::Max(0,$version.Build),[Math]::Max(0,$version.Revision)))
    $manifest.Save((Join-Path $stage 'AppxManifest.xml'))
    $name = "RightClickOfficeToPDF-$Architecture.msix"
    $package = Join-Path $work $name
    Run-Native (Join-Path $sdkBin 'makeappx.exe') @('pack','/d',$stage,'/p',$package)
    if ($CertificateThumbprint) {
        $sign = @('sign','/fd','SHA256','/sha1',$CertificateThumbprint)
        if ($TimestampUrl) { $sign += @('/tr', $TimestampUrl.AbsoluteUri, '/td','SHA256') }
        Run-Native (Join-Path $sdkBin 'signtool.exe') ($sign + $package)
        Run-Native (Join-Path $sdkBin 'signtool.exe') @('verify','/pa',$package)
    } elseif ($TimestampUrl) { throw 'Specify a signing certificate when using a timestamp URL.' }
    Copy-Item -LiteralPath $package -Destination (Join-Path $output $name) -Force
    Write-Host "Built $(Join-Path $output $name)"
} finally {
    $resolvedWork = (Resolve-Path -LiteralPath $work).ProviderPath
    if ((Split-Path -Parent $resolvedWork).TrimEnd('\') -eq $tempRoot.TrimEnd('\') -and (Split-Path -Leaf $resolvedWork) -match '^pdf-[a-z0-9]{8}\.[a-z0-9]{3}$') {
        Remove-Item -LiteralPath $resolvedWork -Recurse -Force
    }
}
