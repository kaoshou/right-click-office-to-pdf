# Run with Windows PowerShell 5.1. Deployment calls are doubles, never the system Appx cmdlets.
$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $PSScriptRoot
. (Join-Path $project 'ModernMenu.ps1')
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Assert($condition, [string]$message) { if (-not $condition) { throw $message } }
function Get-AppxPackage { param($Name) $script:current }
function Add-AppxPackage {
    [CmdletBinding()]
    param([Parameter(Position = 0)]$Path, [switch]$Register)
    Assert (Test-Path -LiteralPath $Path -PathType Leaf) "Missing deployment input $Path"
    if ($Register -and -not $script:devMode) { throw 'Development registration is disabled.' }
    $script:calls += 1
    $script:current = [pscustomobject]@{
        Version = '1.1.0.0'; Status = 'Ok'; Publisher = 'CN=test'; IsDevelopmentMode = $Register.IsPresent
        InstallLocation = (Split-Path -Parent $Path); PackageFullName = 'test-package'; IsResourcePackage = $false
    }
}
function Get-ItemPropertyValue {
    param($LiteralPath, $Name)
    switch ($Name) {
        'PROCESSOR_ARCHITECTURE' { $script:architecture }
        default { throw "Unexpected registry lookup $LiteralPath/$Name" }
    }
}
function New-TestPackage([string]$Architecture, [switch]$Signed) {
    $layout = Join-Path $root ([guid]::NewGuid().ToString('N'))
    [IO.Directory]::CreateDirectory($layout) | Out-Null
    Set-Content -LiteralPath (Join-Path $layout 'AppxManifest.xml') -Value "<Package><Identity Name='Kaoshou.RightClickOfficeToPDF.Menu' Publisher='CN=test' ProcessorArchitecture='$Architecture' Version='1.1.0.0'/></Package>"
    if ($Signed) { [IO.File]::WriteAllBytes((Join-Path $layout 'AppxSignature.p7x'), [byte[]]@(0)) }
    $package = Join-Path $packages "RightClickOfficeToPDF-$Architecture.msix"
    [IO.File]::Delete($package)
    [IO.Compression.ZipFile]::CreateFromDirectory($layout, $package)
}
function Assert-Rejected([scriptblock]$Operation, [string]$Expected) {
    $message = ''
    try { & $Operation } catch { $message = $_.Exception.Message }
    Assert ($message -like "*$Expected*") "Expected rejection matching '$Expected', got '$message'"
}

$tempRoot = [IO.Path]::GetTempPath()
$root = Join-Path $tempRoot ('right-click-menu-test-' + [guid]::NewGuid().ToString('N'))
$packages = Join-Path $root 'ModernMenu'
$install = Join-Path $root 'installed'
[IO.Directory]::CreateDirectory($packages) | Out-Null
[IO.Directory]::CreateDirectory($install) | Out-Null
$run = @{ SourceDir = $root; InstallDir = $install; WindowsBuild = 26100 }
$calls = 0
$current = $null
$architecture = 'AMD64'
$devMode = 0
try {
    Install-ModernMenu -SourceDir $root -InstallDir $install -WindowsBuild 19045
    Install-ModernMenu @run
    Assert ($calls -eq 0) 'An absent package or Windows 10 tried to deploy a menu.'
    New-TestPackage x64 -Signed
    Install-ModernMenu @run
    Install-ModernMenu @run
    Assert ($calls -eq 1) 'The signed package must be installed exactly once.'
    Assert ((Get-Content -LiteralPath (Join-Path $install 'modern-menu.publisher') -Raw).Trim() -eq 'CN=test') 'Missing publisher for uninstall.'
    $current = $null
    $architecture = 'ARM64'
    New-TestPackage arm64
    Assert-Rejected { Install-ModernMenu @run } '未簽章'
    Assert-Rejected { Install-ModernMenu @run -DevelopmentMenu } '開發人員模式'
    Assert ($calls -eq 1) 'Unsigned deployment occurred without explicit development opt-in.'
    $devMode = 1
    Install-ModernMenu @run -DevelopmentMenu
    Install-ModernMenu @run -DevelopmentMenu
    Assert ($calls -eq 2 -and $current.IsDevelopmentMode) 'Development registration must be idempotent.'
    $current.Version = '2.0.0.0'
    Install-ModernMenu @run -DevelopmentMenu
    Assert ($calls -eq 2) 'An older package replaced a new version.'
    # The uninstaller's real registry/SendTo operations are blocked by these local doubles.
    if ($PSEdition -eq 'Desktop' -and [Environment]::Is64BitProcess) {
        & {
            $uninstallTest = @{ Package = $current; Fail = $true; Removed = $false }
            $previousAppData = $env:LOCALAPPDATA
            $env:LOCALAPPDATA = Join-Path $root 'profile'
            $owned = Join-Path $env:LOCALAPPDATA 'RightClickOfficeToPDF'
            $sendTo = [Environment]::GetFolderPath('SendTo')
            try {
                [IO.Directory]::CreateDirectory($owned) | Out-Null
                Set-Content -LiteralPath (Join-Path $owned 'modern-menu.publisher') -Value 'CN=test'
                function Remove-Item {
                    param($LiteralPath, [switch]$Recurse, [switch]$Force)
                    if ($LiteralPath -like 'HKCU:*' -or (Split-Path -Parent $LiteralPath) -eq $sendTo) { return }
                    Assert ($LiteralPath -eq $owned) "Attempted removal outside test directory: $LiteralPath"
                    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $owned -Recurse -Force
                }
                function Get-AppxPackage {
                    param($Name)
                    Assert ($Name -eq 'Kaoshou.RightClickOfficeToPDF.Menu') 'Unexpected package query.'
                    $uninstallTest.Package
                    [pscustomobject]@{ Publisher = 'CN=other'; PackageFullName = 'foreign-package'; IsResourcePackage = $false }
                }
                function Remove-AppxPackage {
                    param($Package)
                    Assert ($Package -eq 'test-package') 'Uninstaller selected the wrong publisher.'
                    if ($uninstallTest.Fail) { throw 'package in use' }
                    $uninstallTest.Removed = $true
                }
                $uninstallTest.Fail = $true
                Assert-Rejected { & (Join-Path $project 'Uninstall-RightClickOfficeToPDF.ps1') } 'package in use'
                Assert ([IO.Directory]::Exists($owned)) 'An unsuccessful unregister lost the install state.'
                $uninstallTest.Fail = $false
                & (Join-Path $project 'Uninstall-RightClickOfficeToPDF.ps1')
                Assert ($uninstallTest.Removed -and (-not [IO.Directory]::Exists($owned))) 'Uninstall did not finish.'
            } finally {
                $env:LOCALAPPDATA = $previousAppData
            }
        }
    }
    Write-Host 'PASS: Windows 10, missing package, signed reinstall, ARM64, developer opt-in, downgrade protection and removal.'
} finally {
    $stage = Get-Item -LiteralPath $root
    if ($stage.Parent.FullName -eq ([IO.DirectoryInfo]$tempRoot).FullName.TrimEnd('\') -and $stage.Name -like 'right-click-menu-test-*') {
        Remove-Item -LiteralPath $stage.FullName -Recurse -Force
    }
}
