param()

. (Join-Path $PSScriptRoot 'OfficePdfText.ps1')

$ModernMenuName = 'Kaoshou.RightClickOfficeToPDF.Menu'

function Get-ExplorerArchitecture {
    $native = Get-ItemPropertyValue -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment' -Name PROCESSOR_ARCHITECTURE -ErrorAction SilentlyContinue
    if (-not $native) {
        $native = if ($env:PROCESSOR_ARCHITEW6432) { $env:PROCESSOR_ARCHITEW6432 } else { $env:PROCESSOR_ARCHITECTURE }
    }
    switch ($native) {
        'AMD64' { 'x64' }
        'ARM64' { 'arm64' }
    }
}

function Install-ModernMenu {
    param([string]$SourceDir, [string]$InstallDir, [switch]$DevelopmentMenu, [int]$WindowsBuild = [Environment]::OSVersion.Version.Build)

    if ($WindowsBuild -lt 22000) { return }
    $architecture = Get-ExplorerArchitecture
    $package = Join-Path $SourceDir "ModernMenu\RightClickOfficeToPDF-$architecture.msix"
    if (-not $architecture -or -not (Test-Path -LiteralPath $package -PathType Leaf)) {
        Write-Warning $OfficePdfText.NoPackage
        return
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($package)
    try {
        $reader = [IO.StreamReader]::new($archive.GetEntry('AppxManifest.xml').Open())
        try { $manifest = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        $hasSignature = $null -ne $archive.GetEntry('AppxSignature.p7x')
    } finally {
        $archive.Dispose()
    }
    $identity = $manifest.Package.Identity
    if ($identity.Name -ne $ModernMenuName -or $identity.ProcessorArchitecture -ne $architecture) {
        throw $OfficePdfText.BadPackage
    }
    if ((-not $hasSignature) -and (-not $DevelopmentMenu)) {
        throw $OfficePdfText.Unsigned
    }

    $publisherFile = Join-Path $InstallDir 'modern-menu.publisher'
    if ((Test-Path -LiteralPath $publisherFile) -and (Get-Content -LiteralPath $publisherFile -Raw).Trim() -ne $identity.Publisher) {
        throw $OfficePdfText.Publisher
    }
    $current = Get-AppxPackage -Name $ModernMenuName -ErrorAction Stop |
        Where-Object { $_.Publisher -eq $identity.Publisher -and -not $_.IsResourcePackage } |
        Select-Object -First 1
    # Record the package identity before deployment so interrupted installs can be removed.
    Set-Content -LiteralPath $publisherFile -Value $identity.Publisher -Encoding UTF8
    if ($current -and [version]$current.Version -gt [version]$identity.Version) {
        Write-Host ($OfficePdfText.Newer -f $current.Version)
        return
    }
    if ($hasSignature) {
        if (-not ($current -and (-not $current.IsDevelopmentMode) -and [version]$current.Version -eq [version]$identity.Version -and $current.Status -eq 'Ok')) {
            Add-AppxPackage -Path $package -ErrorAction Stop
        }
    } else {
        # Explorer may still be using the preceding DLL. Register a new layout on a rebuild.
        $layoutRoot = Join-Path $InstallDir 'ModernMenu.Dev'
        $hash = (Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash
        $layout = Join-Path $layoutRoot $hash
        if (-not [IO.Directory]::Exists($layout)) {
            [IO.Directory]::CreateDirectory($layoutRoot) | Out-Null
            $staging = Join-Path $layoutRoot ([guid]::NewGuid().ToString('N'))
            try {
                [IO.Compression.ZipFile]::ExtractToDirectory($package, $staging)
                foreach ($metadata in 'AppxBlockMap.xml', '[Content_Types].xml') {
                    [IO.File]::Delete((Join-Path $staging $metadata))
                }
                [IO.Directory]::Move($staging, $layout)
            } finally {
                if ([IO.Directory]::Exists($staging)) {
                    $stage = Get-Item -LiteralPath $staging -Force
                    if ($stage.Parent.FullName -eq [IO.Path]::GetFullPath($layoutRoot)) {
                        Remove-Item -LiteralPath $stage.FullName -Recurse -Force
                    }
                }
            }
        }
        if (-not ($current -and $current.InstallLocation -eq $layout -and $current.Status -eq 'Ok')) {
            try {
                Add-AppxPackage -Register (Join-Path $layout 'AppxManifest.xml') -ErrorAction Stop
            } catch {
                throw ($OfficePdfText.RegistrationFailed -f $_.Exception.Message)
            }
        }
    }
    Write-Host $OfficePdfText.Registered
}
