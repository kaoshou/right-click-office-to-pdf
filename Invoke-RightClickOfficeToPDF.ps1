param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Paths
)

# When all files are converted successfully, close the console automatically.
# If you want to always keep the window open, change this to $false.
$AutoCloseOnSuccess = $true
$AutoCloseDelaySeconds = 1

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {}

function Resolve-InputPathsForDisplay {
    param([string[]]$RawPaths)

    $result = New-Object System.Collections.Generic.List[string]

    if ($null -eq $RawPaths) {
        return @()
    }

    for ($i = 0; $i -lt $RawPaths.Count; $i++) {
        $candidate = [string]$RawPaths[$i]

        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        $candidate = $candidate.Trim().Trim('"')

        if (Test-Path -LiteralPath $candidate) {
            try {
                $result.Add((Resolve-Path -LiteralPath $candidate).Path)
            } catch {
                $result.Add($candidate)
            }
            continue
        }

        # Rebuild paths if Windows / shell split a path containing spaces.
        $joined = $candidate
        $found = $false

        for ($j = $i + 1; $j -lt $RawPaths.Count; $j++) {
            $next = ([string]$RawPaths[$j]).Trim().Trim('"')
            $joined = $joined + " " + $next

            if (Test-Path -LiteralPath $joined) {
                try {
                    $result.Add((Resolve-Path -LiteralPath $joined).Path)
                } catch {
                    $result.Add($joined)
                }

                $i = $j
                $found = $true
                break
            }
        }

        if (-not $found) {
            $result.Add($candidate)
        }
    }

    return @($result.ToArray() | Select-Object -Unique)
}

$converter = Join-Path $PSScriptRoot "Convert-OfficeToPdf.ps1"

Clear-Host
Write-Host "==============================================="
Write-Host " Right Click Office to PDF"
Write-Host " 轉換為 PDF 檔案"
Write-Host "==============================================="
Write-Host ""

if (-not (Test-Path -LiteralPath $converter)) {
    Write-Host "Converter script was not found: $converter"
    Write-Host ""
    Write-Host "Press any key to close."
    try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { Read-Host | Out-Null }
    exit 1
}

$inputFiles = @(Resolve-InputPathsForDisplay $Paths)

if ($inputFiles.Count -eq 0) {
    Write-Host "No input files were received."
    Write-Host "沒有收到要轉換的檔案。"
    Write-Host ""
    Write-Host "Press any key to close."
    try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { Read-Host | Out-Null }
    exit 1
}

Write-Host ("Files received: {0}" -f $inputFiles.Count)
Write-Host ""

if ($inputFiles.Count -eq 1) {
    Write-Host ("File: {0}" -f [System.IO.Path]::GetFileName($inputFiles[0]))
}
else {
    for ($i = 0; $i -lt $inputFiles.Count; $i++) {
        Write-Host ("[{0}/{1}] {2}" -f ($i + 1), $inputFiles.Count, [System.IO.Path]::GetFileName($inputFiles[$i]))
    }
}

Write-Host ""
Write-Host "Starting conversion..."
Write-Host ""

& $converter @inputFiles
$exitCode = $LASTEXITCODE

if ($null -eq $exitCode) {
    if ($?) { $exitCode = 0 } else { $exitCode = 1 }
}

Write-Host "----------------------------------------"

if ($exitCode -ne 0) {
    Write-Host "Some files failed."
    Write-Host ""
    Write-Host "Press any key to close."
    try {
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } catch {
        Read-Host | Out-Null
    }
}
else {
    Write-Host "All files completed."

    if ($AutoCloseOnSuccess) {
        Start-Sleep -Seconds $AutoCloseDelaySeconds
    }
    else {
        Write-Host ""
        Write-Host "Press any key to close."
        try {
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        } catch {
            Read-Host | Out-Null
        }
    }
}

exit $exitCode
