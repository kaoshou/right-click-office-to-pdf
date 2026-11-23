param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Paths
)

. (Join-Path $PSScriptRoot 'OfficePdfText.ps1')

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

function Wait-BeforeClosing {
    if ([Console]::IsInputRedirected -or [Environment]::GetCommandLineArgs() -contains '-NonInteractive') { return }
    Write-Host $OfficePdfText.PressKey
    try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch {}
}

$converter = Join-Path $PSScriptRoot "Convert-OfficeToPdf.ps1"
Write-Host $OfficePdfText.Menu
Write-Host ""

if (-not (Test-Path -LiteralPath $converter -PathType Leaf)) {
    Write-Host ($OfficePdfText.MissingConverter -f $converter)
    Wait-BeforeClosing
    exit 1
}

$inputFiles = @(Resolve-InputPathsForDisplay $Paths)
if ($inputFiles.Count -eq 0) {
    Write-Host $OfficePdfText.SelectHint
    Wait-BeforeClosing
    exit 1
}

Write-Host ($OfficePdfText.Selected -f $inputFiles.Count)
foreach ($file in $inputFiles) {
    Write-Host ([System.IO.Path]::GetFileName($file))
}
Write-Host ""

try {
    & $converter -Paths $inputFiles
    $exitCode = $LASTEXITCODE
} catch {
    Write-Host ($OfficePdfText.CannotRun -f $_.Exception.Message)
    $exitCode = 1
}

if ($exitCode -ne 0) {
    Write-Host $OfficePdfText.BatchFailed
    Wait-BeforeClosing
} else {
    Write-Host $OfficePdfText.Finished
    if ($AutoCloseOnSuccess) {
        Start-Sleep -Seconds $AutoCloseDelaySeconds
    } else {
        Wait-BeforeClosing
    }
}
exit $exitCode
