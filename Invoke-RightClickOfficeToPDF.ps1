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

function Wait-BeforeClosing {
    if ([Console]::IsInputRedirected -or [Environment]::GetCommandLineArgs() -contains '-NonInteractive') { return }
    Write-Host "按任意鍵關閉視窗。"
    try { $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch {}
}

$converter = Join-Path $PSScriptRoot "Convert-OfficeToPdf.ps1"
Write-Host "轉換為 PDF 檔案"
Write-Host ""

if (-not (Test-Path -LiteralPath $converter -PathType Leaf)) {
    Write-Host "找不到轉檔程式，請重新安裝：$converter"
    Wait-BeforeClosing
    exit 1
}

$inputFiles = @(Resolve-InputPathsForDisplay $Paths)
if ($inputFiles.Count -eq 0) {
    Write-Host "請在檔案總管選取要轉換的文件，再按右鍵選擇「轉換為 PDF 檔案」。"
    Wait-BeforeClosing
    exit 1
}

Write-Host ("已選取 {0} 個檔案。" -f $inputFiles.Count)
foreach ($file in $inputFiles) {
    Write-Host ([System.IO.Path]::GetFileName($file))
}
Write-Host ""

try {
    & $converter -Paths $inputFiles
    $exitCode = $LASTEXITCODE
} catch {
    Write-Host ("無法執行轉檔程式：{0}" -f $_.Exception.Message)
    $exitCode = 1
}

if ($exitCode -ne 0) {
    Write-Host "部分檔案未能轉換。請查看上方對應檔名的錯誤訊息。"
    Wait-BeforeClosing
} else {
    Write-Host "轉換完成。PDF 已儲存在原檔案旁。"
    if ($AutoCloseOnSuccess) {
        Start-Sleep -Seconds $AutoCloseDelaySeconds
    } else {
        Wait-BeforeClosing
    }
}
exit $exitCode
