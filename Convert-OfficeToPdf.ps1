param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Paths
)

$ErrorActionPreference = "Continue"

$script:TotalCount = 0
$script:SuccessCount = 0
$script:FailCount = 0
$script:SkipCount = 0

function Write-Log {
    param([string]$Message)
    Write-Host $Message
}

function Release-ComObject {
    param($ComObject)
    if ($null -ne $ComObject) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null } catch {}
    }
}

function Resolve-InputPaths {
    param([string[]]$RawPaths)

    # Windows SendTo / batch invocation may occasionally split an unquoted file path such as:
    # C:\Users\name\Downloads\Demo File (1).pptx
    # into multiple tokens. This function rebuilds tokens until an existing path is found.
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
            try { $result.Add((Resolve-Path -LiteralPath $candidate).Path) } catch { $result.Add($candidate) }
            continue
        }

        $joined = $candidate
        $found = $false

        for ($j = $i + 1; $j -lt $RawPaths.Count; $j++) {
            $next = ([string]$RawPaths[$j]).Trim().Trim('"')
            $joined = $joined + " " + $next

            if (Test-Path -LiteralPath $joined) {
                try { $result.Add((Resolve-Path -LiteralPath $joined).Path) } catch { $result.Add($joined) }
                $i = $j
                $found = $true
                break
            }
        }

        if (-not $found) {
            # Keep the original token so the converter can print a meaningful error.
            $result.Add($candidate)
        }
    }

    return @($result.ToArray())
}

function Get-TargetPdfPath {
    param([string]$InputPath)

    $directory = Split-Path -Parent $InputPath
    $filenameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $candidate = Join-Path $directory ($filenameWithoutExt + ".pdf")

    if (-not (Test-Path -LiteralPath $candidate)) {
        return $candidate
    }

    for ($i = 1; $i -le 9999; $i++) {
        $candidate = Join-Path $directory ("{0} ({1}).pdf" -f $filenameWithoutExt, $i)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    return (Join-Path $directory ("{0}-{1}.pdf" -f $filenameWithoutExt, $timestamp))
}

function Test-ComAvailable {
    param([string]$ProgId)
    try {
        $type = [type]::GetTypeFromProgID($ProgId)
        return ($null -ne $type)
    } catch {
        return $false
    }
}

function Find-Soffice {
    $candidates = New-Object System.Collections.Generic.List[string]

    if ($env:SOFFICE_PATH) {
        $candidates.Add($env:SOFFICE_PATH)
    }

    try {
        $cmd = Get-Command "soffice.exe" -ErrorAction SilentlyContinue
        if ($null -ne $cmd -and $cmd.Source) { $candidates.Add($cmd.Source) }
    } catch {}

    $programFiles = @()
    if ($env:ProgramFiles) { $programFiles += $env:ProgramFiles }
    if (${env:ProgramFiles(x86)}) { $programFiles += ${env:ProgramFiles(x86)} }

    foreach ($pf in $programFiles) {
        $candidates.Add((Join-Path $pf "LibreOffice\program\soffice.exe"))
        $candidates.Add((Join-Path $pf "OpenOffice 4\program\soffice.exe"))
        $candidates.Add((Join-Path $pf "Apache OpenOffice 4\program\soffice.exe"))
        $candidates.Add((Join-Path $pf "OxOffice\program\soffice.exe"))
        $candidates.Add((Join-Path $pf "OSSII\OxOffice\program\soffice.exe"))
    }

    foreach ($candidate in $candidates | Select-Object -Unique) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }
    return $null
}

function Convert-WithSoffice {
    param([string]$InputPath)

    $soffice = Find-Soffice
    if (-not $soffice) {
        Write-Log "Fallback failed: soffice.exe was not found."
        return $false
    }

    $tempOutDir = Join-Path ([System.IO.Path]::GetTempPath()) ("RightClickOfficeToPDF-out-" + [Guid]::NewGuid().ToString("N"))
    $tempProfileDir = Join-Path ([System.IO.Path]::GetTempPath()) ("RightClickOfficeToPDF-lo-profile-" + [Guid]::NewGuid().ToString("N"))

    try {
        New-Item -ItemType Directory -Path $tempOutDir -Force | Out-Null
        New-Item -ItemType Directory -Path $tempProfileDir -Force | Out-Null

        $pdfPath = Get-TargetPdfPath $InputPath
        $profileUri = (New-Object System.Uri($tempProfileDir)).AbsoluteUri

        Write-Log "Fallback engine: $soffice"
        $args = @(
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "-env:UserInstallation=$profileUri",
            "--convert-to", "pdf",
            "--outdir", $tempOutDir,
            $InputPath
        )

        $output = & $soffice @args 2>&1
        $exitCode = $LASTEXITCODE

        $expectedTempPdf = Join-Path $tempOutDir ([System.IO.Path]::GetFileNameWithoutExtension($InputPath) + ".pdf")
        $actualTempPdf = $expectedTempPdf

        if (-not (Test-Path -LiteralPath $actualTempPdf)) {
            $foundPdf = Get-ChildItem -LiteralPath $tempOutDir -Filter "*.pdf" -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($null -ne $foundPdf) {
                $actualTempPdf = $foundPdf.FullName
            }
        }

        if (Test-Path -LiteralPath $actualTempPdf) {
            Move-Item -LiteralPath $actualTempPdf -Destination $pdfPath -Force
            Write-Log "OK using LibreOffice/OpenOffice/OxOffice fallback: $pdfPath"
            return $true
        }

        Write-Log "Fallback conversion failed: $InputPath"
        if ($output) { Write-Log ($output | Out-String) }
        Write-Log "Exit code: $exitCode"
        return $false
    } catch {
        Write-Log "Fallback exception: $($_.Exception.Message)"
        return $false
    } finally {
        try { Remove-Item -LiteralPath $tempOutDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        try { Remove-Item -LiteralPath $tempProfileDir -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }
}

function Convert-WordToPdf {
    param([string]$InputPath)

    if (-not (Test-ComAvailable "Word.Application")) {
        Write-Log "Microsoft Word is not available."
        return $false
    }

    $word = $null
    $doc = $null

    try {
        [string]$pdfPath = [string](Get-TargetPdfPath $InputPath)

        $word = New-Object -ComObject Word.Application

        # Keep Word hidden during conversion.
        $word.Visible = $false
        $word.DisplayAlerts = 0

        try { $word.AutomationSecurity = 3 } catch {}

        Write-Log "Using Microsoft Word..."
        Write-Log "Output: $pdfPath"

        # Use a simple Open call for better compatibility across Office versions.
        $doc = $word.Documents.Open($InputPath)

        # Word constants:
        # 17 = wdFormatPDF
        # Use SaveAs2 first. On some Win10 + Office COM environments,
        # SaveAs([ref]$pdfPath, [ref]17) may throw a psobject/Object conversion error.
        try {
            $doc.SaveAs2($pdfPath, 17)
        }
        catch {
            Write-Log "Word SaveAs2 failed, retrying SaveAs with typed reference parameters: $($_.Exception.Message)"

            $fileNameObj = [ref]([object]$pdfPath)
            $formatObj = [ref]([object]17)
            $doc.SaveAs($fileNameObj, $formatObj)
        }

        if (Test-Path -LiteralPath $pdfPath) {
            Write-Log "OK using Microsoft Word: $pdfPath"
            return $true
        }

        Write-Log "Microsoft Word did not create PDF."
        return $false
    }
    catch {
        Write-Log "Microsoft Word conversion failed: $($_.Exception.Message)"
        return $false
    }
    finally {
        if ($null -ne $doc) {
            try { $doc.Close($false) } catch {}
            Release-ComObject $doc
        }

        if ($null -ne $word) {
            try { $word.Quit() } catch {}
            Release-ComObject $word
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}


function Convert-ExcelToPdf {
    param([string]$InputPath)

    if (-not (Test-ComAvailable "Excel.Application")) {
        Write-Log "Microsoft Excel is not available."
        return $false
    }

    $excel = $null
    $workbook = $null

    try {
        $pdfPath = Get-TargetPdfPath $InputPath

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        try { $excel.AutomationSecurity = 3 } catch {}
        try { $excel.AskToUpdateLinks = $false } catch {}
        try { $excel.EnableEvents = $false } catch {}

        Write-Log "Using Microsoft Excel..."
        Write-Log "Output: $pdfPath"

        # 最精簡開啟方式，比帶一堆參數更穩
        $workbook = $excel.Workbooks.Open($InputPath)

        # 0 = xlTypePDF
        $workbook.ExportAsFixedFormat(0, $pdfPath)

        if (Test-Path -LiteralPath $pdfPath) {
            Write-Log "OK using Microsoft Excel: $pdfPath"
            return $true
        }

        Write-Log "Microsoft Excel did not create PDF."
        return $false
    }
    catch {
        Write-Log "Microsoft Excel conversion failed: $($_.Exception.Message)"
        return $false
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch {}
            Release-ComObject $workbook
        }

        if ($null -ne $excel) {
            try { $excel.Quit() } catch {}
            Release-ComObject $excel
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}


function Convert-PowerPointToPdf {
    param([string]$InputPath)

    if (-not (Test-ComAvailable "PowerPoint.Application")) {
        Write-Log "Microsoft PowerPoint is not available."
        return $false
    }

    $powerpoint = $null
    $presentation = $null

    try {
        [string]$pdfPath = [string](Get-TargetPdfPath $InputPath)

        $powerpoint = New-Object -ComObject PowerPoint.Application
        try { $powerpoint.AutomationSecurity = 3 } catch {}

        Write-Log "Using Microsoft PowerPoint..."
        Write-Log "Output: $pdfPath"

        # Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
        $presentation = $powerpoint.Presentations.Open($InputPath, -1, 0, 0)

        # 32 = ppSaveAsPDF
        # Do not use ExportAsFixedFormat here. It can throw int/Object COM conversion errors.
        $presentation.SaveAs($pdfPath, 32)

        if (Test-Path -LiteralPath $pdfPath) {
            Write-Log "OK using Microsoft PowerPoint: $pdfPath"
            return $true
        }

        Write-Log "Microsoft PowerPoint did not create PDF."
        return $false
    }
    catch {
        Write-Log "Microsoft PowerPoint conversion failed: $($_.Exception.Message)"
        return $false
    }
    finally {
        if ($null -ne $presentation) {
            try { $presentation.Close() } catch {}
            Release-ComObject $presentation
        }

        if ($null -ne $powerpoint) {
            try { $powerpoint.Quit() } catch {}
            Release-ComObject $powerpoint
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}


function Convert-OneFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log "File not found: $Path"
        $script:FailCount++
        return
    }

    $fullPath = (Resolve-Path -LiteralPath $Path).Path
    $ext = [System.IO.Path]::GetExtension($fullPath).ToLowerInvariant()
    $usedOffice = $false
    $converted = $false
    $supported = $true

    $script:TotalCount++

    Write-Log "----------------------------------------"
    Write-Log "Input: $fullPath"

    switch ($ext) {
        ".doc"  { $usedOffice = Convert-WordToPdf $fullPath }
        ".docx" { $usedOffice = Convert-WordToPdf $fullPath }
        ".odt"  { $usedOffice = Convert-WordToPdf $fullPath }
        ".xls"  { $usedOffice = Convert-ExcelToPdf $fullPath }
        ".xlsx" { $usedOffice = Convert-ExcelToPdf $fullPath }
        ".ods"  { $usedOffice = Convert-ExcelToPdf $fullPath }
        ".ppt"  { $usedOffice = Convert-PowerPointToPdf $fullPath }
        ".pptx" { $usedOffice = Convert-PowerPointToPdf $fullPath }
        ".odp"  { $usedOffice = Convert-PowerPointToPdf $fullPath }
        default {
            Write-Log "Unsupported file type: $ext"
            $supported = $false
        }
    }

    if (-not $supported) {
        $script:SkipCount++
        return
    }

    if ($usedOffice) {
        $converted = $true
    } else {
        Write-Log "Trying LibreOffice/OpenOffice/OxOffice fallback..."
        $converted = Convert-WithSoffice $fullPath
    }

    if ($converted) {
        $script:SuccessCount++
    } else {
        $script:FailCount++
    }
}

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

if ($null -eq $Paths -or $Paths.Count -eq 0) {
    Write-Host "No input files were received."
    Write-Host "Use the Windows context menu, the SendTo shortcut, or drag files onto RightClickOfficeToPDF.cmd."
    Read-Host "Press Enter to close"
    exit 1
}

$normalizedPaths = Resolve-InputPaths $Paths

foreach ($p in $normalizedPaths) {
    Convert-OneFile $p
}

Write-Log "----------------------------------------"
Write-Log ("Summary: total={0}, success={1}, failed={2}, skipped={3}" -f $script:TotalCount, $script:SuccessCount, $script:FailCount, $script:SkipCount)
Write-Log "Done."

if ($script:FailCount -gt 0) {
    exit 1
}

exit 0
