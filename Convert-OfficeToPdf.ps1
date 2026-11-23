param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Paths,
    [Parameter()][scriptblock]$OnFileResult
)

$ErrorActionPreference = "Stop"

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

        if (Test-Path -LiteralPath $candidate -ErrorAction SilentlyContinue) {
            try { $result.Add((Resolve-Path -LiteralPath $candidate).Path) } catch { $result.Add($candidate) }
            continue
        }

        $joined = $candidate
        $found = $false

        for ($j = $i + 1; $j -lt $RawPaths.Count; $j++) {
            $next = ([string]$RawPaths[$j]).Trim().Trim('"')
            $joined = $joined + " " + $next

            if (Test-Path -LiteralPath $joined -ErrorAction SilentlyContinue) {
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

function Assert-PdfOutput {
    param([string]$Path)

    $stream = [IO.File]::OpenRead($Path)
    try {
        $header = New-Object byte[] 5
        if ($stream.Read($header, 0, 5) -ne 5 -or [Text.Encoding]::ASCII.GetString($header) -ne '%PDF-') {
            throw 'The converter did not write a PDF header.'
        }
        $tail = New-Object byte[] ([int][Math]::Min(1024L, $stream.Length))
        $stream.Position = $stream.Length - $tail.Length
        [void]$stream.Read($tail, 0, $tail.Length)
        if ([Text.Encoding]::ASCII.GetString($tail) -notmatch '%%EOF[\x00\t\n\f\r ]*\z') {
            throw 'The PDF is incomplete (missing end-of-file marker).'
        }
    } finally {
        $stream.Dispose()
    }
}

function Publish-Pdf {
    param([string]$PdfPath, [string]$InputPath)

    # Stage on the destination volume. File.Move claims the final name without
    # overwriting a PDF produced by another Explorer invocation.
    $directory = [IO.Path]::GetDirectoryName($InputPath)
    $name = [IO.Path]::GetFileNameWithoutExtension($InputPath)
    $stage = Join-Path $directory ([IO.Path]::GetRandomFileName() + '.tmp')
    $writer = [IO.File]::Open($stage, 'CreateNew', 'Write', 'None')
    try {
        try {
            $reader = [IO.File]::OpenRead($PdfPath)
            try { $reader.CopyTo($writer) } finally { $reader.Dispose() }
        } finally {
            $writer.Dispose()
        }
        for ($index = 0; ; $index++) {
            $leaf = if ($index -eq 0) { "$name.pdf" } else { "$name ($index).pdf" }
            $target = Join-Path $directory $leaf
            try {
                [IO.File]::Move($stage, $target)
                return $target
            } catch [IO.IOException] {
                if (![IO.File]::Exists($target) -and ![IO.Directory]::Exists($target)) { throw }
            }
        }
    } finally {
        if ([IO.File]::Exists($stage)) { [IO.File]::Delete($stage) }
    }
}

function Convert-WordToPdf {
    param([string]$InputPath, [string]$PdfPath)

    $word = $documents = $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $word.AutomationSecurity = 3
        $documents = $word.Documents
        try {
            # Do not request a write lock or add the file to Recent Documents.
            $doc = $documents.Open($InputPath, $false, $true, $false, '')
        } catch {
            Write-Log "Word open failed; retrying read-only with OpenAndRepair: $($_.Exception.Message)"
            $m = [Type]::Missing
            $doc = $documents.Open($InputPath, $false, $true, $false, '', $m, $false,
                                   $m, $m, $m, $m, $false, $m, $true, $m, $true)
        }
        try {
            $doc.ExportAsFixedFormat($PdfPath, 17) # wdExportFormatPDF
        } catch {
            Write-Log "Word PDF export failed; trying SaveAs: $($_.Exception.Message)"
            [IO.File]::Delete($PdfPath)
            try { $doc.SaveAs2($PdfPath, 17) }
            catch { $doc.SaveAs([ref]([object]$PdfPath), [ref]([object]17)) }
        }
    } finally {
        if ($null -ne $doc) {
            try { $doc.Close($false) } catch {}
            Release-ComObject $doc
        }
        Release-ComObject $documents
        if ($null -ne $word) {
            try { $word.Quit(0) } catch {}
            Release-ComObject $word
        }
    }
}

function Convert-ExcelToPdf {
    param([string]$InputPath, [string]$PdfPath)

    $excel = $workbooks = $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 3
        $excel.AskToUpdateLinks = $false
        $excel.EnableEvents = $false
        $workbooks = $excel.Workbooks
        $workbook = $workbooks.Open($InputPath, 0, $true)
        $workbook.ExportAsFixedFormat(0, $PdfPath) # xlTypePDF
    } finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch {}
            Release-ComObject $workbook
        }
        Release-ComObject $workbooks
        if ($null -ne $excel) {
            try { $excel.Quit() } catch {}
            Release-ComObject $excel
        }
    }
}

function Convert-PowerPointToPdf {
    param([string]$InputPath, [string]$PdfPath)

    $powerpoint = $presentations = $presentation = $null
    # PowerPoint can share an application between Explorer invocations. Serialize
    # the temporary application-wide macro setting and restore it after opening.
    $openLock = [Threading.Mutex]::new($false, 'Local\RightClickOfficeToPDF-PowerPointOpen')
    $locked = $false
    try {
        try { $locked = $openLock.WaitOne(30000) }
        catch [Threading.AbandonedMutexException] { $locked = $true }
        if (!$locked) { throw 'Another PowerPoint conversion is still opening a presentation.' }
        try {
            $powerpoint = New-Object -ComObject PowerPoint.Application
            $presentations = $powerpoint.Presentations
            $previousSecurity = $powerpoint.AutomationSecurity
            $powerpoint.AutomationSecurity = 3
            try {
                # Untitled opens a copy, including when this file is open for editing.
                try { $presentation = $presentations.Open($InputPath, -1, -1, 0) }
                catch {
                    Write-Log "PowerPoint open failed; retrying read-only with repair: $($_.Exception.Message)"
                    $presentation = $presentations.Open2007($InputPath, -1, -1, 0, -1)
                }
            } finally {
                $powerpoint.AutomationSecurity = $previousSecurity
            }
        } finally {
            $openLock.ReleaseMutex()
        }
        try {
            $presentation.SaveAs($PdfPath, 32) # ppSaveAsPDF
        } catch {
            Write-Log "PowerPoint SaveAs failed; trying fixed-format export: $($_.Exception.Message)"
            [IO.File]::Delete($PdfPath)
            # PrintRange is required, even when exporting all slides.
            $presentation.ExportAsFixedFormat($PdfPath, 2, 1, 0, 1, 1, 0, $null)
        }
    } finally {
        $openLock.Dispose()
        if ($null -ne $presentation) {
            try { $presentation.Close() } catch {}
            Release-ComObject $presentation
        }
        Release-ComObject $presentations
        # Releasing our references leaves any interactive PowerPoint session intact.
        Release-ComObject $powerpoint
    }
}

function Convert-OneFile {
    param([string]$Path)

    $script:TotalCount++
    $workDir = $null
    $engine = $null
    $result = [pscustomobject]@{ Path = $Path; Status = 'Failed'; OutputPath = $null; Message = '' }
    try {
        if (!(Test-Path -LiteralPath $Path -PathType Leaf)) { throw 'File not found.' }
        $fullPath = (Resolve-Path -LiteralPath $Path).ProviderPath
        $ext = [IO.Path]::GetExtension($fullPath).ToLowerInvariant()
        $engine = switch ($ext) {
            { $_ -in '.doc', '.docx', '.odt' } { 'Word' }
            { $_ -in '.xls', '.xlsx', '.ods' } { 'Excel' }
            { $_ -in '.ppt', '.pptx', '.odp' } { 'PowerPoint' }
        }
        if (!$engine) {
            $result.Status = 'Skipped'
            $result.Message = "不支援的檔案格式 $ext"
            Write-Log "Skipped: $Path ($($result.Message))"
            $script:SkipCount++
            return
        }
        Write-Log "----------------------------------------"
        Write-Log "Input: $fullPath"
        $workDir = Join-Path ([IO.Path]::GetTempPath()) ('rcpdf-' + [Guid]::NewGuid().ToString('N'))
        [void][IO.Directory]::CreateDirectory($workDir)
        $pdf = Join-Path $workDir 'output.pdf'
        Write-Log "Using Microsoft $engine..."
        & "Convert-${engine}ToPdf" $fullPath $pdf
        Assert-PdfOutput $pdf
        $output = Publish-Pdf $pdf $fullPath
        $result.Status = 'Success'
        $result.OutputPath = $output
        Write-Log "OK: $output"
        $script:SuccessCount++
    } catch {
        $source = if ($engine) { "Microsoft $engine" } else { 'Input' }
        $reason = $_.Exception.GetBaseException().Message
        $result.Message = "[$source] $reason"
        Write-Log "Failed: $Path $($result.Message)"
        $script:FailCount++
    } finally {
        if ($workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue }
        if ($OnFileResult) {
            try { $null = & $OnFileResult $result }
            catch { Write-Log "Could not send file status: $($_.Exception.Message)" }
        }
    }
}

try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

if ($null -eq $Paths -or $Paths.Count -eq 0) {
    Write-Host "No input files were received."
    Write-Host "Select an Office document in File Explorer and use the Convert to PDF menu command."
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
