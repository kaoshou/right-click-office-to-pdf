# Run with Windows PowerShell 5.1. Only the Office COM boundary is simulated.
# The real converter resolves paths, validates files, publishes PDFs and reports batch status.
$ErrorActionPreference = 'Stop'
$converter = Join-Path (Split-Path -Parent $PSScriptRoot) 'Convert-OfficeToPdf.ps1'
$testDir = Join-Path ([IO.Path]::GetTempPath()) ('office-pdf-test-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testDir)
$testPdfData = "%PDF-1.7`nfixture`n%%EOF`n"
$failures = @()

function New-Object {
    [CmdletBinding(DefaultParameterSetName = 'Net')]
    param(
        [Parameter(Position=0, ParameterSetName='Net')][string]$TypeName,
        [Parameter(Position=1, ParameterSetName='Net')][object[]]$ArgumentList,
        [Parameter(ParameterSetName='Com')][string]$ComObject
    )
    if ($ComObject) {
        if ($ComObject -eq 'Word.Application' -and $office -eq 'word') { return $testWord }
        if ($ComObject -eq 'PowerPoint.Application' -and $office -eq 'ppt') { return $ppt }
        throw 'The selected Office backend is unavailable.'
    }
    Microsoft.PowerShell.Utility\New-Object @PSBoundParameters
}

function Test-Case {
    param([string]$Description, [scriptblock]$Body)
    try {
        & $Body
        Write-Host "PASS $Description"
    } catch {
        $script:failures += "$Description - $($_.Exception.Message)"
        Write-Host "FAIL $Description - $($_.Exception.Message)"
    }
}

function New-InputFile([string]$Name) {
    $path = Join-Path $testDir $Name
    [IO.File]::WriteAllText($path, 'original data')
    return $path
}

try {
    $office = 'word'
    $testDocument = [pscustomobject]@{ Source = '' }
    $testDocument | Add-Member ScriptMethod ExportAsFixedFormat {
        param($path,$type)
        if ($type -ne 17) { throw 'Wrong Word PDF export format.' }
        switch ([IO.Path]::GetFileName($this.Source)) {
            'concurrent.docx' {
                [IO.File]::WriteAllText([IO.Path]::ChangeExtension($this.Source, '.pdf'), 'another conversion')
            }
            'truncated.docx' {
                [IO.File]::WriteAllText($path, '%PDF-1.7')
                return
            }
        }
        [IO.File]::WriteAllText($path, $testPdfData)
    }
    $testDocument | Add-Member ScriptMethod Close {}
    $docs = [pscustomobject]@{}
    $docs | Add-Member ScriptMethod Open {
        if (!$args[2]) { throw 'Read-only access required.' }
        $testDocument.Source = $args[0]
        switch ([IO.Path]::GetFileName($args[0])) {
            'locked.docx' { throw 'This document requires a password.' }
            'report [1].docx' {
                if (!$args[13]) { throw 'This document requires repair.' }
            }
        }
        return $testDocument
    }
    $testWord = [pscustomobject]@{ Documents=$docs; Visible=$false; DisplayAlerts=0; AutomationSecurity=2 }
    $testWord | Add-Member ScriptMethod Quit {}

    Test-Case 'concurrent output is preserved and our PDF gets a suffix' {
        $file = New-InputFile 'concurrent.docx'
        & $converter $file
        if ($LASTEXITCODE -ne 0 -or
            [IO.File]::ReadAllText((Join-Path $testDir 'concurrent.pdf')) -ne 'another conversion' -or
            [IO.File]::ReadAllText((Join-Path $testDir 'concurrent (1).pdf')) -ne $testPdfData) {
            throw 'Concurrent PDF overwritten or missing numbered result.'
        }
    }

    Test-Case 'a truncated export is reported as failed and never published' {
        & $converter (New-InputFile 'truncated.docx')
        if ($LASTEXITCODE -eq 0 -or [IO.File]::Exists((Join-Path $testDir 'truncated.pdf'))) {
            throw 'Unusable export reported as successful or left on disk.'
        }
    }

    Test-Case 'a failed file reports its reason and the rest of the batch completes' {
        $locked = New-InputFile 'locked.docx'
        $next = New-InputFile 'next.docx'
        $skip = New-InputFile 'image.jpg'
        $results = [Collections.Generic.List[object]]::new()
        $log = (& $converter -Paths @($locked, $skip, $next) -OnFileResult {
            param($result)
            $results.Add($result)
        } 6>&1) | Out-String -Width 4096
        if ($LASTEXITCODE -eq 0 -or !$log.Contains('This document requires a password.') -or
            !$log.Contains('locked.docx') -or !$log.Contains('Word') -or
            [IO.File]::Exists((Join-Path $testDir 'locked.pdf')) -or
            [IO.File]::ReadAllText((Join-Path $testDir 'next.pdf')) -ne $testPdfData) {
            throw "Batch status $LASTEXITCODE captured: $log"
        }
        if ($results.Count -ne 3 -or $results[0].Status -ne 'Failed' -or
            $results[0].Path -ne $locked -or !$results[0].Message.Contains('This document requires a password.') -or
            !$results[0].Message.Contains('Word') -or $results[0].OutputPath -or
            $results[1].Status -ne 'Skipped' -or $results[1].Path -ne $skip -or
            $results[2].Status -ne 'Success' -or $results[2].Path -ne $next -or
            $results[2].OutputPath -ne (Join-Path $testDir 'next.pdf')) {
            throw 'Expected one ordered GUI result per input with its outcome and details.'
        }
    }

    Test-Case 'Word can repair a read-only input and export without SaveAs' {
        $source = New-InputFile 'report [1].docx'
        & $converter $source
        if ($LASTEXITCODE -ne 0 -or [IO.File]::ReadAllText([IO.Path]::ChangeExtension($source,'.pdf')) -ne $testPdfData) {
            throw 'Word PDF missing.'
        }
        if ([IO.File]::ReadAllText($source) -ne 'original data') { throw 'Source file changed.' }
    }

    Test-Case 'PowerPoint can repair and export while another presentation is open' {
        $office = 'ppt'
        $testPresentation = [pscustomobject]@{}
        $testPresentation | Add-Member ScriptMethod SaveAs { throw 'SaveAs unsupported for this presentation.' }
        $testPresentation | Add-Member ScriptMethod ExportAsFixedFormat {
            if ($args[1] -ne 2 -or $args.Count -lt 8 -or $null -ne $args[7]) { throw 'Wrong PDF export range.' }
            [IO.File]::WriteAllText($args[0], $testPdfData)
        }
        $testPresentation | Add-Member ScriptMethod Close {}
        $collection = [pscustomobject]@{ Count=1 }
        $collection | Add-Member ScriptMethod Open { throw 'Repair required.' }
        $collection | Add-Member ScriptMethod Open2007 {
            if ($args[1] -ne -1 -or $args[4] -ne -1) { throw 'Read-only repair required.' }
            return $testPresentation
        }
        $ppt = [pscustomobject]@{ Presentations=$collection; AutomationSecurity=2; QuitCalled=$false }
        $ppt | Add-Member ScriptMethod Quit { $this.QuitCalled = $true }
        $source = New-InputFile 'slides [2].pptx'
        & $converter $source
        if ($LASTEXITCODE -ne 0 -or [IO.File]::ReadAllText([IO.Path]::ChangeExtension($source,'.pdf')) -ne $testPdfData) {
            throw 'PowerPoint PDF missing.'
        }
        if ($ppt.QuitCalled -or $ppt.AutomationSecurity -ne 2) { throw 'Existing PowerPoint session changed.' }
    }
} finally {
    [IO.Directory]::Delete($testDir, $true)
}
if ($failures.Count) { throw ($failures -join [Environment]::NewLine) }
