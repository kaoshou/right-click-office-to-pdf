param([ValidateSet('en-US','zh-TW')][string]$Language = 'zh-TW')
# Run in an interactive Windows session with Windows PowerShell 5.1.
# The test launches only the copied GUI and a fixture converter. Its profile
# and files are unique to this run, and only its own child windows are queried.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
# A multiline WinForms EDIT can be exposed only through MSAA on some Windows
# builds. Read its standard WM_GETTEXT message on this test child's handle.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class TestWindowText {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint msg,
        IntPtr size, StringBuilder text, uint flags, uint timeout, out IntPtr result);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint msg,
        IntPtr key, IntPtr keyData, uint flags, uint timeout, out IntPtr result);
    public static string Read(IntPtr hwnd) {
        var text = new StringBuilder(4096);
        IntPtr length;
        SendMessageTimeout(hwnd, 0x000D, (IntPtr)text.Capacity, text, 2, 500, out length);
        return text.ToString();
    }
    public static void ArrowKey(IntPtr list, int key) {
        IntPtr result;
        SendMessageTimeout(list, 0x0100, (IntPtr)key, (IntPtr)1, 2, 500, out result);
        SendMessageTimeout(list, 0x0101, (IntPtr)key, (IntPtr)1, 2, 500, out result);
    }
}
'@
$root = Join-Path ([IO.Path]::GetTempPath()) ('rcpdf-ui-' + [Guid]::NewGuid().ToString('N'))
$testApp = Join-Path $root 'test application'
$testFiles = Join-Path $root 'test files'
$children = [Collections.Generic.List[Diagnostics.Process]]::new()
[Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo($Language)
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'OfficePdfText.ps1')
$guiName = $OfficePdfText.Title
$project = Split-Path -Parent $PSScriptRoot

function Wait-For([scriptblock]$Condition, [int]$Seconds, [string]$Description) {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    do {
        if (& $Condition) { return }
        Start-Sleep -Milliseconds 100
    } while ($timer.Elapsed.TotalSeconds -lt $Seconds)
    throw "Timed out waiting for $Description."
}

function Start-TestWindow([string]$Profile, [string[]]$Names) {
    # powershell.exe -File cannot bind a string array in Windows PowerShell.
    # A UTF-8 BOM script preserves the exact array just as the shell launcher does.
    $paths = foreach ($name in $Names) { Join-Path $testFiles $name }
    $quoted = ($paths | ForEach-Object { "'" + $_.Replace("'","''") + "'" }) -join ', '
    $gui = (Join-Path $testApp 'Show-ConversionProgress.ps1').Replace("'","''")
    $entry = Join-Path $testApp ([Guid]::NewGuid().ToString('N') + '.ps1')
    [IO.File]::WriteAllText($entry, "[Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo('$Language')`r`n& '$gui' -Paths @($quoted)", [Text.UTF8Encoding]::new($true))
    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = Join-Path $PSHOME 'powershell.exe'
    $start.Arguments = '-NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "' + $entry + '"'
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.EnvironmentVariables['LOCALAPPDATA'] = $Profile
    $process = [Diagnostics.Process]::Start($start)
    $children.Add($process)
    return $process
}

function Get-TestWindows {
    $desktop = [Windows.Automation.AutomationElement]::RootElement
    foreach ($process in $children) {
        if ($process.HasExited) { continue }
        $condition = [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::ProcessIdProperty, $process.Id)
        $windows = $desktop.FindAll([Windows.Automation.TreeScope]::Children, $condition)
        foreach ($window in $windows) {
            if ($window.Current.Name -eq $guiName) { $window }
        }
    }
}

function Get-Events([string]$Profile) {
    $log = Join-Path $Profile 'events.txt'
    if ([IO.File]::Exists($log)) { [IO.File]::ReadAllLines($log) }
}

function Get-WindowText($Window) {
    $controls = $Window.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition)
    $text = foreach ($control in $controls) {
        $control.Current.Name
        if ($control.Current.Name -eq $OfficePdfText.Details -and $control.Current.NativeWindowHandle) {
            [TestWindowText]::Read([IntPtr]$control.Current.NativeWindowHandle)
        }
    }
    return $text -join ' '
}

try {
    [void][IO.Directory]::CreateDirectory($testApp)
    [void][IO.Directory]::CreateDirectory($testFiles)
    foreach ($file in 'OfficePdfText.ps1', 'Show-ConversionProgress.ps1', 'ProgressControls.cs') {
        Copy-Item -LiteralPath (Join-Path $project $file) -Destination $testApp
    }
    $fake = @'
param([string[]]$Paths, [scriptblock]$OnFileResult)
$ErrorActionPreference = 'Stop'
$log = Join-Path $env:LOCALAPPDATA 'events.txt'
$lockPath = Join-Path $env:LOCALAPPDATA 'working.lock'
foreach ($path in $Paths) {
    $leaf = [IO.Path]::GetFileName($path)
    try { $lock = [IO.File]::Open($lockPath, 'OpenOrCreate', 'ReadWrite', 'None') }
    catch {
        [IO.File]::WriteAllText((Join-Path $env:LOCALAPPDATA 'overlap'), $leaf)
        throw
    }
    try {
        [IO.File]::AppendAllText($log, "start $leaf" + [Environment]::NewLine)
        if ($leaf -eq 'first [1].docx') {
            $gate = Join-Path $env:LOCALAPPDATA 'release'
        } elseif ($leaf -eq 'another.docx') {
            $gate = Join-Path $env:LOCALAPPDATA 'finish'
        } else { $gate = $null }
        if ($gate) {
            $timeout = [Diagnostics.Stopwatch]::StartNew()
            while (![IO.File]::Exists($gate) -and $timeout.Elapsed.TotalSeconds -lt 30) {
                Start-Sleep -Milliseconds 100
            }
        }
        Start-Sleep -Milliseconds 150
        $result = [pscustomobject]@{ Path=$path; Status='Success'; OutputPath=$null; Message='' }
        if ($leaf -eq 'bad.docx') {
            $result.Status = 'Failed'
            $result.Message = 'Word fixture password required'
        } elseif ($leaf -eq 'another.docx') {
            $result.Status = 'Failed'
            $result.Message = 'Word fixture document is damaged'
        } else {
            $result.OutputPath = [IO.Path]::ChangeExtension($path, '.pdf')
            [IO.File]::WriteAllText($result.OutputPath, 'fixture')
        }
        & $OnFileResult $result
        [IO.File]::AppendAllText($log, "end $leaf" + [Environment]::NewLine)
    } finally { $lock.Dispose() }
}
if ($result.Status -eq 'Failed') { exit 1 }
exit 0
'@
    [IO.File]::WriteAllText((Join-Path $testApp 'Convert-OfficeToPdf.ps1'), $fake, [Text.UTF8Encoding]::new($true))
    $longName = ('跨處室公文與會議紀錄' * 5) + " [定稿] O'Brien.pptx"
    foreach ($name in 'first [1].docx','second.pptx','third.doc','bad.docx',$longName,'another.docx') {
        [IO.File]::WriteAllText((Join-Path $testFiles $name), 'fixture')
    }

    $successProfile = [IO.Directory]::CreateDirectory((Join-Path $root 'success')).FullName
    $owner = Start-TestWindow $successProfile @('first [1].docx')
    Wait-For { 'start first [1].docx' -in @(Get-Events $successProfile) } 15 'the first conversion to start'
    $second = Start-TestWindow $successProfile @('first [1].docx','second.pptx')
    $third = Start-TestWindow $successProfile @('third.doc')
    Wait-For { $second.HasExited -and $third.HasExited } 10 'both extra launches to hand off their selections'
    $windows = @(Get-TestWindows)
    if ($windows.Count -ne 1) { throw 'Expected one progress window for the three launches.' }
    Wait-For {
        $activeText = Get-WindowText $windows[0]
        $activeText.Contains(($OfficePdfText.Active -f '').Trim()) -and $activeText.Contains($OfficePdfText.Queued) -and $activeText.Contains('0 / 3')
    } 5 'active file, queued files and progress to be visible'
    [IO.File]::WriteAllText((Join-Path $successProfile 'release'), '')
    Wait-For { @(Get-Events $successProfile | Where-Object { $_ -like 'end *' }).Count -ge 3 } 12 'three file results'
    if (!$owner.WaitForExit(6000)) { throw 'The all-success window did not close within six seconds of the last result.' }
    if ([IO.File]::Exists((Join-Path $successProfile 'overlap'))) { throw 'More than one converter ran at once.' }
    $events = @(Get-Events $successProfile)
    foreach ($name in 'first [1].docx','second.pptx','third.doc') {
        if (@($events | Where-Object { $_ -eq "start $name" }).Count -ne 1 -or
            @($events | Where-Object { $_ -eq "end $name" }).Count -ne 1) {
            throw "Expected exactly one start and finish for $name. Events: $events"
        }
    }
    Write-Host 'PASS simultaneous launches share one serial queue, deduplicate inputs and close automatically.'

    $failureProfile = [IO.Directory]::CreateDirectory((Join-Path $root 'failure')).FullName
    $failedWindow = Start-TestWindow $failureProfile @('bad.docx',$longName,'another.docx')
    Wait-For { 'start another.docx' -in @(Get-Events $failureProfile) } 15 'the file after a failure to complete'
    Wait-For {
        $windows = @(Get-TestWindows)
        if ($failedWindow.HasExited) { throw "Failure window closed with exit code $($failedWindow.ExitCode)." }
        $windows.Count -eq 1 -and (Get-WindowText $windows[0]).Contains('Word fixture password required')
    } 5 'the completed error details to appear'
    $windows = @(Get-TestWindows)
    $window = $windows[0]
    $itemCondition = [Windows.Automation.AndCondition]::new(
        [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty, $longName),
        [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::IsSelectionItemPatternAvailableProperty, $true))
    $row = $window.FindFirst([Windows.Automation.TreeScope]::Descendants, $itemCondition)
    if (!$row) { throw 'The complete long filename is missing from the accessible list items.' }
    $selection = $row.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern)
    $selection.Select()
    $row.SetFocus()
    $longOutput = Join-Path $testFiles ([IO.Path]::ChangeExtension($longName, '.pdf'))
    Wait-For { (Get-WindowText $window).Contains($longOutput) } 4 'the full selected output path'
    [IO.File]::WriteAllText((Join-Path $failureProfile 'finish'), '')
    Wait-For { (Get-WindowText $window).Contains('3 / 3') } 8 'the last failure to be processed'
    if (!$selection.Current.IsSelected -or !(Get-WindowText $window).Contains($longOutput)) {
        throw 'A later conversion displaced the selected item or its output details.'
    }
    if (((Get-Events $failureProfile) -join ',') -ne "start bad.docx,end bad.docx,start $longName,end $longName,start another.docx,end another.docx") {
        throw 'Failure interrupted or reordered the batch.'
    }
    # Send standard list navigation to this window's list only, independent of
    # the foreground app. Both neighbouring rows failed for different reasons.
    $listCondition = [Windows.Automation.PropertyCondition]::new(
        [Windows.Automation.AutomationElement]::NameProperty, $OfficePdfText.FileList)
    $listElement = $window.FindFirst([Windows.Automation.TreeScope]::Descendants, $listCondition)
    $listHandle = [IntPtr]$listElement.Current.NativeWindowHandle
    if ($listHandle -eq [IntPtr]::Zero) { throw 'The accessible file list has no native window handle.' }
    [TestWindowText]::ArrowKey($listHandle, 0x26)
    Wait-For { (Get-WindowText $window).Contains('Word fixture password required') } 4 'up-arrow selection details'
    [TestWindowText]::ArrowKey($listHandle, 0x28)
    Wait-For { (Get-WindowText $window).Contains($longOutput) } 4 'down-arrow selection details'
    if ($failedWindow.WaitForExit(5000)) {
        throw "The window closed after displaying failure details (exit code $($failedWindow.ExitCode))."
    }
    $closeCondition = [Windows.Automation.AndCondition]::new(
        [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty, $OfficePdfText.Close),
        [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::Button))
    $close = $window.FindFirst([Windows.Automation.TreeScope]::Descendants, $closeCondition)
    $close.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
    if (!$failedWindow.WaitForExit(5000)) { throw 'The failed batch window did not close.' }
    Write-Host 'PASS failures preserve the selected long filename, keyboard navigation, details and review window.'
} finally {
    foreach ($process in $children) {
        if (!$process.HasExited) {
            [void]$process.CloseMainWindow()
            if (!$process.WaitForExit(2000)) { $process.Kill(); $process.WaitForExit() }
        }
        $process.Dispose()
    }
    $temp = [IO.DirectoryInfo]([IO.Path]::GetTempPath())
    $owned = [IO.DirectoryInfo]$root
    if ($owned.Exists -and $owned.Parent.FullName -eq $temp.FullName.TrimEnd('\') -and $owned.Name -like 'rcpdf-ui-*') {
        $owned.Delete($true)
    }
}
