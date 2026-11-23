param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Paths
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# Explorer may start a separate process for each file. Publish each request and
# elect the window owner under the same lock, including when the window closes.
$session = [Diagnostics.Process]::GetCurrentProcess().SessionId
$inbox = Join-Path $env:LOCALAPPDATA "RightClickOfficeToPDF\Queue\$session"
$sha = [Security.Cryptography.SHA256]::Create()
try {
    $key = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($inbox.ToUpperInvariant()))).Replace('-', '')
} finally { $sha.Dispose() }
$gate = [Threading.Mutex]::new($false, "Local\RightClickOfficeToPDF-Gate-$key")
$owner = [Threading.Mutex]::new($false, "Local\RightClickOfficeToPDF-Window-$key")
$state = @{ OwnsWindow = $false; Current = $null; Completed = 0; Failures = 0; CloseAfter = $null }

function Enter-Mutex($Mutex, [int]$Timeout = -1) {
    try { return $Mutex.WaitOne($Timeout) }
    catch [Threading.AbandonedMutexException] { return $true }
}

function Receive-Requests {
    $null = Enter-Mutex $gate
    try {
        foreach ($request in [IO.Directory]::GetFiles($inbox, '*.json')) {
            try {
                $batch = [IO.File]::ReadAllText($request) | ConvertFrom-Json
                if (!$batch.Paths -or @($batch.Paths | Where-Object { $_ -isnot [string] -or ![IO.Path]::IsPathRooted($_) }).Count) {
                    throw '無法讀取檔案清單，請重新選取文件。'
                }
            } catch [IO.IOException] {
                continue
            } catch {
                $row = $list.Items.Add('無法讀取的轉檔要求')
                $null = $row.SubItems.Add('未完成')
                $row.Tag = $_.Exception.Message
                $detail.Text = $row.Tag
                $state.Completed++
                $state.Failures++
                [IO.File]::Delete($request)
                continue
            }
            foreach ($path in $batch.Paths) {
                if ($seen.Add($path)) {
                    $row = $list.Items.Add([IO.Path]::GetFileName($path))
                    $null = $row.SubItems.Add('等候中')
                    $row.ToolTipText = $path
                    $row.Tag = $path
                    $queue.Enqueue([pscustomobject]@{ Path = $path; Row = $row })
                }
            }
            [IO.File]::Delete($request)
            $state.CloseAfter = $null
        }
    } finally { $gate.ReleaseMutex() }
}

function Start-NextFile {
    $task = $queue.Dequeue()
    $task.Row.SubItems[1].Text = '轉換中'
    $task.Row.ForeColor = [Drawing.Color]::FromArgb(24, 86, 162)
    $task.Row.EnsureVisible()
    $activity.Text = "正在轉換  $([IO.Path]::GetFileName($task.Path))"
    $activity.ForeColor = $ink
    $busy.Visible = $true
    $closeButton.Text = '最小化'
    $worker = [PowerShell]::Create()
    $worker.Runspace = $runspace
    $box = [hashtable]::Synchronized(@{})
    $null = $worker.AddScript({
        param($entry, $file, $box)
        & $entry -Paths $file -OnFileResult { param($record) $box.Result = $record }
    }).AddArgument($converter).AddArgument($task.Path).AddArgument($box)
    $state.Current = @{ Task = $task; Worker = $worker; Box = $box }
    try { $state.Current.Async = $worker.BeginInvoke() }
    catch { $state.Current.Error = $_.Exception.Message }
}

function Finish-CurrentFile {
    $current = $state.Current
    if (!$current.Error -and !$current.Async.IsCompleted) { return }
    $row = $current.Task.Row
    try {
        if ($current.Error) { throw $current.Error }
        $null = $current.Worker.EndInvoke($current.Async)
        $result = $current.Box.Result
        if (!$result) {
            if ($current.Worker.Streams.Error.Count) { throw $current.Worker.Streams.Error[0] }
            throw '轉檔程式未回報結果。'
        }
        if ($result.Status -ne 'Success') { throw $result.Message }
        $row.SubItems[1].Text = '已完成'
        $row.ForeColor = [Drawing.Color]::FromArgb(21, 107, 74)
        $row.Tag = $result.OutputPath
    } catch {
        $row.SubItems[1].Text = '未完成'
        $row.ForeColor = [Drawing.Color]::FromArgb(157, 48, 48)
        $row.Tag = $_.Exception.Message
        $state.Failures++
        if (!$list.SelectedItems.Count) { $detail.Text = "$($row.Text)  ·  $($row.Tag)" }
    } finally {
        $row.ToolTipText = $row.Tag
        if ($row.Selected) {
            $detail.Text = $row.Tag
            $openFolder.Visible = $row.SubItems[1].Text -eq '已完成'
        }
        $null = $seen.Remove($current.Task.Path)
        $current.Worker.Dispose()
        $state.Current = $null
        $state.Completed++
    }
}

try {
    if (!$Paths) { throw '請先在檔案總管選取要轉換的 Office 檔案。' }
    $fullPaths = @(foreach ($path in $Paths) {
        if ($path) { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path) }
    })
    $null = Enter-Mutex $gate
    try {
        [void][IO.Directory]::CreateDirectory($inbox)
        $request = Join-Path $inbox ([Guid]::NewGuid().ToString('N'))
        try {
            [IO.File]::WriteAllText("$request.tmp", (ConvertTo-Json -Compress -InputObject @{ Paths = $fullPaths }))
            [IO.File]::Move("$request.tmp", "$request.json")
        } finally { [IO.File]::Delete("$request.tmp") }
        $state.OwnsWindow = Enter-Mutex $owner 0
    } finally { $gate.ReleaseMutex() }
    if (!$state.OwnsWindow) { return }

    $converter = Join-Path $PSScriptRoot 'Convert-OfficeToPdf.ps1'
    $queue = [Collections.Generic.Queue[object]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()

    [Windows.Forms.Application]::EnableVisualStyles()
    if (!('OfficePdf.FileList' -as [type])) {
        Add-Type -Path (Join-Path $PSScriptRoot 'ProgressControls.cs') -ReferencedAssemblies System.Windows.Forms,System.Drawing
    }
    $highContrast = [Windows.Forms.SystemInformation]::HighContrast
    $ink = if ($highContrast) { [Drawing.SystemColors]::WindowText } else { [Drawing.ColorTranslator]::FromHtml('#172A40') }
    $muted = if ($highContrast) { $ink } else { [Drawing.ColorTranslator]::FromHtml('#506479') }
    $blue = if ($highContrast) { [Drawing.SystemColors]::Highlight } else { [Drawing.ColorTranslator]::FromHtml('#155FC5') }
    $form = [Windows.Forms.Form]::new()
    $form.Text = 'Office 轉 PDF'
    $form.Font = [Drawing.Font]::new('Microsoft JhengHei UI', 10)
    $form.BackColor = [Drawing.SystemColors]::Window
    $form.ForeColor = $ink
    $form.ShowIcon = $false
    $form.StartPosition = 'CenterScreen'
    $form.ClientSize = [Drawing.Size]::new(646, 445)
    $form.MinimumSize = [Drawing.Size]::new(490, 397)
    $form.AutoScaleMode = 'Dpi'
    $form.Padding = [Windows.Forms.Padding]::new(24, 16, 24, 16)

    $layout = [Windows.Forms.TableLayoutPanel]::new()
    $layout.Dock = 'Fill'
    $layout.ColumnCount = 1
    $form.Controls.Add($layout)
    function Add-Row($Control, [int]$Height) {
        $index = $layout.RowStyles.Count
        $layout.RowCount = $index + 1
        $style = if ($Height) { [Windows.Forms.RowStyle]::new('Absolute', $Height) }
                 else { [Windows.Forms.RowStyle]::new('Percent', 100) }
        $null = $layout.RowStyles.Add($style)
        $Control.Dock = 'Fill'
        $layout.Controls.Add($Control, 0, $index)
    }
    function New-Label {
        $label = [Windows.Forms.Label]::new()
        $label.TextAlign = 'MiddleLeft'
        $label.AutoEllipsis = $true
        return $label
    }
    $header = [Windows.Forms.Panel]::new()
    $header.Margin = [Windows.Forms.Padding]::Empty
    $heading = New-Label
    $heading.Dock = 'Fill'
    $heading.Font = [Drawing.Font]::new($form.Font.FontFamily, 19, [Drawing.FontStyle]::Bold)
    $heading.Text = '轉換為 PDF'
    $count = New-Label
    $count.Dock = 'Right'
    $count.Width = 190
    $count.TextAlign = 'MiddleRight'
    $count.ForeColor = $muted
    $header.Controls.Add($heading)
    $header.Controls.Add($count)
    Add-Row $header 52
    $progress = [OfficePdf.BatchProgress]::new()
    $progress.AccessibleName = '整批轉檔進度'
    $progress.Margin = [Windows.Forms.Padding]::new(0, 6, 0, 8)
    Add-Row $progress 21
    $activityPanel = [Windows.Forms.Panel]::new()
    $activityPanel.Margin = [Windows.Forms.Padding]::new(0, 1, 0, 12)
    $activityPanel.Padding = [Windows.Forms.Padding]::new(12, 0, 12, 0)
    if (!$highContrast) { $activityPanel.BackColor = [Drawing.ColorTranslator]::FromHtml('#EFF5FC') }
    $busy = New-Label
    $busy.Dock = 'Left'
    $busy.Width = 21
    $busy.ForeColor = $blue
    $busy.Text = '●'
    $busy.AccessibleRole = 'Graphic'
    $busy.AccessibleName = '處理中'
    $activity = New-Label
    $activity.Dock = 'Fill'
    $activityPanel.Controls.Add($activity)
    $activityPanel.Controls.Add($busy)
    Add-Row $activityPanel 52
    $list = [OfficePdf.FileList]::new()
    $list.AccessibleName = '轉檔檔案與狀態'
    $list.Margin = [Windows.Forms.Padding]::Empty
    $list.View = 'Details'
    $list.FullRowSelect = $true
    $list.MultiSelect = $false
    $list.ShowItemToolTips = $true
    $list.HideSelection = $false
    $null = $list.Columns.Add('檔案名稱', 370)
    $null = $list.Columns.Add('狀態', 112)
    $list.add_Resize({ $list.Columns[0].Width = [Math]::Max(100, $list.ClientSize.Width - $list.Columns[1].Width - 1) })
    Add-Row $list 0
    $detail = [Windows.Forms.TextBox]::new()
    $detail.ReadOnly = $true
    $detail.Multiline = $true
    $detail.ScrollBars = 'None'
    $detail.BorderStyle = 'None'
    $detail.BackColor = $form.BackColor
    $detail.ForeColor = $muted
    $detail.Margin = [Windows.Forms.Padding]::new(0, 8, 0, 0)
    $detail.AccessibleName = '輸出位置或錯誤原因'
    $detail.add_TextChanged({
        $lines = $detail.GetLineFromCharIndex([Math]::Max(0, $detail.TextLength - 1)) + 1
        $detail.ScrollBars = if ($lines * $detail.Font.Height -gt $detail.ClientSize.Height) { 'Vertical' } else { 'None' }
    })
    $list.add_SelectedIndexChanged({
        $openFolder.Visible = $false
        if ($list.SelectedItems.Count) {
            $selected = $list.SelectedItems[0]
            $detail.Text = $selected.Tag
            $openFolder.Visible = $selected.SubItems[1].Text -eq '已完成'
        }
    })
    Add-Row $detail 48
    $footer = [Windows.Forms.Panel]::new()
    $footer.Margin = [Windows.Forms.Padding]::Empty
    $destination = New-Label
    $destination.Text = 'PDF 儲存在原檔案旁'
    $destination.ForeColor = $muted
    $destination.Dock = 'Fill'
    $footer.Controls.Add($destination)
    $openFolder = [Windows.Forms.Button]::new()
    $openFolder.Text = '開啟資料夾'
    $openFolder.Visible = $false
    $openFolder.Dock = 'Right'
    $openFolder.Width = 115
    $openFolder.FlatStyle = 'Flat'
    $openFolder.FlatAppearance.BorderSize = 0
    $openFolder.ForeColor = $blue
    $openFolder.add_Click({
        if (!$list.SelectedItems.Count) { return }
        $folder = [IO.Path]::GetDirectoryName($list.SelectedItems[0].Tag)
        try {
            if (![IO.Directory]::Exists($folder)) { throw '找不到輸出資料夾。' }
            $start = [Diagnostics.ProcessStartInfo]::new($folder)
            $start.UseShellExecute = $true
            [void][Diagnostics.Process]::Start($start)
        } catch { $detail.Text = $_.Exception.Message }
    })
    $footer.Controls.Add($openFolder)
    $closeButton = [Windows.Forms.Button]::new()
    $closeButton.Text = '關閉'
    $closeButton.Width = 88
    $closeButton.Dock = 'Right'
    if (!$highContrast) {
        $closeButton.FlatStyle = 'Flat'
        $closeButton.FlatAppearance.BorderColor = [Drawing.ColorTranslator]::FromHtml('#CBD5DF')
        $closeButton.FlatAppearance.MouseOverBackColor = [Drawing.ColorTranslator]::FromHtml('#EDF3FA')
        $closeButton.BackColor = [Drawing.Color]::White
    }
    $closeButton.add_Click({
        if ($timer.Enabled -and ($state.Current -or $queue.Count)) { $form.WindowState = 'Minimized' }
        else { $form.Close() }
    })
    $footer.Controls.Add($closeButton)
    Add-Row $footer 34

    $timer = [Windows.Forms.Timer]::new()
    $timer.Interval = 150
    $timer.add_Tick({
        try {
            Receive-Requests
            if ($state.Current) { Finish-CurrentFile }
            if (!$state.Current -and $queue.Count) { Start-NextFile }
            $count.Text = "已處理 $($state.Completed) / $($list.Items.Count) 份"
            $progress.Maximum = [Math]::Max(1, $list.Items.Count)
            $progress.Value = $state.Completed
            if ($state.Current -and !$highContrast) {
                $busy.ForeColor = if ([datetime]::UtcNow.Millisecond -lt 500) { $blue } else { [Drawing.ColorTranslator]::FromHtml('#638FC5') }
            }
            if (!$state.Current -and !$queue.Count) {
                $busy.Visible = $false
                $closeButton.Text = '關閉'
                if ($state.Failures) {
                    $activity.Text = "有 $($state.Failures) 份未完成，點選檔案可查看原因。"
                } else {
                    $activity.Text = '轉換完成，PDF 已儲存在原檔案旁。'
                    if (!$state.CloseAfter) { $state.CloseAfter = [datetime]::UtcNow.AddSeconds(3) }
                    if (!$list.SelectedItems.Count -and [datetime]::UtcNow -ge $state.CloseAfter) { $form.Close() }
                }
            }
        } catch {
            $timer.Stop()
            $detail.Text = $_.Exception.Message
            $activity.Text = '無法更新轉檔進度。'
            $closeButton.Text = '關閉'
        }
    })
    $form.add_FormClosing({
        param($sender, $eventArgs)
        $null = Enter-Mutex $gate
        try {
            if ($timer.Enabled -and ($state.Current -or $queue.Count -or [IO.Directory]::GetFiles($inbox, '*.json').Length)) {
                $eventArgs.Cancel = $true
                $detail.Text = '轉檔仍在進行。你可以先將視窗最小化。'
            } elseif ($state.OwnsWindow) {
                $owner.ReleaseMutex()
                $state.OwnsWindow = $false
            }
        } finally { $gate.ReleaseMutex() }
    })
    $timer.Start()
    [void]$form.ShowDialog()
} catch {
    [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Office 轉 PDF', 'OK', 'Error')
} finally {
    if ($timer) { $timer.Dispose() }
    if ($form) { $form.Dispose() }
    if ($runspace) { $runspace.Dispose() }
    if ($state.OwnsWindow) { $owner.ReleaseMutex() }
    $owner.Dispose()
    $gate.Dispose()
}
