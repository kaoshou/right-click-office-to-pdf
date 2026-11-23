param([Parameter(Mandatory = $true)][long]$SelectionHandle)

$ErrorActionPreference = 'Stop'
$handle = [Microsoft.Win32.SafeHandles.SafeFileHandle]::new([IntPtr]::new($SelectionHandle), $true)
try {
    $stream = [IO.FileStream]::new($handle, [IO.FileAccess]::Read)
    try {
        $reader = [IO.StreamReader]::new($stream, [Text.Encoding]::Unicode, $true)
        try {
            $files = [Collections.Generic.List[string]]::new()
            while ($null -ne ($line = $reader.ReadLine())) {
                if ($line.Length) { $files.Add($line) }
            }
        } finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
} finally { $handle.Dispose() }
& (Join-Path $PSScriptRoot 'Show-ConversionProgress.ps1') -Paths $files.ToArray()
exit $LASTEXITCODE
