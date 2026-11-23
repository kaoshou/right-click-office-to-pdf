# Run with Windows PowerShell of the same architecture as the package.
# Exercises the native COM class directly, without installing it in Explorer.
param([Parameter(Mandatory = $true)][string]$Package)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class MenuTest {
    [DllImport("kernel32", CharSet=CharSet.Unicode, SetLastError=true)]
    static extern IntPtr LoadLibrary(string file);
    [DllImport("kernel32", CharSet=CharSet.Ansi, ExactSpelling=true)]
    static extern IntPtr GetProcAddress(IntPtr module, string name);
    [DllImport("kernel32")] static extern bool FreeLibrary(IntPtr module);
    [DllImport("shell32", CharSet=CharSet.Unicode, PreserveSig=false)]
    static extern void SHParseDisplayName(string path, IntPtr context, out IntPtr item, uint flags, IntPtr attrs);
    [DllImport("shell32", PreserveSig=false)]
    static extern void SHCreateShellItemArrayFromIDLists(uint count, IntPtr[] pidls, out IntPtr items);

    [ComImport, Guid("00000001-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IClassFactory {
        void CreateInstance(IntPtr outer, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out IExplorerCommand value);
        void LockServer(bool value);
    }
    [ComImport, Guid("a08ce4d0-fa25-44ab-b57c-c7b1c323e0b9"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IExplorerCommand {
        void GetTitle(IntPtr items, [MarshalAs(UnmanagedType.LPWStr)] out string title);
        void GetIcon(IntPtr items, out IntPtr icon);
        void GetToolTip(IntPtr items, out IntPtr tip);
        void GetCanonicalName(out Guid name);
        void GetState(IntPtr items, bool allowSlow, out uint state);
        void Invoke(IntPtr items, IntPtr context);
    }
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int GetFactory(ref Guid clsid, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out IClassFactory value);

    public static uint Check(string dll, Guid clsid, string[] paths, bool invoke) {
        IntPtr module = LoadLibrary(dll);
        if (module == IntPtr.Zero) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        var pidls = new IntPtr[paths.Length];
        IntPtr selection = IntPtr.Zero;
        IClassFactory factory = null;
        IExplorerCommand command = null;
        try {
            var factoryMethod = (GetFactory)Marshal.GetDelegateForFunctionPointer(GetProcAddress(module, "DllGetClassObject"), typeof(GetFactory));
            var factoryId = typeof(IClassFactory).GUID;
            Marshal.ThrowExceptionForHR(factoryMethod(ref clsid, ref factoryId, out factory));
            var commandId = typeof(IExplorerCommand).GUID;
            factory.CreateInstance(IntPtr.Zero, ref commandId, out command);
            command.GetCanonicalName(out commandId);
            if (commandId != clsid) throw new Exception("Unexpected command identity");
            string title;
            command.GetTitle(IntPtr.Zero, out title);
            if (String.IsNullOrWhiteSpace(title)) throw new Exception("Empty command title");
            for (int i = 0; i < paths.Length; ++i)
                SHParseDisplayName(paths[i], IntPtr.Zero, out pidls[i], 0, IntPtr.Zero);
            if (paths.Length > 0) SHCreateShellItemArrayFromIDLists((uint)paths.Length, pidls, out selection);
            uint state;
            command.GetState(selection, paths.Length > 16, out state);
            if (invoke) command.Invoke(selection, IntPtr.Zero);
            return state;
        } finally {
            if (selection != IntPtr.Zero) Marshal.Release(selection);
            foreach (IntPtr item in pidls) if (item != IntPtr.Zero) Marshal.FreeCoTaskMem(item);
            if (command != null) Marshal.FinalReleaseComObject(command);
            if (factory != null) Marshal.FinalReleaseComObject(factory);
            FreeLibrary(module);
        }
    }
}
'@
$work = Join-Path ([IO.Path]::GetTempPath()) ('menu-' + [IO.Path]::GetRandomFileName())
New-Item -ItemType Directory -Path $work | Out-Null
try {
    [IO.Compression.ZipFile]::ExtractToDirectory((Resolve-Path -LiteralPath $Package).ProviderPath, $work)
    $manifest = [xml](Get-Content -LiteralPath (Join-Path $work 'AppxManifest.xml') -Encoding UTF8 -Raw)
    $cls = $manifest.SelectSingleNode("//*[local-name()='Class']")
    $verb = $manifest.SelectSingleNode("//*[local-name()='Verb']")
    if ($cls.Id -ne $verb.Clsid) { throw 'Manifest class and verb do not match.' }
    $virtualization = $manifest.SelectSingleNode("//*[local-name()='FileSystemWriteVirtualization']")
    if ($virtualization.InnerText -ne 'disabled' -or !$manifest.SelectSingleNode("//*[local-name()='Capability' and @Name='unvirtualizedResources']")) {
        throw 'Desktop and packaged conversion must share the physical queue and PDF output paths.'
    }
    $dll = Join-Path $work $cls.Path
    $id = [guid]$verb.Clsid
    # The mock records the exact arguments that would reach the real UI.
    $sink = Join-Path $work 'Show-ConversionProgress.ps1'
    foreach ($file in 'Show-ConversionProgress.ps1', 'ProgressControls.cs') {
        if (!(Test-Path -LiteralPath (Join-Path $work $file) -PathType Leaf)) {
            throw "Package is missing a progress UI file: $file"
        }
    }
    Set-Content -LiteralPath $sink -Encoding UTF8 -Value @'
param([string[]]$Paths)
[IO.File]::WriteAllLines((Join-Path $PSScriptRoot 'received.tmp'), $Paths)
[IO.File]::Move((Join-Path $PSScriptRoot 'received.tmp'), (Join-Path $PSScriptRoot 'received.txt'))
'@
    $paths = @(foreach ($i in 1..320) {
        $ext = @('doc','pptx','ods')[$i % 3]
        $name = ("報告 [草稿] & O'Reilly {0} {1}.{2}" -f $i, ('x' * 64), $ext)
        $path = Join-Path $work $name
        [IO.File]::WriteAllText($path, '')
        $path
    })
    $text = Join-Path $work 'readme.txt'
    [IO.File]::WriteAllText($text, '')
    $folder = [IO.Directory]::CreateDirectory((Join-Path $work 'folder.docx')).FullName
    if ([MenuTest]::Check($dll,$id,@(),$false) -ne 2) { throw 'Empty selection should hide the command.' }
    if ([MenuTest]::Check($dll,$id,@($paths[0]),$false) -ne 0) { throw 'Supported file should enable the command.' }
    if ([MenuTest]::Check($dll,$id,@($paths[0],$text),$false) -ne 2) { throw 'Mixed unsupported selection should hide the command.' }
    if ([MenuTest]::Check($dll,$id,@($folder),$false) -ne 2) { throw 'Directories should hide the command.' }
    if (($paths -join ' ').Length -lt 32768) { throw 'Fixture is too short to exercise large selections.' }
    if ([MenuTest]::Check($dll,$id,$paths,$true) -ne 0) { throw 'Supported multi-selection should enable the command.' }
    $result = Join-Path $work 'received.txt'
    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    while (![IO.File]::Exists($result) -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
    $received = [IO.File]::ReadAllLines($result)
    if ($received.Length -ne $paths.Count -or (Compare-Object $paths $received -SyncWindow 0)) {
        throw "File paths changed in transit ($($paths.Count) sent, $($received.Length) received). First sent: $($paths[0]). First received: $($received[0])."
    }
    # The classic installation can use the same package's GUI-subsystem launcher.
    [IO.File]::Delete($result)
    $app = Join-Path $work $manifest.Package.Applications.Application.Executable
    & $app $paths[0] $paths[1]
    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    while (![IO.File]::Exists($result) -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
    $received = [IO.File]::ReadAllLines($result)
    if ($received.Count -ne 2 -or (Compare-Object $paths[0..1] $received -SyncWindow 0)) {
        throw 'Native launcher arguments changed in transit.'
    }
    Write-Host "Native COM selection ($($paths.Count) paths) and classic GUI launcher handoff passed."
} finally {
    $resolved = (Resolve-Path -LiteralPath $work).ProviderPath
    if ((Split-Path -Parent $resolved).TrimEnd('\') -eq [IO.Path]::GetTempPath().TrimEnd('\') -and (Split-Path -Leaf $resolved) -match '^menu-[a-z0-9]{8}\.[a-z0-9]{3}$') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
