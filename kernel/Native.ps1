#Requires -Version 5.1
# ============================================================
#  kernel/Native.ps1
#
#  The one place Win32 / SendKeys / clipboard are touched (P1-11, P1-18
#  decided the shared binding). Dot-source only (no param() block, ASCII
#  source, no class of our own). Steps under modules/ dot-source it with
#      . (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')
#  -- a step may not call another step, but every step may share a kernel
#  library (the same rule P1-27 gives kernel/Key.ps1).
#
#  Everything here is IMPURE and Windows-only, and nothing is compiled
#  until Get-EbiNative is called: a DryRun path never reaches it, so the
#  Linux CI that dry-runs every step never sees user32 or GDI+.
#
#  Foreground rule (STEP-CONTRACT.md 4, P0-R12): a step that sends keys
#  or clicks brings ITS window to the front and verifies it is there
#  (Set-EbiForeground); it never sends to "whatever is in front".
# ============================================================

function Get-EbiNative {
    # Compile the Win32 bindings once per process; returns the type.
    if (-not ('EbiNative' -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class EbiNative {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")] public static extern int GetSystemMetrics(int nIndex);
}
"@
    }
    Add-Type -AssemblyName System.Windows.Forms
    return ('EbiNative' -as [type])
}

function ConvertTo-EbiHandle {
    # Whatever a provides step registered: IntPtr, or an int from a fixture.
    param($Value)
    if ($null -eq $Value) { return [IntPtr]::Zero }
    if ($Value -is [IntPtr]) { return $Value }
    try { return [IntPtr]([int64]$Value) } catch { return [IntPtr]::Zero }
}

function Get-EbiWindowRect {
    # @{ ok; X; Y; W; H } -- ok=$false when the window is gone or empty.
    param([IntPtr]$HWnd)
    [void](Get-EbiNative)
    if ($HWnd -eq [IntPtr]::Zero -or -not [EbiNative]::IsWindow($HWnd)) { return @{ ok = $false; X = 0; Y = 0; W = 0; H = 0 } }
    $rect = New-Object EbiNative+RECT
    if (-not [EbiNative]::GetWindowRect($HWnd, [ref]$rect)) { return @{ ok = $false; X = 0; Y = 0; W = 0; H = 0 } }
    $w = $rect.Right - $rect.Left; $h = $rect.Bottom - $rect.Top
    return @{ ok = ($w -gt 0 -and $h -gt 0); X = $rect.Left; Y = $rect.Top; W = $w; H = $h }
}

function Set-EbiForeground {
    <#
      Bring the window to the front and VERIFY it is there. Returns
      @{ ok; message }. Restores a minimized window first; retries the
      SetForegroundWindow once (Windows refuses it when another process
      holds the foreground lock, and a second try after a short wait
      usually goes through).
    #>
    param([IntPtr]$HWnd, [int]$SettleMs = 300)
    [void](Get-EbiNative)
    if ($HWnd -eq [IntPtr]::Zero -or -not [EbiNative]::IsWindow($HWnd)) { return @{ ok = $false; message = 'the window handle is not a window (closed?)' } }
    for ($try = 1; $try -le 2; $try++) {
        [void][EbiNative]::ShowWindowAsync($HWnd, 9)   # SW_RESTORE
        [void][EbiNative]::SetForegroundWindow($HWnd)
        Start-Sleep -Milliseconds ([Math]::Max(50, $SettleMs))
        if ([EbiNative]::GetForegroundWindow() -eq $HWnd) { return @{ ok = $true; message = '' } }
    }
    return @{ ok = $false; message = ('the window did not come to the front (foreground is ' + [string][EbiNative]::GetForegroundWindow() + ')') }
}

function Send-EbiKeys {
    # SendKeys.SendWait with the wait the caller asks for.
    param([string]$Keys, [int]$WaitMs = 300)
    [void](Get-EbiNative)
    [System.Windows.Forms.SendKeys]::SendWait($Keys)
    if ($WaitMs -gt 0) { Start-Sleep -Milliseconds $WaitMs }
}

function Set-EbiClipboardText {
    param([string]$Text, [int]$WaitMs = 200)
    [void](Get-EbiNative)
    if ([string]::IsNullOrEmpty($Text)) { [System.Windows.Forms.Clipboard]::Clear() } else { [System.Windows.Forms.Clipboard]::SetText($Text) }
    if ($WaitMs -gt 0) { Start-Sleep -Milliseconds $WaitMs }
}

function Get-EbiClipboardText {
    [void](Get-EbiNative)
    try { return [string][System.Windows.Forms.Clipboard]::GetText() } catch { return '' }
}

function Read-EbiPageText {
    <#
      The visible text of the window in front: Ctrl+A, Ctrl+C, Esc, then
      the clipboard (Read-PageText.ps1's routine). The caller has already
      put its window in front (Set-EbiForeground). Returns the text ('' when
      the clipboard came back empty).
    #>
    param([int]$SelectWaitMs = 400, [int]$CopyWaitMs = 400)
    [void](Get-EbiNative)
    try { [System.Windows.Forms.Clipboard]::Clear() } catch { }
    Start-Sleep -Milliseconds 100
    Send-EbiKeys -Keys '^a' -WaitMs $SelectWaitMs
    Send-EbiKeys -Keys '^c' -WaitMs $CopyWaitMs
    Send-EbiKeys -Keys '{ESC}' -WaitMs 100
    return (Get-EbiClipboardText)
}

function Invoke-EbiClick {
    # A left click at screen coordinates.
    param([int]$X, [int]$Y, [int]$SettleMs = 400)
    [void](Get-EbiNative)
    [void][EbiNative]::SetCursorPos($X, $Y)
    Start-Sleep -Milliseconds 100
    [EbiNative]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)   # LEFTDOWN
    Start-Sleep -Milliseconds 50
    [EbiNative]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)   # LEFTUP
    if ($SettleMs -gt 0) { Start-Sleep -Milliseconds $SettleMs }
}

function Get-EbiVirtualScreen {
    # @{ X; Y; W; H } of the virtual screen (all monitors).
    [void](Get-EbiNative)
    return @{ X = [EbiNative]::GetSystemMetrics(76); Y = [EbiNative]::GetSystemMetrics(77); W = [EbiNative]::GetSystemMetrics(78); H = [EbiNative]::GetSystemMetrics(79) }
}

function Resolve-EbiWorkPath {
    # PURE. A step's path input: rooted stays; relative goes under WorkDir;
    # separators normalized (the P0-08 decision, shared by every step
    # that writes a file).
    param([string]$PathValue, [string]$WorkDir)
    if ([string]::IsNullOrWhiteSpace($PathValue)) { return '' }
    if ([System.IO.Path]::IsPathRooted($PathValue)) { return $PathValue }
    if ([string]::IsNullOrWhiteSpace($WorkDir)) { return $PathValue }
    return [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($WorkDir, $PathValue))
}

function Write-EbiTextFile {
    # UTF-8 without BOM, parent directory created. @{ ok; message }.
    param([string]$Path, [string]$Text)
    try {
        $dir = Split-Path -Path $Path -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [System.IO.File]::WriteAllText($Path, $Text, (New-Object System.Text.UTF8Encoding($false)))
        return @{ ok = $true; message = '' }
    } catch { return @{ ok = $false; message = $_.Exception.Message } }
}
