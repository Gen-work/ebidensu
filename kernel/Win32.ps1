# ============================================================
#  kernel/Win32.ps1
#
#  The user32 P/Invoke surface shared by window-facing steps. Dot-source
#  only (no param()). ASCII source.
#
#  Compiled lazily: Add-Type at dot-source time would make every step
#  that includes this library pay the compile even when only its
#  manifest is being read (the contract checker, ebi help, catalog
#  generation). Steps call Get-EbiWin32 inside Invoke-Step instead.
#
#  Steps must not dot-source Common.ps1 for [WinAPI]: that file carries
#  $Global:Shell / $Global:Timing, the exact global state ebi-dance is
#  removing (spec/STEP-CONTRACT.md 3.4).
# ============================================================

function Get-EbiWin32 {
    # Returns the [EbiWin32] type, compiling it on first use.
    $existing = ([System.Management.Automation.PSTypeName]'EbiWin32').Type
    if ($null -ne $existing) { return $existing }

    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[StructLayout(LayoutKind.Sequential)]
public struct EbiRect { public int Left; public int Top; public int Right; public int Bottom; }

public static class EbiWin32 {
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out EbiRect lpRect);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
}
"@
    return ([System.Management.Automation.PSTypeName]'EbiWin32').Type
}

function Test-EbiWindowHandle {
    # $true when the handle is non-zero and Windows still knows the window.
    # This is the consumer-side validity check STEP-CONTRACT.md 3.1 asks
    # for before a step returns session_invalid.
    param($Handle)
    if ($null -eq $Handle) { return $false }
    $h = [IntPtr]$Handle
    if ($h -eq [IntPtr]::Zero) { return $false }
    $win = Get-EbiWin32
    return [bool]$win::IsWindow($h)
}

function Get-EbiWindowRect {
    # Screen rectangle of a window as @{ X; Y; W; H }, or $null when
    # GetWindowRect fails.
    param($Handle)
    $win  = Get-EbiWin32
    $rect = New-Object EbiRect
    if (-not $win::GetWindowRect([IntPtr]$Handle, [ref]$rect)) { return $null }
    return @{
        X = $rect.Left
        Y = $rect.Top
        W = ($rect.Right - $rect.Left)
        H = ($rect.Bottom - $rect.Top)
    }
}

function Set-EbiForegroundWindow {
    # Restore-if-minimized + SetForegroundWindow, then report whether the
    # window really is foreground now. Callers decide what a $false means
    # (browser.ensure -> activate_failed; key-sending steps ->
    # foreground_lost, spec/STEP-CONTRACT.md 4).
    param($Handle, [int]$SettleMs = 400)
    $win = Get-EbiWin32
    $h   = [IntPtr]$Handle
    [void]$win::ShowWindowAsync($h, 9)     # SW_RESTORE
    [void]$win::SetForegroundWindow($h)
    if ($SettleMs -gt 0) { Start-Sleep -Milliseconds $SettleMs }
    return ($win::GetForegroundWindow() -eq $h)
}
