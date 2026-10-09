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
    [DllImport("user32.dll")] public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")] public static extern int GetSystemMetrics(int nIndex);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hWnd);
    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc lpEnumFunc, IntPtr lParam);
    [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);
    public static IntPtr[] TopWindows() {
        var list = new System.Collections.Generic.List<IntPtr>();
        EnumWindows(delegate (IntPtr h, IntPtr l) { list.Add(h); return true; }, IntPtr.Zero);
        return list.ToArray();
    }
    public static string TitleOf(IntPtr hWnd) {
        int n = GetWindowTextLength(hWnd);
        if (n <= 0) { return ""; }
        var sb = new System.Text.StringBuilder(n + 1);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }
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
      @{ ok; message }. Restores a minimized window (and leaves a maximized
      one maximized); retries the
      SetForegroundWindow twice, tapping Alt first (Windows refuses it
      while another process holds the foreground lock).
    #>
    param([IntPtr]$HWnd, [int]$SettleMs = 300)
    [void](Get-EbiNative)
    if ($HWnd -eq [IntPtr]::Zero -or -not [EbiNative]::IsWindow($HWnd)) { return @{ ok = $false; message = 'the window handle is not a window (closed?)' } }
    for ($try = 1; $try -le 3; $try++) {
        # SW_RESTORE (9) only for a MINIMIZED window: on a maximized one it
        # would un-maximize it, and screen geometry measured on the maximized
        # window would land in the wrong place. Otherwise SW_SHOW (5).
        if ([EbiNative]::IsIconic($HWnd)) { [void][EbiNative]::ShowWindowAsync($HWnd, 9) } else { [void][EbiNative]::ShowWindowAsync($HWnd, 5) }
        # Retries: tap Alt first. Windows refuses SetForegroundWindow to a
        # process that did not get the last input (the console the operator
        # just pressed Enter in keeps the lock); a key event from this
        # process lifts that refusal. Seen on the first office run.
        if ($try -gt 1) { [EbiNative]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero); [EbiNative]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero) }
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
    param([int]$SelectWaitMs = 400, [int]$CopyWaitMs = 400, [int]$MaxCopyWaitMs = 8000)
    [void](Get-EbiNative)
    try { [System.Windows.Forms.Clipboard]::Clear() } catch { }
    Start-Sleep -Milliseconds 100
    Send-EbiKeys -Keys '^a' -WaitMs $SelectWaitMs
    Send-EbiKeys -Keys '^c' -WaitMs $CopyWaitMs
    # A long page (a Jenkins file list of thousands of rows) takes the
    # browser seconds to put on the clipboard: wait until the text is there
    # and has stopped growing, up to MaxCopyWaitMs. The first office run
    # read an empty clipboard seven times in a row after a fixed 400 ms.
    $text = Get-EbiClipboardText
    $deadline = (Get-Date).AddMilliseconds([Math]::Max(0, $MaxCopyWaitMs - $CopyWaitMs))
    $last = -1
    while ((Get-Date) -lt $deadline) {
        if ($text.Length -gt 0 -and $text.Length -eq $last) { break }
        $last = $text.Length
        Start-Sleep -Milliseconds 300
        $text = Get-EbiClipboardText
    }
    Send-EbiKeys -Keys '{ESC}' -WaitMs 100
    return $text
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
    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        # one separator style: '\\srv\share\x/DATA/y' reaches DF.exe and Excel as written otherwise
        try { return [System.IO.Path]::GetFullPath($PathValue) } catch { return $PathValue }
    }
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

function Get-EbiTopWindows {
    <#
      Every visible, titled, top-level window (no owner) as
      @{ handle; title; processId; processName; minimized; maximized }, in
      Z order (front first). The candidates screen.find_window filters.
    #>
    [void](Get-EbiNative)
    $out = New-Object System.Collections.ArrayList
    $names = @{}
    foreach ($h in [EbiNative]::TopWindows()) {
        if (-not [EbiNative]::IsWindowVisible($h)) { continue }
        if ([EbiNative]::GetWindow($h, 4) -ne [IntPtr]::Zero) { continue }   # GW_OWNER: a dialog / tool window
        $title = [EbiNative]::TitleOf($h)
        if ([string]::IsNullOrWhiteSpace($title)) { continue }
        [uint32]$procId = 0
        [void][EbiNative]::GetWindowThreadProcessId($h, [ref]$procId)
        $pn = ''
        $key = [string]$procId
        if ($names.Contains($key)) { $pn = $names[$key] } else { try { $pn = (Get-Process -Id ([int]$procId) -ErrorAction Stop).ProcessName } catch { $pn = '' }; $names[$key] = $pn }
        [void]$out.Add(@{ handle = $h; title = $title; processId = [int]$procId; processName = $pn; minimized = [bool][EbiNative]::IsIconic($h); maximized = [bool][EbiNative]::IsZoomed($h) })
    }
    return $out.ToArray()
}

function Set-EbiWindowState {
    # 'maximize' | 'restore' | 'minimize' (ShowWindowAsync 3 / 9 / 6).
    param([IntPtr]$HWnd, [string]$State, [int]$SettleMs = 400)
    [void](Get-EbiNative)
    $cmd = switch ($State) { 'maximize' { 3 } 'minimize' { 6 } default { 9 } }
    [void][EbiNative]::ShowWindowAsync($HWnd, $cmd)
    if ($SettleMs -gt 0) { Start-Sleep -Milliseconds $SettleMs }
}

function Close-EbiWindow {
    # Ask the window to close (WM_CLOSE). @{ ok; message }; a window that is
    # already gone is ok (a release must survive "nothing to release").
    param([IntPtr]$HWnd, [int]$WaitMs = 1500)
    [void](Get-EbiNative)
    if ($HWnd -eq [IntPtr]::Zero -or -not [EbiNative]::IsWindow($HWnd)) { return @{ ok = $true; message = 'already closed' } }
    [void][EbiNative]::PostMessage($HWnd, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
    $deadline = (Get-Date).AddMilliseconds([Math]::Max(100, $WaitMs))
    while ((Get-Date) -lt $deadline) {
        if (-not [EbiNative]::IsWindow($HWnd)) { return @{ ok = $true; message = '' } }
        Start-Sleep -Milliseconds 100
    }
    if ([EbiNative]::IsWindow($HWnd)) { return @{ ok = $false; message = 'the window is still open after WM_CLOSE (a save prompt?)' } }
    return @{ ok = $true; message = '' }
}

function Set-EbiClipboardRich {
    <#
      Put HTML (CF_HTML, already wrapped -- kernel/RichClip.ps1 New-EbiCfHtml),
      RTF and plain text on the clipboard in one data object, so a chat box
      that takes pictures gets the pictures and one that does not still gets
      the words. Needs an STA thread (powershell.exe 5.1's console is).
      @{ ok; message }.
    #>
    param([string]$CfHtml, [string]$Rtf, [string]$Text, [int]$WaitMs = 300)
    [void](Get-EbiNative)
    try {
        $do = New-Object System.Windows.Forms.DataObject
        if (-not [string]::IsNullOrEmpty($CfHtml)) {
            # CF_HTML must reach the clipboard as UTF-8 bytes; a .NET string
            # would be marshalled as UTF-16 and the byte offsets would lie.
            $ms = New-Object System.IO.MemoryStream (, ((New-Object System.Text.UTF8Encoding($false)).GetBytes($CfHtml)))
            $do.SetData('HTML Format', $ms)
        }
        if (-not [string]::IsNullOrEmpty($Rtf)) { $do.SetData([System.Windows.Forms.DataFormats]::Rtf, $Rtf) }
        if (-not [string]::IsNullOrEmpty($Text)) { $do.SetData([System.Windows.Forms.DataFormats]::UnicodeText, $Text) }
        [System.Windows.Forms.Clipboard]::SetDataObject($do, $true, 5, 200)
        if ($WaitMs -gt 0) { Start-Sleep -Milliseconds $WaitMs }
        return @{ ok = $true; message = '' }
    } catch { return @{ ok = $false; message = $_.Exception.Message } }
}

function Set-EbiClipboardImage {
    # One PNG on the clipboard as a bitmap (the "sequence" share mode).
    param([string]$Path, [int]$WaitMs = 300)
    [void](Get-EbiNative)
    Add-Type -AssemblyName System.Drawing
    try {
        $img = [System.Drawing.Image]::FromFile($Path)
        try { [System.Windows.Forms.Clipboard]::SetImage($img) } finally { $img.Dispose() }
        if ($WaitMs -gt 0) { Start-Sleep -Milliseconds $WaitMs }
        return @{ ok = $true; message = '' }
    } catch { return @{ ok = $false; message = $_.Exception.Message } }
}

function Get-EbiWindowVisibleRect {
    <#
      The window's VISIBLE bounds (DWM extended frame bounds): Windows 10
      pads GetWindowRect with invisible resize borders, so a capture of the
      plain rect shows a strip of whatever is behind the window. Falls back
      to GetWindowRect. @{ ok; X; Y; W; H; borderL; borderT; borderR; borderB }
      (the borders are what GetWindowRect adds on each side).
    #>
    param([IntPtr]$HWnd)
    $outer = Get-EbiWindowRect -HWnd $HWnd
    if (-not $outer['ok']) { return @{ ok = $false; X = 0; Y = 0; W = 0; H = 0; borderL = 0; borderT = 0; borderR = 0; borderB = 0 } }
    $vis = New-Object EbiNative+RECT
    $hr = -1
    try { $hr = [EbiNative]::DwmGetWindowAttribute($HWnd, 9, [ref]$vis, 16) } catch { $hr = -1 }   # DWMWA_EXTENDED_FRAME_BOUNDS
    if ($hr -ne 0 -or ($vis.Right - $vis.Left) -le 0) { $outer['borderL'] = 0; $outer['borderT'] = 0; $outer['borderR'] = 0; $outer['borderB'] = 0; return $outer }
    $bl = $vis.Left - $outer['X']; $bt = $vis.Top - $outer['Y']; $br = ($outer['X'] + $outer['W']) - $vis.Right; $bb = ($outer['Y'] + $outer['H']) - $vis.Bottom
    foreach ($b in @($bl, $bt, $br, $bb)) {
        # DWM answers in physical pixels, GetWindowRect in this (DPI-unaware)
        # process's scaled ones: at a scaling other than 100% the "borders"
        # come out absurd -- trust GetWindowRect then.
        if ($b -lt 0 -or $b -gt 16) { $outer['borderL'] = 0; $outer['borderT'] = 0; $outer['borderR'] = 0; $outer['borderB'] = 0; return $outer }
    }
    return @{ ok = $true; X = $vis.Left; Y = $vis.Top; W = ($vis.Right - $vis.Left); H = ($vis.Bottom - $vis.Top);
              borderL = ($vis.Left - $outer['X']); borderT = ($vis.Top - $outer['Y']); borderR = (($outer['X'] + $outer['W']) - $vis.Right); borderB = (($outer['Y'] + $outer['H']) - $vis.Bottom) }
}

function Set-EbiWindowVisibleSize {
    # Move / resize so the VISIBLE window is X,Y,W,H (borders added back).
    param([IntPtr]$HWnd, [int]$X, [int]$Y, [int]$W, [int]$H, [int]$SettleMs = 300)
    [void](Get-EbiNative)
    $v = Get-EbiWindowVisibleRect -HWnd $HWnd
    if (-not $v['ok']) { return $false }
    [void][EbiNative]::MoveWindow($HWnd, $X - [int]$v['borderL'], $Y - [int]$v['borderT'], $W + [int]$v['borderL'] + [int]$v['borderR'], $H + [int]$v['borderT'] + [int]$v['borderB'], $true)
    if ($SettleMs -gt 0) { Start-Sleep -Milliseconds $SettleMs }
    return $true
}

function Invoke-EbiFindText {
    # Ctrl+F the text (through the clipboard: SendKeys mangles kana and
    # some symbols), Esc to close the bar -- Chromium then leaves the focus
    # on the FIRST match (the active one while typing; an Enter here would
    # jump to the second) or the link it sits in.
    param([string]$Text)
    Set-EbiClipboardText -Text $Text
    Send-EbiKeys -Keys '^{f}' -WaitMs 300
    Send-EbiKeys -Keys '^a' -WaitMs 100
    Send-EbiKeys -Keys '^v' -WaitMs 500
    Send-EbiKeys -Keys '{ESC}' -WaitMs 300
}

function Test-EbiKeyRecipe {
    # PURE. Every entry is find:<text> | keys:<SendKeys> | wait:<ms>.
    # @{ ok; message }.
    param($Recipe)
    $i = 0
    foreach ($r in @($Recipe)) {
        $i++
        if ([string]$r -notmatch '^(find:.+|keys:.+|wait:\d+)$') { return @{ ok = $false; message = ('recipe entry ' + $i + ' "' + [string]$r + '" is not find:<text>, keys:<SendKeys> or wait:<ms>') } }
    }
    return @{ ok = $true; message = '' }
}

function Invoke-EbiKeyRecipe {
    <#
      Run a key recipe against a window, re-checking the foreground before
      each entry (P0-R12). Entries: find:<text> (Invoke-EbiFindText),
      keys:<SendKeys>, wait:<ms>. @{ ok; failure; message } -- failure is
      foreground_lost or recipe_invalid.
    #>
    param([IntPtr]$HWnd, $Recipe, [int]$KeyWaitMs = 300)
    $t = Test-EbiKeyRecipe -Recipe $Recipe
    if (-not $t['ok']) { return @{ ok = $false; failure = 'recipe_invalid'; message = $t['message'] } }
    foreach ($r in @($Recipe)) {
        $s = [string]$r
        if ($s -match '^wait:(\d+)$') { Start-Sleep -Milliseconds ([int]$Matches[1]); continue }
        $fg = Set-EbiForeground -HWnd $HWnd -SettleMs 150
        if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message'] } }
        if ($s -match '^find:(.+)$') { Invoke-EbiFindText -Text $Matches[1] }
        elseif ($s -match '^keys:(.+)$') { Send-EbiKeys -Keys $Matches[1] -WaitMs $KeyWaitMs }
    }
    return @{ ok = $true; failure = ''; message = '' }
}

function Invoke-EbiDeselect {
    # Click a blank point of the window (window-relative px) so the Ctrl+A
    # selection Read-EbiPageText leaves behind is gone before a screenshot:
    # Esc does not clear an Edge selection (JenkinsSnap.ps1 / HmSnap.ps1 click
    # for the same reason).
    param([IntPtr]$HWnd, [int]$X, [int]$Y)
    $r = Get-EbiWindowRect -HWnd $HWnd
    if (-not $r['ok']) { return $false }
    Invoke-EbiClick -X ([int]$r['X'] + $X) -Y ([int]$r['Y'] + $Y) -SettleMs 300
    return $true
}
