# ============================================================
#  Common.ps1 - Shared utility functions
# ============================================================

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinAPI {
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    public struct RECT {
        public int Left; public int Top; public int Right; public int Bottom;
    }
}

public class MouseAPI {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
}
"@

$Global:Shell = New-Object -ComObject WScript.Shell

# ── Directory ──
function Ensure-Dir([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

# ── Type conversion ──
function To-Int($value) {
    if ($null -eq $value -or $value -eq '') { return 0 }
    try { return [int]([double]$value) } catch { return 0 }
}

# ── Screenshot ──
function Take-WindowScreenshot($hWnd, $savePath) {
    $rect = New-Object WinAPI+RECT
    [WinAPI]::GetWindowRect($hWnd, [ref]$rect) | Out-Null

    $width  = $rect.Right  - $rect.Left
    $height = $rect.Bottom - $rect.Top

    if ($width -le 0 -or $height -le 0) {
        Write-Host "  [WARN] Window size error: $savePath" -ForegroundColor Red
        return
    }

    $bmp = New-Object System.Drawing.Bitmap($width, $height)
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    $gfx.CopyFromScreen($rect.Left, $rect.Top, 0, 0, [System.Drawing.Size]::new($width, $height))
    $bmp.Save($savePath, [System.Drawing.Imaging.ImageFormat]::Png)
    $gfx.Dispose()
    $bmp.Dispose()
}

function Take-ForegroundScreenshot([string]$savePath) {
    $hWnd = [WinAPI]::GetForegroundWindow()
    if ($hWnd -ne [IntPtr]::Zero) {
        Take-WindowScreenshot $hWnd $savePath
        Write-Host ("  >> Saved: {0}" -f (Split-Path $savePath -Leaf)) -ForegroundColor Green
    } else {
        Write-Host "  [WARN] No foreground window." -ForegroundColor Red
    }
}

# ── Edge window control ──
function Activate-EdgeWindow {
    $null = $Shell.AppActivate("Microsoft Edge")
    Start-Sleep -Milliseconds 700

    $hWnd = [WinAPI]::GetForegroundWindow()
    if ($hWnd -eq [IntPtr]::Zero) {
        Write-Host "  [WARN] Edge window not found." -ForegroundColor Yellow
        return $hWnd
    }

    [WinAPI]::ShowWindowAsync($hWnd, 9) | Out-Null
    [WinAPI]::SetForegroundWindow($hWnd) | Out-Null
    Start-Sleep -Milliseconds 400
    return $hWnd
}

function Switch-ToEdge {
    $Shell.SendKeys("%{TAB}")
    Start-Sleep -Milliseconds 900
    [void](Activate-EdgeWindow)
}

function Click-PageBody {
    $hWnd = [WinAPI]::GetForegroundWindow()
    if ($hWnd -eq [IntPtr]::Zero) { return }

    $rect = New-Object WinAPI+RECT
    [WinAPI]::GetWindowRect($hWnd, [ref]$rect) | Out-Null

    $x = $rect.Left + 150
    $y = $rect.Top + 150

    [MouseAPI]::SetCursorPos($x, $y) | Out-Null
    Start-Sleep -Milliseconds 100
    [MouseAPI]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)  # LEFTDOWN
    Start-Sleep -Milliseconds 50
    [MouseAPI]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)  # LEFTUP
    Start-Sleep -Milliseconds 400
}

function Reset-FocusToBody {
    [void](Activate-EdgeWindow)
    Click-PageBody
}

# ── SendKeys helpers ──
function Set-ClipboardText([string]$text) {
    [System.Windows.Forms.Clipboard]::SetText($text)
    Start-Sleep -Milliseconds 200
}

function Send-Key([string]$keys, [int]$waitMs = 300) {
    [System.Windows.Forms.SendKeys]::SendWait($keys)
    Start-Sleep -Milliseconds $waitMs
}

function Send-Tab([int]$count = 1) {
    for ($i = 0; $i -lt $count; $i++) {
        Send-Key "{TAB}" $Global:Timing.ActionWaitMs
    }
}

function Send-ShiftTab([int]$count = 1) {
    for ($i = 0; $i -lt $count; $i++) {
        Send-Key "+{TAB}" $Global:Timing.ActionWaitMs
    }
}

function Send-Enter { Send-Key "{ENTER}" $Global:Timing.ActionWaitMs }
function Send-CtrlF { Send-Key "^{f}" $Global:Timing.ActionWaitMs }

function Paste-Replace([string]$text) {
    Set-ClipboardText $text
    Send-Key "^{a}" 150
    Send-Key "^v" 300
}

# ── UI helpers ──
function Show-RowHeader($item, [string]$stepName, [string]$sheetName) {
    Write-Host ""
    Write-Host ("=" * 72) -ForegroundColor White
    Write-Host ("  {0} | Sheet={1} | CorrelID={2} | File={3}" -f $stepName, $sheetName, $item.ColB, $item.ColD) -ForegroundColor White
    Write-Host ("=" * 72) -ForegroundColor White
}

function Wait-PagePrepared([string]$message) {
    Write-Host ""
    Write-Host $message -ForegroundColor Yellow
    Write-Host "Enter=OK / q=quit : " -ForegroundColor Magenta -NoNewline
    $resp = Read-Host
    if ($resp -eq 'q') { exit }
}

function Prepare-EdgeStep([string]$message) {
    Wait-PagePrepared $message
    Switch-ToEdge
    Click-PageBody
}
