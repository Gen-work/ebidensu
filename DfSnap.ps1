# ============================================================
#  DfSnap.ps1
#
#  Phase: DfSnap
#
#  For each pending row in the mapping:
#    1. Locate the GIFT and GFIX data files for the correlid.
#    2. Launch df.exe with both file paths as arguments.
#    3. Wait for the window to become ready, then capture a
#       full-screen (or foreground-window) screenshot.
#    4. Save the PNG to snap\DF\<Correl_ID_S>.png.
#    5. Close df.exe and continue to the next row.
#    6. On success, set DF_snap = 1 in the mapping.
#
#  File discovery: looks for files matching <FilePattern> (default
#  "{0}*", where {0} = Correl_ID_S) in -GiftDataDir and -GfixDataDir.
#  If multiple candidates exist, the most-recently-modified file wins.
#
#  Usage:
#    .\DfSnap.ps1 -WorkDir C:\work\proj -DfExePath "C:\tools\df.exe" `
#                 -GiftDataDir C:\work\proj\DATA\GIFT `
#                 -GfixDataDir C:\work\proj\DATA\GFIX
# ============================================================

param(
    [string]$WorkDir,
    [string]$Owner = ([char]0x53B3),
    [string[]]$TargetIds = @(),
    [switch]$Force,

    # df.exe path. Must be set (no sensible default).
    [Parameter(Mandatory=$false)]
    [string]$DfExePath = '',

    # Directories to search for data files.
    [string]$GiftDataDir = '',
    [string]$GfixDataDir = '',

    # File search pattern. {0} = Correl_ID_S.
    [string]$FilePattern = '{0}*',

    # Seconds to wait after df.exe window becomes idle before screenshot.
    [int]$LoadWaitSec = 8,

    # 'fullscreen' = whole primary screen; 'window' = foreground window only.
    [ValidateSet('fullscreen','window')]
    [string]$CaptureMode = 'fullscreen',

    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

try {
    Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
} catch {}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1
}

$forceFlag = [bool]$Force.IsPresent
$dryFlag   = [bool]$DryRun.IsPresent

# Derive default data dirs from WorkDir if not supplied
if ([string]::IsNullOrWhiteSpace($GiftDataDir)) {
    $GiftDataDir = Join-Path $WorkDir 'DATA\GIFT'
}
if ([string]::IsNullOrWhiteSpace($GfixDataDir)) {
    $GfixDataDir = Join-Path $WorkDir 'DATA\GFIX'
}

$snapDir = Join-Path $WorkDir 'snap\DF'

# ── Header ───────────────────────────────────────────────────
$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)

Write-Host ''
Write-Host '===== DfSnap =====' -ForegroundColor Green
Write-Host ("  WorkDir     : {0}" -f $WorkDir)
Write-Host ("  Mapping     : {0}" -f $mappingPath)
Write-Host ("  DfExePath   : {0}" -f $DfExePath)
Write-Host ("  GiftDataDir : {0}" -f $GiftDataDir)
Write-Host ("  GfixDataDir : {0}" -f $GfixDataDir)
Write-Host ("  FilePattern : {0}" -f $FilePattern)
Write-Host ("  LoadWaitSec : {0}" -f $LoadWaitSec)
Write-Host ("  CaptureMode : {0}" -f $CaptureMode)
Write-Host ("  SnapDir     : {0}" -f $snapDir)
Write-Host ("  Force       : {0}" -f $forceFlag)
Write-Host ("  DryRun      : {0}" -f $dryFlag)
Write-Host ''

if (-not $dryFlag -and [string]::IsNullOrWhiteSpace($DfExePath)) {
    Write-Host '[ERROR] -DfExePath is required. Set it in VerifyConfig.psd1 -> Df.ExePath.' -ForegroundColor Red
    exit 1
}
if (-not $dryFlag -and -not [string]::IsNullOrWhiteSpace($DfExePath) -and -not (Test-Path -LiteralPath $DfExePath)) {
    Write-Host ("[ERROR] df.exe not found: {0}" -f $DfExePath) -ForegroundColor Red; exit 1
}
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1
}

# ── Target filter ────────────────────────────────────────────
$targetSet = @{}
foreach ($rawId in @($TargetIds)) {
    if ($null -eq $rawId) { continue }
    foreach ($part in ($rawId.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $targetSet[$v] = $true }
    }
}
function Test-TargetRow($row) {
    if ($targetSet.Count -eq 0) { return $true }
    return ($targetSet.ContainsKey([string]$row.Correl_ID_S) -or
            $targetSet.ContainsKey([string]$row.Correl_ID_M) -or
            $targetSet.ContainsKey([string]$row.JOB_NAME) -or
            $targetSet.ContainsKey([string]$row.Excel_NAME))
}

$allRows  = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)

# Ensure DF_snap column exists
foreach ($r in $allRows) {
    if (-not ($r.PSObject.Properties.Name -contains 'DF_snap')) {
        $r | Add-Member -NotePropertyName 'DF_snap' -NotePropertyValue '0' -Force
    }
}

$workRows = @($allRows | Where-Object { Test-TargetRow $_ })
if ($workRows.Count -eq 0) {
    Write-Host '[INFO] No rows after filter.' -ForegroundColor Yellow
    return
}

# ── Assembly helper for screenshot ───────────────────────────
if (-not $dryFlag) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    try { Add-Type -AssemblyName Microsoft.VisualBasic } catch {}
}

function Find-File([string]$baseDir, [string]$correlIdS) {
    if (-not (Test-Path -LiteralPath $baseDir)) { return $null }
    $pattern = $FilePattern -f $correlIdS
    $hits = @(Get-ChildItem -LiteralPath $baseDir -Filter $pattern -File -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending)
    if ($hits.Count -eq 0) { return $null }
    if ($hits.Count -gt 1) {
        Write-Host ("    [WARN] {0} files match '{1}' in {2}; using newest" -f $hits.Count, $pattern, $baseDir) -ForegroundColor Yellow
    }
    return $hits[0].FullName
}

function Take-Screenshot([string]$outPath, [System.Diagnostics.Process]$proc, [string]$mode) {
    if ($mode -eq 'window') {
        # Bring df.exe to foreground first
        try {
            $hwnd = $proc.MainWindowHandle
            if ($hwnd -ne [IntPtr]::Zero) {
                [Microsoft.VisualBasic.Interaction]::AppActivate($proc.Id) | Out-Null
                Start-Sleep -Milliseconds 300
            }
        } catch {}
        # Capture foreground window via SendMessage/PrintWindow approach.
        # Fallback to full screen if window handle unavailable.
        $hwnd = [IntPtr]::Zero
        try { $hwnd = $proc.MainWindowHandle } catch {}
        if ($hwnd -eq [IntPtr]::Zero) {
            $mode = 'fullscreen'
        } else {
            $sig = @'
using System;
using System.Runtime.InteropServices;
using System.Drawing;

public class WinCapture {
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, uint nFlags);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left,Top,Right,Bottom; }

    public static Bitmap CaptureWindow(IntPtr hWnd) {
        RECT r;
        GetWindowRect(hWnd, out r);
        int w = r.Right - r.Left, h = r.Bottom - r.Top;
        if (w <= 0 || h <= 0) return null;
        var bmp = new Bitmap(w, h);
        using (var g = Graphics.FromImage(bmp))
            PrintWindow(hWnd, g.GetHdc(), 0);
        return bmp;
    }
}
'@
            try {
                if (-not ([System.Management.Automation.PSTypeName]'WinCapture').Type) {
                    Add-Type -TypeDefinition $sig -ReferencedAssemblies System.Drawing
                }
                $bmp = [WinCapture]::CaptureWindow($hwnd)
                if ($null -ne $bmp) {
                    $bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
                    $bmp.Dispose()
                    return
                }
            } catch {}
            # Fall through to fullscreen on error
            $mode = 'fullscreen'
        }
    }

    # fullscreen path
    $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bmp    = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $g      = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
        $bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $g.Dispose()
        $bmp.Dispose()
    }
}

# ── Ensure snap dir ──────────────────────────────────────────
if (-not $dryFlag -and -not (Test-Path -LiteralPath $snapDir)) {
    New-Item -ItemType Directory -Path $snapDir -Force | Out-Null
    Write-Host ("  [INFO] created {0}" -f $snapDir) -ForegroundColor DarkGray
}

# ── Main loop ────────────────────────────────────────────────
$cntDone = 0
$cntSkip = 0
$cntFail = 0

foreach ($row in $workRows) {
    $correlIdS = [string]$row.Correl_ID_S
    $correlIdM = [string]$row.Correl_ID_M
    if ([string]::IsNullOrWhiteSpace($correlIdS)) { continue }

    $curSnap = '0'
    try { $curSnap = [string]$row.DF_snap } catch {}
    if (-not $forceFlag -and $curSnap -eq '1') {
        Write-Host ("[SKIP] {0}: DF_snap already 1" -f $correlIdS) -ForegroundColor DarkGray
        $cntSkip++; continue
    }

    Write-Host ''
    Write-Host ("----- {0} -----" -f $correlIdS) -ForegroundColor Cyan

    # Locate files
    $giftFile = Find-File $GiftDataDir $correlIdS
    $gfixFile = Find-File $GfixDataDir $correlIdS

    if ($null -eq $giftFile) {
        Write-Host ("  [FAIL] GIFT file not found for {0} in {1}" -f $correlIdS, $GiftDataDir) -ForegroundColor Red
        $cntFail++; continue
    }
    if ($null -eq $gfixFile) {
        Write-Host ("  [FAIL] GFIX file not found for {0} in {1}" -f $correlIdS, $GfixDataDir) -ForegroundColor Red
        $cntFail++; continue
    }

    Write-Host ("  GIFT: {0}" -f $giftFile) -ForegroundColor DarkGray
    Write-Host ("  GFIX: {0}" -f $gfixFile) -ForegroundColor DarkGray

    $outPng = Join-Path $snapDir ("{0}.png" -f $correlIdS)

    if ($dryFlag) {
        Write-Host ("  [DRY]  would save: {0}" -f $outPng) -ForegroundColor DarkGray
        $cntSkip++; continue
    }

    $proc = $null
    try {
        $proc = Start-Process -FilePath $DfExePath `
                              -ArgumentList ("`"{0}`"" -f $giftFile), ("`"{0}`"" -f $gfixFile) `
                              -PassThru

        # Wait for window to be ready
        try { $proc.WaitForInputIdle(10000) | Out-Null } catch {}
        if ($LoadWaitSec -gt 0) { Start-Sleep -Seconds $LoadWaitSec }

        Take-Screenshot $outPng $proc $CaptureMode
        Write-Host ("  [OK]   saved: {0}" -f $outPng) -ForegroundColor Green

        # Mark done in allRows (find by Correl_ID_M for per-row update)
        foreach ($r in $allRows) {
            if ([string]$r.Correl_ID_M -eq $correlIdM) {
                $r.DF_snap = '1'
            }
        }
        $cntDone++
    } catch {
        Write-Host ("  [FAIL] {0}" -f $_.Exception.Message) -ForegroundColor Red
        $cntFail++
    } finally {
        if ($null -ne $proc) {
            try {
                if (-not $proc.HasExited) {
                    $proc.CloseMainWindow() | Out-Null
                    if (-not $proc.WaitForExit(3000)) {
                        $proc.Kill()
                    }
                }
            } catch {
                Write-Host ("  [WARN] could not close df.exe: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
            }
            $proc.Dispose()
        }
        Start-Sleep -Milliseconds 1000
    }
}

# Persist mapping if anything changed
if ($cntDone -gt 0) {
    $allRows | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force
    Write-Host ''
    Write-Host ("Mapping saved: {0}" -f $mappingPath) -ForegroundColor DarkGreen
}

Write-Host ''
Write-Host '===== DfSnap Done =====' -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
