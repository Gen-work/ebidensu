# ============================================================
#  MqSnap.ps1  (Phase 4, v2 — self-contained capture/crop)
#
#  For each pending Correl_ID_S in mapping_<Owner>.csv:
#    1. Tab to 照会 button -> Enter (open inquiry form)
#    2. Tab to 相関ID input -> Paste Correl_ID_S -> Enter (search)
#    3. Capture Edge main window, crop border, save snap/GIFT_MQ/<Correl_ID_S>.png
#    4. Update mapping.GIFT_MQ_snap = 1
#    5. Reset to inquiry entry for next row
#
#  Important:
#    - Crop-Snap.ps1 is NOT dot-sourced here. Invoke-CropPng is inline.
#    - Switch params are copied to plain bools before any dot-source.
#    - Screenshot targets Edge main hwnd, not arbitrary foreground window.
#
#  Usage:
#    .\MqSnap.ps1
#    .\MqSnap.ps1 -Force
#    .\MqSnap.ps1 -Interactive
#    .\MqSnap.ps1 -NoResize
# ============================================================

param(
    [string]$WorkDir,
    [string]$Owner               = ([char]0x53B3),  # 厳

    [int]$TabsToInquiry          = 1,
    [int]$TabsToCorrelid         = 4,
    [int]$ActionWaitMs           = 500,
    [int]$ResultWaitSec          = 2,
    [int]$CropPx                 = 6,

    [int]$WindowWidth            = 1050,
    [int]$WindowHeight           = 761,
    [switch]$NoResize,

    [string[]]$TargetIds         = @(),
    [string]$CommonScript        = "",
    [switch]$Interactive,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# Resolve switch flags BEFORE any dot-source can clobber switch variables.
$forceFlag       = [bool]$Force.IsPresent
$interactiveFlag = [bool]$Interactive.IsPresent
$noResizeFlag    = [bool]$NoResize.IsPresent

# Optional narrow run. Accepts: -TargetIds JIGPL48S or -TargetIds JIGPL48S,JIDSL48S
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
            $targetSet.ContainsKey([string]$row.JOB_NAME))
}


# ── Unblock all PS1 files in this folder (avoid UNC-path security warning) ──
try {
    Get-ChildItem -LiteralPath $PSScriptRoot -Filter "*.ps1" -File -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
} catch {}

# ── Interactive fallback ──
if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host "WorkDir path" }

Write-Host ""
Write-Host "===== MqSnap (Phase 4 v2) =====" -ForegroundColor Green
Write-Host ("  WorkDir     : {0}" -f $WorkDir)
Write-Host ("  Owner       : {0}" -f $Owner)
Write-Host ("  Window      : {0}" -f $(if ($noResizeFlag) { "no resize" } else { "${WindowWidth}x${WindowHeight}" }))
Write-Host ("  CropPx      : {0}" -f $CropPx)
Write-Host ("  Force       : {0}, Interactive : {1}" -f $forceFlag, $interactiveFlag)
if ($targetSet.Count -gt 0) { Write-Host ("  TargetIds   : {0}" -f (($targetSet.Keys | Sort-Object) -join ", ")) }
Write-Host ""

# ── Validate WorkDir & mapping ──
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found." -ForegroundColor Red
    exit 1
}
$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host ("[ERROR] mapping file not found: {0}" -f $mappingPath) -ForegroundColor Red
    exit 1
}

$snapField = "GIFT_MQ_snap"
$snapDir   = Join-Path $WorkDir "snap\GIFT_MQ"
if (-not (Test-Path -LiteralPath $snapDir)) {
    New-Item -ItemType Directory -Path $snapDir -Force | Out-Null
}
Write-Host ("[INFO] OutDir : {0}" -f $snapDir)
Write-Host ("[INFO] Field  : mapping.{0}" -f $snapField)
Write-Host ""

# ============================================================
# Locate & dot-source Common.ps1 (only shared primitive dependency)
# ============================================================
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($CommonScript)) { $candidates += $CommonScript }
$candidates += @(
    (Join-Path $PSScriptRoot "Common.ps1"),
    (Join-Path $PSScriptRoot "Common(1).ps1"),
    (Join-Path $PSScriptRoot "..\VerifyTool\Common.ps1"),
    (Join-Path (Split-Path $PSScriptRoot -Parent) "VerifyTool\Common.ps1")
)

$commonPath = $null
foreach ($c in $candidates) {
    if (Test-Path -LiteralPath $c) { $commonPath = (Resolve-Path -LiteralPath $c).Path; break }
}
if (-not $commonPath) {
    Write-Host "[ERROR] Common.ps1 not found. Tried:" -ForegroundColor Red
    foreach ($c in $candidates) { Write-Host ("        - {0}" -f $c) -ForegroundColor Red }
    Write-Host "        Pass -CommonScript <path> to specify." -ForegroundColor Red
    exit 1
}

Write-Host ("[INFO] Loading Common.ps1 : {0}" -f $commonPath)
$savedEAP = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'
. $commonPath
$ErrorActionPreference = $savedEAP
if (-not (Get-Command -Name 'Wait-PagePrepared' -ErrorAction SilentlyContinue)) {
    Write-Host "[ERROR] Common.ps1 dot-source failed (function not found)." -ForegroundColor Red
    exit 1
}

# ── Globals used by Common helpers ──
$Global:Timing = @{
    ActionWaitMs   = $ActionWaitMs
    ResultWaitSec  = $ResultWaitSec
}
if (-not $Global:Shell) { $Global:Shell = New-Object -ComObject WScript.Shell }

Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

# ============================================================
# Inline helpers (do not dot-source Crop-Snap.ps1)
# ============================================================
function Invoke-CropPng {
    param(
        [Parameter(Mandatory=$true)][string]$path,
        [int]$cropPx = 6
    )

    if (-not (Test-Path -LiteralPath $path)) { throw "File not found: $path" }

    $bytes   = [System.IO.File]::ReadAllBytes($path)
    $ms      = New-Object System.IO.MemoryStream(, $bytes)
    $tmpPath = "$path.crop.tmp"

    try {
        $orig = [System.Drawing.Image]::FromStream($ms)
        try {
            $newW = $orig.Width  - 2 * $cropPx
            $newH = $orig.Height - 2 * $cropPx
            if ($newW -le 0 -or $newH -le 0) {
                throw ("Image too small ({0}x{1}) to crop {2} px per side" -f $orig.Width, $orig.Height, $cropPx)
            }

            $bmp = New-Object System.Drawing.Bitmap($newW, $newH)
            try {
                $gfx = [System.Drawing.Graphics]::FromImage($bmp)
                try {
                    $srcRect = New-Object System.Drawing.Rectangle($cropPx, $cropPx, $newW, $newH)
                    $dstRect = New-Object System.Drawing.Rectangle(0, 0, $newW, $newH)
                    $gfx.DrawImage($orig, $dstRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
                } finally {
                    $gfx.Dispose()
                }
                $bmp.Save($tmpPath, [System.Drawing.Imaging.ImageFormat]::Png)
            } finally {
                $bmp.Dispose()
            }
        } finally {
            $orig.Dispose()
        }
    } finally {
        $ms.Dispose()
    }

    Move-Item -LiteralPath $tmpPath -Destination $path -Force
}

function Bring-ShellToFront {
    try {
        $hwnd = (Get-Process -Id $PID).MainWindowHandle
        if ($hwnd -ne [IntPtr]::Zero) {
            [WinAPI]::ShowWindowAsync($hwnd, 9) | Out-Null
            [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
            Start-Sleep -Milliseconds 200
        }
    } catch {}
}

function Activate-EdgeMainWindow {
    $edgeProc = Get-Process msedge -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } |
                Select-Object -First 1
    if (-not $edgeProc) { return [IntPtr]::Zero }

    $hwnd = $edgeProc.MainWindowHandle
    [WinAPI]::ShowWindowAsync($hwnd, 9) | Out-Null
    [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 300
    return $hwnd
}

function Move-EdgeAwayFromBorder {
    $hWnd = Activate-EdgeMainWindow
    if ($hWnd -eq [IntPtr]::Zero) {
        Write-Host "  [WARN] Edge main window not found; cannot move/resize." -ForegroundColor Yellow
        return
    }

    try {
        if ($noResizeFlag) {
            $rect = New-Object WinAPI+RECT
            [WinAPI]::GetWindowRect($hWnd, [ref]$rect) | Out-Null
            $curW = $rect.Right  - $rect.Left
            $curH = $rect.Bottom - $rect.Top
            [WinAPI]::MoveWindow($hWnd, 40, 40, $curW, $curH, $true) | Out-Null
        } else {
            [WinAPI]::MoveWindow($hWnd, 40, 40, $WindowWidth, $WindowHeight, $true) | Out-Null
        }
        Start-Sleep -Milliseconds 400
    } catch {
        Write-Host ("  [WARN] MoveWindow failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

function Save-EdgeMainScreenshot {
    param([Parameter(Mandatory=$true)][string]$outPath)

    $mainHwnd = Activate-EdgeMainWindow
    if ($mainHwnd -eq [IntPtr]::Zero) {
        Write-Host "    [WARN] Edge main window not found, fallback to foreground." -ForegroundColor Yellow
        Take-ForegroundScreenshot $outPath
        return
    }

    $rectCap = New-Object WinAPI+RECT
    [WinAPI]::GetWindowRect($mainHwnd, [ref]$rectCap) | Out-Null
    $wCap = $rectCap.Right  - $rectCap.Left
    $hCap = $rectCap.Bottom - $rectCap.Top
    Write-Host ("    capture: hwnd={0} size={1}x{2}" -f $mainHwnd, $wCap, $hCap) -ForegroundColor DarkGray
    Take-WindowScreenshot $mainHwnd $outPath
}

function Show-MqRowHeader($item, [int]$idx, [int]$total) {
    Write-Host ""
    Write-Host ("=" * 72) -ForegroundColor White
    Write-Host ("  [MQ {0}/{1}] JOB:{2} | Correl_ID_S:{3} | TO:{4}" -f `
        $idx, $total, $item.JOB_NAME, $item.Correl_ID_S, $item.TO_code) -ForegroundColor White
    Write-Host ("=" * 72) -ForegroundColor White
}

function Update-MappingSnapField([string]$correlIdM, [string]$field) {
    foreach ($r in $mapping) {
        if ($r.Correl_ID_M -eq $correlIdM) { $r.$field = "1"; break }
    }
    $mapping | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force
}

# ============================================================
# Load mapping & filter pending
# ============================================================
Write-Host ""
Write-Host "[Step 1] Loading mapping..." -ForegroundColor Cyan
$mapping = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)
Write-Host ("  Total rows : {0}" -f $mapping.Count)

$pendingItems = @()
$doneCount = 0
foreach ($r in $mapping) {
    if (-not (Test-TargetRow $r)) { continue }
    $cur = $r.$snapField
    if ($cur -eq "1" -and -not $forceFlag) {
        $doneCount++
        continue
    }
    $pendingItems += $r
}
Write-Host ("  Pending    : {0}, Already done : {1}" -f $pendingItems.Count, $doneCount)
if ($pendingItems.Count -eq 0) {
    Write-Host "[INFO] Nothing to do. Use -Force to redo." -ForegroundColor Yellow
    return
}

# ============================================================
# Main loop (single page, no appl grouping)
# ============================================================
$mqUrl = "https://bizver.hm.jp.honda.com/vergift/index.html"

Bring-ShellToFront
Write-Host ""
Write-Host ("####################################################################") -ForegroundColor Cyan
Write-Host ("##  GIFT MQ - {0} item(s)" -f $pendingItems.Count) -ForegroundColor Cyan
Write-Host ("####################################################################") -ForegroundColor Cyan
Write-Host ""
Write-Host "  Open GIFT MQ in Edge:" -ForegroundColor Yellow
Write-Host ("    URL : {0}" -f $mqUrl) -ForegroundColor Yellow
Wait-PagePrepared "Press Enter when GIFT MQ page is ready."

Switch-ToEdge
Move-EdgeAwayFromBorder
Click-PageBody

$totalDone    = 0
$totalSkipped = 0
$idx = 0

foreach ($item in $pendingItems) {
    $idx++
    Show-MqRowHeader $item $idx $pendingItems.Count

    if ($interactiveFlag) {
        Bring-ShellToFront
        Write-Host "  Enter=run / s=skip / q=quit : " -ForegroundColor Magenta -NoNewline
        $resp = Read-Host
        if ($resp -eq 'q') { Write-Host "[ABORT] User quit." -ForegroundColor Yellow; return }
        if ($resp -eq 's') { Write-Host "  -> skipped" -ForegroundColor DarkYellow; $totalSkipped++; continue }
        Switch-ToEdge
        Click-PageBody
    }

    # Step 1-4: navigate to inquiry form
    Reset-FocusToBody
    Send-Tab $TabsToInquiry
    Send-Enter
    Start-Sleep -Seconds 1

    # Step 5-7: enter Correl_ID and search
    Send-Tab $TabsToCorrelid
    Paste-Replace $item.Correl_ID_S
    Send-Enter
    Start-Sleep -Seconds $Global:Timing.ResultWaitSec

    # Step 8-10: screenshot + crop
    $outPath = Join-Path $snapDir ("{0}.png" -f $item.Correl_ID_S)
    Save-EdgeMainScreenshot $outPath

    try {
        Invoke-CropPng -path $outPath -cropPx $CropPx
    } catch {
        Write-Host ("    [WARN] Crop failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }

    # Step 11: mark done
    Update-MappingSnapField $item.Correl_ID_M $snapField
    Write-Host ("    -> {0} = 1, saved {1}" -f $snapField, (Split-Path -Leaf $outPath)) -ForegroundColor Green
    $totalDone++

    # Step 12: reset for next iteration
    Reset-FocusToBody
    Send-Tab $TabsToInquiry
    Send-Enter
    Start-Sleep -Milliseconds 600
}

Bring-ShellToFront
Write-Host ""
Write-Host "===== MqSnap Done =====" -ForegroundColor Green
Write-Host ("  Snapped : {0}" -f $totalDone)
if ($totalSkipped -gt 0) {
    Write-Host ("  Skipped : {0}" -f $totalSkipped) -ForegroundColor DarkGray
}
