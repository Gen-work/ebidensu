# ============================================================
#  MqSnap.ps1  (Phase: GiftMqSnap)  v3 -- MappingStore + F2 detection
#
#  For each pending Correl_ID_S in mapping_<Owner>.csv:
#    1. Tab to the inquiry button -> Enter (open inquiry form)
#    2. Tab to the Correlid input -> Paste Correl_ID_S -> Enter (search)
#    3. (detection on) poll page text, classify page kind, archive .txt
#    4. Capture Edge main window, crop border, save snap\GIFT_MQ\<id>.png
#    5. (detection on) parse + verdict (F2): ok -> field=1, ng -> field=2
#    6. Persist atomically via MappingStore; append a progress.jsonl event
#
#  SnapVerify (F2) -- spec docs/SnapVerify-Plan.md milestone M2:
#    - GIFT_MQ_snap value domain: 0/empty = pending, 1 = ok, 2 = NG.
#      '2' STILL counts as pending and is re-offered next run (plan 2.1).
#    - NG conditions (Test-MqRecord, plan 2.4): no matching row / "No Data!" /
#      RecvDate outside the time window / non-zero Rtncd|Rsncd. Newest-wins by
#      RecvDate when a correl has several rows.
#    - Time window: one batch prompt at start (Resolve-SnapRunTime); empty
#      Expected_Time cells on pending rows are filled and persisted (plan 2.2).
#    - Page-kind sentinel (plan 3.6): an off-page text (OuterFrame/Empty/
#      Unknown) stops and asks the operator (r=retry / s=skip / q=quit).
#    - SnapEnabled=$false reverts to pure screenshot (legacy behavior).
#
#  Conventions:
#    - All mapping I/O goes through MappingStore (atomic writes); progress
#      events go to status\progress.jsonl via ProgressLog.
#    - Crop-Snap.ps1 is NOT dot-sourced; Invoke-CropPng is inline.
#    - Switch params are copied to plain bools before any dot-source.
#    - Screenshot targets the Edge main hwnd, not the foreground window.
#    - Pure parse/verdict logic lives in SnapVerify.ps1 (unit-tested);
#      this script is only the COM/SendKeys wiring.
#
#  Usage:
#    .\MqSnap.ps1 -WorkDir <dir> -Owner <owner>
#    .\MqSnap.ps1 ... -Force
#    .\MqSnap.ps1 ... -Interactive
#    .\MqSnap.ps1 ... -NoResize
#    .\MqSnap.ps1 ... -SnapEnabled $false      # pure screenshot, no detection
# ============================================================
#Requires -Version 5.1

param(
    [string]$WorkDir,
    [string]$Owner               = '',

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
    [switch]$Force,

    # ---- SnapVerify (F2) detection wiring (defaults match VerifyConfig.psd1) ----
    [bool]$SnapEnabled           = $true,
    [int]$ToleranceMinutes       = 30,
    [bool]$SaveText              = $true,
    [int]$PollTimeoutSec         = 10,
    [int]$PollIntervalMs         = 500,
    [string]$TimeColumn          = 'Expected_Time',
    [string]$TimeFormat          = 'yyyy/MM/dd HH:mm:ss',
    [string]$RunTime             = '',   # '' = prompt; else 'n' / 'yyyy/MM/dd HH:mm:ss'
    [string]$RunTolerance        = '',   # non-interactive tolerance override

    # ---- M5/F5 pixel localisation (Config.SnapVerify.Localize); off by default ----
    [hashtable]$Localize         = @{}
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

# Resolve switch flags BEFORE any dot-source can clobber switch variables.
$forceFlag       = [bool]$Force.IsPresent
$interactiveFlag = [bool]$Interactive.IsPresent
$noResizeFlag    = [bool]$NoResize.IsPresent
$snapVerifyOn    = [bool]$SnapEnabled
$saveTextOn      = [bool]$SaveText

$scriptDir = $PSScriptRoot

# -- Unblock all PS1 files in this folder (avoid UNC-path security warning) --
try {
    Get-ChildItem -LiteralPath $scriptDir -Filter "*.ps1" -File -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
} catch {}

# -- Interactive fallback --
if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host "WorkDir path" }

Write-Host ""
Write-Host "===== MqSnap (Phase GiftMqSnap, v3) =====" -ForegroundColor Green
Write-Host ("  WorkDir     : {0}" -f $WorkDir)
Write-Host ("  Owner       : {0}" -f $Owner)
Write-Host ("  Window      : {0}" -f $(if ($noResizeFlag) { "no resize" } else { "${WindowWidth}x${WindowHeight}" }))
Write-Host ("  CropPx      : {0}" -f $CropPx)
Write-Host ("  Force       : {0}, Interactive : {1}" -f $forceFlag, $interactiveFlag)
Write-Host ("  Detection   : {0} (tol +-{1} min)" -f $snapVerifyOn, $ToleranceMinutes)
if (@($TargetIds).Count -gt 0) { Write-Host ("  TargetIds   : {0}" -f ((@($TargetIds)) -join ", ")) }
Write-Host ""

# -- Validate WorkDir & mapping --
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
# Dot-source shared libraries (all no-param() -> safe per CLAUDE.md)
# ============================================================
if ([string]::IsNullOrWhiteSpace($CommonScript)) { $CommonScript = Join-Path $scriptDir "Common.ps1" }
if (-not (Test-Path -LiteralPath $CommonScript)) {
    Write-Host ("[ERROR] Common.ps1 not found: {0}" -f $CommonScript) -ForegroundColor Red
    Write-Host "        Pass -CommonScript <path> to specify." -ForegroundColor Red
    exit 1
}

Write-Host ("[INFO] Loading Common.ps1 : {0}" -f $CommonScript)
$savedEAP = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'
. $CommonScript
$ErrorActionPreference = $savedEAP
. (Join-Path $scriptDir "MappingStore.ps1")
. (Join-Path $scriptDir "ProgressLog.ps1")
. (Join-Path $scriptDir "SnapVerify.ps1")
$snapLocalizeScript = Join-Path $scriptDir "SnapLocalize.ps1"
if (Test-Path -LiteralPath $snapLocalizeScript) { . $snapLocalizeScript }
$pageTextScript = Join-Path $scriptDir "Read-PageText.ps1"

if (-not (Get-Command -Name 'Wait-PagePrepared' -ErrorAction SilentlyContinue)) {
    Write-Host "[ERROR] Common.ps1 dot-source failed (Wait-PagePrepared not found)." -ForegroundColor Red
    exit 1
}
if (-not (Get-Command -Name 'Test-MqRecord' -ErrorAction SilentlyContinue)) {
    Write-Host "[ERROR] SnapVerify.ps1 dot-source failed (Test-MqRecord not found)." -ForegroundColor Red
    exit 1
}

# -- Globals used by Common helpers --
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

    if ($cropPx -le 0) { return }
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

function Click-MqPageCenter {
    $hWnd = [WinAPI]::GetForegroundWindow()
    if ($hWnd -eq [IntPtr]::Zero) { return }

    $rect = New-Object WinAPI+RECT
    [WinAPI]::GetWindowRect($hWnd, [ref]$rect) | Out-Null

    $x = [int]($rect.Left + (($rect.Right  - $rect.Left) / 2))
    $y = [int]($rect.Top  + (($rect.Bottom - $rect.Top)  / 2))

    [MouseAPI]::SetCursorPos($x, $y) | Out-Null
    Start-Sleep -Milliseconds 100
    [MouseAPI]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)  # LEFTDOWN
    Start-Sleep -Milliseconds 50
    [MouseAPI]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)  # LEFTUP
    Start-Sleep -Milliseconds 400
}

# Grab the foreground Edge page text once (Ctrl+A/Ctrl+C via Read-PageText).
# Click the MQ page center first so frameset pages focus frame_main rather than
# the left navigation frame before select-all/copy parsing.
function Get-MqPageTextOnce {
    Click-MqPageCenter
    $txt = & $pageTextScript -SelectWaitMs $ActionWaitMs -CopyWaitMs $ActionWaitMs
    if ($null -eq $txt) { return '' }
    return [string]$txt
}

# Poll the page text (A2) until it is a recognised MQ terminal page (result
# with the correl id present, or "No Data!"), an off-page OuterFrame, or the
# poll timeout elapses. Empty/Unknown keep polling (page may still be loading).
# Returns @{ Text = <string>; Kind = <pageKind> }.
function Wait-MqPageReady {
    param([string]$CorrelId)

    $text = ''
    $kind = 'Empty'
    $deadline = (Get-Date).AddSeconds([Math]::Max(1, $PollTimeoutSec))
    do {
        $text = Get-MqPageTextOnce
        $kind = Get-SnapPageKind -Phase 'Mq' -Text $text
        if ($kind -eq 'MqNoData') { break }
        if ($kind -eq 'MqResult' -and $text.Contains($CorrelId)) { break }
        if ($kind -eq 'OuterFrame') { break }
        Start-Sleep -Milliseconds ([Math]::Max(100, $PollIntervalMs))
    } while ((Get-Date) -lt $deadline)

    return @{ Text = $text; Kind = $kind }
}

function Show-MqRowHeader($item, [int]$idx, [int]$total) {
    Write-Host ""
    Write-Host ("=" * 72) -ForegroundColor White
    Write-Host ("  [MQ {0}/{1}] JOB:{2} | Correl_ID_S:{3} | TO:{4}" -f `
        $idx, $total, (Get-RowProp $item 'JOB_NAME'), (Get-RowProp $item 'Correl_ID_S'), (Get-RowProp $item 'TO_code')) -ForegroundColor White
    Write-Host ("=" * 72) -ForegroundColor White
}

# ============================================================
# Load mapping & filter pending (MappingStore = single source of truth)
# ============================================================
Write-Host "[Step 1] Loading mapping..." -ForegroundColor Cyan

# $allRows is the FULL set (never dropped on write). $pendingItems holds
# references INTO $allRows, so mutating a pending row then Export-MappingAtomic
# $allRows persists exactly that change.
$allRows = Import-Mapping $mappingPath
Ensure-MappingColumns -Rows $allRows -Extra @(
    @{ Name = $snapField;  Default = '0' },
    @{ Name = $TimeColumn; Default = ''  }
) | Out-Null
Write-Host ("  Total rows : {0}" -f $allRows.Count)

# A snap field is "done" ONLY when exactly '1'. Empty/'0'/'2'(NG) are all
# pending: NG='2' is re-offered every run (plan 2.1). We deliberately do NOT
# use Get-PendingRows here -- its Test-SnapDone treats any non-'0' value
# (including '2') as done, which would hide NG rows.
function Test-MqSnapDone([string]$Value) { return ($Value -eq '1') }

$targetList   = ConvertTo-TargetIdList $TargetIds
$pendingItems = @()
$doneCount    = 0
foreach ($r in $allRows) {
    if (-not (Test-TargetRow $r $targetList)) { continue }
    if ($forceFlag) { $pendingItems += $r; continue }
    if (Test-MqSnapDone (Get-RowProp $r $snapField)) { $doneCount++ } else { $pendingItems += $r }
}
$pendingItems = @($pendingItems)
Write-Host ("  Pending    : {0}, Already done : {1}" -f $pendingItems.Count, $doneCount)

Write-ProgressEvent -WorkDir $WorkDir -Phase 'GiftMqSnap' -Action 'start' -Status 'info' `
    -Message ("pending={0} force={1} targets=[{2}] detect={3}" -f $pendingItems.Count, $forceFlag, ($targetList -join ','), $snapVerifyOn)

if ($pendingItems.Count -eq 0) {
    Write-Host "[INFO] Nothing to do. Use -Force to redo." -ForegroundColor Yellow
    return
}

# ============================================================
# Batch run-time inquiry (plan 2.2) -- one prompt, applied to pending rows
# ============================================================
$timeMode     = 'none'
$runTolerance = $ToleranceMinutes
if ($snapVerifyOn) {
    Bring-ShellToFront
    Write-Host ""
    Write-Host "[Time window] MQ records are checked against a run time +- tolerance." -ForegroundColor Cyan
    Write-Host "  [Enter]             = use current time" -ForegroundColor Gray
    Write-Host "  yyyy/MM/dd HH:mm:ss = use that time" -ForegroundColor Gray
    Write-Host "  n                   = no time check this run" -ForegroundColor Gray
    $tIn   = if ([string]::IsNullOrWhiteSpace($RunTime)) { Read-Host "  Run time" }                       else { $RunTime }
    $tolIn = if ([string]::IsNullOrWhiteSpace($RunTime)) { Read-Host "  Tolerance minutes (Enter=default)" } else { $RunTolerance }

    $rt = Resolve-SnapRunTime -TimeInput $tIn -ToleranceInput $tolIn -DefaultTolerance $ToleranceMinutes
    if (-not $rt.Ok) {
        Write-Host ("  [WARN] {0} -- continuing with NO time check." -f $rt.Error) -ForegroundColor Yellow
        $timeMode = 'none'
    } else {
        $timeMode     = $rt.TimeMode
        $runTolerance = $rt.ToleranceMinutes
        if ($timeMode -eq 'fixed') {
            $timeStr = $rt.Time.ToString($TimeFormat)
            $filled  = Set-EmptyRunTimeCells -Rows $pendingItems -Field $TimeColumn -Value $timeStr
            if ($filled -gt 0) {
                Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
                Write-Host ("  {0} set on {1} pending row(s) = {2}; existing values kept." -f $TimeColumn, $filled, $timeStr) -ForegroundColor Green
            } else {
                Write-Host ("  All pending rows already have {0}; using each row's value (run time {1})." -f $TimeColumn, $timeStr) -ForegroundColor Green
            }
        } else {
            Write-Host "  No time check this run." -ForegroundColor Yellow
        }
    }
    Write-Host ("  Tolerance: +-{0} min" -f $runTolerance) -ForegroundColor Gray
}

# ============================================================
# Open GIFT MQ in Edge, then main loop (single page, no appl grouping)
# ============================================================
$mqUrl = "https://<mq-host>/vergift/index.html"

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

$ngList       = [System.Collections.Generic.List[string]]::new()
$totalDone    = 0
$totalNg      = 0
$totalSkipped = 0
$userQuit     = $false
$maxAttempts  = 3
$idx          = 0

foreach ($item in $pendingItems) {
    $idx++
    $correl = [string](Get-RowProp $item 'Correl_ID_S')
    $jobName = [string](Get-RowProp $item 'JOB_NAME')
    Show-MqRowHeader $item $idx $pendingItems.Count

    if ($interactiveFlag) {
        Bring-ShellToFront
        Write-Host "  Enter=run / s=skip / q=quit : " -ForegroundColor Magenta -NoNewline
        $resp = Read-Host
        if ($resp -eq 'q') { Write-Host "[ABORT] User quit." -ForegroundColor Yellow; $userQuit = $true; break }
        if ($resp -eq 's') { Write-Host "  -> skipped" -ForegroundColor DarkYellow; $totalSkipped++; continue }
        # Shell is foreground here (Bring-ShellToFront above), so Switch-ToEdge's
        # Alt+Tab correctly lands on Edge. This is the ONLY safe place for it.
        Switch-ToEdge
        Click-PageBody
    }

    $resolved = $false
    $attempt  = 0
    do {
        $attempt++

        # Per-row refocus is Reset-FocusToBody ONLY (Activate-EdgeWindow ->
        # AppActivate by title + Click-PageBody) -- no blind Alt+Tab. Do NOT add
        # a Switch-ToEdge here: its Alt+Tab toggles relative to the *current*
        # foreground, and after the previous row's screenshot Edge is already
        # foreground, so Alt+Tab would flip to the console/ISE and Click-PageBody
        # would then click that window instead of the MQ page (v3 regression).
        Reset-FocusToBody
        Send-Tab $TabsToInquiry
        Send-Enter
        Start-Sleep -Seconds 1

        # Enter Correl_ID and search
        Send-Tab $TabsToCorrelid
        Paste-Replace $correl
        Send-Enter
        Start-Sleep -Seconds $Global:Timing.ResultWaitSec

        # Capture page text (A2 poll) when detection is on
        $pageText = ''
        $pageKind = 'Unknown'
        if ($snapVerifyOn) {
            [void](Activate-EdgeMainWindow)
            $ready    = Wait-MqPageReady -CorrelId $correl
            $pageText = [string]$ready.Text
            $pageKind = [string]$ready.Kind
            Write-Host ("    pageKind: {0}" -f $pageKind) -ForegroundColor DarkGray
        }

        # Screenshot + crop (always)
        $outPath = Join-Path $snapDir ("{0}.png" -f $correl)
        # Ctrl+A poll leaves text selected; one click deselects before capture.
        if ($snapVerifyOn) { Click-PageBody }
        Save-EdgeMainScreenshot $outPath
        try {
            Invoke-CropPng -path $outPath -cropPx $CropPx
        } catch {
            Write-Host ("    [WARN] Crop failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        }

        # -- Detection OFF: legacy behavior (screenshot -> field=1) --
        if (-not $snapVerifyOn) {
            $item.$snapField = '1'
            Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'GiftMqSnap' -CorrelIdS $correl -JobName $jobName `
                -Action 'snap' -Status 'ok' -Message ("snap\GIFT_MQ\{0}.png (detection off)" -f $correl)
            Write-Host ("    -> {0} = 1, saved {1}" -f $snapField, (Split-Path -Leaf $outPath)) -ForegroundColor Green
            $totalDone++
            $resolved = $true
            break
        }

        # Archive page text (A1)
        if ($saveTextOn -and -not [string]::IsNullOrWhiteSpace($pageText)) {
            try {
                $txtPath = Join-Path $snapDir ("{0}.txt" -f $correl)
                $enc = New-Object System.Text.UTF8Encoding($false)
                [System.IO.File]::WriteAllText($txtPath, $pageText, $enc)
            } catch {
                Write-Host ("    [WARN] save text failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
            }
        }

        # Sentinel (A3): off-page kinds -> stop and ask the operator
        if ($pageKind -eq 'OuterFrame' -or $pageKind -eq 'Empty' -or $pageKind -eq 'Unknown') {
            $preview = if ($pageText.Length -gt 200) { $pageText.Substring(0, 200) } else { $pageText }
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'GiftMqSnap' -CorrelIdS $correl -JobName $jobName `
                -Action 'verify' -Status 'warn' -Message ("pageKind={0}: {1}" -f $pageKind, $preview)
            Bring-ShellToFront
            Write-Host ("  [SENTINEL] Unexpected page ({0}) for {1}." -f $pageKind, $correl) -ForegroundColor Yellow
            Write-Host "    Fix Edge focus/frame, then r=retry this row / s=skip / q=quit : " -ForegroundColor Magenta -NoNewline
            $ans = Read-Host
            if ($ans -eq 'q') { Write-Host "[ABORT] User quit." -ForegroundColor Yellow; $userQuit = $true; $resolved = $true; break }
            if ($ans -eq 's') { Write-Host "  -> skipped" -ForegroundColor DarkYellow; $totalSkipped++; $resolved = $true; break }
            if ($attempt -ge $maxAttempts) {
                Write-Host ("  -> still unexpected after {0} attempts; skipping." -f $maxAttempts) -ForegroundColor Yellow
                $totalSkipped++
                $resolved = $true
                break
            }
            continue   # retry this row
        }

        # -- F2 verdict (plan 2.4): ok -> field=1, ng -> field=2 --
        $rowExpected = $null
        if ($timeMode -ne 'none') {
            $rowExpected = ConvertTo-ExpectedDateTime -Value (Get-RowProp $item $TimeColumn) -Format $TimeFormat
        }

        $verdict = $null
        $parsed  = @{ Rows = @() }
        try {
            $parsed  = ConvertFrom-MqPageText $pageText
            $verdict = Test-MqRecord -Parsed $parsed -CorrelId $correl -Expected $rowExpected `
                        -ToleranceMin $runTolerance -IsNoData ($pageKind -eq 'MqNoData')
        } catch {
            # Detection bugs must never block the screenshot workflow: keep the
            # shot, mark done, and record a warning for review.
            Write-Host ("    [WARN] detection failed: {0} (marking snap done)" -f $_.Exception.Message) -ForegroundColor Yellow
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'GiftMqSnap' -CorrelIdS $correl -JobName $jobName `
                -Action 'verify' -Status 'warn' -Message ("detect error: {0}" -f $_.Exception.Message)
            $verdict = @{ Verdict = 'ok'; Reason = 'detection error -> screenshot kept' }
        }

        # M5/F5: best-effort <correl>.loc.json sidecar for the matched MQ row.
        if ([bool]$Localize['Enabled'] -and (Get-Command -Name 'Write-SnapLocalize' -ErrorAction SilentlyContinue)) {
            $locPath = Write-SnapLocalize -Page 'Mq' -Localize $Localize -SnapDir $snapDir `
                -Correl $correl -PngPath $outPath -Rows $parsed.Rows -Expected $rowExpected `
                -ToleranceMin $runTolerance `
                -CropLeft ([int]$Localize['CropLeft']) -CropTop ([int]$Localize['CropTop'])
            if ($locPath) { Write-Host ("    loc: snap\GIFT_MQ\{0}.loc.json" -f $correl) -ForegroundColor DarkGray }
        }

        if ($verdict.Verdict -eq 'ok') {
            $item.$snapField = '1'
            Write-Host ("    -> OK: {0}" -f $verdict.Reason) -ForegroundColor Green
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'GiftMqSnap' -CorrelIdS $correl -JobName $jobName `
                -Action 'verify' -Status 'ok' -Message $verdict.Reason
            $totalDone++
        } else {
            $item.$snapField = '2'
            Write-Host ("    -> NG: {0}" -f $verdict.Reason) -ForegroundColor Red
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'GiftMqSnap' -CorrelIdS $correl -JobName $jobName `
                -Action 'verify' -Status 'ng' -Message $verdict.Reason
            $ngList.Add(("{0} : {1}" -f $correl, $verdict.Reason))
            $totalNg++
        }

        Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
        Write-Host ("    saved {0}" -f (Split-Path -Leaf $outPath)) -ForegroundColor DarkGray
        $resolved = $true
    } until ($resolved)

    if ($userQuit) { break }
}

Bring-ShellToFront
Write-Host ""
Write-Host "===== MqSnap Done =====" -ForegroundColor Green
Write-Host ("  Snapped OK : {0}" -f $totalDone)
if ($totalNg -gt 0) {
    Write-Host ("  NG         : {0}" -f $totalNg) -ForegroundColor Red
}
if ($totalSkipped -gt 0) {
    Write-Host ("  Skipped    : {0}" -f $totalSkipped) -ForegroundColor DarkGray
}

if ($ngList.Count -gt 0) {
    Write-Host ""
    Write-Host ("===== MQ NG summary ({0}) =====" -f $ngList.Count) -ForegroundColor Red
    foreach ($n in $ngList) { Write-Host ("  [NG] {0}" -f $n) -ForegroundColor Red }
    Write-Host "  (NG rows stay pending; re-run GiftMqSnap to retry.)" -ForegroundColor DarkYellow
}
