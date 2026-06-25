# ============================================================
#  HmSnap.ps1  (Phase: GiftHmSnap / GfixHmSnap)  v3 -- MappingStore + F1 detection
#
#  For each pending Correl_ID_S in mapping_<Owner>.csv (grouped by TO_code,
#  one HM page per appl):
#    1. Tab to the Correlid input -> Paste Correl_ID_S -> Shift+Tab to search
#       button -> Enter (search)
#    2. (detection on) poll page text, classify page kind, archive .txt
#    3. Capture Edge main window, crop border, save snap\<Stage>_HM\<id>.png
#    4. (detection on) parse + verdict (F1): ok -> field=1, ng -> field=2,
#       ask -> operator decides (o=1 / n=2 / s=skip)
#    5. Persist atomically via MappingStore; append a progress.jsonl event
#
#  SnapVerify (F1) -- spec docs/SnapVerify-Plan.md milestone M4:
#    - <Stage>_HM_snap value domain: 0/empty = pending, 1 = ok, 2 = NG.
#      '2' STILL counts as pending and is re-offered next run (plan 2.1).
#    - Verdict (Test-HmAbend, plan 2.3): window-inside NG, window-outside only
#      warns; newest-wins by StartTime when a correl has several rows (retried
#      run that ended normally after an earlier abend -> ok). 0 rows / 0
#      window rows / no-time-mode abend -> 'ask' the operator (plan 4.F1).
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
#    .\HmSnap.ps1 -Stage GIFT -WorkDir <dir> -Owner <owner>
#    .\HmSnap.ps1 -Stage GFIX ...
#    .\HmSnap.ps1 -Stage GIFT ... -Force
#    .\HmSnap.ps1 -Stage GIFT ... -Interactive
#    .\HmSnap.ps1 -Stage GIFT ... -NoResize
#    .\HmSnap.ps1 -Stage GIFT ... -SnapEnabled $false   # pure screenshot
# ============================================================
#Requires -Version 5.1

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('GIFT','GFIX')]
    [string]$Stage,

    [string]$WorkDir,
    [string]$Owner               = '',

    [int]$TabsToCorrelid         = 1,
    [int]$TabsBackFromSearch     = 1,
    [int]$TabsBackToInput        = 4,
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

    # ---- SnapVerify (F1) detection wiring (defaults match VerifyConfig.psd1) ----
    [bool]$SnapEnabled           = $true,
    [bool]$TimeCheck             = $false,  # $false = existence/abend checks only (no run-time window)
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

$scriptDir   = $PSScriptRoot
$phaseName   = if ($Stage -eq 'GIFT') { 'GiftHmSnap' } else { 'GfixHmSnap' }

# -- Unblock all PS1 files in this folder (avoid UNC-path security warning) --
try {
    Get-ChildItem -LiteralPath $scriptDir -Filter "*.ps1" -File -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
} catch {}

# -- Interactive fallback --
if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host "WorkDir path" }

Write-Host ""
Write-Host "===== HmSnap (Phase $phaseName, v3) =====" -ForegroundColor Green
Write-Host ("  Stage       : {0}" -f $Stage)
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

$snapField  = "${Stage}_HM_snap"      # GIFT_HM_snap or GFIX_HM_snap
$snapFolder = "${Stage}_HM"           # GIFT_HM or GFIX_HM
$snapDir    = Join-Path $WorkDir ("snap\{0}" -f $snapFolder)
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
if (-not (Get-Command -Name 'Test-HmAbend' -ErrorAction SilentlyContinue)) {
    Write-Host "[ERROR] SnapVerify.ps1 dot-source failed (Test-HmAbend not found)." -ForegroundColor Red
    exit 1
}

# -- Globals used by Common helpers --
$Global:Timing = @{
    ActionWaitMs   = $ActionWaitMs
    ResultWaitSec  = $ResultWaitSec
}
$Global:TabCounts = @{
    HM_ToCorrelid        = $TabsToCorrelid
    HM_ShiftTabToSearch  = $TabsBackFromSearch
    HM_BackToInput       = $TabsBackToInput
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

# Grab the foreground Edge page text once (Ctrl+A/Ctrl+C via Read-PageText).
# Click-PageBody first so the select-all lands on the page, not an input field.
function Get-HmPageTextOnce {
    Click-PageBody
    $txt = & $pageTextScript -SelectWaitMs $ActionWaitMs -CopyWaitMs $ActionWaitMs
    if ($null -eq $txt) { return '' }
    return [string]$txt
}

# Poll the page text (A2) until it is a recognised HM batch-status page (the
# table/title is present and the correl id shows), an off-page OuterFrame, or
# the poll timeout elapses. Empty/Unknown keep polling (page may be loading).
# Returns @{ Text = <string>; Kind = <pageKind> }.
function Wait-HmPageReady {
    param([string]$CorrelId)

    $text = ''
    $kind = 'Empty'
    $deadline = (Get-Date).AddSeconds([Math]::Max(1, $PollTimeoutSec))
    do {
        $text = Get-HmPageTextOnce
        $kind = Get-SnapPageKind -Phase 'Hm' -Text $text
        if ($kind -eq 'HmResult' -and $text.Contains($CorrelId)) { break }
        if ($kind -eq 'OuterFrame') { break }
        Start-Sleep -Milliseconds ([Math]::Max(100, $PollIntervalMs))
    } while ((Get-Date) -lt $deadline)

    return @{ Text = $text; Kind = $kind }
}

function Show-HmRowHeader($item, [string]$stepName, [int]$idx, [int]$total) {
    Write-Host ""
    Write-Host ("=" * 72) -ForegroundColor White
    Write-Host ("  [{0}/{1}] {2} | JOB:{3} | Correl_ID_S:{4} | TO:{5}" -f `
        $idx, $total, $stepName, (Get-RowProp $item 'JOB_NAME'), (Get-RowProp $item 'Correl_ID_S'), (Get-RowProp $item 'TO_code')) -ForegroundColor White
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
function Test-HmSnapDone([string]$Value) { return ($Value -eq '1') }

$targetList   = ConvertTo-TargetIdList $TargetIds
$pendingItems = @()
$doneCount    = 0
foreach ($r in $allRows) {
    if (-not (Test-TargetRow $r $targetList)) { continue }
    if ($forceFlag) { $pendingItems += $r; continue }
    if (Test-HmSnapDone (Get-RowProp $r $snapField)) { $doneCount++ } else { $pendingItems += $r }
}
$pendingItems = @($pendingItems)
Write-Host ("  Pending    : {0}, Already done : {1}" -f $pendingItems.Count, $doneCount)

Write-ProgressEvent -WorkDir $WorkDir -Phase $phaseName -Action 'start' -Status 'info' `
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
if ($snapVerifyOn -and $TimeCheck) {
    Bring-ShellToFront
    Write-Host ""
    Write-Host "[Time window] HM rows are checked against a run time +- tolerance." -ForegroundColor Cyan
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
# Group pending by TO_code (one HM page per appl) and run the main loop
# ============================================================
$grouped = $pendingItems | Group-Object TO_code | Sort-Object Name
Write-Host ""
Write-Host "  Group plan :"
foreach ($g in $grouped) {
    Write-Host ("    {0,-6} -> {1} item(s)" -f $g.Name, $g.Count) -ForegroundColor DarkGray
}

# HM page title hint (built from [char] so this source stays ASCII).
$hmPageLabel = [char]0x30D0 + [char]0x30C3 + [char]0x30C1 + [char]0x51E6 + [char]0x7406 + `
               [char]0x72B6 + [char]0x6CC1 + [char]0x4E00 + [char]0x89A7   # batch shori jokyo ichiran

$ngList       = [System.Collections.Generic.List[string]]::new()
$totalDone    = 0
$totalNg      = 0
$totalSkipped = 0
$userQuit     = $false
$maxAttempts  = 3

foreach ($g in $grouped) {
    $appl       = $g.Name
    $items      = @($g.Group)
    $applLower  = $appl.ToLower()
    $hmUrlGuess = "https://<hm-host>/{0}9a21/x/x0/{1}X0011A.do" -f $applLower, $appl

    Bring-ShellToFront
    Write-Host ""
    Write-Host ("####################################################################") -ForegroundColor Cyan
    Write-Host ("##  {0} HM - appl: {1}  ({2} item(s))" -f $Stage, $appl, $items.Count) -ForegroundColor Cyan
    Write-Host ("####################################################################") -ForegroundColor Cyan
    Write-Host ""
    Write-Host ("  Open {0} HM in Edge:" -f $appl) -ForegroundColor Yellow
    Write-Host ("    URL hint : {0}" -f $hmUrlGuess) -ForegroundColor Yellow
    Write-Host ("    (login if needed, navigate to {0} search page)" -f $hmPageLabel) -ForegroundColor Yellow
    Wait-PagePrepared ("Press Enter when {0} HM search page is ready." -f $appl)

    Switch-ToEdge
    Move-EdgeAwayFromBorder
    Click-PageBody

    $idx = 0
    foreach ($item in $items) {
        $idx++
        $correl  = [string](Get-RowProp $item 'Correl_ID_S')
        $jobName = [string](Get-RowProp $item 'JOB_NAME')
        Show-HmRowHeader $item ("{0}-HM {1}" -f $Stage, $appl) $idx $items.Count

        if ($interactiveFlag) {
            Bring-ShellToFront
            Write-Host "  Enter=run / s=skip / q=quit : " -ForegroundColor Magenta -NoNewline
            $resp = Read-Host
            if ($resp -eq 'q') { Write-Host "[ABORT] User quit." -ForegroundColor Yellow; $userQuit = $true; break }
            if ($resp -eq 's') { Write-Host "  -> skipped" -ForegroundColor DarkYellow; $totalSkipped++; continue }
            # Shell is foreground here (Bring-ShellToFront above), so Switch-ToEdge's
            # Alt+Tab correctly lands on Edge.
            Switch-ToEdge
            Click-PageBody
        }

        $resolved = $false
        $attempt  = 0
        do {
            $attempt++

            # Per-row refocus is Reset-FocusToBody (Activate-EdgeWindow by title +
            # Click-PageBody) then the HM search key sequence: Tab to the correlid
            # input, paste, Shift+Tab back to the search button, Enter. Do NOT add
            # a Switch-ToEdge here (its Alt+Tab is relative to the foreground and
            # would flip to the console after the previous row's screenshot).
            Reset-FocusToBody
            Send-Tab $Global:TabCounts.HM_ToCorrelid
            Paste-Replace $correl
            Send-ShiftTab $Global:TabCounts.HM_ShiftTabToSearch
            Send-Enter
            Start-Sleep -Seconds $Global:Timing.ResultWaitSec

            # Capture page text (A2 poll) when detection is on
            $pageText = ''
            $pageKind = 'Unknown'
            if ($snapVerifyOn) {
                [void](Activate-EdgeMainWindow)
                $ready    = Wait-HmPageReady -CorrelId $correl
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
                Write-ProgressEvent -WorkDir $WorkDir -Phase $phaseName -CorrelIdS $correl -JobName $jobName `
                    -Action 'snap' -Status 'ok' -Message ("snap\{0}\{1}.png (detection off)" -f $snapFolder, $correl)
                Write-Host ("    -> {0} = 1, saved {1}" -f $snapField, (Split-Path -Leaf $outPath)) -ForegroundColor Green
                $totalDone++
                Send-Tab $Global:TabCounts.HM_BackToInput
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
                Write-ProgressEvent -WorkDir $WorkDir -Phase $phaseName -CorrelIdS $correl -JobName $jobName `
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
                # The sentinel prompt ran with the shell foreground; hand focus
                # back to Edge before retrying so Reset-FocusToBody/keystrokes hit
                # the page, not the console.
                Switch-ToEdge
                continue   # retry this row
            }

            # -- F1 verdict (plan 2.3): ok -> field=1, ng -> field=2, ask -> operator --
            $rowExpected = $null
            if ($timeMode -ne 'none') {
                $rowExpected = ConvertTo-ExpectedDateTime -Value (Get-RowProp $item $TimeColumn) -Format $TimeFormat
            }

            $verdict = $null
            $hmRows  = @()
            try {
                $hmRows  = @(ConvertFrom-HmPageText $pageText)
                $verdict = Test-HmAbend -Rows $hmRows -CorrelId $correl -Expected $rowExpected -ToleranceMin $runTolerance
            } catch {
                # Detection bugs must never block the screenshot workflow: keep the
                # shot, mark done, and record a warning for review.
                Write-Host ("    [WARN] detection failed: {0} (marking snap done)" -f $_.Exception.Message) -ForegroundColor Yellow
                Write-ProgressEvent -WorkDir $WorkDir -Phase $phaseName -CorrelIdS $correl -JobName $jobName `
                    -Action 'verify' -Status 'warn' -Message ("detect error: {0}" -f $_.Exception.Message)
                $verdict = @{ Verdict = 'ok'; Reason = 'detection error -> screenshot kept'; Warnings = @() }
            }

            # M5/F5: best-effort <correl>.loc.json sidecar for the matched HM row.
            if ([bool]$Localize['Enabled'] -and (Get-Command -Name 'Write-SnapLocalize' -ErrorAction SilentlyContinue)) {
                $locPath = Write-SnapLocalize -Page 'Hm' -Localize $Localize -SnapDir $snapDir `
                    -Correl $correl -PngPath $outPath -Rows $hmRows -Expected $rowExpected `
                    -ToleranceMin $runTolerance `
                    -CropLeft ([int]$Localize['CropLeft']) -CropTop ([int]$Localize['CropTop'])
                if ($locPath) { Write-Host ("    loc: snap\{0}\{1}.loc.json" -f $snapFolder, $correl) -ForegroundColor DarkGray }
            }

            # Surface any warnings (historic / retried abends) without changing the verdict
            $warnings = @()
            if ($verdict.ContainsKey('Warnings')) { $warnings = @($verdict.Warnings) }
            foreach ($w in $warnings) { Write-Host ("    [warn] {0}" -f $w) -ForegroundColor Yellow }
            if ($warnings.Count -gt 0) {
                Write-ProgressEvent -WorkDir $WorkDir -Phase $phaseName -CorrelIdS $correl -JobName $jobName `
                    -Action 'verify' -Status 'warn' -Message ($warnings -join ' | ')
            }

            # Decide what to persist. $applyValue = $null means "leave pending".
            $applyValue = $null
            $countAs    = ''

            if ($verdict.Verdict -eq 'ok') {
                $applyValue = '1'; $countAs = 'ok'
                Write-Host ("    -> OK: {0}" -f $verdict.Reason) -ForegroundColor Green
                Write-ProgressEvent -WorkDir $WorkDir -Phase $phaseName -CorrelIdS $correl -JobName $jobName `
                    -Action 'verify' -Status 'ok' -Message $verdict.Reason
            }
            elseif ($verdict.Verdict -eq 'ng') {
                $applyValue = '2'; $countAs = 'ng'
                Write-Host ("    -> NG: {0}" -f $verdict.Reason) -ForegroundColor Red
                Write-ProgressEvent -WorkDir $WorkDir -Phase $phaseName -CorrelIdS $correl -JobName $jobName `
                    -Action 'verify' -Status 'ng' -Message $verdict.Reason
                $ngList.Add(("{0} : {1}" -f $correl, $verdict.Reason))
            }
            else {
                # 'ask' (or any unexpected verdict) -> the operator decides.
                Bring-ShellToFront
                Write-Host ("  [ASK] {0}: {1}" -f $correl, $verdict.Reason) -ForegroundColor Yellow
                Write-Host "    o=OK(1) / n=NG(2) / s=skip(pending) / q=quit : " -ForegroundColor Magenta -NoNewline
                $ans = Read-Host
                if ($ans -eq 'q') {
                    Write-Host "[ABORT] User quit." -ForegroundColor Yellow
                    $userQuit = $true; $resolved = $true; break
                }
                elseif ($ans -eq 's') {
                    Write-Host "  -> left pending" -ForegroundColor DarkYellow
                    $applyValue = $null; $countAs = 'skip'
                    Write-ProgressEvent -WorkDir $WorkDir -Phase $phaseName -CorrelIdS $correl -JobName $jobName `
                        -Action 'verify' -Status 'skip' -Message ("ask -> operator skipped: {0}" -f $verdict.Reason)
                }
                elseif ($ans -eq 'n') {
                    $applyValue = '2'; $countAs = 'ng'
                    Write-Host "  -> NG (operator)" -ForegroundColor Red
                    Write-ProgressEvent -WorkDir $WorkDir -Phase $phaseName -CorrelIdS $correl -JobName $jobName `
                        -Action 'verify' -Status 'ng' -Message ("ask -> operator NG: {0}" -f $verdict.Reason)
                    $ngList.Add(("{0} : operator NG ({1})" -f $correl, $verdict.Reason))
                }
                else {
                    $applyValue = '1'; $countAs = 'ok'
                    Write-Host "  -> OK (operator)" -ForegroundColor Green
                    Write-ProgressEvent -WorkDir $WorkDir -Phase $phaseName -CorrelIdS $correl -JobName $jobName `
                        -Action 'verify' -Status 'ok' -Message ("ask -> operator OK: {0}" -f $verdict.Reason)
                }
                # The ASK prompt above ran with the shell in the foreground
                # (Bring-ShellToFront). Return focus to Edge so the Send-Tab below
                # and the next row's Reset-FocusToBody land on the HM page, not the
                # console -- otherwise every later keystroke goes to the CLI. The
                # 'q' branch already broke out, so we only reach here for o/n/s and
                # the shell is known-foreground (Switch-ToEdge's Alt+Tab is valid).
                Switch-ToEdge
            }

            # Persist the chosen value (skip leaves the row pending).
            if ($null -ne $applyValue) {
                $item.$snapField = $applyValue
                Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
                Write-Host ("    saved {0}" -f (Split-Path -Leaf $outPath)) -ForegroundColor DarkGray
            }
            if     ($countAs -eq 'ok')   { $totalDone++ }
            elseif ($countAs -eq 'ng')   { $totalNg++ }
            elseif ($countAs -eq 'skip') { $totalSkipped++ }

            Send-Tab $Global:TabCounts.HM_BackToInput
            $resolved = $true
        } until ($resolved)

        if ($userQuit) { break }
    }

    if ($userQuit) { break }

    Bring-ShellToFront
    Write-Host ""
    Write-Host ("  Appl {0} complete." -f $appl) -ForegroundColor Green
}

Bring-ShellToFront
Write-Host ""
Write-Host "===== HmSnap ($Stage) Done =====" -ForegroundColor Green
Write-Host ("  Snapped OK : {0}" -f $totalDone)
if ($totalNg -gt 0) {
    Write-Host ("  NG         : {0}" -f $totalNg) -ForegroundColor Red
}
if ($totalSkipped -gt 0) {
    Write-Host ("  Skipped    : {0}" -f $totalSkipped) -ForegroundColor DarkGray
}

if ($ngList.Count -gt 0) {
    Write-Host ""
    Write-Host ("===== HM NG summary ({0}) =====" -f $ngList.Count) -ForegroundColor Red
    foreach ($n in $ngList) { Write-Host ("  [NG] {0}" -f $n) -ForegroundColor Red }
    Write-Host "  (NG rows stay pending; re-run $phaseName to retry.)" -ForegroundColor DarkYellow
}
