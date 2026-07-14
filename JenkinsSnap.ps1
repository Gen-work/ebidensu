# ============================================================
#  JenkinsSnap.ps1  (Phases: GiftJenkins / GfixJenkins / GiftJenkinsNoFile)
#  v2 -- MappingStore + ProgressLog + F3/F4 instant detection
#
#  For each pending Correl_ID_S in mapping_<Owner>.csv (grouped by TO_code so
#  Edge is navigated to a system's Jenkins folder once per group):
#    1. Ctrl+F search CORREL_ID_S (leave the find bar open for the highlight)
#    2. Capture the Edge main window, crop the border, save snap\<folder>\<id>.png
#    3. (GiftRecv/GfixRecv) download the matching receive file into DATA\GIFT|GFIX
#    4. (detection on) poll page text, classify page kind, archive .txt
#    5. (detection on) parse the file list + verdict (F3/F4): ok -> field=1, ng -> 2
#    6. Persist atomically via MappingStore; append a progress.jsonl event
#
#  SnapVerify (F3/F4) -- spec docs/SnapVerify-Plan.md milestones M3/M6:
#    - <field>_snap value domain: 0/empty = pending, 1 = ok, 2 = NG.
#      '2' STILL counts as pending and is re-offered next run (plan 2.1).
#    - NG conditions (Test-JenkinsFile, plan 4.F3): file not in the list, or
#      found but its DateTime is outside the Expected_Time +- tolerance window.
#    - Time window: one batch prompt at start (Resolve-SnapRunTime); empty
#      Expected_Time cells on pending rows are filled and persisted (plan 2.2).
#    - Page-kind sentinel (plan 3.6): an off-page text (OuterFrame/Empty/
#      Unknown) stops and asks the operator (r=retry / s=skip / q=quit).
#    - Detection covers GiftRecv/GfixRecv (F3) and NoGfix (F4). NoGfix NG
#      writes a .note.json sidecar when localisation is available.
#    - SnapEnabled=$false reverts all modes to pure screenshot.
#
#  Mark.Boxes template-hit sidecar (opt-in via -MarkBoxes, from
#  Config.Mark.Boxes[<folder>]): right after the screenshot is saved, any box
#  carrying a 'Template' key is matched against the just-captured PNG and the
#  hit cached to <correl>.tplhit.json (SnapLocalize.ps1's Write-MarkTemplateHits).
#  Mark.ps1 reads this sidecar first instead of re-matching the archived PNG
#  later, falling back to a live match when it is missing/stale. Empty
#  -MarkBoxes = feature inert, byte-for-byte the old screenshot-only behavior.
#
#  Conventions:
#    - All mapping I/O goes through MappingStore (atomic writes); progress
#      events go to status\progress.jsonl via ProgressLog.
#    - Pure parse/verdict logic lives in SnapVerify.ps1 (unit-tested); this
#      script is only the COM/SendKeys wiring.
#    - Switch params are copied to plain bools before any dot-source.
# ============================================================
#Requires -Version 5.1
param(
    [ValidateSet('GiftRecv','GfixRecv','NoGfix')]
    [string]$Mode           = 'GiftRecv',
    [string]$WorkDir        = '',
    [string]$Owner          = '',
    [string[]]$TargetIds    = @(),
    [switch]$RefreshUrls,
    [switch]$Force,
    [switch]$DryRun,
    [switch]$Interactive,
    [switch]$NoResize,
    [int]$WindowWidth       = 1050,
    [int]$WindowHeight      = 761,
    [int]$CropPx            = 6,
    # Per-side overrides in px. -1 (default) = inherit CropPx for that side
    # (uniform crop; existing -CropPx-only usage is unchanged).
    [int]$CropLeft          = -1,
    [int]$CropTop           = -1,
    [int]$CropRight         = -1,
    [int]$CropBottom        = -1,
    [int]$ActionWaitMs      = 500,
    [int]$ResultWaitMs      = 500,
    [string]$CommonScript   = '',

    # ---- SnapVerify (F3) detection wiring (defaults match VerifyConfig.psd1) ----
    [bool]$SnapEnabled      = $true,
    [bool]$TimeCheck        = $false,  # $false = file existence checks only (no run-time window)
    [int]$ToleranceMinutes  = 30,
    [bool]$SaveText         = $true,
    [int]$PollTimeoutSec    = 10,
    [int]$PollIntervalMs    = 500,
    [string]$TimeColumn     = 'Expected_Time',
    [string]$TimeFormat     = 'yyyy/MM/dd HH:mm:ss',
    [string]$RunTime        = '',   # '' = prompt; else 'n' / 'yyyy/MM/dd HH:mm:ss'
    [string]$RunTolerance   = '',   # non-interactive tolerance override

    # ---- M5/F5 pixel localisation (Config.SnapVerify.Localize); off by default ----
    [hashtable]$Localize    = @{},

    # ---- Snap-time template-hit sidecar (Config.Mark.Boxes[<folder>]) ----
    # Opt-in per box via a 'Template' key, same as Mark.ps1; boxes without one
    # are ignored here. Empty array = feature inert (no sidecar attempted).
    [object[]]$MarkBoxes             = @(),
    [string]$MarkTemplateDir         = '',
    [int]$MarkImageMatchTolerance    = 15
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$forceFlag    = [bool]$Force.IsPresent
$dryFlag      = [bool]$DryRun.IsPresent
$noResizeFlag = [bool]$NoResize.IsPresent
$refreshFlag  = [bool]$RefreshUrls.IsPresent
$snapVerifyOn = [bool]$SnapEnabled
$saveTextOn   = [bool]$SaveText

$scriptDir = Split-Path $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($CommonScript)) {
    $CommonScript = Join-Path $scriptDir 'Common.ps1'
}
. $CommonScript
. (Join-Path $scriptDir 'MappingStore.ps1')
. (Join-Path $scriptDir 'ProgressLog.ps1')
. (Join-Path $scriptDir 'SnapVerify.ps1')
$jkFindHighlight = Join-Path $scriptDir 'Find-ActiveHighlightRow.ps1'
if (Test-Path -LiteralPath $jkFindHighlight) { . $jkFindHighlight }
$jkSnapLocalize = Join-Path $scriptDir 'SnapLocalize.ps1'
if (Test-Path -LiteralPath $jkSnapLocalize) { . $jkSnapLocalize }
. (Join-Path $scriptDir 'JenkinsDownload.ps1')

$jkLocateByImage = Join-Path $scriptDir 'Locate-ByImage.ps1'
if (-not (Test-Path -LiteralPath $jkLocateByImage)) { $jkLocateByImage = '' }

$pageTextScript = Join-Path $scriptDir 'Read-PageText.ps1'

if (-not (Get-Command -Name 'Test-JenkinsFile' -ErrorAction SilentlyContinue)) {
    Write-Host '[ERROR] SnapVerify.ps1 dot-source failed (Test-JenkinsFile not found).' -ForegroundColor Red
    exit 1
}

$Global:Timing = @{ ActionWaitMs = $ActionWaitMs; ResultWaitMs = $ResultWaitMs }

if (-not $WorkDir) { throw '-WorkDir is required' }

# -- mode config ---------------------------------------------------------------
# Field    : CSV column that tracks completion for this mode
# Folder   : subfolder under snap\ for screenshots
# GroupCol : CSV column to group rows by (navigate to Jenkins once per group)
# SearchCol: CSV column value to Ctrl+F search on the Jenkins page
$modeCfg = switch ($Mode) {
    'GiftRecv' { @{
        Field     = 'GIFT_Jenkins_snap'
        Folder    = 'GIFT_Jenkins'
        GroupCol  = 'TO_code'
        SearchCol = 'CORREL_ID_S'
        JOB       = 'JOB_NAME'
    }}
    'GfixRecv' { @{
        Field     = 'GFIX_Jenkins_snap'
        Folder    = 'GFIX_Jenkins'
        GroupCol  = 'TO_code'
        SearchCol = 'CORREL_ID_S'
        JOB       = 'JOB_NAME'
    }}
    'NoGfix'   { @{
        Field     = 'GIFT_noGfixfile_snap'
        Folder    = 'GIFT_noGfixfile'
        GroupCol  = 'TO_code'
        SearchCol = 'CORREL_ID_S'
        JOB       = 'JOB_NAME'
    }}
}

$snapField  = $modeCfg.Field
$snapFolder = $modeCfg.Folder
$groupCol   = $modeCfg.GroupCol
$searchCol  = $modeCfg.SearchCol
$job        = $modeCfg.JOB

# F3/F4 detection applies to receive modes and NoGfix.  In NoGfix mode a
# matching Jenkins file is not a stop-the-world NG; it marks the snap field as
# 2 and writes a note sidecar so MarkGift can add the past-data annotation.
$detectMode = $snapVerifyOn -and ($Mode -in 'GiftRecv','GfixRecv','NoGfix')

# -- mapping -------------------------------------------------------------------
$mappingFile = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)

# MappingStore: single source of truth for read/filter/write.
# $allRows is the FULL set (so we never drop non-target rows on write);
# $pending holds references INTO $allRows, so mutating a pending row and
# then Export-MappingAtomic $allRows persists exactly that change.
$allRows = Import-Mapping $mappingFile
Ensure-MappingColumns -Rows $allRows -Extra @(
    @{ Name = $snapField;  Default = '0' },
    @{ Name = $TimeColumn; Default = ''  }
) | Out-Null

# A snap field is "done" ONLY when exactly '1'. Empty/'0'/'2'(NG) are all
# pending: NG='2' is re-offered every run (plan 2.1). We deliberately do NOT
# use Get-PendingRows here -- its Test-SnapDone treats any non-'0' value
# (including '2') as done, which would hide NG rows.
function Test-JenkinsSnapDone([string]$Value) { return ($Value -eq '1') }

$targets = ConvertTo-TargetIdList $TargetIds
$pending = @()
$doneCount = 0
foreach ($r in $allRows) {
    if (-not (Test-TargetRow $r $targets)) { continue }
    if ($forceFlag) { $pending += $r; continue }
    if (Test-JenkinsSnapDone (Get-RowProp $r $snapField)) { $doneCount++ } else { $pending += $r }
}
$pending = @($pending)

Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -Action 'start' -Status 'info' `
    -Message ("pending={0} force={1} targets=[{2}] detect={3}" -f $pending.Count, $forceFlag, ($targets -join ','), $detectMode)

if ($pending.Count -eq 0) {
    Write-Host "[$Mode] No pending rows." -ForegroundColor Green
    exit 0
}

Write-Host "`n===== JenkinsSnap $Mode =====" -ForegroundColor Green
Write-Host "Pending rows: $($pending.Count)" -ForegroundColor Cyan
Write-Host ("Detection   : {0} (tol +-{1} min)" -f $detectMode, $ToleranceMinutes) -ForegroundColor Cyan

# -- paths ---------------------------------------------------------------------
$snapDir = Join-Path (Join-Path $WorkDir 'snap') $snapFolder
$dataRoot = Join-Path $WorkDir 'DATA'
Ensure-Dir $snapDir
if ($Mode -in 'GiftRecv','GfixRecv') { Ensure-Dir $dataRoot }

# -- URL cache (PS5.1-safe: no -AsHashtable) ----------------------------------
$urlCacheFile = Join-Path $WorkDir 'jenkins_urls.json'
$urlCache = @{}
if (Test-Path $urlCacheFile) {
    try {
        $parsed = Get-Content $urlCacheFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $parsed.PSObject.Properties | ForEach-Object { $urlCache[$_.Name] = $_.Value }
    } catch {}
}
$urlDirty = $false

# -- inline helpers ------------------------------------------------------------
function Invoke-CropPng(
    [string]$path, [int]$crop,
    # -1 (default) = inherit crop for that side (uniform crop).
    [int]$cropLeft = -1, [int]$cropTop = -1, [int]$cropRight = -1, [int]$cropBottom = -1
) {
    if ($cropLeft   -lt 0) { $cropLeft   = $crop }
    if ($cropTop    -lt 0) { $cropTop    = $crop }
    if ($cropRight  -lt 0) { $cropRight  = $crop }
    if ($cropBottom -lt 0) { $cropBottom = $crop }
    if (($cropLeft -le 0 -and $cropTop -le 0 -and $cropRight -le 0 -and $cropBottom -le 0) -or -not (Test-Path -LiteralPath $path)) { return }
    try {
        $orig = [System.Drawing.Image]::FromFile($path)
        $w = $orig.Width  - $cropLeft - $cropRight
        $h = $orig.Height - $cropTop  - $cropBottom
        if ($w -le 0 -or $h -le 0) { $orig.Dispose(); return }
        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $g.DrawImage($orig, -$cropLeft, -$cropTop)
        $g.Dispose()
        $orig.Dispose()
        $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
    } catch {
        Write-Host ("  [WARN] crop failed: {0}" -f $_) -ForegroundColor Yellow
    }
}

# Get-EdgeMainWindowHandle / Activate-EdgeWindow (process-handle-first,
# AppActivate-by-title fallback) now live in Common.ps1 -- this used to be a
# JenkinsSnap-only fix; promoted so GfixLogDownload/MqSnap/HmSnap share it too.

function Move-EdgeToWorkPos([IntPtr]$hWnd) {
    if ($hWnd -eq [IntPtr]::Zero) { return }
    [WinAPI]::MoveWindow($hWnd, 0, 0, $WindowWidth, $WindowHeight, $true) | Out-Null
    Start-Sleep -Milliseconds 200
}

function Bring-ConsoleToFront {
    $hwnd = (Get-Process -Id $PID).MainWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) {
        [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
        [WinAPI]::ShowWindowAsync($hwnd, 9) | Out-Null
    }
}

function Get-CurrentEdgeUrl {
    Send-Key '^l' 300
    Send-Key '^a' 150
    Send-Key '^c' 300
    Send-Key '{ESC}' 200
    $url = [System.Windows.Forms.Clipboard]::GetText()
    if ($url -match '^https?://') { return $url.Trim() }
    return ''
}

# Click the CENTER of the Jenkins page. This (a) gives the document keyboard
# focus so the select-all/copy reads the page rather than the find bar, and
# (b) collapses any prior Ctrl+A select-all so it is not captured in a later
# screenshot -- Esc does NOT clear an Edge text selection, only a click does.
# The generic Click-PageBody clicks (Left+150, Top+150), which on a Jenkins
# folder page lands in the LEFT sidebar; when a build is queued the "Build Queue"
# widget shows a job hyperlink there, so that click navigates Edge into the job.
# The page center carries no link, so it is safe -- same approach as MqSnap's
# Click-MqPageCenter.
function Click-JenkinsPageCenter {
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
# Click the page center first so the select-all lands on the document (not the
# find bar) and so any earlier selection is cleared before we re-select.
function Get-JenkinsPageTextOnce {
    Click-JenkinsPageCenter
    $txt = & $pageTextScript -SelectWaitMs $ActionWaitMs -CopyWaitMs $ResultWaitMs
    if ($null -eq $txt) { return '' }
    return [string]$txt
}

# Poll the page text (A2) until it is a recognised Jenkins file-list page with
# the search term present, an off-page OuterFrame, or the poll timeout elapses.
# A timeout on a JenkinsResult page without the term is the F3 NG case (file
# absent) -- it falls through to the verdict. Returns @{ Text=...; Kind=... }.
function Wait-JenkinsPageReady {
    param(
        [string]$SearchTerm,
        [bool]$RequireTerm = $true
    )

    $text = ''
    $kind = 'Empty'
    $deadline = (Get-Date).AddSeconds([Math]::Max(1, $PollTimeoutSec))
    do {
        $text = Get-JenkinsPageTextOnce
        $kind = Get-SnapPageKind -Phase 'Jenkins' -Text $text
        if ($kind -eq 'OuterFrame') { break }
        if ($kind -eq 'JenkinsResult') {
            # NoGfix (RequireTerm=$false) expects the correl to be ABSENT, so a
            # loaded file-list page is "ready" as soon as it classifies as a
            # Jenkins result; waiting for the term to appear would always burn
            # the full timeout (and re-read the clipboard ~20 times per row).
            if (-not $RequireTerm) { break }
            if ($text.Contains($SearchTerm)) { break }
        }
        Start-Sleep -Milliseconds ([Math]::Max(100, $PollIntervalMs))
    } while ((Get-Date) -lt $deadline)

    return @{ Text = $text; Kind = $kind }
}

# -- Group pending rows by TO_code ---------------------------------------------
# Build an ordered list of distinct TO_code values (preserving first-seen order)
$groupOrder = [System.Collections.Generic.List[string]]::new()
$groupMap   = @{}   # TO_code -> list of rows
foreach ($row in $pending) {
    $grp = [string]$row.$groupCol
    if ([string]::IsNullOrWhiteSpace($grp)) { $grp = '(unknown)' }
    if (-not $groupMap.ContainsKey($grp)) {
        $groupMap[$grp] = [System.Collections.Generic.List[object]]::new()
        $groupOrder.Add($grp)
    }
    $groupMap[$grp].Add($row)
}

Write-Host ''
Write-Host 'Groups by TO_code:' -ForegroundColor DarkGray
foreach ($g in $groupOrder) {
    Write-Host ("  {0,-12} : {1} rows" -f $g, $groupMap[$g].Count) -ForegroundColor DarkGray
}

# -- Batch run-time inquiry (plan 2.2) -- one prompt, applied to pending rows ----
$timeMode     = 'none'
$runTolerance = $ToleranceMinutes
if ($detectMode -and -not $dryFlag -and $TimeCheck) {
    Bring-ConsoleToFront
    Write-Host ''
    Write-Host "[Time window] Jenkins files are checked against a run time +- tolerance." -ForegroundColor Cyan
    Write-Host "  [Enter]             = use current time" -ForegroundColor Gray
    Write-Host "  yyyy/MM/dd HH:mm:ss = use that time" -ForegroundColor Gray
    Write-Host "  n                   = no time check this run" -ForegroundColor Gray
    $tIn   = if ([string]::IsNullOrWhiteSpace($RunTime)) { Read-Host "  Run time" }                           else { $RunTime }
    $tolIn = if ([string]::IsNullOrWhiteSpace($RunTime)) { Read-Host "  Tolerance minutes (Enter=default)" }  else { $RunTolerance }

    $rt = Resolve-SnapRunTime -TimeInput $tIn -ToleranceInput $tolIn -DefaultTolerance $ToleranceMinutes
    if (-not $rt.Ok) {
        Write-Host ("  [WARN] {0} -- continuing with NO time check." -f $rt.Error) -ForegroundColor Yellow
        $timeMode = 'none'
    } else {
        $timeMode     = $rt.TimeMode
        $runTolerance = $rt.ToleranceMinutes
        if ($timeMode -eq 'fixed') {
            $timeStr = $rt.Time.ToString($TimeFormat)
            $filled  = Set-EmptyRunTimeCells -Rows $pending -Field $TimeColumn -Value $timeStr
            if ($filled -gt 0) {
                Export-MappingAtomic -Rows $allRows -Path $mappingFile | Out-Null
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

# -- Process each TO_code group ------------------------------------------------
$cntDone = 0
$cntSkip = 0
$cntFail = 0
$cntNg   = 0
$ngList  = [System.Collections.Generic.List[string]]::new()
$userQuit    = $false
$maxAttempts = 3

foreach ($toCode in $groupOrder) {
    if ($userQuit) { break }
    $rows = $groupMap[$toCode]

    Write-Host ''
    Write-Host ("===== TO_code: {0}  ({1} rows) =====" -f $toCode, $rows.Count) -ForegroundColor Cyan

    # -- DryRun: report the plan only; never open Edge or prompt. --------------
    if ($dryFlag) {
        foreach ($row in $rows) {
            $correl = [string]$row.Correl_ID_S
            Write-Host ("  [DryRun] would snap {0} (search: {1}) -> snap\{2}\{0}.png" -f `
                $correl, [string]$row.$searchCol, $snapFolder) -ForegroundColor DarkGray
            $cntSkip++
        }
        continue
    }

    # -- Navigate Edge to this system's Jenkins folder -------------------------
    $cacheKey  = "{0}_{1}" -f $Mode, $toCode
    $cachedUrl = if ($urlCache.ContainsKey($cacheKey)) { $urlCache[$cacheKey] } else { '' }

    if (-not $refreshFlag -and $cachedUrl) {
        Write-Host ("  [cached URL] {0}" -f $cachedUrl) -ForegroundColor DarkGray
        Write-Host ("  Navigating Edge to cached Jenkins folder URL for {0}." -f $toCode) -ForegroundColor Yellow
        Write-Host '  Press Enter to confirm, r+Enter to refresh URL, q=quit.' -ForegroundColor Magenta
        $resp = Read-Host
        if ($resp -eq 'q') { $userQuit = $true; break }

        if ($resp -eq 'r') {
            # force re-navigate
            $cachedUrl = ''
        } else {
            $edgeHwnd = Activate-EdgeWindow
            if (-not $noResizeFlag) { Move-EdgeToWorkPos $edgeHwnd }
            # Navigate to cached URL
            Send-Key '^l' 300
            [System.Windows.Forms.Clipboard]::SetText($cachedUrl)
            Send-Key '^v' 200
            Send-Key '{ENTER}' ($ResultWaitMs * 4)
        }
    }

    if (-not $cachedUrl) {
        Write-Host ''
        Write-Host (">>> Open Edge to the [{0}] Jenkins folder page." -f $toCode) -ForegroundColor Yellow
        Write-Host '    (e.g. for IDS: the IDS Jenkins job list page)' -ForegroundColor Yellow
        Write-Host '    Press Enter when ready. (q=quit)' -ForegroundColor Magenta
        $resp = Read-Host
        if ($resp -eq 'q') { $userQuit = $true; break }

        $edgeHwnd = Activate-EdgeWindow
        if (-not $noResizeFlag) { Move-EdgeToWorkPos $edgeHwnd }

        # Capture and cache the URL
        $capturedUrl = Get-CurrentEdgeUrl
        if ($capturedUrl) {
            $urlCache[$cacheKey] = $capturedUrl
            $urlDirty = $true
            Write-Host ("  [URL saved] {0}" -f $capturedUrl) -ForegroundColor Green
        } else {
            Write-Host '  [WARN] Could not capture URL. Continuing without cache.' -ForegroundColor Yellow
        }
    }

    # -- Per-row screenshot loop -----------------------------------------------
    foreach ($row in $rows) {
        if ($userQuit) { break }

        $correl     = [string]$row.Correl_ID_S
        $searchJob  = [string]$row.$job
        $searchTerm = [string]$row.$searchCol

        Write-Host ''
        Write-Host ("  [$correl] search: $searchTerm") -ForegroundColor White

        $snapPath = Join-Path $snapDir "$correl.png"

        $resolved = $false
        $attempt  = 0
        do {
            $attempt++

            $edgeHwnd = Activate-EdgeWindow
            if ($edgeHwnd -eq [IntPtr]::Zero) {
                Write-Host "    [FAIL] Edge not found" -ForegroundColor Red
                Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                    -JobName $searchJob -Action 'snap' -Status 'fail' -Message 'Edge window not found'
                $cntFail++; $resolved = $true; break
            }

            # Click the page center BEFORE Ctrl+F: it clears any select-all left
            # by the previous row's page-text read (Esc does not clear an Edge
            # selection) so the highlight is not captured here, and it focuses the
            # page. The center is used rather than Click-PageBody's
            # (Left+150, Top+150), which lands on the left-sidebar "Build Queue"
            # job link when a build is queued and navigates Edge into that job
            # before we capture.
            Click-JenkinsPageCenter
            # Ctrl+F search for CORREL_ID_S; leave the find bar open so the
            # screenshot shows the highlighted match (ESC is sent after capture).
            Send-CtrlF
            Paste-Replace $searchTerm
            Start-Sleep -Milliseconds $ResultWaitMs

            try {
                Take-WindowScreenshot $edgeHwnd $snapPath
                Invoke-CropPng $snapPath $CropPx -cropLeft $CropLeft -cropTop $CropTop -cropRight $CropRight -cropBottom $CropBottom
                # Dismiss find bar now that the screenshot is captured.
                Send-Key '{ESC}' 200
                Write-Host ("    Saved: snap\{0}\{1}.png" -f $snapFolder, $correl) -ForegroundColor Green

                # Opt-in per Mark.Boxes[$snapFolder] entry ('Template' key):
                # run the same image-match now, against the page as just
                # captured, and cache the hit so Mark.ps1 does not need to
                # re-scan the archived PNG later. Best-effort -- never blocks
                # the snap on failure (see Write-MarkTemplateHits).
                if ($MarkBoxes.Count -gt 0 -and (Get-Command -Name 'Write-MarkTemplateHits' -ErrorAction SilentlyContinue)) {
                    $tplHitPath = Write-MarkTemplateHits -SnapDir $snapDir -Correl $correl -PngPath $snapPath `
                        -Boxes $MarkBoxes -TemplateDir $MarkTemplateDir -Tolerance $MarkImageMatchTolerance `
                        -LocateScript $jkLocateByImage
                    if ($tplHitPath) { Write-Host ("    tplhit: snap\{0}\{1}.tplhit.json" -f $snapFolder, $correl) -ForegroundColor DarkGray }
                }
            } catch {
                Write-Host ("    [FAIL] screenshot: {0}" -f $_.Exception.Message) -ForegroundColor Red
                Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                    -JobName $searchJob -Action 'screenshot' -Status 'fail' -Message $_.Exception.Message
                $cntFail++; $resolved = $true; break
            }

            # Read page text once -- needed for detection (F3) and/or the file
            # download (GiftRecv/GfixRecv). Reuse the same text for both.
            $pageText = ''
            $pageKind = 'Unknown'
            $needPageText = $detectMode -or ($Mode -in 'GiftRecv','GfixRecv')
            if ($needPageText) {
                if ($detectMode) {
                    $ready    = Wait-JenkinsPageReady -SearchTerm $searchTerm -RequireTerm ($Mode -ne 'NoGfix')
                    $pageText = [string]$ready.Text
                    $pageKind = [string]$ready.Kind
                    Write-Host ("    pageKind: {0}" -f $pageKind) -ForegroundColor DarkGray
                } else {
                    $pageText = Get-JenkinsPageTextOnce
                }
            }

            # Archive page text (A1) -- before the sentinel so NG/ambiguous pages
            # still leave an offline record.
            if ($detectMode -and $saveTextOn -and -not [string]::IsNullOrWhiteSpace($pageText)) {
                try {
                    $txtPath = Join-Path $snapDir ("{0}.txt" -f $correl)
                    $enc = New-Object System.Text.UTF8Encoding($false)
                    [System.IO.File]::WriteAllText($txtPath, $pageText, $enc)
                } catch {
                    Write-Host ("    [WARN] save text failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                }
            }

            # Sentinel (A3): off-page kinds -> stop and ask the operator.
            if ($detectMode -and ($pageKind -eq 'OuterFrame' -or $pageKind -eq 'Empty' -or $pageKind -eq 'Unknown')) {
                $preview = if ($pageText.Length -gt 200) { $pageText.Substring(0, 200) } else { $pageText }
                Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                    -JobName $searchJob -Action 'verify' -Status 'warn' -Message ("pageKind={0}: {1}" -f $pageKind, $preview)
                Bring-ConsoleToFront
                Write-Host ("  [SENTINEL] Unexpected page ({0}) for {1}." -f $pageKind, $correl) -ForegroundColor Yellow
                Write-Host "    Fix Edge focus/frame, then r=retry this row / s=skip / q=quit : " -ForegroundColor Magenta -NoNewline
                $ans = Read-Host
                if ($ans -eq 'q') { $userQuit = $true; $resolved = $true; break }
                if ($ans -eq 's') { Write-Host "  -> skipped" -ForegroundColor DarkYellow; $cntSkip++; $resolved = $true; break }
                if ($attempt -ge $maxAttempts) {
                    Write-Host ("  -> still unexpected after {0} attempts; skipping." -f $maxAttempts) -ForegroundColor Yellow
                    $cntSkip++; $resolved = $true; break
                }
                continue   # retry this row
            }

            # Download the matching Jenkins receive file into DATA\GIFT or DATA\GFIX.
            if ($Mode -in 'GiftRecv','GfixRecv') {
                try {
                    $folderUrl = Get-CurrentEdgeUrl
                    if ([string]::IsNullOrWhiteSpace($folderUrl)) { $folderUrl = $cachedUrl }

                    if ([string]::IsNullOrWhiteSpace($pageText) -or [string]::IsNullOrWhiteSpace($folderUrl)) {
                        Write-Host '    [WARN] download skipped: page text or Jenkins folder URL is empty' -ForegroundColor Yellow
                        Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                            -JobName $searchJob -Action 'download' -Status 'warn' -Message 'page text or folder URL is empty'
                    } else {
                        $dl = Invoke-JenkinsFileDownload -WorkDir $WorkDir -Mode $Mode -FolderUrl $folderUrl `
                            -PageText $pageText -CorrelId $correl -JobName $searchJob -Force:$forceFlag `
                            -ParserScript (Join-Path $scriptDir 'Parse-JenkinsList.ps1')
                        Write-Host ("    Jenkins files: found={0} matched={1} downloaded={2} skipped={3} failed={4} -> DATA\{5}" -f `
                            $dl.Found, $dl.Matched, $dl.Downloaded, $dl.Skipped, $dl.Failed, $dl.DataKind) -ForegroundColor DarkGray
                        foreach ($f in @($dl.Files)) {
                            $color = if ($f.Status -eq 'ok') { 'Gray' } elseif ($f.Status -eq 'skip') { 'DarkGray' } else { 'Yellow' }
                            Write-Host ("      [{0}] {1}" -f $f.Status, $f.Name) -ForegroundColor $color
                        }
                        $dlStatus = if ($dl.Failed -gt 0) { 'fail' } elseif ($dl.Matched -eq 0) { 'warn' } else { 'ok' }
                        $dlMessage = "DATA\$($dl.DataKind): found=$($dl.Found) matched=$($dl.Matched) downloaded=$($dl.Downloaded) skipped=$($dl.Skipped) failed=$($dl.Failed)"
                        Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                            -JobName $searchJob -Action 'download' -Status $dlStatus -Message $dlMessage
                    }
                } catch {
                    Write-Host ("    [WARN] download failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                    Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                        -JobName $searchJob -Action 'download' -Status 'fail' -Message $_.Exception.Message
                }
            }

            # -- Detection OFF: legacy screenshot -> 1 ------------------------
            if (-not $detectMode) {
                try {
                    $row.$snapField = '1'
                    Export-MappingAtomic -Rows $allRows -Path $mappingFile | Out-Null
                    Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                        -JobName $searchJob -Action 'snap' -Status 'ok' `
                        -Message ("snap\{0}\{1}.png" -f $snapFolder, $correl)
                    $cntDone++
                } catch {
                    Write-Host ("    [WARN] mapping update failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                    Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                        -JobName $searchJob -Action 'mapping' -Status 'fail' -Message $_.Exception.Message
                }
                $resolved = $true; break
            }

            # -- F3/F4 verdict: ok -> field=1, ng -> field=2 ------------------
            $rowExpected = $null
            if ($timeMode -ne 'none') {
                $rowExpected = ConvertTo-ExpectedDateTime -Value (Get-RowProp $row $TimeColumn) -Format $TimeFormat
            }

            $verdict = $null
            try {
                $files   = ConvertFrom-JenkinsListText $pageText
                $expectExists = ($Mode -ne 'NoGfix')
                $verdict = Test-JenkinsFile -Files $files -CorrelId $searchTerm -Expected $rowExpected `
                            -ToleranceMin $runTolerance -ExpectExists $expectExists
            } catch {
                # Detection bugs must never block the screenshot workflow: keep the
                # shot, mark done, and record a warning for review.
                Write-Host ("    [WARN] detection failed: {0} (marking snap done)" -f $_.Exception.Message) -ForegroundColor Yellow
                Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                    -JobName $searchJob -Action 'verify' -Status 'warn' -Message ("detect error: {0}" -f $_.Exception.Message)
                $verdict = @{ Verdict = 'ok'; Reason = 'detection error -> screenshot kept' }
            }

            if ($verdict.Verdict -eq 'ok') {
                $row.$snapField = '1'
                Write-Host ("    -> OK: {0}" -f $verdict.Reason) -ForegroundColor Green
                Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                    -JobName $searchJob -Action 'verify' -Status 'ok' -Message $verdict.Reason
                $cntDone++
            } else {
                $row.$snapField = '2'
                Write-Host ("    -> NG: {0}" -f $verdict.Reason) -ForegroundColor Red
                Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                    -JobName $searchJob -Action 'verify' -Status 'ng' -Message $verdict.Reason
                $ngList.Add(("{0} : {1}" -f $correl, $verdict.Reason))
                $cntNg++
            }

            # M5/F5: best-effort <correl>.loc.json sidecar (orange Ctrl+F highlight row).
            $locPath = $null
            if ([bool]$Localize['Enabled'] -and (Get-Command -Name 'Write-SnapLocalize' -ErrorAction SilentlyContinue)) {
                $locPath = Write-SnapLocalize -Page 'Jenkins' -Localize $Localize -SnapDir $snapDir `
                    -Correl $correl -PngPath $snapPath
                if ($locPath) { Write-Host ("    loc: snap\{0}\{1}.loc.json" -f $snapFolder, $correl) -ForegroundColor DarkGray }
            }

            # M6/F4: for NoGfix past-data hits, carry the localisation payload
            # forward to ReplaceEvidence/Mark via <correl>.note.json.
            if ($Mode -eq 'NoGfix' -and $verdict.Verdict -eq 'ng' -and $locPath -and (Test-Path -LiteralPath $locPath)) {
                try {
                    $loc = Get-Content -LiteralPath $locPath -Raw -Encoding UTF8 | ConvertFrom-Json
                    $fileDt = ''
                    if ($verdict.File -and $verdict.File.DateTime) { $fileDt = ([datetime]$verdict.File.DateTime).ToString('yyyy/MM/dd HH:mm:ss') }
                    $note = [ordered]@{
                        kind         = 'NoGfixPastData'
                        correl       = $correl
                        folder       = $snapFolder
                        reason       = [string]$verdict.Reason
                        fileDateTime = $fileDt
                        pixelRect    = [ordered]@{ x = [int]$loc.x; y = [int]$loc.y; w = [int]$loc.w; h = [int]$loc.h }
                        imageWidth   = [int]$loc.imageWidth
                        imageHeight  = [int]$loc.imageHeight
                        created      = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
                    }
                    $notePath = Join-Path $snapDir ("{0}.note.json" -f $correl)
                    $enc = New-Object System.Text.UTF8Encoding($false)
                    [System.IO.File]::WriteAllText($notePath, ($note | ConvertTo-Json -Depth 6), $enc)
                    Write-Host ("    note: snap\{0}\{1}.note.json" -f $snapFolder, $correl) -ForegroundColor DarkGray
                } catch {
                    Write-Host ("    [WARN] note sidecar failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                }
            }

            # M6/F4: a NoGfix row that now reads OK (the past-data file is gone)
            # must not keep a stale note.json from an earlier run, or the next
            # ReplaceEvidence would re-stamp a past-data annotation that no longer
            # applies. Only clear on a clean OK (never on ng -- an un-localised
            # ng would otherwise lose a still-valid sidecar).
            if ($Mode -eq 'NoGfix' -and $verdict.Verdict -eq 'ok') {
                $staleNote = Join-Path $snapDir ("{0}.note.json" -f $correl)
                if (Test-Path -LiteralPath $staleNote) {
                    try { Remove-Item -LiteralPath $staleNote -Force } catch {}
                }
            }

            try {
                Export-MappingAtomic -Rows $allRows -Path $mappingFile | Out-Null
            } catch {
                Write-Host ("    [WARN] mapping update failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                    -JobName $searchJob -Action 'mapping' -Status 'fail' -Message $_.Exception.Message
            }
            $resolved = $true
        } until ($resolved)
    }
}

# -- save URL cache ------------------------------------------------------------
if ($urlDirty) {
    $urlCache | ConvertTo-Json -Depth 3 | Set-Content $urlCacheFile -Encoding UTF8
    Write-Host ''
    Write-Host "[$Mode] URL cache saved: jenkins_urls.json" -ForegroundColor DarkGray
}

# -- summary -------------------------------------------------------------------
Write-Host ''
Write-Host ("===== JenkinsSnap $Mode Done =====") -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
if ($cntNg -gt 0) {
    Write-Host ("  NG      : {0}" -f $cntNg) -ForegroundColor Red
}
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })

if ($ngList.Count -gt 0) {
    Write-Host ''
    Write-Host ("===== Jenkins $Mode NG summary ({0}) =====" -f $ngList.Count) -ForegroundColor Red
    foreach ($n in $ngList) { Write-Host ("  [NG] {0}" -f $n) -ForegroundColor Red }
    Write-Host "  (NG rows stay pending; re-run to retry.)" -ForegroundColor DarkYellow
}

Bring-ConsoleToFront
