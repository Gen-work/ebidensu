#Requires -Version 5.1
# ============================================================
#  GfixLogDownload.ps1   (Phase: GfixLogDownload)
#  UTF-8, NO BOM, ASCII source. The one Japanese runtime string (the
#  GoAnywhere "download" link text) is built via [char] so this file is
#  codepage-agnostic.
#
#  Per-IF_NO TRANSACTION (spec 5, reworked for duplicate IF_NO safety):
#    Ctrl+A the LIST page ONCE -> parse every job row (GfixJobList.ps1) ->
#    group pending mapping rows by normalized IF_NO -> for each IF_NO,
#    download EVERY matching job (not just the first Find-in-page hit).
#
#  Why per-IF_NO, not per-correl: one IF_NO commonly feeds more than one
#  downstream SS_CODE receive job (e.g. JIDSC02S and JIDSC03S both off
#  IF5001_001), so the GoAnywhere list carries duplicate rows with the
#  identical project name. Ctrl+F-searching that project name text cannot
#  tell the rows apart -- it used to land on the SAME physical row for
#  every correl sharing that IF_NO, so one correl's real log was never
#  downloaded at all (silently, since "a new file appeared" was enough to
#  mark GFIX_log=1 with no content check). Job numbers ARE unique, so this
#  version opens each one by its job number instead, downloads every
#  candidate log for a needed IF_NO, and only THEN asks GfixLog.ps1's
#  content matcher (Find-GfixLogForCorrel -- already used at ReplaceGfix
#  time, unit-tested, no changes needed) which correl each log actually
#  belongs to. GFIX_log is only ever set to 1 after a real content match.
#
#  Missing detection: an IF_NO with zero matching rows in today's list is
#  a hard miss (nothing to navigate to) -- every correl needing it is
#  marked GFIX_log=2 immediately, an error is logged, and the run
#  continues with the next IF_NO. No interactive prompt for this case:
#  there is no GoAnywhere state to retry against.
#
#  Detail-page open: Ctrl+F <job number>, Enter, Esc, Enter (job number is
#  itself the row's link target -- no Shift+Tab needed).
#
#  Log file naming: <JobNo>_<timestamp>_<originalName>.log (never named
#  after a correl -- correctness comes from content matching, not from
#  which row happened to be opened first).
#
#  On download failure for a job number (no new log AND none already
#  present for that job number):
#    - Interactive (default): PAUSE. Operator downloads / drops the log
#      into work\log\, re-shows the GoAnywhere LIST page, then Enter to
#      retry, s to skip that job number, q to quit.
#    - -NonInteractive: mark fail and STOP the run (skip the rest) so a
#      wrong page state cannot cascade.
#
#  Every step is recorded in status\progress.jsonl. Mapping writes are
#  atomic (MappingStore.Export-MappingAtomic).
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner = '',
    [string[]]$TargetIds = @(),
    [switch]$Force,
    [switch]$NonInteractive,
    [int]$ActionWaitMs = 500,
    [int]$ResultWaitMs = 500,
    [string]$CommonScript = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$scriptDir = Split-Path $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($CommonScript)) { $CommonScript = Join-Path $scriptDir 'Common.ps1' }
. $CommonScript
. (Join-Path $scriptDir 'MappingStore.ps1')
. (Join-Path $scriptDir 'ProgressLog.ps1')
. (Join-Path $scriptDir 'GfixLog.ps1')
. (Join-Path $scriptDir 'GfixJobList.ps1')

$pageTextScript = Join-Path $scriptDir 'Read-PageText.ps1'

$Global:Timing = @{ ActionWaitMs = $ActionWaitMs; ResultWaitMs = $ResultWaitMs }

$forceFlag   = [bool]$Force.IsPresent
$interactive = -not [bool]$NonInteractive.IsPresent

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) { Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1 }

$mappingFile = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
$logDir      = Join-Path $WorkDir 'log'
Ensure-Dir $logDir
$downloadDir = Join-Path $env:USERPROFILE 'Downloads'

# GoAnywhere "download" link text: katakana 'daunro-do' (U+30C0 U+30A6 U+30F3 U+30ED U+30FC U+30C9)
$dlSearch = [char]0x30C0 + [char]0x30A6 + [char]0x30F3 + [char]0x30ED + [char]0x30FC + [char]0x30C9

$allRows = Import-Mapping $mappingFile
Ensure-MappingColumns -Rows $allRows | Out-Null
$targets = ConvertTo-TargetIdList $TargetIds

# GFIX_log is a plain value column, NOT a bitmask: '2' (NG -- content match
# failed / IF_NO missing from today's list) still counts as pending, same as
# SendVsGift's NG=2 convention. Get-PendingRows/Test-SnapDone treats any
# non-'0' value as done, which would hide '2' rows -- so, like Mq/Hm/JenkinsSnap,
# use a local "done == exactly '1'" filter instead.
function Test-GfixLogDone([string]$Value) { return ($Value -eq '1') }
$pending = [System.Collections.Generic.List[object]]::new()
foreach ($r in $allRows) {
    if (-not (Test-TargetRow $r $targets)) { continue }
    if ($forceFlag) { $pending.Add($r); continue }
    if (-not (Test-GfixLogDone (Get-RowProp $r 'GFIX_log'))) { $pending.Add($r) }
}
$pending = @($pending.ToArray())

Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -Action 'start' -Status 'info' `
    -Message ("pending={0} interactive={1}" -f $pending.Count, $interactive)

if ($pending.Count -eq 0) { Write-Host '[GfixLogDownload] No pending rows.' -ForegroundColor Green; exit 0 }

# -- local helpers --
function Get-JobNameForRow($row) {
    if ($row.PSObject.Properties.Name -contains 'JOB_NAME') { return [string]$row.JOB_NAME }
    return ''
}
function Get-NewDownloadedLog([datetime]$sinceTime, [hashtable]$beforeSet) {
    $hits = @(Get-ChildItem -LiteralPath $downloadDir -Filter '*.log' -ErrorAction SilentlyContinue | Where-Object {
        (-not $beforeSet.ContainsKey($_.Name)) -or ($_.LastWriteTime -gt $sinceTime) -or ($_.CreationTime -gt $sinceTime)
    } | Sort-Object LastWriteTime -Descending)
    if ($hits.Count -gt 0) { return $hits[0] }
    return $null
}
function Test-JobLogExists([string]$jobNo) {
    return (@(Get-ChildItem -LiteralPath $logDir -Filter ("*{0}*" -f $jobNo) -File -ErrorAction SilentlyContinue).Count -gt 0)
}
function Move-JobLogToWork($file, [string]$jobNo) {
    $ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $targetName = "{0}_{1}_{2}" -f $jobNo, $ts, $file.Name
    Move-Item -LiteralPath $file.FullName -Destination (Join-Path $logDir $targetName) -Force
    return $targetName
}
function Complete-CorrelRow($row, [string]$correl, [string]$jobName, [string]$msg) {
    $row.GFIX_log = '1'
    Export-MappingAtomic -Rows $allRows -Path $mappingFile | Out-Null
    Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -CorrelIdS $correl -JobName $jobName `
        -Action 'log' -Status 'ok' -Message $msg
}
function Fail-CorrelRow($row, [string]$correl, [string]$jobName, [string]$msg) {
    $row.GFIX_log = '2'
    Export-MappingAtomic -Rows $allRows -Path $mappingFile | Out-Null
    Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -CorrelIdS $correl -JobName $jobName `
        -Action 'log' -Status 'fail' -Message $msg
}

Write-Host "`n===== GFIX log download (GoAnywhere) =====" -ForegroundColor Green
Write-Host "Pending rows: $($pending.Count)" -ForegroundColor Cyan
Write-Host ''
Write-Host 'Prepare the GoAnywhere page in Edge BEFORE starting:' -ForegroundColor Yellow
Write-Host '  - Filter the BIZ code (example: JRV)'
Write-Host '  - Sort newer=up / older=down'
Write-Host '  *** Set rows-per-page to 100 (default 20 is not enough!) ***' -ForegroundColor Red
Write-Host '  - Keep focus on the LIST page'
Write-Host 'Then press Enter. (q to quit)' -ForegroundColor Magenta
if ((Read-Host) -eq 'q') { exit 0 }
Switch-ToEdge
$hWnd = [WinAPI]::GetForegroundWindow()
[WinAPI]::ShowWindowAsync($hWnd, 3) | Out-Null   # SW_SHOWMAXIMIZED
Start-Sleep -Milliseconds 300

# -- capture + parse the LIST page once (job number is the identity we key on) --
$jobListRows = @()
$captureOk = $false
while (-not $captureOk) {
    Send-Key '{ESC}' 200
    $listText = & $pageTextScript -SelectWaitMs $ActionWaitMs -CopyWaitMs $ActionWaitMs
    $jobListRows = @(ConvertFrom-GfixJobListText ([string]$listText))
    if ($jobListRows.Count -gt 0) { $captureOk = $true; break }
    Write-Host '[WARN] parsed 0 job rows from the list page (Ctrl+A/Ctrl+C).' -ForegroundColor Yellow
    if (-not $interactive) { Write-Host '[ERROR] non-interactive: aborting.' -ForegroundColor Red; exit 1 }
    Write-Host 'Make sure the GoAnywhere completed-jobs LIST page is focused in Edge,' -ForegroundColor Yellow
    Write-Host 'then: [Enter]=retry capture, q=quit' -ForegroundColor Magenta
    if ((Read-Host) -eq 'q') { exit 0 }
}
Write-Host ("[GfixLogDownload] parsed {0} job row(s) from the list." -f $jobListRows.Count) -ForegroundColor DarkGray

$cntDone = 0; $cntFail = 0; $cntSkip = 0
$stopRun = $false
$resolved = @{}     # Correl_ID_S -> $true once GFIX_log has been finalized
$correlToRow = @{}  # Correl_ID_S -> mapping row (for HardMiss lookback below)

# -- rows with empty IF are skipped up front (nothing to plan for); everything
#    else feeds the pure planner (Get-GfixLogDownloadPlan, GfixJobList.ps1) --
$plannerRows = [System.Collections.Generic.List[object]]::new()
foreach ($row in $pending) {
    $correl = [string]$row.Correl_ID_S
    $jobName = Get-JobNameForRow $row
    $correlToRow[$correl] = $row
    $ifRaw = [string]$row.IF
    if ([string]::IsNullOrWhiteSpace($ifRaw)) {
        Write-Host ("[SKIP] {0}: empty IF" -f $correl) -ForegroundColor DarkGray
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -CorrelIdS $correl -JobName $jobName -Action 'search' -Status 'skip' -Message 'empty IF'
        $resolved[$correl] = $true
        $cntSkip++; continue
    }
    $ifNorm = ($ifRaw -replace '-', '_').Trim()
    $plannerRows.Add([pscustomobject]@{ CorrelIdS = $correl; IfNorm = $ifNorm })
}

$plan = Get-GfixLogDownloadPlan -PendingRows $plannerRows.ToArray() -JobListRows $jobListRows -ReceiveOnly $true

foreach ($miss in @($plan.HardMiss)) {
    $row = $correlToRow[$miss.CorrelIdS]
    $jobName = Get-JobNameForRow $row
    $msg = ("no GoAnywhere job found for IF={0} in today's list" -f $miss.IfNorm)
    Write-Host ("[{0}] [FAIL] {1}" -f $miss.CorrelIdS, $msg) -ForegroundColor Red
    Fail-CorrelRow $row $miss.CorrelIdS $jobName $msg
    $resolved[$miss.CorrelIdS] = $true
    $cntFail++
}

$neededJobNumbers = @($plan.NeededJobNumbers)
if ($neededJobNumbers.Count -gt 0) {
    Write-Host ("[GfixLogDownload] {0} distinct job number(s) to fetch." -f $neededJobNumbers.Count) -ForegroundColor DarkGray
}

# -- download every needed job number (idempotent: skip ones already in log\) --
foreach ($jobNo in $neededJobNumbers) {
    if ($stopRun) { break }
    if (Test-JobLogExists $jobNo) {
        Write-Host ("[JobNo {0}] already in log\, skipping download." -f $jobNo) -ForegroundColor DarkGray
        continue
    }
    Write-Host ("`n[JobNo {0}] opening detail..." -f $jobNo) -ForegroundColor White

    try {
        # 1) open the job's detail page by its unique job number
        Send-CtrlF
        Paste-Replace $jobNo
        Send-Enter
        Send-Key '{ESC}' 300
        Send-Enter
        Start-Sleep -Milliseconds 1200

        # 2) snapshot Downloads, then click download
        $before = @{}
        Get-ChildItem -LiteralPath $downloadDir -Filter '*.log' -ErrorAction SilentlyContinue |
            ForEach-Object { $before[$_.Name] = $_.LastWriteTime }
        $beforeTime = Get-Date

        Send-CtrlF
        Paste-Replace $dlSearch
        Send-Enter
        Send-Key '{ESC}' 200
        Send-Enter
        Start-Sleep -Milliseconds 1800

        # 3) detect new log
        $new = Get-NewDownloadedLog $beforeTime $before
        if ($null -ne $new) {
            $name = Move-JobLogToWork $new $jobNo
            Write-Host ("  moved: {0} -> log\{1}" -f $new.Name, $name) -ForegroundColor Green
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -Action 'download' -Status 'ok' -Message ("job {0} -> {1}" -f $jobNo, $name)
            # return to the list page for the next job
            Send-Tab 2
            Send-Enter
            Start-Sleep -Milliseconds 800
            continue
        }

        # 4) failure handling -- do NOT blindly continue
        Write-Host ("  [FAIL] no new log for job {0}" -f $jobNo) -ForegroundColor Red
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -Action 'download' -Status 'fail' -Message ("job {0}: no new .log and none in work\log" -f $jobNo)

        if (-not $interactive) {
            Write-Host '  [STOP] non-interactive: stopping run to avoid blind navigation.' -ForegroundColor Red
            $stopRun = $true
            continue
        }

        $jobResolved = $false
        while (-not $jobResolved) {
            Write-Host ''
            Write-Host ("  Log not found for job {0}." -f $jobNo) -ForegroundColor Yellow
            Write-Host ("  Download it (or drop a *{0}*.log into {1}), re-show the GoAnywhere LIST page," -f $jobNo, $logDir) -ForegroundColor Yellow
            Write-Host '  then: [Enter]=retry, s=skip this job, q=quit' -ForegroundColor Magenta
            $resp = Read-Host
            if ($resp -eq 'q') { $stopRun = $true; break }
            if ($resp -eq 's') {
                Write-Host ("  [SKIP] job {0} left undownloaded" -f $jobNo) -ForegroundColor DarkGray
                Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -Action 'manual' -Status 'skip' -Message ("operator skipped job {0}" -f $jobNo)
                break
            }
            $retryNew = Get-NewDownloadedLog $beforeTime $before
            if ($null -ne $retryNew) {
                $name = Move-JobLogToWork $retryNew $jobNo
                Write-Host ("  moved: {0} -> log\{1}" -f $retryNew.Name, $name) -ForegroundColor Green
                Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -Action 'download' -Status 'ok' -Message ("manual: job {0} -> {1}" -f $jobNo, $name)
                $jobResolved = $true
                Write-Host '  Log saved. Return Edge to the GoAnywhere LIST page, then press Enter.' -ForegroundColor Magenta
                if ((Read-Host) -eq 'q') { $stopRun = $true }
                continue
            }
            if (Test-JobLogExists $jobNo) {
                Write-Host ("  found in log\ for job {0}" -f $jobNo) -ForegroundColor Green
                $jobResolved = $true
                Write-Host '  Log saved. Return Edge to the GoAnywhere LIST page, then press Enter.' -ForegroundColor Magenta
                if ((Read-Host) -eq 'q') { $stopRun = $true }
                continue
            }
            Write-Host '  still not found.' -ForegroundColor Yellow
        }
    } catch {
        Write-Host ("  [FAIL] job {0}: {1}" -f $jobNo, $_.Exception.Message) -ForegroundColor Red
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -Action 'exception' -Status 'fail' -Message ("job {0}: {1}" -f $jobNo, $_.Exception.Message)
        if ($interactive) {
            Write-Host '  Fix the page manually, then [Enter]=continue, q=quit' -ForegroundColor Magenta
            if ((Read-Host) -eq 'q') { $stopRun = $true }
        } else {
            $stopRun = $true
        }
    }
}

# Return focus to this console
$consoleHwnd = (Get-Process -Id $PID).MainWindowHandle
if ($consoleHwnd -ne [IntPtr]::Zero) {
    [WinAPI]::SetForegroundWindow($consoleHwnd) | Out-Null
    [WinAPI]::ShowWindowAsync($consoleHwnd, 9) | Out-Null
}

# -- finalize: content-match every not-yet-resolved correl against whatever
#    actually made it into log\ (GfixLog.ps1, no changes needed there --
#    it already scans every *.log in the folder and picks by Command: line
#    content when no <correl>_*.log file exists) --
Write-Host "`n===== Matching downloaded logs to correls =====" -ForegroundColor Green
$ngList = [System.Collections.Generic.List[string]]::new()
foreach ($row in $pending) {
    $correl = [string]$row.Correl_ID_S
    if ($resolved.ContainsKey($correl)) { continue }
    $jobName = Get-JobNameForRow $row
    $toCode  = [string]$row.TO_code
    $ssCode  = Get-RowProp $row 'SS_CODE'
    $res = Find-GfixLogForCorrel -LogDir $logDir -ToCode $toCode -CorrelIdS $correl -SsCode $ssCode
    if (-not [string]::IsNullOrWhiteSpace([string]$res.Warning)) {
        Write-Host ("  [WARN] {0}: {1}" -f $correl, $res.Warning) -ForegroundColor Yellow
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -CorrelIdS $correl -JobName $jobName -Action 'match' -Status 'info' -Message $res.Warning
    }
    if ($null -ne $res.Chosen) {
        $msg = ("matched {0}" -f (Split-Path -Leaf $res.Chosen.File))
        Write-Host ("  [OK]   {0}: {1}" -f $correl, $msg) -ForegroundColor Green
        Complete-CorrelRow $row $correl $jobName $msg
        $cntDone++
    } else {
        Write-Host ("  [FAIL] {0}: {1}" -f $correl, $res.Error) -ForegroundColor Red
        Fail-CorrelRow $row $correl $jobName $res.Error
        $ngList.Add(("{0}: {1}" -f $correl, $res.Error))
        $cntFail++
    }
    $resolved[$correl] = $true
}

Write-Host ''
Write-Host '===== GfixLogDownload Done =====' -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
if ($ngList.Count -gt 0) {
    Write-Host '  NG detail:' -ForegroundColor Yellow
    foreach ($ng in $ngList) { Write-Host ("    - {0}" -f $ng) -ForegroundColor Yellow }
}
if ($stopRun) { Write-Host '  (run stopped early)' -ForegroundColor Yellow }
