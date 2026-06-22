#Requires -Version 5.1
# ============================================================
#  GfixLogDownload.ps1   (Phase: GfixLogDownload)
#  UTF-8, NO BOM, ASCII source. The one Japanese runtime string (the
#  GoAnywhere "download" link text) is built via [char] so this file is
#  codepage-agnostic.
#
#  Per-correl TRANSACTION (spec 5):
#    search IF -> open detail -> snapshot Downloads -> click download ->
#    wait for a NEW .log -> move to work\log\ -> return to list.
#  Detection: a .log in Downloads whose name is new OR whose write/create
#  time is after the pre-click snapshot. Fallback: an existing
#  work\log\<Correl_ID_S>_*.log (manually placed / already downloaded).
#
#  On failure (no new log AND none in work\log):
#    - GFIX_log is NOT set.
#    - A PS script cannot read the page state, so we do NOT blindly
#      Tab/Enter into the next row.
#    - Interactive (default): PAUSE. Operator downloads / drops the log
#      into work\log\ and re-shows the GoAnywhere LIST page, then Enter to
#      retry, s to skip (mark fail), q to quit.
#    - -NonInteractive: mark fail and STOP the run (skip the rest) so a
#      wrong page state cannot cascade.
#
#  Log file naming: <Correl_ID_S>_<timestamp>_<originalName>.log
#  GFIX_log = 1 only after a log is really found/moved. Per-row atomic
#  mapping write; every step is recorded in status\progress.jsonl.
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
$pending = @(Get-PendingRows -Rows $allRows -Field 'GFIX_log' -Force $forceFlag -Targets $targets)

Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -Action 'start' -Status 'info' `
    -Message ("pending={0} interactive={1}" -f $pending.Count, $interactive)

if ($pending.Count -eq 0) { Write-Host '[GfixLogDownload] No pending rows.' -ForegroundColor Green; exit 0 }

# -- local helpers --
function Get-NewDownloadedLog([datetime]$sinceTime, [hashtable]$beforeSet) {
    $hits = @(Get-ChildItem -LiteralPath $downloadDir -Filter '*.log' -ErrorAction SilentlyContinue | Where-Object {
        (-not $beforeSet.ContainsKey($_.Name)) -or ($_.LastWriteTime -gt $sinceTime) -or ($_.CreationTime -gt $sinceTime)
    } | Sort-Object LastWriteTime -Descending)
    if ($hits.Count -gt 0) { return $hits[0] }
    return $null
}
function Get-ExistingCorrelLog([string]$correl) {
    $hits = @(Get-ChildItem -LiteralPath $logDir -Filter ("{0}_*.log" -f $correl) -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending)
    if ($hits.Count -gt 0) { return $hits[0] }
    return $null
}
function Move-LogToWork($file, [string]$correl) {
    $ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $targetName = "{0}_{1}_{2}" -f $correl, $ts, $file.Name
    Move-Item -LiteralPath $file.FullName -Destination (Join-Path $logDir $targetName) -Force
    return $targetName
}
function Complete-Row($row, [string]$correl, [string]$jobName, [string]$msg) {
    $row.GFIX_log = '1'
    Export-MappingAtomic -Rows $allRows -Path $mappingFile | Out-Null
    Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -CorrelIdS $correl -JobName $jobName `
        -Action 'log' -Status 'ok' -Message $msg
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

$cntDone = 0; $cntFail = 0; $cntSkip = 0
$stopRun = $false

foreach ($row in $pending) {
    if ($stopRun) { break }

    $correl = [string]$row.Correl_ID_S
    $ifRaw  = [string]$row.IF
    $jobName = ''
    if ($row.PSObject.Properties.Name -contains 'JOB_NAME') { $jobName = [string]$row.JOB_NAME }
    if ([string]::IsNullOrWhiteSpace($ifRaw)) {
        Write-Host ("[SKIP] {0}: empty IF" -f $correl) -ForegroundColor DarkGray
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -CorrelIdS $correl -JobName $jobName -Action 'search' -Status 'skip' -Message 'empty IF'
        $cntSkip++; continue
    }
    $ifNorm = ($ifRaw -replace '-', '_').Trim()
    Write-Host ("`n[{0}] IF search: {1}" -f $correl, $ifNorm) -ForegroundColor White

    try {
        # 1) find IF and open its detail
        Send-CtrlF
        Paste-Replace $ifNorm
        Send-Enter
        Send-Key '{ESC}' 300
        Send-ShiftTab 1
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
        $found = $false
        if ($null -ne $new) {
            $name = Move-LogToWork $new $correl
            Write-Host ("  moved: {0} -> log\{1}" -f $new.Name, $name) -ForegroundColor Green
            Complete-Row $row $correl $jobName ("moved {0}" -f $name)
            $found = $true; $cntDone++
        } else {
            $existing = Get-ExistingCorrelLog $correl
            if ($null -ne $existing) {
                Write-Host ("  already in log\: {0}" -f $existing.Name) -ForegroundColor DarkGray
                Complete-Row $row $correl $jobName ("already present {0}" -f $existing.Name)
                $found = $true; $cntDone++
            }
        }

        if ($found) {
            # success: return to the list page for the next row
            Send-Tab 2
            Send-Enter
            Start-Sleep -Milliseconds 800
            continue
        }

        # 4) failure handling -- do NOT blindly continue
        Write-Host ("  [FAIL] no new log for {0}" -f $correl) -ForegroundColor Red
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -CorrelIdS $correl -JobName $jobName `
            -Action 'download' -Status 'fail' -Message 'no new .log and none in work\log'

        if (-not $interactive) {
            Write-Host '  [STOP] non-interactive: stopping run to avoid blind navigation.' -ForegroundColor Red
            $cntFail++; $stopRun = $true
            continue
        }

        $resolved = $false
        while (-not $resolved) {
            Write-Host ''
            Write-Host ("  Log not found for {0}." -f $correl) -ForegroundColor Yellow
            Write-Host ("  Download it (or drop {0}_*.log into {1}), re-show the GoAnywhere LIST page," -f $correl, $logDir) -ForegroundColor Yellow
            Write-Host '  then: [Enter]=retry, s=skip(mark fail), q=quit' -ForegroundColor Magenta
            $resp = Read-Host
            if ($resp -eq 'q') { $stopRun = $true; break }
            if ($resp -eq 's') {
                Write-Host ("  [SKIP] {0} left as not-done" -f $correl) -ForegroundColor DarkGray
                Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -CorrelIdS $correl -JobName $jobName -Action 'manual' -Status 'skip' -Message 'operator skipped'
                $cntFail++; break
            }
            # retry: re-check Downloads (since the original snapshot) and work\log
            $retryNew = Get-NewDownloadedLog $beforeTime $before
            if ($null -ne $retryNew) {
                $name = Move-LogToWork $retryNew $correl
                Write-Host ("  moved: {0} -> log\{1}" -f $retryNew.Name, $name) -ForegroundColor Green
                Complete-Row $row $correl $jobName ("manual moved {0}" -f $name)
                $cntDone++; $resolved = $true
                Write-Host '  Log saved. Return Edge to the GoAnywhere LIST page, then press Enter.' -ForegroundColor Magenta
                if ((Read-Host) -eq 'q') { $stopRun = $true }
                continue
            }
            $retryExisting = Get-ExistingCorrelLog $correl
            if ($null -ne $retryExisting) {
                Write-Host ("  found in log\: {0}" -f $retryExisting.Name) -ForegroundColor Green
                Complete-Row $row $correl $jobName ("manual present {0}" -f $retryExisting.Name)
                $cntDone++; $resolved = $true
                Write-Host '  Log saved. Return Edge to the GoAnywhere LIST page, then press Enter.' -ForegroundColor Magenta
                if ((Read-Host) -eq 'q') { $stopRun = $true }
                continue
            }
            Write-Host '  still not found.' -ForegroundColor Yellow
        }
    } catch {
        Write-Host ("  [FAIL] {0}" -f $_.Exception.Message) -ForegroundColor Red
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'GfixLogDownload' -CorrelIdS $correl -JobName $jobName -Action 'exception' -Status 'fail' -Message $_.Exception.Message
        $cntFail++
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

Write-Host ''
Write-Host '===== GfixLogDownload Done =====' -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
if ($stopRun) { Write-Host '  (run stopped early)' -ForegroundColor Yellow }
