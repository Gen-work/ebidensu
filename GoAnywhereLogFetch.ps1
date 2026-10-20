#Requires -Version 5.1
# ============================================================
#  GoAnywhereLogFetch.ps1   (STANDALONE, not a phase; has param():
#  call via -File or &, never dot-source)
#  UTF-8, NO BOM, ASCII source. The one Japanese runtime string (the
#  GoAnywhere "download" link text) is built via [char].
#
#  One-off: fetch the GoAnywhere receive-job log of every correl listed in
#  a local config JSON and copy each one into its receive system's LOG
#  folder (destRoot\<systems.<recv>.dest>\). See docs/GoAnywhereLogFetch.md.
#
#    1. The operator filters the GoAnywhere completed-jobs LIST page
#       (date range, 100 rows), one filter at a time (e.g. LPS, ALC, CRY).
#       Per filter the script Ctrl+A/Ctrl+C's the page and merges the rows
#       (job number = identity).
#    2. Candidates per correl = Receive rows in its goFolder whose project
#       name carries its IF number (VER1 / DEV1 / sender prefix ignored).
#    3. The candidates ON THAT PAGE are downloaded before the next filter
#       (Ctrl+F only finds a job number on the visible page) into
#       <config folder>\goanywhere_logs\ (Ctrl+F job number -> detail ->
#       download link; same keystrokes as GfixLogDownload.ps1). Each job
#       once; a staged job is not downloaded again, so a rerun resumes.
#    4. Ownership comes from the log CONTENT: '/gfix/recv/<correl, last
#       char free>'. Several logs naming one correl (reruns) -> the newest
#       is copied, the older ones are listed. No log naming it -> the
#       candidates are listed and the operator may pick one by number
#       (status MANUAL in the report) or leave it NG.
#    5. Copy (never move, never overwrite without -Force) to the share,
#       then write goanywhere_logs\report.csv.
#
#  -ListTextFile <txt>: use saved list-page text instead of capturing.
#  -DryRun: print the plan and stop (no browser, no copy).
# ============================================================

param(
    [string]$ConfigPath = '',
    [string[]]$ListTextFile = @(),
    [switch]$DryRun,
    [switch]$Force,
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

$dryRunFlag = [bool]$DryRun.IsPresent
$forceFlag  = [bool]$Force.IsPresent

$scriptDir = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'GoAnywhereLogFetchLib.ps1')

if ([string]::IsNullOrWhiteSpace($ConfigPath)) { $ConfigPath = Read-Host 'Config JSON path' }
$ConfigPath = $ConfigPath.Trim('"')
if (-not (Test-Path -LiteralPath $ConfigPath)) { Write-Host "[ERROR] config not found: $ConfigPath" -ForegroundColor Red; exit 1 }
$cfg = Read-GoAnywhereLogFetchConfig $ConfigPath
$stageDir = Join-Path (Split-Path -Parent (Resolve-Path -LiteralPath $ConfigPath).Path) 'goanywhere_logs'

Write-Host "`n===== GoAnywhere log fetch =====" -ForegroundColor Green
Write-Host ("Config : {0}" -f $ConfigPath)
Write-Host ("Dest   : {0}" -f $cfg.DestRoot)
Write-Host ("Stage  : {0}" -f $stageDir)
Write-Host ("Targets: {0}" -f @($cfg.Targets).Count)
foreach ($t in $cfg.Targets) {
    $known = $cfg.Systems.ContainsKey($t.Recv)
    $color = $(if ($known) { 'Gray' } else { 'Red' })
    $destSub = $(if ($known) { $cfg.Systems[$t.Recv].dest } else { '(unknown recv!)' })
    Write-Host ("  {0,-10} IF{1}  {2,-5} -> {3}" -f $t.Correl, $t.IfNo, $t.Recv, $destSub) -ForegroundColor $color
}

if (-not (Test-Path -LiteralPath $stageDir) -and -not $dryRunFlag) { New-Item -ItemType Directory -Path $stageDir | Out-Null }
function Get-StagedLog([string]$jobNo) {
    if (-not (Test-Path -LiteralPath $stageDir)) { return @() }
    return @(Get-ChildItem -LiteralPath $stageDir -Filter ("*{0}*.log" -f $jobNo) -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1)
}
function Write-FetchPlan($plan) {
    foreach ($it in $plan.Items) {
        if ($it.Candidates.Count -eq 0) {
            Write-Host ("  {0,-10} IF{1} {2,-5}: no job yet" -f $it.Correl, $it.IfNo, $it.Recv) -ForegroundColor Yellow
            continue
        }
        Write-Host ("  {0,-10} IF{1} {2,-5}: {3} candidate(s)" -f $it.Correl, $it.IfNo, $it.Recv, $it.Candidates.Count)
        foreach ($c in $it.Candidates) {
            $staged = $(if (@(Get-StagedLog $c.JobNo).Count -gt 0) { 'staged' } else { '' })
            Write-Host ("      {0}  {1}  {2}  {3}  {4}" -f $c.JobNo, $c.StartTime, $c.ProjectName, $c.Folder, $staged) -ForegroundColor DarkGray
        }
    }
}

# Download the given job numbers from the LIST page Edge shows NOW
# (Ctrl+F only finds a job number on the visible page, so this runs once
# per capture, before the operator changes the filter). Returns $false
# when the operator asked to stop.
function Invoke-FetchJobs([string[]]$jobNos) {
    $downloadDir = Join-Path $env:USERPROFILE 'Downloads'
    # GoAnywhere "download" link text: katakana 'daunro-do' (U+30C0 U+30A6 U+30F3 U+30ED U+30FC U+30C9)
    $dlSearch = [char]0x30C0 + [char]0x30A6 + [char]0x30F3 + [char]0x30ED + [char]0x30FC + [char]0x30C9
    function Get-NewDownloadedLog([datetime]$sinceTime, [hashtable]$beforeSet) {
        $hits = @(Get-ChildItem -LiteralPath $downloadDir -Filter '*.log' -ErrorAction SilentlyContinue | Where-Object {
            (-not $beforeSet.ContainsKey($_.Name)) -or ($_.LastWriteTime -gt $sinceTime) -or ($_.CreationTime -gt $sinceTime)
        } | Sort-Object LastWriteTime -Descending)
        if ($hits.Count -gt 0) { return $hits[0] }
        return $null
    }
    function Move-ToStage($file, [string]$jobNo) {
        $name = $file.Name
        if ($name -notlike ("*{0}*" -f $jobNo)) { $name = "{0}_{1}" -f $jobNo, $name }
        Move-Item -LiteralPath $file.FullName -Destination (Join-Path $stageDir $name) -Force
        return $name
    }
    Switch-ToEdge
    Start-Sleep -Milliseconds 300
    foreach ($jobNo in $jobNos) {
        Write-Host ("[JobNo {0}] opening detail..." -f $jobNo) -ForegroundColor White
        try {
            Send-CtrlF
            Paste-Replace $jobNo
            Send-Enter
            Send-Key '{ESC}' 300
            Send-Enter
            Start-Sleep -Milliseconds 1200

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

            $new = Get-NewDownloadedLog $beforeTime $before
            if ($null -ne $new) {
                Write-Host ("  staged: {0}" -f (Move-ToStage $new $jobNo)) -ForegroundColor Green
                # back to the LIST page for the next job
                Send-Tab 2
                Send-Enter
                Start-Sleep -Milliseconds 800
                continue
            }
            Write-Host ("  [FAIL] no new .log in Downloads for job {0}" -f $jobNo) -ForegroundColor Red
            while ($true) {
                Write-Host ("  Download it by hand (or drop *{0}*.log into {1}), put Edge back on the LIST page," -f $jobNo, $stageDir) -ForegroundColor Yellow
                Write-Host '  then: [Enter]=check again, s=skip this job, q=stop' -ForegroundColor Magenta
                $resp = Read-Host
                if ($resp -eq 'q') { return $false }
                if ($resp -eq 's') { break }
                $late = Get-NewDownloadedLog $beforeTime $before
                if ($null -ne $late) { Write-Host ("  staged: {0}" -f (Move-ToStage $late $jobNo)) -ForegroundColor Green }
                if (@(Get-StagedLog $jobNo).Count -gt 0) { break }
                Write-Host '  still not found.' -ForegroundColor Yellow
            }
            Switch-ToEdge
            Start-Sleep -Milliseconds 300
        } catch {
            Write-Host ("  [FAIL] job {0}: {1}" -f $jobNo, $_.Exception.Message) -ForegroundColor Red
            Write-Host '  Fix the page by hand (LIST page), then [Enter]=continue, q=stop' -ForegroundColor Magenta
            if ((Read-Host) -eq 'q') { return $false }
            Switch-ToEdge
            Start-Sleep -Milliseconds 300
        }
    }
    return $true
}

# -- 1+2+3. job list -> plan -> download, once per filter --
$listRows = @()
$stopRun = $false
if (@($ListTextFile).Count -gt 0) {
    foreach ($f in $ListTextFile) {
        $txt = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8)
        $listRows = @(Merge-GoAnywhereRows -Existing $listRows -New @(ConvertFrom-GoAnywhereJobList $txt))
    }
    Write-Host ("`n[list] {0} job row(s) from {1} file(s)." -f $listRows.Count, @($ListTextFile).Count) -ForegroundColor DarkGray
    $plan = Get-GoAnywhereFetchPlan -Targets $cfg.Targets -Rows $listRows -Systems $cfg.Systems
    Write-Host "`n===== Plan =====" -ForegroundColor Green
    Write-FetchPlan $plan
    if ($dryRunFlag) { Write-Host "`n[DryRun] stop here." -ForegroundColor Yellow; exit 0 }
    $unstaged = @($plan.JobNumbers | Where-Object { @(Get-StagedLog $_).Count -eq 0 })
    if ($unstaged.Count -gt 0) {
        Write-Host ("`n[list] {0} job(s) not staged: {1}. Run without -ListTextFile to download them, or drop <jobNo>.log into {2}." -f $unstaged.Count, ($unstaged -join ' '), $stageDir) -ForegroundColor Yellow
    }
} else {
    if ([string]::IsNullOrWhiteSpace($CommonScript)) { $CommonScript = Join-Path $scriptDir 'Common.ps1' }
    . $CommonScript
    $Global:Timing = @{ ActionWaitMs = $ActionWaitMs; ResultWaitMs = $ResultWaitMs }
    $pageTextScript = Join-Path $scriptDir 'Read-PageText.ps1'

    Write-Host ''
    Write-Host 'GoAnywhere completed-jobs LIST page in Edge, one filter at a time' -ForegroundColor Yellow
    Write-Host '(e.g. project name LPS, then ALC, then CRY; date range covering the test days):'
    Write-Host '  *** Rows per page = 100 ***' -ForegroundColor Red
    Write-Host 'Set the FIRST filter, then press Enter here. (q to quit)' -ForegroundColor Magenta
    if ((Read-Host) -eq 'q') { exit 0 }
    $round = 0
    while ($true) {
        $round++
        Switch-ToEdge
        Start-Sleep -Milliseconds 300
        Send-Key '{ESC}' 200
        $txt = [string](& $pageTextScript -SelectWaitMs $ActionWaitMs -CopyWaitMs $ActionWaitMs)
        $pageRows = @(ConvertFrom-GoAnywhereJobList $txt)
        $listRows = @(Merge-GoAnywhereRows -Existing $listRows -New $pageRows)
        $plan = Get-GoAnywhereFetchPlan -Targets $cfg.Targets -Rows $listRows -Systems $cfg.Systems
        $onPage = @{}
        foreach ($r in $pageRows) { $onPage[$r.JobNo] = $true }
        $fetchNow = @($plan.JobNumbers | Where-Object { $onPage.ContainsKey($_) -and @(Get-StagedLog $_).Count -eq 0 })

        Write-Host ("`n===== Filter {0}: {1} row(s) on this page, {2} in total =====" -f $round, $pageRows.Count, $listRows.Count) -ForegroundColor Green
        Write-FetchPlan $plan
        if ($pageRows.Count -eq 0) { Write-Host '[WARN] 0 job rows parsed -- is Edge on the LIST page?' -ForegroundColor Yellow }
        Write-Host ("To download from THIS page: {0}" -f $fetchNow.Count) -ForegroundColor Cyan

        if (-not $dryRunFlag -and $fetchNow.Count -gt 0) {
            Write-Host 'Enter = download them now,  s = skip downloading this page,  q = quit' -ForegroundColor Magenta
            $ans = Read-Host
            if ($ans -eq 'q') { exit 0 }
            if ($ans -ne 's') {
                if (-not (Invoke-FetchJobs $fetchNow)) { $stopRun = $true }
                $consoleHwnd = (Get-Process -Id $PID).MainWindowHandle
                if ($consoleHwnd -ne [IntPtr]::Zero) { [WinAPI]::SetForegroundWindow($consoleHwnd) | Out-Null }
            }
        }
        if ($stopRun) { break }
        $left = @($plan.Missing).Count
        Write-Host ("Targets still without a job: {0}" -f $left) -ForegroundColor $(if ($left -gt 0) { 'Yellow' } else { 'Green' })
        Write-Host 'c = next filter (set it in Edge first, then c),  Enter = finish (match + copy),  q = quit' -ForegroundColor Magenta
        $ans = Read-Host
        if ($ans -eq 'q') { exit 0 }
        if ($ans -ne 'c') { break }
    }
    if ($dryRunFlag) { Write-Host "`n[DryRun] stop here." -ForegroundColor Yellow; exit 0 }
}

# -- 4. match by content + 5. copy --
Write-Host "`n===== Match + copy =====" -ForegroundColor Green
$report = [System.Collections.Generic.List[object]]::new()
foreach ($it in $plan.Items) {
    $status = ''; $jobNo = ''; $start = ''; $dest = ''; $note = ''
    if (-not $it.KnownRecv) {
        $status = 'NG'; $note = ("recv '{0}' is not in config systems" -f $it.Recv)
    } elseif ($it.Candidates.Count -eq 0) {
        $status = 'NG'; $note = 'no job with this IF number in the captured list'
    } else {
        $logs = [System.Collections.Generic.List[object]]::new()
        foreach ($c in $it.Candidates) {
            $f = @(Get-StagedLog $c.JobNo)
            if ($f.Count -eq 0) { continue }
            $logs.Add([pscustomobject]@{ JobNo = $c.JobNo; StartTime = $c.StartTime; File = $f[0].FullName; Name = $f[0].Name
                Text = [IO.File]::ReadAllText($f[0].FullName, [Text.Encoding]::UTF8) })
        }
        $sel = Select-GoAnywhereLogForCorrel -Correl $it.Correl -Logs $logs.ToArray()
        $pick = $null; $picked = 'OK'
        if ($sel.Status -eq 'nolog') {
            $status = 'NG'; $note = 'no candidate log downloaded'
        } elseif ($sel.Status -eq 'nocontent') {
            # the content does not name the correl: the operator may still
            # point at one (the receive command can be spelled differently)
            $note = ("{0} candidate log(s), none names /gfix/recv/{1}?" -f $logs.Count, $it.Correl.Substring(0, [Math]::Max(0, $it.Correl.Length - 1)))
            Write-Host ("  {0} (IF{1} {2}): {3}" -f $it.Correl, $it.IfNo, $it.Recv, $note) -ForegroundColor Yellow
            for ($i = 0; $i -lt $logs.Count; $i++) {
                Write-Host ("    {0}) {1}  {2}  {3}" -f ($i + 1), $logs[$i].JobNo, $logs[$i].StartTime, $logs[$i].File) -ForegroundColor DarkGray
            }
            Write-Host '    number = copy that log for this correl anyway,  Enter = leave it NG' -ForegroundColor Magenta
            $ans = [string](Read-Host)
            $n = 0
            if ([int]::TryParse($ans.Trim(), [ref]$n) -and $n -ge 1 -and $n -le $logs.Count) {
                $pick = $logs[$n - 1]; $picked = 'MANUAL'; $note = 'picked by hand (content did not name the correl)'
            } else {
                $status = 'NG'; $note = $note + ' : ' + (($logs | ForEach-Object { $_.JobNo }) -join ' ')
            }
        } else {
            $pick = $sel.Chosen
            if (@($sel.Others).Count -gt 0) { $note = ('older run(s) not copied: ' + ((@($sel.Others) | ForEach-Object { "{0} ({1})" -f $_.JobNo, $_.StartTime }) -join ', ')) }
        }
        if ($null -ne $pick) {
            $jobNo = $pick.JobNo; $start = $pick.StartTime
            $destDir = Join-Path $cfg.DestRoot $cfg.Systems[$it.Recv].dest
            $name = Format-GoAnywhereLogName -Format $cfg.NameFormat -Correl $it.Correl -JobNo $jobNo -StartTime $start -OrigName $pick.Name
            $dest = Join-Path $destDir $name
            if (-not (Test-Path -LiteralPath $destDir)) {
                $status = 'NG'; $note = ("destination folder not found: {0}" -f $destDir)
            } elseif ((Test-Path -LiteralPath $dest) -and -not $forceFlag) {
                $status = 'EXISTS'; $note = (@('already there (use -Force to overwrite)', $note) | Where-Object { $_ }) -join '; '
            } else {
                Copy-Item -LiteralPath $pick.File -Destination $dest -Force
                $status = $picked
            }
        }
    }
    $color = switch ($status) { 'OK' { 'Green' } 'MANUAL' { 'Green' } 'EXISTS' { 'DarkYellow' } default { 'Red' } }
    Write-Host ("  [{0,-6}] {1,-10} IF{2} {3,-5} {4} {5}" -f $status, $it.Correl, $it.IfNo, $it.Recv, $jobNo, $dest) -ForegroundColor $color
    if ($note) { Write-Host ("           {0}" -f $note) -ForegroundColor DarkGray }
    $report.Add([pscustomobject]@{ Correl = $it.Correl; IF = $it.IfNo; Recv = $it.Recv; Status = $status; JobNo = $jobNo; StartTime = $start; Dest = $dest; Note = $note })
}

$reportPath = Join-Path $stageDir 'report.csv'
$report | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
$ok = @($report | Where-Object { $_.Status -eq 'OK' -or $_.Status -eq 'MANUAL' }).Count
$ex = @($report | Where-Object { $_.Status -eq 'EXISTS' }).Count
$ng = @($report | Where-Object { $_.Status -eq 'NG' }).Count
Write-Host ''
Write-Host ("Done: OK {0} / already there {1} / NG {2}   report: {3}" -f $ok, $ex, $ng, $reportPath) -ForegroundColor $(if ($ng -gt 0) { 'Yellow' } else { 'Green' })
if ($stopRun) { Write-Host '(download stopped early -- rerun to resume; staged jobs are not fetched again)' -ForegroundColor Yellow }
