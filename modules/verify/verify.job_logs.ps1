# modules/verify/verify.job_logs.ps1
# GoAnywhere job logs -> the runs of ONE transfer (kadou.collect, 10/26
# JJODWDBx operation check). The Receive project's log names the file it
# moved ("Overriding variable 'FileName' with value '<correl>.YYMMDD.HHMMSScc'"),
# so the logs of a day's rounds are told apart by content, not by the job
# number or the list row. Every run is judged on the four lines the
# reference evidence shows: uploaded to <recvDir>/<file>, the follow-up
# shell's return code 0, the GoAnywhere copy deleted, End Date and Time.
# Read-only: it never touches the files it reads.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')

$Manifest = @{
  id         = 'verify.job_logs'
  group      = 'verify'
  summary    = 'Pick the GoAnywhere job logs of one transfer file and judge each run (upload, return code 0, delete, end)'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    paths      = @{ type='list';   required=$true; desc='job log files (browser.download_each paths); logs of other transfers are passed over' }
    key        = @{ type='string'; required=$true; desc='the transfer file name or its prefix: a log is this transfer''s when its FileName equals key or starts with key + "."' }
    recvDir    = @{ type='string'; default=''; desc='the directory the file must be uploaded to (/appl/IDS/IDSVer1/gfix/recv); empty = any' }
    shell      = @{ type='string'; default=''; desc='the follow-up shell the Command line must run (IDSLB283run.sh); empty = any' }
    ss         = @{ type='string'; default=''; desc='the last argument of the Command line (U / S); empty = any' }
    day        = @{ type='string'; default=''; desc='yyyy-MM-dd: only runs that started that day count; empty = every day' }
    preferText = @{ type='string'; default=''; desc='a page text (the Jenkins receive folder): the chosen run is the earliest ok run whose file name appears in it; empty = the earliest ok run' }
    encoding   = @{ type='string'; default='mixed'; enum=@('mixed', 'utf8', 'cp932') }
  }
  outputs    = @{
    runs       = @{ type='list';   desc='@{ path; jobNumber; project; fileName; start; end; bytes; uploadedTo; command; returnCode; deleted; code; reasons } per run, oldest first' }
    runCount   = @{ type='int';    desc='runs of this transfer (that day)' }
    okCount    = @{ type='int' }
    ngCount    = @{ type='int' }
    code       = @{ type='string'; desc='ok = at least one run and every run ok; ng = a run failed a check; unknown = no run' }
    reason     = @{ type='string'; desc='one line for people: what was found, what failed' }
    hasChosen  = @{ type='bool';   desc='a run was chosen for the evidence' }
    chosenPath = @{ type='path';   desc='the chosen run''s log file; empty when none' }
    chosenFile = @{ type='string'; desc='the chosen run''s transfer file name; empty when none' }
    chosenStart = @{ type='string'; desc='the chosen run''s start (yyyy-MM-dd HH:mm:ss)' }
    listed     = @{ type='bool';   desc='with preferText: the chosen run''s file is in that text' }
    summary    = @{ type='list';   desc='one line per run: "HH:mm:ss  <file>  ok" / "... ng: <why>" (for a gate)' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $true }
  )
  example    = @{ use = 'verify.job_logs'; with = @{ paths = @('log/GoAnywhere/jobs/1000003451754.log'); key = 'JIDSU86S'; recvDir = '/appl/IDS/IDSVer1/gfix/recv'; shell = 'IDSLB283run.sh'; ss = 'U'; day = '{{run.date}}' } }
  notes      = 'A verdict is never a failure: ng and unknown are outputs. file_not_found only when a listed path is missing (a download that did not land). The return code line is GoAnywhere''s "The Command executed successfully with return code ''0''"; a log without it (the shell never ran) is ng.'
}

function VerifyJobLogs-Parse {
    <#
      PURE. One job log's text -> @{ jobNumber; project; fileName; start;
      end; bytes; uploadedTo; command; returnCode; deleted; errors }.
      Missing pieces are '' (deleted: $false).
    #>
    param([string]$Text)
    $r = @{ jobNumber = ''; project = ''; fileName = ''; start = ''; end = ''; bytes = ''; uploadedTo = ''; command = ''; returnCode = ''; deleted = $false; errors = @() }
    $errors = New-Object System.Collections.ArrayList
    foreach ($line in @(Get-EbiTextLines -Text $Text)) {
        $l = [string]$line
        $m = [regex]::Match($l, 'Start Date and Time:\s*(\d{4}-\d{2}-\d{2} \d{1,2}:\d{2}:\d{2})')
        if ($m.Success -and $r['start'] -eq '') { $r['start'] = $m.Groups[1].Value; continue }
        $m = [regex]::Match($l, 'End Date and Time:\s*(\d{4}-\d{2}-\d{2} \d{1,2}:\d{2}:\d{2})')
        if ($m.Success) { $r['end'] = $m.Groups[1].Value; continue }
        $m = [regex]::Match($l, 'Job Number:\s*(\d+)')
        if ($m.Success -and $r['jobNumber'] -eq '') { $r['jobNumber'] = $m.Groups[1].Value; continue }
        $m = [regex]::Match($l, 'Project Name:\s*(\S+)')
        if ($m.Success -and $r['project'] -eq '') { $r['project'] = $m.Groups[1].Value; continue }
        $m = [regex]::Match($l, "Overriding variable 'FileName' with value '([^']*)'")
        if ($m.Success) { $r['fileName'] = $m.Groups[1].Value.Trim(); continue }
        $m = [regex]::Match($l, "Uploading '[^']*' \(([\d,]+) bytes\)")
        if ($m.Success -and $r['bytes'] -eq '') { $r['bytes'] = $m.Groups[1].Value.Replace(',', '') }
        $m = [regex]::Match($l, "File '[^']*' successfully uploaded to '([^']*)'")
        if ($m.Success) { $r['uploadedTo'] = $m.Groups[1].Value.Trim(); continue }
        $m = [regex]::Match($l, "Command:\s*'([^']*)'")
        if ($m.Success) { $r['command'] = ([regex]::Replace($m.Groups[1].Value, '\s+', ' ')).Trim(); continue }
        $m = [regex]::Match($l, "executed successfully with return code '(-?\d+)'")
        if ($m.Success) { $r['returnCode'] = $m.Groups[1].Value; continue }
        $m = [regex]::Match($l, "return code '(-?\d+)'")
        if ($m.Success -and $r['returnCode'] -eq '') { $r['returnCode'] = $m.Groups[1].Value }
        if ($l -match "\) was deleted\.") { $r['deleted'] = $true; continue }
        if ($l -match '^\S+ \S+\s+(ERROR|SEVERE|FATAL)\s') { [void]$errors.Add($l.Trim()) }
    }
    $r['errors'] = $errors.ToArray()
    return $r
}

function VerifyJobLogs-IsKey {
    # PURE. Is this FileName the transfer named by key (equal, or key + ".")?
    param([string]$FileName, [string]$Key)
    $f = $FileName.Trim().ToUpperInvariant(); $k = $Key.Trim().ToUpperInvariant()
    if ($k -eq '' -or $f -eq '') { return $false }
    return ($f -eq $k -or $f.StartsWith($k + '.'))
}

function VerifyJobLogs-Judge {
    <#
      PURE. A parsed log -> @{ code; reasons } against the expected receive
      directory / shell / SS. Every check that fails adds one reason.
    #>
    param($Run, [string]$RecvDir = '', [string]$Shell = '', [string]$Ss = '')
    $why = New-Object System.Collections.ArrayList
    $file = [string]$Run['fileName']
    $to = [string]$Run['uploadedTo']
    if ($to -eq '') { [void]$why.Add('no "successfully uploaded" line') }
    elseif ($RecvDir -ne '') {
        $want = $RecvDir.TrimEnd('/') + '/' + $file
        if ($to -ne $want) { [void]$why.Add(('uploaded to ' + $to + ', not ' + $want)) }
    }
    $cmd = [string]$Run['command']
    if ($cmd -eq '') { [void]$why.Add('no Command line (the follow-up shell did not run)') }
    else {
        $parts = @($cmd -split ' ' | Where-Object { $_ -ne '' })
        if ($Shell -ne '' -and -not ([string]$parts[0]).EndsWith('/' + $Shell) -and [string]$parts[0] -ne $Shell) { [void]$why.Add(('the Command runs ' + $parts[0] + ', not ' + $Shell)) }
        if ($parts.Count -lt 2 -or -not ([string]$parts[1]).EndsWith($file)) { [void]$why.Add(('the Command is not given ' + $file)) }
        if ($Ss -ne '' -and ($parts.Count -lt 3 -or [string]$parts[$parts.Count - 1] -ne $Ss)) { [void]$why.Add(('the Command ends in "' + $(if ($parts.Count -ge 3) { $parts[$parts.Count - 1] } else { '' }) + '", not ' + $Ss)) }
    }
    $rc = [string]$Run['returnCode']
    if ($rc -eq '') { [void]$why.Add('no return code line') } elseif ($rc -ne '0') { [void]$why.Add(('return code ' + $rc)) }
    if (-not [bool]$Run['deleted']) { [void]$why.Add('no "was deleted" line') }
    if ([string]$Run['end'] -eq '') { [void]$why.Add('no End Date and Time (the job did not finish)') }
    foreach ($e in @($Run['errors'])) { [void]$why.Add(('error line: ' + $e)) }
    return @{ code = $(if ($why.Count -eq 0) { 'ok' } else { 'ng' }); reasons = $why.ToArray() }
}

function VerifyJobLogs-Select {
    <#
      PURE. Parsed + judged runs (each with path) -> the outputs. Runs of
      other transfers and other days are dropped; the rest sorted by start.
    #>
    param($Runs, [string]$Key, [string]$Day = '', [string]$PreferText = '')
    $mine = @(@($Runs) | Where-Object { (VerifyJobLogs-IsKey -FileName ([string]$_['fileName']) -Key $Key) -and ($Day -eq '' -or ([string]$_['start']).StartsWith($Day)) } | Sort-Object { [string]$_['start'] }, { [string]$_['fileName'] })
    $ok = @($mine | Where-Object { $_['code'] -eq 'ok' })
    $ng = @($mine | Where-Object { $_['code'] -ne 'ok' })
    $chosen = $null; $listed = $false
    if ($PreferText -ne '') {
        foreach ($r in $ok) { if ($PreferText.Contains([string]$r['fileName'])) { $chosen = $r; $listed = $true; break } }
    }
    if ($null -eq $chosen -and $ok.Count -gt 0) { $chosen = $ok[0] }
    $summary = @(foreach ($r in $mine) {
        $t = ([string]$r['start']); if ($t.Length -ge 19) { $t = $t.Substring(11) }
        $t + '  ' + [string]$r['fileName'] + '  ' + $(if ($r['code'] -eq 'ok') { 'ok' } else { 'ng: ' + (@($r['reasons']) -join '; ') })
    })
    $code = if ($mine.Count -eq 0) { 'unknown' } elseif ($ng.Count -gt 0) { 'ng' } else { 'ok' }
    $reason = if ($mine.Count -eq 0) { ('no job log of ' + $Key + $(if ($Day -ne '') { ' on ' + $Day } else { '' }) + ' among the logs read') }
              elseif ($ng.Count -gt 0) { ('{0} of {1} run(s) of {2} failed a check: {3}' -f $ng.Count, $mine.Count, $Key, ((@($ng | ForEach-Object { [string]$_['fileName'] + ' (' + (@($_['reasons']) -join '; ') + ')' })) -join ' | ')) }
              else { ('{0} run(s) of {1}, all ok (upload, return code 0, delete, end)' -f $mine.Count, $Key) }
    if ($null -ne $chosen -and $PreferText -ne '' -and -not $listed) { $reason += ' -- the chosen run''s file is NOT in the receive folder listing' }
    $out = @(foreach ($r in $mine) { @{ path = [string]$r['path']; jobNumber = [string]$r['jobNumber']; project = [string]$r['project']; fileName = [string]$r['fileName']; start = [string]$r['start']; end = [string]$r['end']; bytes = [string]$r['bytes']; uploadedTo = [string]$r['uploadedTo']; command = [string]$r['command']; returnCode = [string]$r['returnCode']; deleted = [bool]$r['deleted']; code = [string]$r['code']; reasons = @($r['reasons']) } })
    return @{
        runs = $out; runCount = $mine.Count; okCount = $ok.Count; ngCount = $ng.Count; code = $code; reason = $reason
        hasChosen = ($null -ne $chosen)
        chosenPath = $(if ($null -ne $chosen) { [string]$chosen['path'] } else { '' })
        chosenFile = $(if ($null -ne $chosen) { [string]$chosen['fileName'] } else { '' })
        chosenStart = $(if ($null -ne $chosen) { [string]$chosen['start'] } else { '' })
        listed = $listed; summary = $summary
    }
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $runs = New-Object System.Collections.ArrayList
    foreach ($p in @($In['paths'])) {
        if ([string]::IsNullOrWhiteSpace([string]$p)) { continue }
        $full = Resolve-EbiWorkPath -PathValue ([string]$p) -WorkDir $work
        if (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
            if ($Ctx['DryRun']) { $Ctx.Log.Info(('would read {0}' -f $full)); continue }
            $r = VerifyJobLogs-Select -Runs @() -Key ([string]$In['key'])
            $r['ok'] = $false; $r['failure'] = 'file_not_found'; $r['message'] = $full
            return $r
        }
        $bytes = [System.IO.File]::ReadAllBytes($full)
        $text = switch ([string]$In['encoding']) {
            'cp932' { (Get-EbiCodePage -CodePage 932).GetString($bytes) }
            'utf8'  { (New-Object System.Text.UTF8Encoding($false)).GetString($bytes) }
            default { (ConvertFrom-EbiMixedBytes -Bytes $bytes)['text'] }
        }
        $run = VerifyJobLogs-Parse -Text $text
        $run['path'] = $full
        $j = VerifyJobLogs-Judge -Run $run -RecvDir ([string]$In['recvDir']) -Shell ([string]$In['shell']) -Ss ([string]$In['ss'])
        $run['code'] = $j['code']; $run['reasons'] = $j['reasons']
        [void]$runs.Add($run)
    }
    $r = VerifyJobLogs-Select -Runs $runs.ToArray() -Key ([string]$In['key']) -Day ([string]$In['day']) -PreferText ([string]$In['preferText'])
    foreach ($l in @($r['summary'])) { $Ctx.Log.Info($l) }
    $r['ok'] = $true
    return $r
}
