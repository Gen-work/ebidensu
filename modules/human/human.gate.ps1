# modules/human/human.gate.ps1
# The verdict gate (P1-34, WORKFLOW-SCHEMA 5.2, P0-R13). ALWAYS runs: when
# the code is not in askWhen it passes the code through untouched with
# action=pass and no panel; when it is, the panel asks the operator and the
# answer becomes the code flow.checkpoint writes. `s` leaves the item
# pending (code '' + action=skip; the checkpoint's `when` skips it); `q`
# is operator_quit, which the runner turns into cancelled.

. (Join-Path $PSScriptRoot '..\..\kernel\Gate.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')

$Manifest = @{
  id         = 'human.gate'
  group      = 'human'
  summary    = 'Ask the operator to confirm a verdict when it is in askWhen; pass otherwise'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    code     = @{ type='string'; required=$true; desc='the verdict so far: ok | ng | unknown | empty' }
    askWhen  = @{ type='list';   default=@('unknown'); desc='codes that need a person' }
    reason   = @{ type='string'; default=''; desc='why (verify.assert''s reason), shown on the panel' }
    evidence = @{ type='any';    default=''; desc='a path, or a list of paths / values, to look at' }
    key      = @{ type='string'; default=''; desc='which item; empty = the current item' }
    checklist = @{ type='list';  default=@(); desc='what to look at, one line each, shown as "check: ..." (10/15: the review points, kept in the profile so they can grow)' }
    alert     = @{ type='any';   default=''; desc='a line, or a list of lines, drawn RED on the panel (an automatic check that disagrees, e.g. verify.find_token alert); empty lines are left out' }
    front     = @{ type='list';  default=@(); desc='windows to bring to the front once the panel is drawn: each entry a piece of a window title, or a process name (EXCEL, OUTLOOK); every match comes forward (10/19: all the review-mail drafts). The answer is still typed in the console' }
    skipFails = @{ type='bool';  default=$false; desc='true: the operator''s s answer is the failure operator_skip (onError.byFailure.operator_skip: skip -> the whole item is left pending), not just an empty verdict the next steps run past. For a gate whose "no" makes the rest of the item pointless (a proposed start time)' }
    enterOk   = @{ type='bool';  default=$true; desc='false: Enter alone is not "ok" -- type o (10/23: line counts differed, Enter took another job''s file)' }
    retime    = @{ type='bool';  default=$false; desc='offer t 10:15 = the start time was wrong: action retime + clock (a flow.redo after the gate goes back to the start)' }
    recheck   = @{ type='bool';  default=$false; desc='offer r = a file was put right by hand: action recheck (a flow.redo after the gate goes back to the listing)' }
  }
  outputs    = @{
    code   = @{ type='string'; desc='ok | ng | unknown | empty -- what flow.checkpoint should write' }
    action = @{ type='string'; desc='pass | ok | ng | keep | skip | retime | recheck' }
    note   = @{ type='string'; desc='free text typed with m' }
    clock  = @{ type='string'; desc='retime: the start time typed after t (HH:mm:ss); else empty' }
  }
  failures   = @(
    @{ id = 'operator_quit'; transient = $false }
    @{ id = 'operator_skip'; transient = $false }
  )
  example    = @{ use = 'human.gate'; with = @{ code = '{{steps.verdict.out.code}}'; askWhen = @('unknown'); reason = '{{steps.verdict.out.reason}}'; evidence = '{{steps.shot.out.path}}' } }
  notes      = 'Under DryRun or without a console the panel is printed and the verdict is kept as is (action=keep). onError.policy=ask reuses this panel''s rendering (kernel/Gate.ps1), so the two never drift.'
}

function HumanGate-Actions {
    param([bool]$EnterOk = $true, [bool]$Retime = $false, [bool]$Recheck = $false)
    $a = New-Object System.Collections.ArrayList
    [void]$a.Add(@{ key = 'o'; label = $(if ($EnterOk) { 'ok (Enter)' } else { 'ok (type o)' }) })
    [void]$a.Add(@{ key = 'n'; label = 'ng' })
    if ($Retime) { [void]$a.Add(@{ key = 't'; label = 't 10:15 = other start time'; text = $true }) }
    if ($Recheck) { [void]$a.Add(@{ key = 'r'; label = 'r = files fixed, check again' }) }
    [void]$a.Add(@{ key = 'k'; label = 'keep as is' })
    [void]$a.Add(@{ key = 's'; label = 'skip (leave pending)' })
    [void]$a.Add(@{ key = 'm'; label = 'm <note>' })
    [void]$a.Add(@{ key = 'q'; label = 'quit' })
    return $a.ToArray()
}

function HumanGate-Clock {
    # PURE. What was typed after t -> 'HH:mm:ss', or '' when it is not a
    # time of day (10:15, 9:05, 10:15:45; 24h).
    param([string]$Text)
    $m = [regex]::Match(([string]$Text).Trim(), '^(\d{1,2}):(\d{2})(?::(\d{2}))?$')
    if (-not $m.Success) { return '' }
    $h = [int]$m.Groups[1].Value; $mi = [int]$m.Groups[2].Value; $se = if ($m.Groups[3].Success) { [int]$m.Groups[3].Value } else { 0 }
    if ($h -gt 23 -or $mi -gt 59 -or $se -gt 59) { return '' }
    return ('{0:00}:{1:00}:{2:00}' -f $h, $mi, $se)
}

function HumanGate-Decide {
    # PURE. An answer -> @{ code; action }. The gate's own condition and
    # the answer mapping live here so a test can drive them without a console.
    param([string]$Code, $AskWhen, [string]$Answer)
    $ask = @(@($AskWhen) | ForEach-Object { [string]$_ })
    if (-not ($ask -contains $Code)) { return @{ code = $Code; action = 'pass' } }
    switch ($Answer) {
        'o' { return @{ code = 'ok'; action = 'ok' } }
        'n' { return @{ code = 'ng'; action = 'ng' } }
        's' { return @{ code = ''; action = 'skip' } }
        'k' { return @{ code = $Code; action = 'keep' } }
        default { return @{ code = $Code; action = 'keep' } }
    }
}

function HumanGate-WhatLines {
    # PURE. The WHAT HAPPENED lines and the red fragments: "!! <alert>"
    # first (what an automatic check disagrees with), then the verdict, the
    # reason, and "check: <point>" per checklist line. Alert fragments are
    # the wrapped pieces of each alert line, so every panel row of a long
    # alert is drawn red, not only its first.
    param([string]$Code, [string]$Reason = '', $Checklist = @(), $Alert = '')
    $what = New-Object System.Collections.ArrayList
    $frags = New-Object System.Collections.ArrayList
    foreach ($a in @($Alert)) {
        if ($null -eq $a -or ([string]$a).Trim() -eq '') { continue }
        $line = '!! ' + ([string]$a).Trim()
        [void]$what.Add($line)
        foreach ($w in @(ConvertTo-EbiGateWrapped -Text $line -Width ((Get-EbiGateWidth) - 6))) { if ($w.Trim() -ne '') { [void]$frags.Add($w.Trim()) } }
    }
    [void]$what.Add(('verdict so far: ' + $(if ($Code -eq '') { '(empty)' } else { $Code })))
    if ($Reason -ne '') { [void]$what.Add($Reason) }
    foreach ($c in @($Checklist)) { if ($null -ne $c -and ([string]$c).Trim() -ne '') { [void]$what.Add('check: ' + ([string]$c).Trim()) } }
    return @{ what = $what.ToArray(); alert = $frags.ToArray() }
}

function Invoke-Step {
    param($In, $Ctx)
    $code = [string]$In['code']
    $askWhen = @($In['askWhen'])
    $first = HumanGate-Decide -Code $code -AskWhen $askWhen -Answer ''
    if ($first['action'] -eq 'pass') { return @{ ok = $true; code = $code; action = 'pass'; note = ''; clock = '' } }
    $enterOk = [bool]$In['enterOk']; $retime = [bool]$In['retime']; $recheck = [bool]$In['recheck']
    $key = [string]$In['key']
    if ($key -eq '' -and $null -ne $Ctx['Item']) { $key = Get-EbiKeyDisplay -Item $Ctx['Item'] -KeyColumns @($Ctx['KeyColumns']) }
    $ev = New-Object System.Collections.ArrayList
    foreach ($e in @($In['evidence'])) { if ($null -ne $e -and [string]$e -ne '') { [void]$ev.Add([string]$e) } }
    $panel = HumanGate-WhatLines -Code $code -Reason ([string]$In['reason']) -Checklist $In['checklist'] -Alert $In['alert']
    $what = $panel['what']
    $front = @()
    $nx = New-Object System.Collections.ArrayList
    [void]$nx.Add($(if ($enterOk) { 'Enter/o: it is ok' } else { 'o: it is ok (Enter alone does nothing here: look first)' }))
    [void]$nx.Add('n: it is ng')
    if ($retime) { [void]$nx.Add('t 10:15: the start time was wrong -- this job again from that time (the files just downloaded are set aside)') }
    if ($recheck) { [void]$nx.Add('r: you put the right file in the folder -- list the files and compare again') }
    foreach ($l in @('k: keep the verdict as it is', 's: skip this item, leave it pending', 'm <text>: add a note, then answer again', 'q: cancel the whole run')) { [void]$nx.Add($l) }
    $next = $nx.ToArray()
    if (@(@($In['front']) | Where-Object { [string]$_ -ne '' }).Count -gt 0 -and -not $Ctx['DryRun']) {
        $wins = @()
        try { $wins = @(Select-EbiWindowsByText -Windows (Get-EbiTopWindows) -Texts $In['front']) } catch { $wins = @() }
        $front = @($wins | ForEach-Object { [IntPtr]$_['handle'] })
        if ($front.Count -gt 0) { $next = @(('(' + $front.Count + ' window(s) brought to the front: look there, then answer here)')) + $next }
        else { $Ctx.Log.Warn(('no window matches front: {0}' -f (@($In['front']) -join ', '))) }
    }
    $note = ''
    $actions = HumanGate-Actions -EnterOk $enterOk -Retime $retime -Recheck $recheck
    $clock = ''
    while ($true) {
        $r = Show-EbiGate -Title ('GATE ' + $key) -What $what -Next $next -Evidence $ev.ToArray() -Actions $actions -Default $(if ($enterOk) { 'o' } else { '' }) -Auto 'k' -DryRun ([bool]$Ctx['DryRun']) -Alert $panel['alert'] -FrontAfter $front
        if ([string]$r['action'] -eq 'm') { $note = [string]$r['note']; continue }
        if ([string]$r['action'] -eq 't') {
            $clock = HumanGate-Clock -Text ([string]$r['note'])
            if ($clock -eq '') { $what = @($panel['what']) + @(('t needs a time of day, like t 10:15 or t 10:15:45 (typed: "' + [string]$r['note'] + '")')); continue }
        }
        break
    }
    if ($r['auto']) { $Ctx.Log.Info('nobody to ask (dry run or no console): verdict kept as is') }
    if ([string]$r['action'] -eq 'q') { return @{ ok = $false; failure = 'operator_quit'; message = 'operator answered q at the gate'; code = $code; action = 'quit'; note = $note; clock = '' } }
    if ([string]$r['action'] -eq 't') { $Ctx.Log.Info(('{0}: start time {1} typed at the gate -- this job again from there' -f $key, $clock)); return @{ ok = $true; code = $code; action = 'retime'; note = $note; clock = $clock } }
    if ([string]$r['action'] -eq 'r') { $Ctx.Log.Info(('{0}: files put right by hand -- listing and comparing again' -f $key)); return @{ ok = $true; code = $code; action = 'recheck'; note = $note; clock = '' } }
    $d = HumanGate-Decide -Code $code -AskWhen $askWhen -Answer ([string]$r['action'])
    if ($d['action'] -eq 'skip' -and [bool]$In['skipFails']) { return @{ ok = $false; failure = 'operator_skip'; message = 'operator skipped this item at the gate'; code = ''; action = 'skip'; note = $note; clock = '' } }
    return @{ ok = $true; code = $d['code']; action = $d['action']; note = $note; clock = '' }
}
