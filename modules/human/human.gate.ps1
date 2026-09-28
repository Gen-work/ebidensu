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
  }
  outputs    = @{
    code   = @{ type='string'; desc='ok | ng | unknown | empty -- what flow.checkpoint should write' }
    action = @{ type='string'; desc='pass | ok | ng | keep | skip' }
    note   = @{ type='string'; desc='free text typed with m' }
  }
  failures   = @(
    @{ id = 'operator_quit'; transient = $false }
  )
  example    = @{ use = 'human.gate'; with = @{ code = '{{steps.verdict.out.code}}'; askWhen = @('unknown'); reason = '{{steps.verdict.out.reason}}'; evidence = '{{steps.shot.out.path}}' } }
  notes      = 'Under DryRun or without a console the panel is printed and the verdict is kept as is (action=keep). onError.policy=ask reuses this panel''s rendering (kernel/Gate.ps1), so the two never drift.'
}

function HumanGate-Actions { return @(@{ key = 'o'; label = 'ok (Enter)' }, @{ key = 'n'; label = 'ng' }, @{ key = 'k'; label = 'keep as is' }, @{ key = 's'; label = 'skip (leave pending)' }, @{ key = 'm'; label = 'm <note>' }, @{ key = 'q'; label = 'quit' }) }

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

function Invoke-Step {
    param($In, $Ctx)
    $code = [string]$In['code']
    $askWhen = @($In['askWhen'])
    $first = HumanGate-Decide -Code $code -AskWhen $askWhen -Answer ''
    if ($first['action'] -eq 'pass') { return @{ ok = $true; code = $code; action = 'pass'; note = '' } }
    $key = [string]$In['key']
    if ($key -eq '' -and $null -ne $Ctx['Item']) { $key = Get-EbiKeyDisplay -Item $Ctx['Item'] -KeyColumns @($Ctx['KeyColumns']) }
    $ev = New-Object System.Collections.ArrayList
    foreach ($e in @($In['evidence'])) { if ($null -ne $e -and [string]$e -ne '') { [void]$ev.Add([string]$e) } }
    $what = @(('verdict so far: ' + $(if ($code -eq '') { '(empty)' } else { $code })), [string]$In['reason'])
    $note = ''
    while ($true) {
        $r = Show-EbiGate -Title ('GATE ' + $key) -What $what -Next @('Enter/o: it is ok', 'n: it is ng', 'k: keep the verdict as it is', 's: skip this item, leave it pending', 'm <text>: add a note, then answer again', 'q: cancel the whole run') -Evidence $ev.ToArray() -Actions (HumanGate-Actions) -Default 'o' -Auto 'k' -DryRun ([bool]$Ctx['DryRun'])
        if ([string]$r['action'] -eq 'm') { $note = [string]$r['note']; continue }
        break
    }
    if ($r['auto']) { $Ctx.Log.Info('nobody to ask (dry run or no console): verdict kept as is') }
    if ([string]$r['action'] -eq 'q') { return @{ ok = $false; failure = 'operator_quit'; message = 'operator answered q at the gate'; code = $code; action = 'quit'; note = $note } }
    $d = HumanGate-Decide -Code $code -AskWhen $askWhen -Answer ([string]$r['action'])
    return @{ ok = $true; code = $d['code']; action = $d['action']; note = $note }
}
