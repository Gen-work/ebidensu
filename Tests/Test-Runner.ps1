#Requires -Version 5.1
# Test-Runner.ps1 -- the P0-07 spike runner against fixture steps.
#
# Fixture steps are written to a temp directory at run time (never
# committed: a committed fixture would be picked up by the contract checker
# and the parse check). Every assertion is about the runner's own contract
# surface: Session registration / lookup / release, input preparation,
# failure handling, trace, teardown guarantee, quit. The last block runs the
# REAL spike workflow in DryRun against modules/ -- the only execution the
# three P0-08 steps get before an office PC.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Runner.ps1')

Reset-Tests 'Runner'

$utf8 = New-Object System.Text.UTF8Encoding($false)
$tmp  = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-runner-' + [guid]::NewGuid().ToString('N'))
$mods = Join-Path $tmp 'modules'
$work = Join-Path $tmp 'work'
New-Item -ItemType Directory -Path (Join-Path $mods 'fx') -Force | Out-Null
New-Item -ItemType Directory -Path $work -Force | Out-Null

function Write-Fixture {
    param([string]$Name, [string]$Body)
    [System.IO.File]::WriteAllText((Join-Path (Join-Path $mods 'fx') ($Name + '.ps1')), $Body, $utf8)
}
function Write-Workflow {
    param([string]$Name, [string]$Json)
    $p = Join-Path $tmp ($Name + '.json')
    [System.IO.File]::WriteAllText($p, $Json, $utf8)
    return $p
}
function Get-Rec {
    param($Summary, [string]$Id)
    foreach ($r in $Summary.steps) { if ($r.id -eq $Id) { return $r } }
    return $null
}

# ---------------------------------------------------------------- fixtures
Write-Fixture 'fx.provide' @'
$Manifest = @{
  id='fx.provide'; group='fx'; summary='provide a window-kind resource'; tier='core'; effects='ui'
  needs=@(); provides=@('window'); releases=@(); idempotent=$true
  inputs=@{ value=@{ type='int'; default=42 } }
  outputs=@{ note=@{ type='string' } }
  failures=@( @{ id='boom'; transient=$false } )
  example=@{ use='fx.provide'; with=@{ as='w' } }
}
function Invoke-Step { param($In, $Ctx)
  if ($In.ContainsKey('as') -and -not [string]::IsNullOrEmpty([string]$In['as'])) {
    return @{ ok=$true; note='registered'; resource=[int]$In['value'] }
  }
  return @{ ok=$true; note='unregistered' }
}
'@
Write-Fixture 'fx.consume' @'
$Manifest = @{
  id='fx.consume'; group='fx'; summary='read a window-kind resource'; tier='core'; effects='read'
  needs=@(); provides=@(); releases=@(); idempotent=$true
  inputs=@{ window=@{ type='session'; sessionKind='window'; required=$true }; out=@{ type='path'; default='out/x.txt' } }
  outputs=@{ seen=@{ type='int' }; out=@{ type='path' }; dry=@{ type='bool' } }
  failures=@( @{ id='not_found'; transient=$false } )
  example=@{ use='fx.consume'; with=@{ window='w' } }
}
function Invoke-Step { param($In, $Ctx)
  return @{ ok=$true; seen=[int]$Ctx.Session[[string]$In['window']]; out=[string]$In['out']; dry=[bool]$Ctx.DryRun }
}
'@
Write-Fixture 'fx.release' @'
$Manifest = @{
  id='fx.release'; group='fx'; summary='release a window-kind resource'; tier='core'; effects='ui'
  needs=@(); provides=@(); releases=@('window'); idempotent=$true
  inputs=@{ window=@{ type='session'; sessionKind='window'; required=$true } }
  outputs=@{}
  failures=@( @{ id='boom'; transient=$false } )
  example=@{ use='fx.release'; with=@{ window='w' } }
}
function Invoke-Step { param($In, $Ctx) return @{ ok=$true } }
'@
Write-Fixture 'fx.fail' @'
$Manifest = @{
  id='fx.fail'; group='fx'; summary='always fails'; tier='core'; effects='pure'
  needs=@(); provides=@(); releases=@(); idempotent=$true
  inputs=@{}; outputs=@{}
  failures=@( @{ id='not_found'; transient=$false } )
  example=@{ use='fx.fail' }
}
function Invoke-Step { param($In, $Ctx) return @{ ok=$false; failure='not_found'; message='nope' } }
'@
Write-Fixture 'fx.helper' @'
$Manifest = @{
  id='fx.helper'; group='fx'; summary='uses a prefixed helper'; tier='core'; effects='pure'
  needs=@(); provides=@(); releases=@(); idempotent=$true
  inputs=@{ v=@{ type='string'; required=$true } }; outputs=@{ echoed=@{ type='string' } }
  failures=@( @{ id='boom'; transient=$false } )
  example=@{ use='fx.helper'; with=@{ v='x' } }
}
function FxHelper-Shout { param([string]$s) return ($s.ToUpperInvariant() + '!') }
function Invoke-Step { param($In, $Ctx) return @{ ok=$true; echoed=(FxHelper-Shout ([string]$In['v'])) } }
'@
Write-Fixture 'fx.quit' @'
$Manifest = @{
  id='fx.quit'; group='fx'; summary='operator quits'; tier='core'; effects='ui'
  needs=@(); provides=@(); releases=@(); idempotent=$true
  inputs=@{}; outputs=@{ action=@{ type='string' } }
  failures=@( @{ id='no_console'; transient=$false } )
  example=@{ use='fx.quit' }
}
function Invoke-Step { param($In, $Ctx) return @{ ok=$true; action='quit' } }
'@
Write-Fixture 'fx.noresource' @'
$Manifest = @{
  id='fx.noresource'; group='fx'; summary='provides but forgets resource'; tier='core'; effects='ui'
  needs=@(); provides=@('window'); releases=@(); idempotent=$true
  inputs=@{}; outputs=@{}
  failures=@( @{ id='boom'; transient=$false } )
  example=@{ use='fx.noresource'; with=@{ as='w' } }
}
function Invoke-Step { param($In, $Ctx) return @{ ok=$true } }
'@
Write-Fixture 'fx.throw' @'
$Manifest = @{
  id='fx.throw'; group='fx'; summary='throws'; tier='core'; effects='pure'
  needs=@(); provides=@(); releases=@(); idempotent=$true
  inputs=@{}; outputs=@{}
  failures=@( @{ id='boom'; transient=$false } )
  example=@{ use='fx.throw' }
}
function Invoke-Step { param($In, $Ctx) throw 'kaboom' }
'@
Write-Fixture 'fx.badfailure' @'
$Manifest = @{
  id='fx.badfailure'; group='fx'; summary='returns an undeclared failure id'; tier='core'; effects='pure'
  needs=@(); provides=@(); releases=@(); idempotent=$true
  inputs=@{}; outputs=@{}
  failures=@( @{ id='boom'; transient=$false } )
  example=@{ use='fx.badfailure' }
}
function Invoke-Step { param($In, $Ctx) return @{ ok=$false; failure='made_up' } }
'@

# ---------------------------------------------------------------- 1. happy chain
$wf1 = Write-Workflow 'chain' @'
{ "schema": 1, "id": "t.chain",
  "setup": [
    { "id": "p", "use": "fx.provide", "with": { "as": "w" } },
    { "id": "c", "use": "fx.consume", "with": { "window": "w" } },
    { "id": "h", "use": "fx.helper",  "with": { "v": "hi" } }
  ] }
'@
$s1 = Invoke-EbiWorkflow -Path $wf1 -WorkDir $work -ModulesRoot $mods -RunId 'r1'
Assert-True  $s1.ok                                'chain: run ok'
Assert-Equal 3 $s1.steps.Count                     'chain: three step records'
Assert-Equal 42 $s1.ctx.Session['w']               'chain: resource registered under the as name (default value filled)'
Assert-Equal 42 (Get-Rec $s1 'c').outputs['seen']  'chain: consumer looked the name up in Session'
Assert-True  (-not (Get-Rec $s1 'p').outputs.ContainsKey('resource')) 'chain: resource stripped from recorded outputs'
Assert-Equal 'registered' (Get-Rec $s1 'p').outputs['note'] 'chain: provider saw $In.as'
Assert-Equal (Join-Path $work 'out/x.txt') (Get-Rec $s1 'c').outputs['out'] 'chain: path input rooted under WorkDir'
Assert-Equal 'HI!' (Get-Rec $s1 'h').outputs['echoed'] 'chain: prefixed helper resolves when the captured Invoke-Step runs'
$ev1 = @(Read-TraceEvents -WorkDir $work -RunId 'r1')
$stepEvents = @($ev1 | Where-Object { $_.action -ne 'run' })
Assert-Equal 3 $stepEvents.Count                   'chain: one trace event per step call'
Assert-Equal 'setup' $stepEvents[0].tags.segment   'chain: trace event tagged with its segment'
$startEnd = @($ev1 | Where-Object { $_.action -eq 'run' })
Assert-Equal 2 $startEnd.Count                     'chain: run start + run end events'
Assert-Equal 'ok' $startEnd[1].status              'chain: run end status ok'

# ---------------------------------------------------------------- 2. failure stops the segment
$wf2 = Write-Workflow 'fail' @'
{ "schema": 1, "id": "t.fail",
  "setup": [
    { "id": "p", "use": "fx.provide", "with": { "as": "w" } },
    { "id": "f", "use": "fx.fail" },
    { "id": "c", "use": "fx.consume", "with": { "window": "w" } }
  ] }
'@
$s2 = Invoke-EbiWorkflow -Path $wf2 -WorkDir $work -ModulesRoot $mods -RunId 'r2'
Assert-True  (-not $s2.ok)                         'fail: run not ok'
Assert-Equal 'f' $s2.failedStep                    'fail: failed step id recorded'
Assert-Equal 'not_found' $s2.failure               'fail: failure id from the step'
Assert-Equal 'nope' $s2.message                    'fail: message from the step'
Assert-Equal 2 $s2.steps.Count                     'fail: the step after the failure did not run'
Assert-True  ($null -eq (Get-Rec $s2 'c'))         'fail: no record for the unreached step'

# ---------------------------------------------------------------- 3. Session rules
$wf3 = Write-Workflow 'as-on-nonprovider' '{ "schema": 1, "id": "t.as", "setup": [ { "id": "f", "use": "fx.fail", "with": { "as": "x" } } ] }'
$s3 = Invoke-EbiWorkflow -Path $wf3 -WorkDir $work -ModulesRoot $mods -RunId 'r3'
Assert-Equal 'internal_error' $s3.failure          'as on a step that provides nothing is a contract violation'
Assert-True  ($s3.message -like '*provides nothing*') 'as on non-provider: message names the rule'
Assert-Equal 0 $s3.ctx.Session.Count               'as on non-provider: step never ran, nothing registered'

$wf4 = Write-Workflow 'missing' '{ "schema": 1, "id": "t.missing", "setup": [ { "id": "c", "use": "fx.consume", "with": { "window": "ghost" } } ] }'
$s4 = Invoke-EbiWorkflow -Path $wf4 -WorkDir $work -ModulesRoot $mods -RunId 'r4'
Assert-Equal 'session_missing' $s4.failure         'unregistered session name -> session_missing before the step runs'

$wf5 = Write-Workflow 'conflict' @'
{ "schema": 1, "id": "t.conflict",
  "setup": [
    { "id": "p1", "use": "fx.provide", "with": { "as": "w" } },
    { "id": "p2", "use": "fx.provide", "with": { "as": "w" } }
  ] }
'@
$s5 = Invoke-EbiWorkflow -Path $wf5 -WorkDir $work -ModulesRoot $mods -RunId 'r5'
Assert-Equal 'session_conflict' $s5.failure        'registering a live name again -> session_conflict'
Assert-Equal 'p2' $s5.failedStep                   'conflict: the second registration is the one that fails'
Assert-Equal 42 $s5.ctx.Session['w']               'conflict: the first resource is untouched, not overwritten'

$wf6 = Write-Workflow 'release' @'
{ "schema": 1, "id": "t.release",
  "setup": [
    { "id": "p1", "use": "fx.provide", "with": { "as": "w", "value": 7 } },
    { "id": "rel", "use": "fx.release", "with": { "window": "w" } },
    { "id": "p2", "use": "fx.provide", "with": { "as": "w", "value": 9 } }
  ] }
'@
$s6 = Invoke-EbiWorkflow -Path $wf6 -WorkDir $work -ModulesRoot $mods -RunId 'r6'
Assert-True  $s6.ok                                'release: run ok'
Assert-Equal 9 $s6.ctx.Session['w']                'release: name freed by the releases step, then re-registered'

$wf7 = Write-Workflow 'noresource' '{ "schema": 1, "id": "t.nores", "setup": [ { "id": "p", "use": "fx.noresource", "with": { "as": "w" } } ] }'
$s7 = Invoke-EbiWorkflow -Path $wf7 -WorkDir $work -ModulesRoot $mods -RunId 'r7'
Assert-Equal 'internal_error' $s7.failure          'as given but no resource returned -> internal_error'
Assert-Equal 0 $s7.ctx.Session.Count               'no resource: nothing registered'

# ---------------------------------------------------------------- 4. return-value contract
$wf8 = Write-Workflow 'throw' '{ "schema": 1, "id": "t.throw", "setup": [ { "id": "t", "use": "fx.throw" } ] }'
$s8 = Invoke-EbiWorkflow -Path $wf8 -WorkDir $work -ModulesRoot $mods -RunId 'r8'
Assert-Equal 'internal_error' $s8.failure          'a throwing step becomes internal_error'
Assert-True  ($s8.message -like '*kaboom*')        'throw: the exception message is kept'

$wf9 = Write-Workflow 'badfailure' '{ "schema": 1, "id": "t.badf", "setup": [ { "id": "b", "use": "fx.badfailure" } ] }'
$s9 = Invoke-EbiWorkflow -Path $wf9 -WorkDir $work -ModulesRoot $mods -RunId 'r9'
Assert-Equal 'internal_error' $s9.failure          'an undeclared failure id is a contract violation'
Assert-True  ($s9.message -like '*made_up*')       'bad failure: names the offending id'

$wf10 = Write-Workflow 'undeclared' '{ "schema": 1, "id": "t.undecl", "setup": [ { "id": "h", "use": "fx.helper", "with": { "v": "a", "bogus": 1 } } ] }'
$s10 = Invoke-EbiWorkflow -Path $wf10 -WorkDir $work -ModulesRoot $mods -RunId 'r10'
Assert-Equal 'internal_error' $s10.failure         'an undeclared input is rejected'
Assert-True  ($s10.message -like '*bogus*')        'undeclared input: names the key'

$wf11 = Write-Workflow 'required' '{ "schema": 1, "id": "t.req", "setup": [ { "id": "h", "use": "fx.helper" } ] }'
$s11 = Invoke-EbiWorkflow -Path $wf11 -WorkDir $work -ModulesRoot $mods -RunId 'r11'
Assert-True  ($s11.message -like '*required input "v"*') 'a missing required input is named'

# ---------------------------------------------------------------- 5. teardown guarantee + quit
$wf12 = Write-Workflow 'teardown' @'
{ "schema": 1, "id": "t.teardown",
  "setup":    [ { "id": "f", "use": "fx.fail" } ],
  "teardown": [ { "id": "t", "use": "fx.provide", "with": { "as": "afterwards" } } ] }
'@
$s12 = Invoke-EbiWorkflow -Path $wf12 -WorkDir $work -ModulesRoot $mods -RunId 'r12'
Assert-True  (-not $s12.setupOk)                   'teardown: setup failed'
Assert-True  $s12.teardownOk                       'teardown: still ran after the setup failure'
Assert-True  $s12.ctx.Session.ContainsKey('afterwards') 'teardown: its step really executed'
Assert-Equal 'teardown' (Get-Rec $s12 't').segment 'teardown: record tagged with its segment'

$wf13 = Write-Workflow 'quit' @'
{ "schema": 1, "id": "t.quit",
  "setup":    [ { "id": "q", "use": "fx.quit" }, { "id": "h", "use": "fx.helper", "with": { "v": "x" } } ],
  "teardown": [ { "id": "t", "use": "fx.provide", "with": { "as": "afterwards" } } ] }
'@
$s13 = Invoke-EbiWorkflow -Path $wf13 -WorkDir $work -ModulesRoot $mods -RunId 'r13'
Assert-True  $s13.cancelled                        'quit: run marked cancelled'
Assert-True  (-not $s13.ok)                        'quit: run not ok'
Assert-Equal 'cancelled' $s13.failure              'quit: reserved failure id'
Assert-True  ($null -eq (Get-Rec $s13 'h'))        'quit: the following step did not run'
Assert-True  $s13.ctx.Session.ContainsKey('afterwards') 'quit: teardown ran'
$ev13 = @(Read-TraceEvents -WorkDir $work -RunId 'r13' | Where-Object { $_.action -eq 'run' })
Assert-Equal 'cancelled' $ev13[1].status           'quit: run end event says cancelled'

# ---------------------------------------------------------------- 6. DryRun passes through
$s14 = Invoke-EbiWorkflow -Path $wf1 -WorkDir $work -ModulesRoot $mods -RunId 'r14' -DryRun $true
Assert-True  $s14.ok                               'dryrun: fixture chain ok'
Assert-True  ([bool](Get-Rec $s14 'c').outputs['dry']) 'dryrun: $Ctx.DryRun seen by the step'

# ---------------------------------------------------------------- 7. workflow shape checks
function Assert-Throws {
    param([string]$Json, [string]$Needle, [string]$Msg)
    $p = Write-Workflow ('bad-' + [guid]::NewGuid().ToString('N')) $Json
    $threw = $false; $text = ''
    try { [void](Read-EbiWorkflow -Path $p) } catch { $threw = $true; $text = $_.Exception.Message }
    Assert-True ($threw -and ($text -like ('*' + $Needle + '*'))) $Msg
}
Assert-Throws '{ "id": "x", "setup": [] }'                       'schema'    'a workflow without schema is refused'
Assert-Throws '{ "schema": 2, "id": "x", "setup": [] }'          'schema 2'  'a newer schema is refused'
Assert-Throws '{ "schema": 1, "setup": [] }'                     '"id"'      'a workflow without id is refused'
Assert-Throws '{ "schema": 1, "id": "x", "setup": [ { "use": "fx.fail" } ] }' '"id"' 'a step call without id is refused (ledger key)'
Assert-Throws '{ "schema": 1, "id": "x", "setup": [ { "id": "a", "use": "fx.fail" }, { "id": "a", "use": "fx.fail" } ] }' 'duplicate' 'duplicate step ids in one segment are refused'
$wfMissing = Write-Workflow 'missingstep' '{ "schema": 1, "id": "t.ms", "setup": [ { "id": "z", "use": "fx.nonexistent" } ] }'
$threw = $false
try { [void](Invoke-EbiWorkflow -Path $wfMissing -WorkDir $work -ModulesRoot $mods -RunId 'r15') } catch { $threw = ($_.Exception.Message -like '*fx.nonexistent*') }
Assert-True $threw 'a workflow referencing a step file that does not exist fails before anything runs'

# ---------------------------------------------------------------- 8. the real spike workflow, DryRun, against modules/
$spike = Join-Path (Join-Path $repoRoot 'workflows') 'spike.capture.json'
$sS = Invoke-EbiWorkflow -Path $spike -WorkDir $work -RunId 'spike' -DryRun $true
Assert-True  $sS.ok                                'spike (dryrun): human.prepare -> browser.ensure -> screen.capture_window all ok'
Assert-Equal 3 $sS.steps.Count                     'spike (dryrun): three step calls'
Assert-True  $sS.ctx.Session.ContainsKey('mainWindow') 'spike (dryrun): browser.ensure registered mainWindow through the resource key'
Assert-Equal (Join-Path $work 'capture/spike/window.png') (Get-Rec $sS 'shot').outputs['path'] 'spike (dryrun): saveAs rooted under WorkDir'
Assert-True  (-not (Test-Path -LiteralPath (Join-Path $work 'capture'))) 'spike (dryrun): nothing written to capture/'

# ---------------------------------------------------------------- cleanup
try { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue } catch { }

exit (Complete-Tests)
