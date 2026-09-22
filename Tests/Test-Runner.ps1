#Requires -Version 5.1
# Test-Runner.ps1 -- the P0-07 minimal runner spike (kernel/Runner.ps1).
#
# The card's completion criterion: a workflow JSON with only a three-step
# "setup" runs, and the window handle travels ensure -> capture through
# $Ctx.Session rather than a global variable. Steps are fixtures written to
# a temp modules root at run time (same approach as Test-StepContract.ps1),
# so nothing under modules/ is touched and the real contract checker never
# sees them.

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'kernel/Runner.ps1')

Reset-Tests 'Runner'

# ---------------------------------------------------------------- fixtures
$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ebidensu-runner-' + [guid]::NewGuid().ToString('N'))
$modules = Join-Path $tmpRoot 'modules'
$work    = Join-Path $tmpRoot 'work'
New-Item -ItemType Directory -Path (Join-Path $modules 'fake') -Force | Out-Null
New-Item -ItemType Directory -Path $work -Force | Out-Null

function Write-Fixture {
    param([string]$Name, [string]$Body)
    $path = Join-Path (Join-Path $modules 'fake') $Name
    [System.IO.File]::WriteAllText($path, $Body, (New-Object System.Text.UTF8Encoding($false)))
}

function Write-Workflow {
    param([string]$Name, [string]$Json)
    $path = Join-Path $tmpRoot $Name
    [System.IO.File]::WriteAllText($path, $Json, (New-Object System.Text.UTF8Encoding($false)))
    return $path
}

# fake.ensure: provides a 'window'; the handle is a plain int so the test can
# recognise it on the other side. Returns resource = $null under DryRun.
Write-Fixture 'fake.ensure.ps1' @'
$Manifest = @{
  id = 'fake.ensure'; group = 'fake'; summary = 'fixture: register a window'; tier = 'core'
  effects = 'ui'; needs = @(); provides = @('window'); releases = @(); idempotent = $true
  inputs  = @{ title = @{ type='string'; required=$true } }
  outputs = @{ title = @{ type='string' }; sawAs = @{ type='bool' } }
  failures = @( @{ id = 'no_window'; transient = $true } )
  example = @{ use='fake.ensure'; with=@{ title='x' } }
}
function FakeEnsure-Handle { return 4242 }
function Invoke-Step {
    param($In, $Ctx)
    if ($Ctx['DryRun']) { return @{ ok = $true; resource = $null; title = $In['title']; sawAs = $In.Contains('as') } }
    return @{ ok = $true; resource = (FakeEnsure-Handle); title = $In['title']; sawAs = $In.Contains('as') }
}
'@

# fake.capture: consumes a 'window' session input and echoes what it received.
Write-Fixture 'fake.capture.ps1' @'
$Manifest = @{
  id = 'fake.capture'; group = 'fake'; summary = 'fixture: use a window'; tier = 'core'
  effects = 'write'; needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs  = @{ window = @{ type='session'; sessionKind='window'; required=$true }
               saveAs = @{ type='path'; required=$true } }
  outputs = @{ path = @{ type='path' }; handle = @{ type='int' }; dryRun = @{ type='bool' } }
  failures = @( @{ id = 'window_gone'; transient = $true } )
  example = @{ use='fake.capture'; with=@{ window='w'; saveAs='x.png' } }
}
function FakeCapture-Path { param($p) return $p }
function Invoke-Step {
    param($In, $Ctx)
    $h = $In['window']
    return @{ ok = $true; path = (FakeCapture-Path $In['saveAs']); handle = $(if ($null -eq $h) { -1 } else { [int]$h }); dryRun = [bool]$Ctx['DryRun'] }
}
'@

# fake.needbook: wants a 'workbook', so a 'window' name is a kind mismatch.
Write-Fixture 'fake.needbook.ps1' @'
$Manifest = @{
  id = 'fake.needbook'; group = 'fake'; summary = 'fixture: wants a workbook'; tier = 'core'
  effects = 'read'; needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs  = @{ book = @{ type='session'; sessionKind='workbook'; required=$true } }
  outputs = @{}
  failures = @( @{ id = 'nope'; transient = $false } )
  example = @{ use='fake.needbook'; with=@{ book='wb' } }
}
function Invoke-Step { param($In, $Ctx) return @{ ok = $true } }
'@

# fake.release: releases the 'window' it is handed.
Write-Fixture 'fake.release.ps1' @'
$Manifest = @{
  id = 'fake.release'; group = 'fake'; summary = 'fixture: release a window'; tier = 'core'
  effects = 'ui'; needs = @(); provides = @(); releases = @('window'); idempotent = $true
  inputs  = @{ window = @{ type='session'; sessionKind='window'; required=$true } }
  outputs = @{ released = @{ type='int' } }
  failures = @( @{ id = 'window_gone'; transient = $true } )
  example = @{ use='fake.release'; with=@{ window='w' } }
}
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; released = [int]$In['window'] } }
'@

# fake.pure: echoes; also the probe for $Ctx shape.
Write-Fixture 'fake.pure.ps1' @'
$Manifest = @{
  id = 'fake.pure'; group = 'fake'; summary = 'fixture: echo'; tier = 'core'
  effects = 'pure'; needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs  = @{ text = @{ type='string'; default='' } }
  outputs = @{ echo = @{ type='string' }; hasSession = @{ type='bool' }; hasLog = @{ type='bool' }; runId = @{ type='string' } }
  failures = @( @{ id = 'never'; transient = $false } )
  example = @{ use='fake.pure'; with=@{ text='hi' } }
}
function Invoke-Step {
    param($In, $Ctx)
    $Ctx.Log.Info('pure step ran')
    return @{ ok = $true; echo = [string]$In['text']; hasSession = ($Ctx['Session'] -is [hashtable]); hasLog = ($null -ne $Ctx['Log']); runId = [string]$Ctx['RunId'] }
}
'@

# fake.fail: every way a step can go wrong, selected by "mode".
Write-Fixture 'fake.fail.ps1' @'
$Manifest = @{
  id = 'fake.fail'; group = 'fake'; summary = 'fixture: fail on demand'; tier = 'core'
  effects = 'pure'; needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs  = @{ mode = @{ type='string'; required=$true } }
  outputs = @{ partial = @{ type='int' } }
  failures = @( @{ id = 'boom'; transient = $false } )
  example = @{ use='fake.fail'; with=@{ mode='declared' } }
}
function Invoke-Step {
    param($In, $Ctx)
    switch ([string]$In['mode']) {
        'declared'   { return @{ ok = $false; failure = 'boom'; message = 'asked to'; partial = 7 } }
        'undeclared' { return @{ ok = $false; failure = 'not_in_manifest' } }
        'throw'      { throw 'kaboom' }
        'notable'    { return 'a string' }
        'nook'       { return @{ partial = 1 } }
        'resource'   { return @{ ok = $true; resource = 1 } }
        'warn'       { return @{ ok = $true; partial = 2; warnings = @( @{ code = 'odd'; message = 'one odd line'; data = @{ lines = @(3) } } ) } }
    }
    return @{ ok = $true; partial = 0 }
}
'@

# fake.noresource: provides a window but forgets to hand it over.
Write-Fixture 'fake.noresource.ps1' @'
$Manifest = @{
  id = 'fake.noresource'; group = 'fake'; summary = 'fixture: provides but no resource key'; tier = 'core'
  effects = 'ui'; needs = @(); provides = @('window'); releases = @(); idempotent = $true
  inputs  = @{}
  outputs = @{}
  failures = @( @{ id = 'no_window'; transient = $true } )
  example = @{ use='fake.noresource'; with=@{} }
}
function Invoke-Step { param($In, $Ctx) return @{ ok = $true } }
'@

# fake.noinvoke: a manifest and nothing else.
Write-Fixture 'fake.noinvoke.ps1' @'
$Manifest = @{ id = 'fake.noinvoke'; group = 'fake'; summary = 'fixture: no entry point'; tier = 'core'; effects = 'pure'; inputs = @{}; outputs = @{}; failures = @( @{ id = 'x'; transient = $false } ); example = @{ use = 'fake.noinvoke' } }
'@

function Get-Rec {
    param($Result, [string]$Id)
    foreach ($r in $Result['steps']) { if ($r['id'] -eq $Id) { return $r } }
    return $null
}

function Get-StepEvents {
    param([string]$RunId, [string]$Id)
    $out = New-Object System.Collections.ArrayList
    foreach ($e in @(Read-TraceEvents -WorkDir $work -RunId $RunId)) {
        if ($e.action -eq 'step' -and $e.tags.step -eq $Id) { [void]$out.Add($e) }
    }
    return $out.ToArray()
}

try {
    # ------------------------------------------------ pure helpers
    Assert-Equal '' (Get-EbiStepPath -ModulesRoot $modules -Use 'ensure') 'a bare verb is not a step id'
    Assert-Equal '' (Get-EbiStepPath -ModulesRoot $modules -Use 'fake.') 'a trailing dot is not a step id'
    $p = Get-EbiStepPath -ModulesRoot $modules -Use 'screen.capture_window'
    Assert-True ($p.EndsWith(('screen' + [IO.Path]::DirectorySeparatorChar + 'screen.capture_window.ps1'))) 'step path is <root>/<group>/<use>.ps1'
    Assert-True ((Get-EbiDefaultModulesRoot).EndsWith('modules')) 'default modules root is the repo modules/ dir'
    Assert-True (Test-Path -LiteralPath (Get-EbiDefaultModulesRoot)) 'default modules root exists'

    $obj = '{"a":1,"b":{"c":[1,2,{"d":"x"}]},"e":null}' | ConvertFrom-Json
    $h = ConvertTo-EbiHashtable $obj
    Assert-True ($h -is [hashtable]) 'JSON object becomes a hashtable'
    Assert-True ($h['b'] -is [hashtable]) 'nested object becomes a hashtable'
    Assert-True ($h['b']['c'] -is [array]) 'array stays an array'
    Assert-Equal 'x' $h['b']['c'][2]['d'] 'object inside array becomes a hashtable'
    Assert-True ($h.Contains('e') -and $null -eq $h['e']) 'null is kept as a null key'
    $h = ConvertTo-EbiHashtable ('{"one":[{"id":"a"}],"none":[]}' | ConvertFrom-Json)
    Assert-True ($h['one'] -is [array] -and $h['one'].Count -eq 1) 'a one-element array is still an array (not unrolled)'
    Assert-True ($h['one'][0] -is [hashtable]) 'the one element is a hashtable'
    Assert-True ($h['none'] -is [array] -and $h['none'].Count -eq 0) 'an empty array is an empty array'

    Assert-True (Test-EbiValueHasTemplate @{ a = @{ b = @('x', '{{item.key}}') } }) 'a template nested in with is found'
    Assert-True (-not (Test-EbiValueHasTemplate @{ a = 'plain'; n = 3 })) 'plain values are not templates'

    $probs = @(Get-EbiWorkflowSpikeProblems -Workflow @{ id = 'w'; setup = @( @{ id = 'a'; use = 'fake.pure' }, @{ id = 'a'; use = 'fake.pure' } ) })
    Assert-True (($probs -join ' ') -like '*used twice*') 'duplicate step id within a section is a problem'
    $probs = @(Get-EbiWorkflowSpikeProblems -Workflow @{ id = 'w'; setup = @( @{ use = 'fake.pure' } ) })
    Assert-True (($probs -join ' ') -like '*"id" is required*') 'a step call without id is a problem'
    $probs = @(Get-EbiWorkflowSpikeProblems -Workflow @{ id = 'w'; each = @(); source = @{ table = 'x' } })
    Assert-Equal 2 $probs.Count 'source and each are each refused'
    $probs = @(Get-EbiWorkflowSpikeProblems -Workflow @{ id = 'w'; setup = @( @{ id = 'a'; use = 'fake.pure'; with = @{ text = '{{vars.x}}' } } ) })
    Assert-True (($probs -join ' ') -like '*template*') 'a template in with is refused'
    $probs = @(Get-EbiWorkflowSpikeProblems -Workflow @{ id = 'w'; setup = @( @{ id = 'a'; use = 'fake.pure'; onError = @{ policy = 'retry' } } ) })
    Assert-True (($probs -join ' ') -like '*onError*') 'a per-step onError is refused'
    $probs = @(Get-EbiWorkflowSpikeProblems -Workflow @{ id = 'w'; setup = @( @{ id = 'a'; use = 'fake.pure'; with = @{ text = 'hi' } } ); teardown = @( @{ id = 'z'; use = 'fake.pure' } ) })
    Assert-Equal 0 $probs.Count 'a plain setup+teardown workflow has no problems'

    # Resolve-EbiStepInputs on its own
    $m = @{ id = 'x'; provides = @('window'); inputs = @{ title = @{ type = 'string' } } }
    $r = Resolve-EbiStepInputs -Manifest $m -With @{ title = 't'; as = 'w' } -Session @{}
    Assert-True ($r['ok']) 'as on a provides step resolves'
    Assert-Equal 'w' $r['As'] 'as is captured'
    Assert-True (-not $r['In'].Contains('as')) 'as does not reach In'
    $r = Resolve-EbiStepInputs -Manifest @{ id = 'x'; provides = @(); inputs = @{} } -With @{ as = 'w' } -Session @{}
    Assert-Equal 'contract_violation' $r['failure'] 'as on a step that provides nothing is a contract violation'
    $r = Resolve-EbiStepInputs -Manifest $m -With @{ title = 't'; as = 'w' } -Session @{ w = @{ kind = 'window'; value = 1 } }
    Assert-Equal 'session_name_taken' $r['failure'] 'a live name cannot be registered twice'

    # Test-EbiStepReturn on its own
    $m2 = @{ id = 'y'; provides = @(); failures = @( @{ id = 'boom'; transient = $false } ) }
    $c = Test-EbiStepReturn -Manifest $m2 -Return @{ ok = $true; a = 1; warnings = @('w') } -WantsResource $false
    Assert-True ($c['ok'] -and $c['Outputs'].Count -eq 1 -and $c['Outputs']['a'] -eq 1) 'reserved keys are split from outputs'
    Assert-Equal 1 @($c['Warnings']).Count 'warnings are handed back separately'
    $c = Test-EbiStepReturn -Manifest $m2 -Return ([ordered]@{ ok = $true }) -WantsResource $false
    Assert-Equal 'contract_violation' $c['failure'] 'an ordered dictionary is not the [hashtable] the contract asks for'

    # ------------------------------------------------ 1. the card's acceptance
    $wf = Write-Workflow 'happy.json' @'
{ "id": "spike.happy", "title": "three-step setup", "version": "1.0.0", "profile": "none",
  "setup": [
    { "id": "ensure",  "use": "fake.ensure",  "with": { "title": "Edge", "as": "mainWindow" } },
    { "id": "shot",    "use": "fake.capture", "with": { "window": "mainWindow", "saveAs": "capture/x.png" } },
    { "id": "echo",    "use": "fake.pure",    "with": { "text": "hi" } }
  ] }
'@
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-happy'
    Assert-True ($res['ok']) 'happy: three-step setup runs ok'
    Assert-Equal 'spike.happy' $res['workflowId'] 'happy: workflow id is reported'
    Assert-Equal 3 @($res['steps']).Count 'happy: one record per step call'
    $shot = Get-Rec $res 'shot'
    Assert-Equal 4242 $shot['outputs']['handle'] 'happy: the handle reached capture THROUGH $Ctx.Session'
    Assert-Equal 'capture/x.png' $shot['outputs']['path'] 'happy: plain with-values pass through'
    $ens = Get-Rec $res 'ensure'
    Assert-Equal 'False' ([string]$ens['outputs']['sawAs']) 'happy: as never reaches the step In'
    Assert-True (-not $ens['outputs'].Contains('resource')) 'happy: resource is not an output'
    Assert-True ($res['session'].Contains('mainWindow')) 'happy: the name is registered in Session'
    Assert-Equal 'window' $res['session']['mainWindow']['kind'] 'happy: Session entry carries the kind'
    Assert-Equal 4242 $res['session']['mainWindow']['value'] 'happy: Session entry carries the value'
    Assert-Equal 'ensure' $res['session']['mainWindow']['registeredBy'] 'happy: Session entry names the registering call'
    $echo = Get-Rec $res 'echo'
    Assert-Equal 'True' ([string]$echo['outputs']['hasSession']) 'happy: $Ctx.Session is a hashtable from day one'
    Assert-Equal 'True' ([string]$echo['outputs']['hasLog']) 'happy: $Ctx.Log exists'
    Assert-Equal 'r-happy' $echo['outputs']['runId'] 'happy: $Ctx.RunId is the run id'
    Assert-True (-not (Test-Path -LiteralPath 'variable:Global:Shell')) 'happy: no global variable was needed'

    $events = @(Read-TraceEvents -WorkDir $work -RunId 'r-happy')
    Assert-True ($events.Count -ge 8) 'happy: trace has run start/end plus start+ok per step'
    Assert-Equal 'run' $events[0].action 'happy: first trace event is the run start'
    Assert-Equal 'ok' $events[$events.Count - 1].status 'happy: last trace event is the run result'
    $ensEv = @(Get-StepEvents -RunId 'r-happy' -Id 'ensure' | Where-Object { $_.status -eq 'ok' })
    Assert-Equal 1 $ensEv.Count 'happy: ensure has one ok event'
    Assert-Equal 'Edge' $ensEv[0].data.outputs.title 'happy: outputs land in trace data'
    Assert-True (-not ($ensEv[0].data.outputs.PSObject.Properties.Name -contains 'resource')) 'happy: the handle never appears in the trace'
    Assert-Equal 'spike.happy' $ensEv[0].tags.workflow 'happy: trace tags carry the workflow id'
    Assert-Equal 'setup' $ensEv[0].phase 'happy: trace phase is the section'

    # ------------------------------------------------ 2. captured Invoke-Step survives later loads
    $wf = Write-Workflow 'twice.json' @'
{ "id": "spike.twice", "setup": [
    { "id": "e1", "use": "fake.ensure",  "with": { "title": "a", "as": "w1" } },
    { "id": "c1", "use": "fake.capture", "with": { "window": "w1", "saveAs": "1.png" } },
    { "id": "e2", "use": "fake.ensure",  "with": { "title": "b", "as": "w2" } },
    { "id": "c2", "use": "fake.capture", "with": { "window": "w2", "saveAs": "2.png" } }
  ] }
'@
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-twice'
    Assert-True ($res['ok']) 'twice: a step loaded earlier still calls its own helper after another step was loaded'
    Assert-Equal 'b' (Get-Rec $res 'e2')['outputs']['title'] 'twice: the second ensure ran the ensure code, not capture'
    Assert-Equal 4242 (Get-Rec $res 'c2')['outputs']['handle'] 'twice: two names, two registrations'

    # ------------------------------------------------ 3. duplicate name, skip after failure
    $wf = Write-Workflow 'dup.json' @'
{ "id": "spike.dup", "setup": [
    { "id": "e1", "use": "fake.ensure", "with": { "title": "a", "as": "w" } },
    { "id": "e2", "use": "fake.ensure", "with": { "title": "b", "as": "w" } },
    { "id": "c",  "use": "fake.capture", "with": { "window": "w", "saveAs": "x.png" } }
  ] }
'@
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-dup'
    Assert-True (-not $res['ok']) 'dup: run fails'
    Assert-Equal 'session_name_taken' (Get-Rec $res 'e2')['failure'] 'dup: re-registering a live name fails before Invoke-Step'
    Assert-Equal 'e1' $res['session']['w']['registeredBy'] 'dup: the first registration is kept, not overwritten'
    Assert-Equal 'skip' (Get-Rec $res 'c')['status'] 'dup: later setup steps are skipped, not run'
    Assert-Equal 'session_name_taken' $res['failure'] 'dup: the first failure is the run failure'
    $skipEv = @(Get-StepEvents -RunId 'r-dup' -Id 'c')
    Assert-Equal 'skip' $skipEv[0].status 'dup: the skip is traced'

    # ------------------------------------------------ 4/5. missing name, wrong kind
    $wf = Write-Workflow 'missing.json' @'
{ "id": "spike.missing", "setup": [ { "id": "c", "use": "fake.capture", "with": { "window": "nope", "saveAs": "x.png" } } ] }
'@
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-missing'
    Assert-Equal 'session_missing' (Get-Rec $res 'c')['failure'] 'missing: an unregistered name is session_missing'

    $wf = Write-Workflow 'kind.json' @'
{ "id": "spike.kind", "setup": [
    { "id": "e", "use": "fake.ensure",   "with": { "title": "a", "as": "w" } },
    { "id": "b", "use": "fake.needbook", "with": { "book": "w" } }
  ] }
'@
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-kind'
    Assert-Equal 'session_kind_mismatch' (Get-Rec $res 'b')['failure'] 'kind: a window handed to a workbook input is session_kind_mismatch'

    # ------------------------------------------------ 6. release frees the name
    $wf = Write-Workflow 'release.json' @'
{ "id": "spike.release", "setup": [
    { "id": "e1", "use": "fake.ensure",  "with": { "title": "a", "as": "w" } },
    { "id": "r",  "use": "fake.release", "with": { "window": "w" } },
    { "id": "e2", "use": "fake.ensure",  "with": { "title": "b", "as": "w" } }
  ],
  "teardown": [ { "id": "r2", "use": "fake.release", "with": { "window": "w" } } ] }
'@
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-release'
    Assert-True ($res['ok']) 'release: release then re-register the same name works'
    Assert-Equal 4242 (Get-Rec $res 'r')['outputs']['released'] 'release: the release step received the instance'
    Assert-Equal 0 $res['session'].Count 'release: teardown released the last window; Session is empty'

    # ------------------------------------------------ 7. teardown after a setup failure
    $wf = Write-Workflow 'teardown.json' @'
{ "id": "spike.teardown",
  "setup":    [ { "id": "f", "use": "fake.fail", "with": { "mode": "declared" } },
                { "id": "after", "use": "fake.pure" } ],
  "teardown": [ { "id": "t", "use": "fake.pure", "with": { "text": "bye" } } ] }
'@
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-teardown'
    Assert-True (-not $res['ok']) 'teardown: a declared failure fails the run'
    Assert-Equal 'boom' (Get-Rec $res 'f')['failure'] 'teardown: the declared failure id is reported as-is'
    Assert-Equal 7 (Get-Rec $res 'f')['outputs']['partial'] 'teardown: partial outputs of a failed step are kept for the trace'
    Assert-Equal 'skip' (Get-Rec $res 'after')['status'] 'teardown: the rest of setup is skipped'
    Assert-Equal 'ok' (Get-Rec $res 't')['status'] 'teardown: teardown still ran'
    Assert-Equal 'bye' (Get-Rec $res 't')['outputs']['echo'] 'teardown: teardown step received its with'

    # ------------------------------------------------ 8-12. contract violations
    foreach ($case in @(
        @{ mode = 'undeclared'; want = 'contract_violation'; msg = 'an undeclared failure id is a contract violation' },
        @{ mode = 'throw';      want = 'internal_error';     msg = 'an exception is internal_error' },
        @{ mode = 'notable';    want = 'contract_violation'; msg = 'a non-hashtable return is a contract violation' },
        @{ mode = 'nook';       want = 'contract_violation'; msg = 'a return without ok is a contract violation' },
        @{ mode = 'resource';   want = 'contract_violation'; msg = 'resource from a step that provides nothing is a contract violation' }
    )) {
        $wf = Write-Workflow ('fail-' + $case['mode'] + '.json') ('{ "id": "spike.fail", "setup": [ { "id": "f", "use": "fake.fail", "with": { "mode": "' + $case['mode'] + '" } } ] }')
        $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId ('r-fail-' + $case['mode'])
        Assert-Equal $case['want'] (Get-Rec $res 'f')['failure'] $case['msg']
    }
    $res = Invoke-EbiWorkflow -Path (Join-Path $tmpRoot 'fail-throw.json') -WorkDir $work -ModulesRoot $modules -RunId 'r-fail-throw2'
    Assert-Equal 'kaboom' (Get-Rec $res 'f')['message'] 'an exception message is carried'

    $wf = Write-Workflow 'nores.json' '{ "id": "spike.nores", "setup": [ { "id": "n", "use": "fake.noresource", "with": { "as": "w" } } ] }'
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-nores'
    Assert-Equal 'contract_violation' (Get-Rec $res 'n')['failure'] 'a provides step called with as but returning no resource is a contract violation'
    Assert-Equal 0 $res['session'].Count 'nothing is registered when the resource key is missing'

    $wf = Write-Workflow 'asplain.json' '{ "id": "spike.asplain", "setup": [ { "id": "p", "use": "fake.pure", "with": { "as": "w" } } ] }'
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-asplain'
    Assert-Equal 'contract_violation' (Get-Rec $res 'p')['failure'] 'as on a step without provides is a contract violation'

    # ------------------------------------------------ 13. warnings
    $wf = Write-Workflow 'warn.json' '{ "id": "spike.warn", "setup": [ { "id": "w", "use": "fake.fail", "with": { "mode": "warn" } } ] }'
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-warn'
    Assert-True ($res['ok']) 'warn: warnings do not fail the step'
    Assert-Equal 1 @((Get-Rec $res 'w')['warnings']).Count 'warn: the warning is on the record'
    $wEv = @(Get-StepEvents -RunId 'r-warn' -Id 'w' | Where-Object { $_.status -eq 'ok' })
    Assert-Equal 'odd' $wEv[0].data.warnings[0].code 'warn: the warning reached the trace as structured data'
    Assert-Equal 3 $wEv[0].data.warnings[0].data.lines[0] 'warn: nested warning data survives the trace round trip'

    # ------------------------------------------------ 14. missing step / no Invoke-Step
    $wf = Write-Workflow 'nostep.json' '{ "id": "spike.nostep", "setup": [ { "id": "x", "use": "fake.absent" } ] }'
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-nostep'
    Assert-Equal 'step_not_found' (Get-Rec $res 'x')['failure'] 'a use with no file is step_not_found'

    $wf = Write-Workflow 'noinvoke.json' '{ "id": "spike.noinvoke", "setup": [ { "id": "p", "use": "fake.pure" }, { "id": "x", "use": "fake.noinvoke" } ] }'
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-noinvoke'
    Assert-Equal 'step_not_found' (Get-Rec $res 'x')['failure'] 'a step file without Invoke-Step is step_not_found, not the previous step''s entry point'

    # ------------------------------------------------ 15. refusals happen before anything runs
    $wf = Write-Workflow 'each.json' '{ "id": "spike.each", "setup": [ { "id": "p", "use": "fake.pure" } ], "each": [ { "id": "q", "use": "fake.pure" } ] }'
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-each'
    Assert-Equal 'unsupported_in_spike' $res['failure'] 'each: refused'
    Assert-Equal 0 @($res['steps']).Count 'each: no step ran'
    Assert-Equal 0 @(Read-TraceEvents -WorkDir $work -RunId 'r-each').Count 'each: nothing was traced'

    $wf = Write-Workflow 'tmpl.json' '{ "id": "spike.tmpl", "setup": [ { "id": "p", "use": "fake.pure", "with": { "text": "{{item.key}}" } } ] }'
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-tmpl'
    Assert-Equal 'unsupported_in_spike' $res['failure'] 'template: refused rather than passed literally'

    $res = Invoke-EbiWorkflow -Path (Join-Path $tmpRoot 'does-not-exist.json') -WorkDir $work -ModulesRoot $modules -RunId 'r-nofile'
    Assert-Equal 'unsupported_in_spike' $res['failure'] 'a missing workflow file is reported, not thrown'

    $wf = Write-Workflow 'bad.json' '{ not json'
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $work -ModulesRoot $modules -RunId 'r-bad'
    Assert-Equal 'unsupported_in_spike' $res['failure'] 'malformed JSON is reported, not thrown'

    # ------------------------------------------------ 16. DryRun reaches the step and a null resource is allowed
    $res = Invoke-EbiWorkflow -Path (Join-Path $tmpRoot 'happy.json') -WorkDir $work -ModulesRoot $modules -RunId 'r-dry' -DryRun
    Assert-True ($res['ok']) 'dry: the happy workflow runs under DryRun'
    Assert-Equal 'True' ([string](Get-Rec $res 'shot')['outputs']['dryRun']) 'dry: $Ctx.DryRun is true inside the step'
    Assert-Equal -1 (Get-Rec $res 'shot')['outputs']['handle'] 'dry: a null resource is registered and handed on as null'
    Assert-True ($res['session'].Contains('mainWindow')) 'dry: the name is still registered'

    # ------------------------------------------------ 16b. relative paths (the first office-PC failure)
    # PowerShell's location and the process working directory are different
    # things; .NET file APIs use the latter. The runner must resolve every
    # path it is given against PowerShell's location before .NET sees it.
    Push-Location -LiteralPath $tmpRoot
    try {
        $res = Invoke-EbiWorkflow -Path 'happy.json' -WorkDir 'work-rel' -ModulesRoot 'modules' -RunId 'r-rel'
        Assert-True ($res['ok']) 'relative: a workflow path relative to the PowerShell location is read'
        Assert-Equal 3 @($res['steps']).Count 'relative: all three steps ran'
        $relWork = Join-Path $tmpRoot 'work-rel'
        Assert-True (Test-Path -LiteralPath (Get-TraceFile $relWork 'r-rel')) 'relative: a relative WorkDir lands under the PowerShell location, not the process cwd'
        Assert-True (((Get-Rec $res 'shot')['outputs']['path']) -eq 'capture/x.png') 'relative: with-values are untouched'
    } finally {
        Pop-Location
    }
    Assert-Equal '' (ConvertTo-EbiAbsolutePath '') 'absolute: empty stays empty'
    Assert-True ([System.IO.Path]::IsPathRooted((ConvertTo-EbiAbsolutePath 'does-not-exist-yet'))) 'absolute: a missing relative path is still made absolute'

    # ------------------------------------------------ 17. a run id is minted when none is given
    $res = Invoke-EbiWorkflow -Path (Join-Path $tmpRoot 'happy.json') -WorkDir $work -ModulesRoot $modules
    Assert-True ($res['runId'] -match '^\d{8}-\d{6}-[0-9a-f]{4}$') 'a minted run id is yyyyMMdd-HHmmss-xxxx'
    Assert-True (Test-Path -LiteralPath (Get-TraceFile $work $res['runId'])) 'the minted run id has its own trace file'

    # ------------------------------------------------ P0-08: the shipped spike workflow dry-runs on the real steps
    $repoRoot = Split-Path $here -Parent
    $shipped  = Join-Path (Join-Path $repoRoot 'workflows') 'spike.capture_window.json'
    Assert-True (Test-Path -LiteralPath $shipped) 'P0-08: workflows/spike.capture_window.json ships'
    $res = Invoke-EbiWorkflow -Path $shipped -WorkDir $work -RunId 'r-p008-dry' -DryRun
    Assert-True ($res['ok']) 'P0-08: prepare -> ensure -> capture dry-runs end to end on the real modules/ tree'
    Assert-Equal 3 @($res['steps']).Count 'P0-08: three step calls'
    Assert-Equal 'enter' (Get-Rec $res 'prepare')['outputs']['action'] 'P0-08: human.prepare answers Enter under DryRun without blocking'
    Assert-True ($res['session'].Contains('mainWindow')) 'P0-08: browser.ensure registered mainWindow (null handle under DryRun)'
    $shotPath = [string](Get-Rec $res 'shot')['outputs']['path']
    Assert-True ($shotPath.StartsWith($work)) 'P0-08: a relative saveAs resolves under the work dir'
    Assert-True ($shotPath.EndsWith('window.png')) 'P0-08: the PNG name is the one the workflow asked for'
    Assert-True (-not (Test-Path -LiteralPath $shotPath)) 'P0-08: DryRun wrote no file'
    $names = @(Get-ChildItem -LiteralPath (Join-Path $repoRoot 'modules') -Filter '*.ps1' -Recurse | ForEach-Object { $_.Name })
    foreach ($n in @('human.prepare.ps1', 'browser.ensure.ps1', 'screen.capture_window.ps1')) {
        Assert-True ($names -contains $n) ('P0-08: ' + $n + ' is in modules/')
    }

    # the step's own pure helper
    . (Join-Path (Join-Path (Join-Path $repoRoot 'modules') 'screen') 'screen.capture_window.ps1')
    $wd = [System.IO.Path]::GetTempPath().TrimEnd([System.IO.Path]::DirectorySeparatorChar)   # rooted on whichever OS runs the test
    Assert-Equal ([System.IO.Path]::GetFullPath([System.IO.Path]::Combine($wd, 'capture/a.png'))) (ScreenCaptureWindow-ResolvePath -SaveAs 'capture/a.png' -WorkDir $wd) 'capture_window: relative saveAs joins the work dir (separators normalized)'
    $rooted = Join-Path ([System.IO.Path]::GetTempPath()) 'a.png'   # rooted on whichever OS runs the test
    Assert-Equal $rooted (ScreenCaptureWindow-ResolvePath -SaveAs $rooted -WorkDir 'C:\w') 'capture_window: a rooted saveAs is kept'
    Assert-Equal '' (ScreenCaptureWindow-ResolvePath -SaveAs '' -WorkDir 'C:\w') 'capture_window: an empty saveAs stays empty'
    Assert-True ((ScreenCaptureWindow-ToHandle 4242) -eq [IntPtr]4242) 'capture_window: an int resource becomes an IntPtr'
    Assert-True ((ScreenCaptureWindow-ToHandle $null) -eq [IntPtr]::Zero) 'capture_window: a null resource is IntPtr.Zero'

    # ------------------------------------------------ the fixtures pass the real contract checker
    . (Join-Path $here 'StepContract.ps1')
    $specText = [System.IO.File]::ReadAllText((Join-Path (Split-Path $here -Parent) 'docs/ebi-dance/spec/STEP-CONTRACT.md'))
    $kinds = Get-StepContractMustReleaseKinds -Text $specText
    foreach ($name in @('fake.ensure', 'fake.capture', 'fake.release', 'fake.pure', 'fake.fail', 'fake.needbook', 'fake.noresource')) {
        $findings = @(Test-StepFileContract -Path (Join-Path (Join-Path $modules 'fake') ($name + '.ps1')) -MustRelease $kinds)
        Assert-Equal 0 $findings.Count ('fixture ' + $name + ' passes the step contract checker' + $(if ($findings.Count -gt 0) { ': ' + (($findings | ForEach-Object { $_['rule'] + ' ' + $_['message'] }) -join '; ') } else { '' }))
    }
} finally {
    if (Test-Path -LiteralPath $tmpRoot) { Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

$rc = Complete-Tests
exit $rc
