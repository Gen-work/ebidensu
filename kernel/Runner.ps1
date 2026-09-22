# ============================================================
#  kernel/Runner.ps1
#
#  P0-07 spike: the smallest runner that can execute a workflow JSON.
#  Dot-source only (no param()). ASCII source.
#
#  What it does (spec/WORKFLOW-SCHEMA.md 1-2, spec/STEP-CONTRACT.md 3):
#    - reads a workflow JSON (schema 1) and runs its 'setup' then its
#      'teardown', one step call after another, printing each result
#    - loads every referenced step from modules/<group>/<id>.ps1 by
#      dot-sourcing it INTO Invoke-EbiWorkflow's own scope and capturing
#      $Manifest + ${function:Invoke-Step} right away, before the next
#      file overwrites them. Step helpers (prefixed, STEP-CONTRACT 7)
#      stay defined in that scope, and every call happens from a child
#      scope of it, so they resolve. P1-02 (Registry) formalizes this.
#    - fills input defaults, roots 'path' inputs under WorkDir, checks
#      required inputs and rejects undeclared ones
#    - owns $Ctx.Session (STEP-CONTRACT 3.4 points 3 and 7): 'as' is
#      lifted out of 'with', checked for session_conflict, handed back
#      to the step as $In.as; the step's 'resource' return key is stored
#      under that name and stripped before anything is traced; a
#      type='session' input whose name is unknown fails session_missing
#      before the step runs; after a 'releases' step succeeds the name
#      is removed again
#    - writes one trace event per step call (kernel/Trace.ps1)
#    - runs 'teardown' on the three in-process exit paths
#      (WORKFLOW-SCHEMA 1.1) via try/finally
#    - treats an ok result carrying action='quit' as the reserved
#      'cancelled' failure (WORKFLOW-SCHEMA 5.2)
#
#  What it deliberately does NOT do yet: {{}} templates (P1-01), 'each'
#  and 'source' (P1-03), onError / ledger / resume (P1-04), 'needs'
#  checks, 'when'. A workflow relying on those fails loudly here rather
#  than half-working.
# ============================================================

. (Join-Path $PSScriptRoot 'Trace.ps1')

function Get-EbiReservedFailureIds {
    # spec/STEP-CONTRACT.md 3.1 (P0-R15). Produced by the runner, never
    # required in a manifest; a step may itself return session_invalid.
    return @('internal_error', 'needs_unmet', 'session_missing', 'session_invalid', 'session_conflict', 'cancelled')
}

function Get-EbiReturnReservedKeys {
    # Keys of a step's return hashtable that are protocol, not outputs.
    return @('ok', 'failure', 'message', 'warnings', 'resource')
}

function ConvertTo-EbiHashtable {
    # ConvertFrom-Json output (PSCustomObject graph) -> nested hashtables
    # and object[] arrays. P1-35 (kernel/Json.ps1) replaces this.
    #
    # Arrays are returned behind a leading comma so a one-element array
    # stays an array and an empty one stays empty when the caller assigns
    # the result -- PowerShell would otherwise unroll them.
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = ConvertTo-EbiHashtable $p.Value }
        return $h
    }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $Value.Keys) { $h[$k] = ConvertTo-EbiHashtable $Value[$k] }
        return $h
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $list = New-Object System.Collections.ArrayList
        foreach ($item in $Value) { [void]$list.Add((ConvertTo-EbiHashtable $item)) }
        return ,($list.ToArray())
    }
    return $Value
}

function Read-EbiWorkflow {
    # Load + shape-check a workflow file. Throws on anything the spike
    # cannot run; ebi lint (P1-08) will report these without running.
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw ('workflow not found: ' + $Path) }
    $raw = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding($false)))
    $obj = $raw | ConvertFrom-Json
    $wf  = ConvertTo-EbiHashtable $obj
    if (-not ($wf -is [hashtable])) { throw ('workflow root must be a JSON object: ' + $Path) }

    if (-not $wf.ContainsKey('schema')) { throw ('workflow has no "schema" field (WORKFLOW-SCHEMA.md 1, P0-R15): ' + $Path) }
    $schema = 0
    try { $schema = [int]$wf['schema'] } catch { throw ('workflow "schema" is not an integer: ' + $Path) }
    if ($schema -ne 1) { throw ('workflow schema ' + $schema + ' is not supported by this runner (supports 1): ' + $Path) }

    if (-not $wf.ContainsKey('id') -or [string]::IsNullOrWhiteSpace([string]$wf['id'])) { throw ('workflow has no "id": ' + $Path) }

    foreach ($seg in @('setup', 'each', 'teardown')) {
        if (-not $wf.ContainsKey($seg) -or $null -eq $wf[$seg]) { $wf[$seg] = @(); continue }
        $calls = $wf[$seg]
        if ($calls -is [hashtable]) { throw ('"' + $seg + '" must be an array of step calls: ' + $Path) }
        $seen = @{}
        foreach ($call in $calls) {
            if (-not ($call -is [hashtable])) { throw ('"' + $seg + '" holds a non-object entry: ' + $Path) }
            foreach ($req in @('id', 'use')) {
                if (-not $call.ContainsKey($req) -or [string]::IsNullOrWhiteSpace([string]$call[$req])) {
                    throw ('a step call in "' + $seg + '" has no "' + $req + '" (WORKFLOW-SCHEMA.md 2): ' + $Path)
                }
            }
            $cid = [string]$call['id']
            if ($seen.ContainsKey($cid)) { throw ('duplicate step id "' + $cid + '" in "' + $seg + '" (ledger key, WORKFLOW-SCHEMA.md 2.1): ' + $Path) }
            $seen[$cid] = $true
        }
    }
    return $wf
}

function New-EbiLogger {
    # $Ctx.Log with Info / Warn / Debug ScriptMethods (STEP-CONTRACT 3.2).
    # A PSObject with ScriptMethods rather than a class: no PowerShell
    # class in dot-sourced libraries (STEP-CONTRACT 1.1).
    param([bool]$Verbose = $false)
    $log = New-Object PSObject
    $log | Add-Member -MemberType NoteProperty -Name Verbose -Value $Verbose
    $log | Add-Member -MemberType ScriptMethod -Name Info  -Value { param($m) Write-Host ('    ' + [string]$m) -ForegroundColor DarkGray }
    $log | Add-Member -MemberType ScriptMethod -Name Warn  -Value { param($m) Write-Host ('    [WARN] ' + [string]$m) -ForegroundColor Yellow }
    $log | Add-Member -MemberType ScriptMethod -Name Debug -Value { param($m) if ($this.Verbose) { Write-Host ('    [dbg] ' + [string]$m) -ForegroundColor DarkGray } }
    return $log
}

function New-EbiContext {
    # The read-only $Ctx every step receives (STEP-CONTRACT 3.2). Session
    # is the one mutable member, and only the runner writes it (3.4).
    param(
        [string]$WorkDir,
        [string]$RunId,
        [bool]$DryRun = $false,
        [hashtable]$Profile = @{},
        [bool]$Verbose = $false
    )
    return @{
        WorkDir = $WorkDir
        RunId   = $RunId
        Profile = $Profile
        Log     = (New-EbiLogger -Verbose $Verbose)
        DryRun  = $DryRun
        Session = @{}
    }
}

function Resolve-EbiStepPath {
    # 'browser.find' -> <ModulesRoot>/browser/browser.find.ps1
    param([string]$ModulesRoot, [string]$StepId)
    $dot = $StepId.IndexOf('.')
    if ($dot -le 0) { throw ('step id "' + $StepId + '" is not <group>.<verb>') }
    $group = $StepId.Substring(0, $dot)
    return (Join-Path (Join-Path $ModulesRoot $group) ($StepId + '.ps1'))
}

function Get-EbiManifestList {
    # A manifest array field ('provides', 'releases', 'failures') as a flat
    # sequence -- explicit loop, never @($hashtable[$key]) (banned shape).
    #
    # Returned UNWRAPPED (no leading comma), exactly like Tests/StepContract.ps1
    # ConvertTo-StepContractArray: callers wrap the call in @(). A comma guard
    # here plus @() at the call site nests the array one level deeper, and
    # then '-contains' and 'is hashtable' silently see the wrong thing.
    param($Manifest, [string]$Key)
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $Manifest -or -not $Manifest.ContainsKey($Key)) { return $out.ToArray() }
    $v = $Manifest[$Key]
    if ($null -eq $v) { return $out.ToArray() }
    if ($v -is [string] -or $v -is [hashtable]) { [void]$out.Add($v); return $out.ToArray() }
    if ($v -is [System.Collections.IEnumerable]) { foreach ($i in $v) { [void]$out.Add($i) }; return $out.ToArray() }
    [void]$out.Add($v)
    return $out.ToArray()
}

function New-EbiStepFailure {
    param([string]$Failure, [string]$Message)
    return @{ ok = $false; failure = $Failure; message = $Message }
}

function Get-EbiStepInputs {
    # Build $In from the call's 'with' against the manifest's inputs:
    # defaults filled, required checked, undeclared rejected, 'path'
    # rooted under WorkDir, 'session' names checked for existence.
    # Returns @{ ok=$true; in=<hashtable> } or a failure record.
    param($Manifest, [hashtable]$With, $Ctx, $AsName, [bool]$Provides)

    $inputs = @{}
    if ($Manifest.ContainsKey('inputs') -and ($Manifest['inputs'] -is [hashtable])) { $inputs = $Manifest['inputs'] }

    foreach ($k in $With.Keys) {
        if (-not $inputs.ContainsKey([string]$k)) {
            return (New-EbiStepFailure 'internal_error' ('input "' + $k + '" is not declared by ' + [string]$Manifest['id'] + ' (ebi lint will catch this statically)'))
        }
    }

    $in = @{}
    foreach ($name in $inputs.Keys) {
        $spec = $inputs[$name]
        if (-not ($spec -is [hashtable])) { $spec = @{} }
        $has = $false; $val = $null
        if ($With.ContainsKey($name)) { $val = $With[$name]; $has = $true }
        elseif ($spec.ContainsKey('default')) { $val = $spec['default']; $has = $true }
        elseif ($spec.ContainsKey('required') -and [bool]$spec['required']) {
            return (New-EbiStepFailure 'internal_error' ('required input "' + $name + '" of ' + [string]$Manifest['id'] + ' was not given'))
        }
        if (-not $has) { continue }

        $type = ''
        if ($spec.ContainsKey('type')) { $type = [string]$spec['type'] }
        if ($type -eq 'path' -and ($val -is [string]) -and -not [string]::IsNullOrWhiteSpace($val)) {
            if (-not [System.IO.Path]::IsPathRooted($val)) { $val = Join-Path $Ctx.WorkDir $val }
        }
        if ($type -eq 'session') {
            $sname = [string]$val
            if (-not $Ctx.Session.ContainsKey($sname)) {
                return (New-EbiStepFailure 'session_missing' ('input "' + $name + '" names Session resource "' + $sname + '", which nothing registered with as: before this step'))
            }
        }
        $in[$name] = $val
    }
    if ($Provides) { $in['as'] = $AsName }   # $null when the call did not register (STEP-CONTRACT 3.4 point 7)
    return @{ ok = $true; in = $in }
}

function Invoke-EbiStepCall {
    # Run one step call. Returns a record:
    #   @{ id; use; status ('ok'|'fail'); failure; message; outputs; cancelled }
    param($Entry, [hashtable]$Call, $Ctx)

    $man = $Entry.Manifest
    $use = [string]$Call['use']
    $rec = @{ id = [string]$Call['id']; use = $use; status = 'fail'; failure = ''; message = ''; outputs = @{}; cancelled = $false }

    $with = @{}
    if ($Call.ContainsKey('with') -and ($Call['with'] -is [hashtable])) { $with = $Call['with'].Clone() }

    $provides = @(Get-EbiManifestList -Manifest $man -Key 'provides')
    $releases = @(Get-EbiManifestList -Manifest $man -Key 'releases')

    $asName = $null
    if ($with.ContainsKey('as')) {
        $asName = [string]$with['as']
        $with.Remove('as')
        if ($provides.Count -eq 0) {
            $rec.failure = 'internal_error'; $rec.message = ('"as" given but ' + $use + ' provides nothing (STEP-CONTRACT 3.4 point 3)')
            return $rec
        }
        if ([string]::IsNullOrWhiteSpace($asName)) {
            $rec.failure = 'internal_error'; $rec.message = '"as" must be a non-empty name'
            return $rec
        }
        if ($Ctx.Session.ContainsKey($asName)) {
            $rec.failure = 'session_conflict'; $rec.message = ('Session name "' + $asName + '" is still registered; release it before registering again (STEP-CONTRACT 3.4 point 3)')
            return $rec
        }
    }

    $prep = Get-EbiStepInputs -Manifest $man -With $with -Ctx $Ctx -AsName $asName -Provides ($provides.Count -gt 0)
    if (-not $prep.ok) { $rec.failure = $prep.failure; $rec.message = $prep.message; return $rec }
    $in = $prep.in

    $ret = $null
    try {
        $ret = & $Entry.Invoke $in $Ctx
    } catch {
        $ret = New-EbiStepFailure 'internal_error' ('step threw: ' + $_.Exception.Message)
    }

    if ($null -eq $ret -or -not ($ret -is [hashtable]) -or -not $ret.ContainsKey('ok')) {
        $rec.failure = 'internal_error'; $rec.message = ($use + ' did not return a hashtable with an ok key (STEP-CONTRACT 3.1)')
        return $rec
    }

    if ([bool]$ret['ok']) {
        if ($null -ne $asName) {
            if (-not $ret.ContainsKey('resource')) {
                $rec.failure = 'internal_error'; $rec.message = ($use + ' was asked to register "' + $asName + '" but returned no resource (STEP-CONTRACT 3.4 point 7)')
                return $rec
            }
            $Ctx.Session[$asName] = $ret['resource']
        } elseif ($ret.ContainsKey('resource')) {
            $rec.failure = 'internal_error'; $rec.message = ($use + ' returned a resource but the call had no "as"; unregistered resources must be released inside the step (STEP-CONTRACT 3.4 point 3)')
            return $rec
        }
        # Release side: drop the names this step freed.
        if ($releases.Count -gt 0 -and $man.ContainsKey('inputs') -and ($man['inputs'] -is [hashtable])) {
            foreach ($name in $man['inputs'].Keys) {
                $spec = $man['inputs'][$name]
                if (-not ($spec -is [hashtable])) { continue }
                if (-not $spec.ContainsKey('type') -or [string]$spec['type'] -ne 'session') { continue }
                $kind = ''
                if ($spec.ContainsKey('sessionKind')) { $kind = [string]$spec['sessionKind'] }
                if (($releases -contains $kind) -and $in.ContainsKey($name)) {
                    $Ctx.Session.Remove([string]$in[$name])
                }
            }
        }
        $rec.status = 'ok'
    } else {
        $fid = ''
        if ($ret.ContainsKey('failure')) { $fid = [string]$ret['failure'] }
        $known = New-Object System.Collections.ArrayList
        foreach ($f in @(Get-EbiManifestList -Manifest $man -Key 'failures')) {
            if ($f -is [hashtable] -and $f.ContainsKey('id')) { [void]$known.Add([string]$f['id']) }
        }
        foreach ($r in (Get-EbiReservedFailureIds)) { [void]$known.Add($r) }
        if ([string]::IsNullOrWhiteSpace($fid) -or -not ($known -contains $fid)) {
            $rec.failure = 'internal_error'
            $rec.message = ($use + ' returned failure "' + $fid + '", which is neither in its manifest nor reserved (STEP-CONTRACT 3.1)')
            return $rec
        }
        $rec.failure = $fid
        if ($ret.ContainsKey('message')) { $rec.message = [string]$ret['message'] }
    }

    $reserved = Get-EbiReturnReservedKeys
    $outputs = @{}
    foreach ($k in $ret.Keys) { if (-not ($reserved -contains [string]$k)) { $outputs[$k] = $ret[$k] } }
    $rec.outputs = $outputs
    if ($ret.ContainsKey('warnings') -and $null -ne $ret['warnings']) { $rec['warnings'] = $ret['warnings'] }

    if ($rec.status -eq 'ok' -and $outputs.ContainsKey('action') -and [string]$outputs['action'] -eq 'quit') {
        $rec.cancelled = $true
        $rec.failure = 'cancelled'
        $rec.message = ('operator chose quit at step "' + $rec.id + '"')
    }
    return $rec
}

function Write-EbiStepLine {
    param($Rec)
    $tag = if ($Rec.status -eq 'ok') { 'ok  ' } else { 'FAIL' }
    if ($Rec.cancelled) { $tag = 'QUIT' }
    $color = if ($Rec.status -eq 'ok' -and -not $Rec.cancelled) { 'Green' } else { 'Red' }
    $extra = ''
    if ($Rec.status -eq 'ok') {
        if ($Rec.outputs.Count -gt 0) { $extra = ($Rec.outputs | ConvertTo-Json -Compress -Depth 5) }
    } else {
        $extra = ($Rec.failure + ': ' + $Rec.message)
    }
    Write-Host ('  [' + $tag + '] ' + $Rec.id.PadRight(14) + ' ' + $Rec.use.PadRight(24) + ' ' + $extra) -ForegroundColor $color
}

function Invoke-EbiSegment {
    # Run one segment's calls in order; stop at the first failure or quit.
    # Returns $true when every call succeeded.
    param([string]$Segment, $Calls, $Steps, $Ctx, $Summary)

    if ($null -eq $Calls) { return $true }
    foreach ($call in $Calls) {
        $entry = $Steps[[string]$call['use']]
        $rec = Invoke-EbiStepCall -Entry $entry -Call $call -Ctx $Ctx
        $rec['segment'] = $Segment
        [void]$Summary.steps.Add($rec)
        Write-EbiStepLine -Rec $rec

        $data = @{ use = $rec.use; outputs = $rec.outputs }
        if ($rec.status -ne 'ok' -or $rec.cancelled) { $data['failure'] = $rec.failure }
        if ($rec.ContainsKey('warnings')) { $data['warnings'] = $rec['warnings'] }
        $status = if ($rec.cancelled) { 'skip' } elseif ($rec.status -eq 'ok') { 'ok' } else { 'fail' }
        Write-TraceEvent -WorkDir $Ctx.WorkDir -RunId $Ctx.RunId -Phase $Summary.workflow -Tags @{ segment = $Segment } `
            -Action $rec.id -Status $status -Message $rec.message -Data $data

        if ($rec.status -ne 'ok' -or $rec.cancelled) {
            if ($null -eq $Summary.failedStep) {
                $Summary.failedStep = $rec.id
                $Summary.failure    = $rec.failure
                $Summary.message    = $rec.message
                $Summary.cancelled  = $rec.cancelled
            }
            return $false
        }
    }
    return $true
}

function Invoke-EbiWorkflow {
    <#
      Run a workflow file. Returns a summary hashtable:
        ok, cancelled, runId, workflow, failedStep, failure, message,
        setupOk, teardownOk, steps (records, see Invoke-EbiStepCall), ctx
    #>
    param(
        [string]$Path,
        [string]$WorkDir,
        [string]$ModulesRoot = '',
        [string]$RunId = '',
        [bool]$DryRun = $false,
        [hashtable]$Profile = @{},
        [bool]$Verbose = $false
    )

    if ([string]::IsNullOrWhiteSpace($WorkDir)) { throw 'WorkDir is required' }
    if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = (Get-Date).ToString('yyyyMMdd-HHmmss') }
    if ([string]::IsNullOrWhiteSpace($ModulesRoot)) { $ModulesRoot = Join-Path (Split-Path $PSScriptRoot -Parent) 'modules' }

    $wf  = Read-EbiWorkflow -Path $Path
    $ctx = New-EbiContext -WorkDir $WorkDir -RunId $RunId -DryRun $DryRun -Profile $Profile -Verbose $Verbose
    $summary = @{
        ok = $false; cancelled = $false; runId = $RunId; workflow = [string]$wf['id']
        failedStep = $null; failure = $null; message = $null
        setupOk = $false; teardownOk = $true
        steps = (New-Object System.Collections.ArrayList); ctx = $ctx
    }

    # Load every referenced step ONCE, here, so its helpers live in this
    # scope for the whole run (see the header).
    $steps = @{}
    $allCalls = New-Object System.Collections.ArrayList
    foreach ($seg in @('setup', 'teardown')) { foreach ($c in $wf[$seg]) { [void]$allCalls.Add($c) } }
    foreach ($call in $allCalls) {
        $use = [string]$call['use']
        if ($steps.ContainsKey($use)) { continue }
        $stepPath = Resolve-EbiStepPath -ModulesRoot $ModulesRoot -StepId $use
        if (-not (Test-Path -LiteralPath $stepPath)) { throw ('step "' + $use + '" not found at ' + $stepPath) }
        $Manifest = $null
        # Forget the previous file's Invoke-Step so a step that forgot to
        # define its own cannot silently inherit it.
        if (Test-Path -LiteralPath 'function:Invoke-Step') { Remove-Item -LiteralPath 'function:Invoke-Step' }
        . $stepPath
        if ($null -eq $Manifest -or -not ($Manifest -is [hashtable])) { throw ('step "' + $use + '" defines no $Manifest hashtable') }
        $cmd = Get-Command -Name 'Invoke-Step' -CommandType Function -ErrorAction SilentlyContinue
        if ($null -eq $cmd) { throw ('step "' + $use + '" defines no Invoke-Step') }
        $invoke = $cmd.ScriptBlock
        $steps[$use] = @{ Manifest = $Manifest; Invoke = $invoke; Path = $stepPath }
    }

    $eachCount = 0
    foreach ($c in $wf['each']) { $eachCount++ }
    if ($eachCount -gt 0) { $ctx.Log.Warn('"each" has ' + $eachCount + ' step(s); the P0-07 spike runs setup/teardown only, each is skipped (P1-03)') }

    Write-Host ''
    Write-Host ('== ' + $summary.workflow + '  run ' + $RunId + $(if ($DryRun) { '  (dry run)' } else { '' })) -ForegroundColor Cyan
    Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase $summary.workflow -Action 'run' -Status 'start' `
        -Message $Path -Data @{ dryRun = $DryRun; modulesRoot = $ModulesRoot }

    try {
        $summary.setupOk = Invoke-EbiSegment -Segment 'setup' -Calls $wf['setup'] -Steps $steps -Ctx $ctx -Summary $summary
    } finally {
        # WORKFLOW-SCHEMA 1.1: teardown on every in-process exit path.
        $summary.teardownOk = Invoke-EbiSegment -Segment 'teardown' -Calls $wf['teardown'] -Steps $steps -Ctx $ctx -Summary $summary
    }

    $summary.ok = ($summary.setupOk -and $summary.teardownOk -and -not $summary.cancelled)
    $final = if ($summary.cancelled) { 'cancelled' } elseif ($summary.ok) { 'ok' } else { 'fail' }
    Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase $summary.workflow -Action 'run' -Status $final `
        -Message $(if ($summary.ok) { 'completed' } else { [string]$summary.message }) `
        -Data @{ steps = $summary.steps.Count; failedStep = $summary.failedStep; failure = $summary.failure }
    Write-Host ('== ' + $final.ToUpperInvariant() + '  ' + $summary.steps.Count + ' step call(s), trace: ' + (Get-TraceFile $WorkDir $RunId)) `
        -ForegroundColor $(if ($summary.ok) { 'Green' } else { 'Red' })
    return $summary
}
