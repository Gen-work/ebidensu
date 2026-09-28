#Requires -Version 5.1
# ============================================================
#  kernel/Runner.ps1
#
#  The ebi-dance workflow runner. Dot-source only (no param() block,
#  ASCII source, no class -- CLAUDE.md conventions).
#
#  P1-03: the main body. What runs:
#    - a workflow JSON (spec/WORKFLOW-SCHEMA.md 1-3), validated up front:
#      schema, id, section shapes, unique ids per section, source.select
#      forms, when forms, once forms; a bad workflow is refused as
#      'workflow_invalid' before any step runs
#    - "setup" once, "each" once per selected worklist row, "teardown"
#      once in a finally block (the 1.1 guarantee: normal end, a step
#      failure, cancelled and an unexpected exception all reach teardown;
#      Ctrl+C is explicitly NOT covered)
#    - {{...}} templates in "with" (kernel/Context.ps1): vars / profile /
#      page / run / item / steps scopes, expanded per call before the
#      inputs schema check; "as" and session names stay literal
#    - source: the worklist is a Session resource of kind 'worklist'
#      (registered by a setup step's with.as, P0-R11); rows are filtered
#      by kernel/Worklist.ps1's Select-EbiWorklistRows (the same function
#      table.select uses), ordered by groupBy / orderBy, limited
#    - "when" (Context.ps1's Test-EbiWhen): a skipped step is recorded as
#      status 'skipped' with every manifest output null plus skipped=true
#      (5.1), and later steps can reference it
#    - "once": "group": runs on the group's first item; its outputs are
#      replayed into the steps scope of every later item of the group
#      (6.3), no new record; "once": "groupEnd" is P1-04
#    - a step failure inside "each" ends THAT item (its remaining steps
#      are skipped) and the run continues with the next item; onError
#      policies are P1-04. A human step's operator_quit becomes the
#      reserved 'cancelled': no further items, teardown still runs (1.1)
#    - the resource channel of STEP-CONTRACT.md 3.4 point 7 and the 3.1
#      return contract, through kernel/Registry.ps1
#    - one trace event per step (kernel/Trace.ps1), phase = section,
#      key = the item's key display
#
#  Not here yet, refused loudly rather than half-done:
#    - "onError" (top-level or per step)   -> P1-04 (policies, ledger,
#      resume, once:groupEnd)
#
#  JSON comes and goes through kernel/Json.ps1 only (R8).
#
#  Hashtable access in this file is by index ($h['k']), never by dot:
#  under Set-StrictMode (which Tests/Run-Tests.ps1 turns on) a missing
#  key read with dot syntax throws.
# ============================================================

. (Join-Path $PSScriptRoot 'Trace.ps1')
. (Join-Path $PSScriptRoot 'Registry.ps1')
. (Join-Path $PSScriptRoot 'Context.ps1')
. (Join-Path $PSScriptRoot 'Worklist.ps1')

function Get-EbiWorkflowSchemaVersion { return 1 }

# --- runner-level reserved failure ids (STEP-CONTRACT.md 3.1 table) ---
function Get-EbiRunnerFailureIds {
    return @('internal_error', 'contract_violation', 'step_not_found', 'input_invalid',
             'session_missing', 'session_kind_mismatch', 'session_name_taken',
             'cancelled', 'workflow_invalid', 'unsupported_in_spike')
}

function New-EbiRunId {
    # yyyyMMdd-HHmmss plus 4 hex chars: sortable, and two runs started in
    # the same second still get their own run/<runId>/ directory.
    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $rand  = [guid]::NewGuid().ToString('N').Substring(0, 4)
    return ('{0}-{1}' -f $stamp, $rand)
}

function ConvertTo-EbiAbsolutePath {
    # PowerShell-relative -> absolute, using PowerShell's current location
    # (not the process working directory, see Invoke-EbiWorkflow). An
    # existing path is resolved through the provider so '.\x' and 'x/../y'
    # collapse; a not-yet-existing one (a fresh WorkDir) is joined by hand.
    # Empty stays empty.
    param([string]$PathValue)
    if ([string]::IsNullOrWhiteSpace($PathValue)) { return $PathValue }
    if (Test-Path -LiteralPath $PathValue) {
        return (Resolve-Path -LiteralPath $PathValue).ProviderPath
    }
    if ([System.IO.Path]::IsPathRooted($PathValue)) { return $PathValue }
    $base = (Get-Location).ProviderPath
    return [System.IO.Path]::GetFullPath((Join-Path $base $PathValue))
}

function Test-EbiValueHasTemplate {
    # True when a value (or anything nested in it) contains '{{'.
    param($Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [string]) { return ($Value.Contains('{{')) }
    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($k in $Value.Keys) { if (Test-EbiValueHasTemplate $Value[$k]) { return $true } }
        return $false
    }
    if ($Value -is [System.Collections.IList]) {
        foreach ($item in $Value) { if (Test-EbiValueHasTemplate $item) { return $true } }
        return $false
    }
    return $false
}

function Get-EbiWorkflowProblems {
    <#
      PURE. Everything wrong with a workflow's SHAPE, found before anything
      runs (WORKFLOW-SCHEMA.md 1-3, 5, 7; the runner-side half of ebi
      lint's list in 9). Returns an array of one-line strings; empty means
      the workflow can run. Manifest-dependent checks (unknown use, bad
      with) are per call, at run time -- they need the step loaded.
    #>
    param([hashtable]$Workflow)
    $problems = New-Object System.Collections.ArrayList
    if ($null -eq $Workflow) { [void]$problems.Add('workflow is empty'); return $problems.ToArray() }

    if (-not $Workflow.Contains('schema') -or $null -eq $Workflow['schema']) {
        [void]$problems.Add('top-level "schema" is required (write "schema": ' + (Get-EbiWorkflowSchemaVersion) + ')')
    } else {
        $sv = $Workflow['schema']
        $isInt = ($sv -is [int] -or $sv -is [long] -or (($sv -is [double]) -and [math]::Floor([double]$sv) -eq [double]$sv))
        if (-not $isInt) { [void]$problems.Add('"schema" must be an integer') }
        elseif ([int]$sv -gt (Get-EbiWorkflowSchemaVersion)) { [void]$problems.Add(('"schema" {0} is newer than this runner supports ({1})' -f [int]$sv, (Get-EbiWorkflowSchemaVersion))) }
    }
    if (-not $Workflow.Contains('id') -or [string]::IsNullOrWhiteSpace([string]$Workflow['id'])) {
        [void]$problems.Add('top-level "id" is required')
    }
    if ($Workflow.Contains('onError') -and $null -ne $Workflow['onError']) {
        [void]$problems.Add('top-level "onError" is not supported yet (P1-04); the runner stops the item at the first failure')
    }
    if ($Workflow.Contains('vars') -and $null -ne $Workflow['vars'] -and -not ($Workflow['vars'] -is [System.Collections.IDictionary])) {
        [void]$problems.Add('"vars" must be an object')
    }

    $groupBy = ''
    if ($Workflow.Contains('source') -and $null -ne $Workflow['source']) {
        $src = $Workflow['source']
        if (-not ($src -is [System.Collections.IDictionary])) {
            [void]$problems.Add('"source" must be an object')
        } else {
            if (-not $src.Contains('table') -or [string]::IsNullOrWhiteSpace([string]$src['table'])) {
                [void]$problems.Add('source.table is required: the session name a setup step registered the worklist under (with.as)')
            }
            if (-not $src.Contains('select') -or -not ($src['select'] -is [System.Collections.IDictionary])) {
                [void]$problems.Add('source.select is required: { "field": ..., "pendingWhen": ... }')
            } else {
                $sel = $src['select']
                $pw = if ($sel.Contains('pendingWhen')) { [string]$sel['pendingWhen'] } else { '' }
                $parsed = ConvertFrom-EbiPendingWhen -Text $pw
                if (-not $parsed['ok']) { [void]$problems.Add('source.select: ' + $parsed['message']) }
                elseif ($parsed['kind'] -ne 'always' -and (-not $sel.Contains('field') -or [string]::IsNullOrWhiteSpace([string]$sel['field']))) {
                    [void]$problems.Add('source.select.field is required unless pendingWhen is "always"')
                }
            }
            if ($src.Contains('groupBy') -and $null -ne $src['groupBy']) { $groupBy = [string]$src['groupBy'] }
            if ($src.Contains('limit') -and $null -ne $src['limit'] -and -not ($src['limit'] -is [int] -or $src['limit'] -is [long] -or $src['limit'] -is [double])) {
                [void]$problems.Add('source.limit must be a number')
            }
        }
    } elseif ($Workflow.Contains('each') -and $null -ne $Workflow['each']) {
        [void]$problems.Add('"each" needs a "source" to iterate')
    }

    foreach ($section in @('setup', 'each', 'teardown')) {
        if (-not $Workflow.Contains($section) -or $null -eq $Workflow[$section]) { continue }
        $calls = $Workflow[$section]
        if (-not ($calls -is [System.Collections.IList])) {
            [void]$problems.Add(('"{0}" must be an array of step calls' -f $section)); continue
        }
        $seen = @{}
        $n = 0
        foreach ($call in $calls) {
            $n++
            if (-not ($call -is [System.Collections.IDictionary])) {
                [void]$problems.Add(('{0}[{1}]: a step call must be an object' -f $section, $n)); continue
            }
            $id  = if ($call.Contains('id'))  { [string]$call['id'] }  else { '' }
            $use = if ($call.Contains('use')) { [string]$call['use'] } else { '' }
            $where = if ($id -ne '') { $section + '/' + $id } else { ('{0}[{1}]' -f $section, $n) }
            if ([string]::IsNullOrWhiteSpace($id))  { [void]$problems.Add(('{0}[{1}]: "id" is required' -f $section, $n)) }
            if ([string]::IsNullOrWhiteSpace($use)) { [void]$problems.Add(('{0}[{1}]: "use" is required' -f $section, $n)) }
            if ($id -ne '' -and $seen.Contains($id)) { [void]$problems.Add(('{0}: step id "{1}" is used twice' -f $section, $id)) }
            $seen[$id] = $true
            if ($call.Contains('onError') -and $null -ne $call['onError']) {
                [void]$problems.Add(('{0}: "onError" is not supported yet (P1-04)' -f $where))
            }
            if ($call.Contains('with') -and $null -ne $call['with']) {
                if (-not ($call['with'] -is [System.Collections.IDictionary])) {
                    [void]$problems.Add(('{0}: "with" must be an object' -f $where))
                } elseif ($call['with'].Contains('as') -and (Test-EbiValueHasTemplate $call['with']['as'])) {
                    [void]$problems.Add(('{0}: "as" is a literal session name, not a template (STEP-CONTRACT 3.4 point 3)' -f $where))
                }
            }
            if ($call.Contains('when') -and $null -ne $call['when']) {
                $w = ConvertFrom-EbiWhen -Text ([string]$call['when'])
                if (-not $w['ok']) { [void]$problems.Add(('{0}: {1}' -f $where, $w['message'])) }
            }
            if ($call.Contains('once') -and $null -ne $call['once']) {
                $once = [string]$call['once']
                if ($once -ne 'group' -and $once -ne 'groupEnd') {
                    [void]$problems.Add(('{0}: "once" must be "group" or "groupEnd", got "{1}"' -f $where, $once))
                } elseif ($section -ne 'each') {
                    [void]$problems.Add(('{0}: "once" is only meaningful in "each"' -f $where))
                } elseif ($groupBy -eq '') {
                    [void]$problems.Add(('{0}: "once": "{1}" needs source.groupBy' -f $where, $once))
                } elseif ($once -eq 'groupEnd') {
                    [void]$problems.Add(('{0}: "once": "groupEnd" is not supported yet (P1-04)' -f $where))
                }
            }
        }
    }
    return $problems.ToArray()
}

function New-EbiLog {
    # $Ctx.Log.Info(...) / .Warn(...) / .Debug(...) per STEP-CONTRACT.md 3.2.
    # ScriptMethods on a PSObject so the method-call syntax holds on PS 5.1.
    $log = New-Object PSObject
    $log | Add-Member -MemberType ScriptMethod -Name Info  -Value { param($m) Write-Host ('    [info ] {0}' -f $m) -ForegroundColor Gray }
    $log | Add-Member -MemberType ScriptMethod -Name Warn  -Value { param($m) Write-Host ('    [warn ] {0}' -f $m) -ForegroundColor Yellow }
    $log | Add-Member -MemberType ScriptMethod -Name Debug -Value { param($m) Write-Verbose ('[debug] {0}' -f $m) }
    return $log
}

function New-EbiContext {
    param([string]$WorkDir, [string]$RunId, [bool]$DryRun, [hashtable]$Profile = @{})
    if ($null -eq $Profile) { $Profile = @{} }
    return @{
        WorkDir = $WorkDir
        RunId   = $RunId
        Profile = $Profile
        Log     = (New-EbiLog)
        DryRun  = $DryRun
        Session = @{}          # name -> @{ kind; value; registeredBy }
    }
}

function Write-EbiStepLine {
    param([string]$Status, [string]$Section, [string]$Id, [string]$Use, [string]$Detail, [string]$Key = '')
    $color = switch ($Status) { 'ok' { 'Green' } 'fail' { 'Red' } 'skip' { 'DarkGray' } 'skipped' { 'DarkGray' } default { 'Gray' } }
    $where = if ($Key -ne '') { $Section + '[' + $Key + ']/' + $Id } else { $Section + '/' + $Id }
    Write-Host ('  [{0,-4}] {1}  {2}  {3}' -f $Status, $where, $Use, $Detail) -ForegroundColor $color
}

function Format-EbiOutputs {
    param([hashtable]$Outputs)
    if ($null -eq $Outputs -or $Outputs.Count -eq 0) { return '' }
    $parts = New-Object System.Collections.ArrayList
    foreach ($k in ($Outputs.Keys | Sort-Object)) {
        $v = $Outputs[$k]
        $text = if ($null -eq $v) { 'null' } elseif ($v -is [string] -or $v.GetType().IsPrimitive) { [string]$v } else { '...' }
        if ($text.Length -gt 40) { $text = $text.Substring(0, 37) + '...' }
        [void]$parts.Add(('{0}={1}' -f $k, $text))
    }
    return ('{' + ($parts.ToArray() -join ' ') + '}')
}

function New-EbiSkippedOutputs {
    # WORKFLOW-SCHEMA 5.1: every manifest output present and null, plus
    # skipped = $true. Without a manifest (step not loaded) just the flag.
    param($Manifest)
    $out = @{ skipped = $true }
    if ($null -ne $Manifest -and $Manifest.Contains('outputs') -and ($Manifest['outputs'] -is [System.Collections.IDictionary])) {
        foreach ($k in $Manifest['outputs'].Keys) { if ([string]$k -ne 'skipped') { $out[[string]$k] = $null } }
    }
    return $out
}

function Invoke-EbiStepCall {
    <#
      One step call, start to finish: load -> when -> expand templates ->
      inputs schema + session names -> Invoke-Step -> return contract ->
      register / release -> record + trace. Returns the record:
        @{ section; id; use; item; group; status; failure; message;
           outputs; warnings }
      status: ok | fail | skipped (by when). The caller decides what a
      failure means for the section (stop setup, end the item, ...).

      MUST be dot-sourced by the caller (. Invoke-EbiStepCall ...): it
      dot-sources Import-EbiStep, and the step's helper functions have to
      land in the scope that outlives this call (Registry.ps1 header).

      $State: the run's shared state (ctx, registry, workDir, runId,
      workflowId). $Scope: the template scope for this call. $Steps: the
      steps hashtable the outputs are written into (id -> outputs), the
      same object the scope's 'steps' points at.
    #>
    param([hashtable]$State, [string]$Section, [hashtable]$Call, [hashtable]$Scope, [hashtable]$Steps, [string]$Key = '', [string]$Group = '')

    $ebiCall = @{
        ctx = $State['ctx']; registry = $State['registry']; workDir = $State['workDir']; runId = $State['runId']
        id = [string]$Call['id']; use = [string]$Call['use']
    }
    $ebiCall['tags'] = @{ workflow = $State['workflowId']; step = $ebiCall['id']; use = $ebiCall['use'] }
    if ($Group -ne '') { $ebiCall['tags']['group'] = $Group }
    $ebiCall['rec'] = @{ section = $Section; id = $ebiCall['id']; use = $ebiCall['use']; item = $Key; group = $Group; status = 'ok'; failure = ''; message = ''; outputs = @{}; warnings = @() }

    Write-TraceEvent -WorkDir $ebiCall['workDir'] -RunId $ebiCall['runId'] -Phase $Section -Key $Key -Tags $ebiCall['tags'] -Action 'step' -Status 'start'

    # -- load (once per use) ------------------------------------------------
    if ($null -eq (Get-EbiStep -Registry $ebiCall['registry'] -Use $ebiCall['use'])) {
        $ebiCall['imported'] = . Import-EbiStep -Registry $ebiCall['registry'] -Use $ebiCall['use']
        if (-not $ebiCall['imported']['ok']) {
            $ebiCall['rec']['status'] = 'fail'; $ebiCall['rec']['failure'] = [string]$ebiCall['imported']['failure']; $ebiCall['rec']['message'] = [string]$ebiCall['imported']['message']
            Write-EbiStepLine -Status 'fail' -Section $Section -Id $ebiCall['id'] -Use $ebiCall['use'] -Key $Key -Detail ($ebiCall['rec']['failure'] + ': ' + $ebiCall['rec']['message'])
            Write-TraceEvent -WorkDir $ebiCall['workDir'] -RunId $ebiCall['runId'] -Phase $Section -Key $Key -Tags $ebiCall['tags'] -Action 'step' -Status 'fail' -Message $ebiCall['rec']['message'] -Data @{ failure = $ebiCall['rec']['failure'] }
            return $ebiCall['rec']
        }
    }
    $ebiCall['entry']    = Get-EbiStep -Registry $ebiCall['registry'] -Use $ebiCall['use']
    $ebiCall['manifest'] = $ebiCall['entry']['Manifest']

    # -- when ---------------------------------------------------------------
    if ($Call.Contains('when') -and $null -ne $Call['when']) {
        $ebiCall['when'] = ConvertFrom-EbiWhen -Text ([string]$Call['when'])
        if (-not (Test-EbiWhen -Parsed $ebiCall['when'] -Scope $Scope)) {
            $ebiCall['rec']['status'] = 'skipped'; $ebiCall['rec']['message'] = ('when "' + [string]$Call['when'] + '" is false')
            $ebiCall['rec']['outputs'] = New-EbiSkippedOutputs -Manifest $ebiCall['manifest']
            $Steps[$ebiCall['id']] = $ebiCall['rec']['outputs']
            Write-EbiStepLine -Status 'skipped' -Section $Section -Id $ebiCall['id'] -Use $ebiCall['use'] -Key $Key -Detail $ebiCall['rec']['message']
            Write-TraceEvent -WorkDir $ebiCall['workDir'] -RunId $ebiCall['runId'] -Phase $Section -Key $Key -Tags $ebiCall['tags'] -Action 'step' -Status 'skipped' -Message $ebiCall['rec']['message'] -Data @{ outputs = $ebiCall['rec']['outputs'] }
            return $ebiCall['rec']
        }
    }

    # -- templates, then inputs: strip "as", schema check, session names ----
    $ebiCall['with'] = if ($Call.Contains('with') -and $null -ne $Call['with']) { [hashtable]$Call['with'] } else { @{} }
    $ebiCall['expanded'] = Expand-EbiTemplate -Value $ebiCall['with'] -Scope $Scope
    if (-not $ebiCall['expanded']['ok']) {
        $ebiCall['rec']['status'] = 'fail'; $ebiCall['rec']['failure'] = 'input_invalid'
        $ebiCall['rec']['message'] = ('template {{' + [string]$ebiCall['expanded']['path'] + '}}: ' + [string]$ebiCall['expanded']['message'])
        Write-EbiStepLine -Status 'fail' -Section $Section -Id $ebiCall['id'] -Use $ebiCall['use'] -Key $Key -Detail ($ebiCall['rec']['failure'] + ': ' + $ebiCall['rec']['message'])
        Write-TraceEvent -WorkDir $ebiCall['workDir'] -RunId $ebiCall['runId'] -Phase $Section -Key $Key -Tags $ebiCall['tags'] -Action 'step' -Status 'fail' -Message $ebiCall['rec']['message'] -Data @{ failure = $ebiCall['rec']['failure']; template = [string]$ebiCall['expanded']['path'] }
        return $ebiCall['rec']
    }
    $ebiCall['withValues'] = [hashtable]$ebiCall['expanded']['value']
    $ebiCall['resolved'] = Resolve-EbiStepInputs -Manifest $ebiCall['manifest'] -With $ebiCall['withValues'] -Session $ebiCall['ctx']['Session']
    if (-not $ebiCall['resolved']['ok']) {
        $ebiCall['rec']['status'] = 'fail'; $ebiCall['rec']['failure'] = [string]$ebiCall['resolved']['failure']; $ebiCall['rec']['message'] = [string]$ebiCall['resolved']['message']
        Write-EbiStepLine -Status 'fail' -Section $Section -Id $ebiCall['id'] -Use $ebiCall['use'] -Key $Key -Detail ($ebiCall['rec']['failure'] + ': ' + $ebiCall['rec']['message'])
        Write-TraceEvent -WorkDir $ebiCall['workDir'] -RunId $ebiCall['runId'] -Phase $Section -Key $Key -Tags $ebiCall['tags'] -Action 'step' -Status 'fail' -Message $ebiCall['rec']['message'] -Data @{ failure = $ebiCall['rec']['failure'] }
        return $ebiCall['rec']
    }

    # -- call -----------------------------------------------------------------
    $ebiCall['ret'] = $null
    $ebiCall['threw'] = ''
    try {
        $ebiCall['ret'] = & $ebiCall['entry']['Invoke'] $ebiCall['resolved']['In'] $ebiCall['ctx']
    } catch {
        $ebiCall['threw'] = $_.Exception.Message
    }

    if ($ebiCall['threw'] -ne '') {
        $ebiCall['rec']['status'] = 'fail'; $ebiCall['rec']['failure'] = 'internal_error'; $ebiCall['rec']['message'] = $ebiCall['threw']
    } else {
        $ebiCall['check'] = Test-EbiStepReturn -Manifest $ebiCall['manifest'] -Return $ebiCall['ret'] -WantsResource ($ebiCall['resolved']['As'] -ne '')
        if (-not $ebiCall['check']['ok']) {
            $ebiCall['rec']['status'] = 'fail'; $ebiCall['rec']['failure'] = [string]$ebiCall['check']['failure']; $ebiCall['rec']['message'] = [string]$ebiCall['check']['message']
        } else {
            $ebiCall['rec']['outputs']  = $ebiCall['check']['Outputs']
            $ebiCall['rec']['warnings'] = $ebiCall['check']['Warnings']
            if (-not [bool]$ebiCall['ret']['ok']) {
                $ebiCall['rec']['status']  = 'fail'
                $ebiCall['rec']['failure'] = [string]$ebiCall['ret']['failure']
                $ebiCall['rec']['message'] = if ($ebiCall['ret'].Contains('message')) { [string]$ebiCall['ret']['message'] } else { '' }
                # A human step's q: the reserved id, bypassing onError (5.2).
                if ($ebiCall['rec']['failure'] -eq 'operator_quit') {
                    $ebiCall['rec']['failure'] = 'cancelled'
                    if ($ebiCall['rec']['message'] -eq '') { $ebiCall['rec']['message'] = 'operator quit' }
                }
            } else {
                # 3.4 point 7: register / release. "resource" never reaches
                # outputs, trace or the records.
                if ($ebiCall['resolved']['As'] -ne '') {
                    $ebiCall['provides'] = @(Get-EbiManifestArray -Manifest $ebiCall['manifest'] -Key 'provides')
                    $ebiCall['ctx']['Session'][$ebiCall['resolved']['As']] = @{ kind = $ebiCall['provides'][0]; value = $ebiCall['check']['Resource']; registeredBy = $ebiCall['id'] }
                }
                $ebiCall['releases'] = @(Get-EbiManifestArray -Manifest $ebiCall['manifest'] -Key 'releases')
                if ($ebiCall['releases'].Count -gt 0) {
                    $ebiCall['sessionInputs'] = Get-EbiSessionInputs -Manifest $ebiCall['manifest']
                    foreach ($ebiName in $ebiCall['sessionInputs'].Keys) {
                        if (-not $ebiCall['withValues'].Contains($ebiName)) { continue }
                        if ($ebiCall['releases'] -contains [string]$ebiCall['sessionInputs'][$ebiName]) {
                            $ebiCall['ctx']['Session'].Remove([string]$ebiCall['withValues'][$ebiName])
                        }
                    }
                }
                $Steps[$ebiCall['id']] = $ebiCall['rec']['outputs']
            }
        }
    }

    $ebiCall['detail'] = if ($ebiCall['rec']['status'] -eq 'ok') { Format-EbiOutputs $ebiCall['rec']['outputs'] } else { $ebiCall['rec']['failure'] + ': ' + $ebiCall['rec']['message'] }
    $ebiCall['warnCount'] = @($ebiCall['rec']['warnings']).Count
    if ($ebiCall['warnCount'] -gt 0) { $ebiCall['detail'] += ('  ({0} warning{1})' -f $ebiCall['warnCount'], $(if ($ebiCall['warnCount'] -eq 1) { '' } else { 's' })) }
    Write-EbiStepLine -Status $ebiCall['rec']['status'] -Section $Section -Id $ebiCall['id'] -Use $ebiCall['use'] -Key $Key -Detail $ebiCall['detail']
    $ebiCall['data'] = @{ outputs = $ebiCall['rec']['outputs'] }
    if ($ebiCall['warnCount'] -gt 0) { $ebiCall['data']['warnings'] = $ebiCall['rec']['warnings'] }
    if ($ebiCall['rec']['status'] -ne 'ok') { $ebiCall['data']['failure'] = $ebiCall['rec']['failure'] }
    Write-TraceEvent -WorkDir $ebiCall['workDir'] -RunId $ebiCall['runId'] -Phase $Section -Key $Key -Tags $ebiCall['tags'] -Action 'step' -Status $ebiCall['rec']['status'] -Message $ebiCall['rec']['message'] -Data $ebiCall['data']
    return $ebiCall['rec']
}

function New-EbiSkipRecord {
    # A step that is not run because an earlier one failed (setup, or the
    # rest of an item). Recorded and traced as 'skip'.
    param([string]$Section, [hashtable]$Call, [string]$Key, [string]$Group, [string]$Why, [hashtable]$State)
    $id = [string]$Call['id']; $use = [string]$Call['use']
    $rec = @{ section = $Section; id = $id; use = $use; item = $Key; group = $Group; status = 'skip'; failure = ''; message = $Why; outputs = @{}; warnings = @() }
    Write-EbiStepLine -Status 'skip' -Section $Section -Id $id -Use $use -Key $Key -Detail $Why
    $tags = @{ workflow = $State['workflowId']; step = $id; use = $use }
    Write-TraceEvent -WorkDir $State['workDir'] -RunId $State['runId'] -Phase $Section -Key $Key -Tags $tags -Action 'step' -Status 'skip' -Message $Why
    return $rec
}

function Invoke-EbiWorkflow {
    <#
      Run a workflow JSON.

      Parameters:
        Path, WorkDir, ModulesRoot, RunId, DryRun   as before
        Profile     the loaded profile data (hashtable: worklist /
                    vocabulary / pages / grammar / rules ...). Loading it
                    from profiles/<name>/ is P2-01; until then the caller
                    passes it (ebi.ps1 passes @{}).
        Vars        overrides for the workflow's "vars" (CLI --var)
        Operator    run.operator (default $env:USERNAME)
        Only        key display strings; only these rows are processed
        Limit       row cap (0 = the workflow's own source.limit)

      Returns a hashtable:
        ok          every step call returned ok (or was when-skipped)
        runId       run/<runId>/trace.jsonl holds the events
        workflowId
        failure     '' or the first failure id ('workflow_invalid' when
                    the workflow was refused; 'cancelled' when the
                    operator quit)
        message
        steps       one record per step call, in execution order:
                    @{ section; id; use; item; group; status
                       ('ok'|'fail'|'skip'|'skipped'); failure; message;
                       outputs; warnings }
        items       one record per selected row: @{ key; group; status
                    ('ok'|'fail'|'cancelled'|'skip'); failure; message }
        selected    how many rows were processed; total: rows in the table
        session     the live $Ctx.Session (in-process only)
    #>
    param(
        [string]$Path,
        [string]$WorkDir,
        [string]$ModulesRoot = '',
        [string]$RunId = '',
        [switch]$DryRun,
        [hashtable]$Profile = @{},
        [hashtable]$Vars = @{},
        [string]$Operator = '',
        $Only = $null,
        [int]$Limit = 0
    )

    $dryRunFlag = [bool]$DryRun.IsPresent
    if ([string]::IsNullOrWhiteSpace($ModulesRoot)) { $ModulesRoot = Get-EbiDefaultModulesRoot }
    if ([string]::IsNullOrWhiteSpace($RunId))       { $RunId = New-EbiRunId }
    if ([string]::IsNullOrWhiteSpace($WorkDir))     { $WorkDir = (Get-Location).ProviderPath }
    if ([string]::IsNullOrWhiteSpace($Operator))    { $Operator = [string]$env:USERNAME }
    if ($null -eq $Profile) { $Profile = @{} }
    if ($null -eq $Vars) { $Vars = @{} }
    # Everything below hands paths to .NET ([IO.File], [IO.Path]::Combine in
    # steps, Trace.ps1's AppendAllText). .NET resolves a relative path
    # against the PROCESS working directory, which PowerShell never syncs
    # with Set-Location -- on PS 5.1 that is wherever the console started.
    # So '.\workflows\x.json' passed Test-Path and then failed ReadAllText
    # on the first office-PC run. Absolute from here on.
    $Path        = ConvertTo-EbiAbsolutePath $Path
    $WorkDir     = ConvertTo-EbiAbsolutePath $WorkDir
    $ModulesRoot = ConvertTo-EbiAbsolutePath $ModulesRoot

    $result = @{
        ok = $false; runId = $RunId; workflowId = ''; failure = ''; message = ''
        steps = @(); items = @(); selected = 0; total = 0; session = @{}
    }
    $records = New-Object System.Collections.ArrayList
    $itemRecords = New-Object System.Collections.ArrayList

    # ---- read + validate the shape ---------------------------------------------
    $read = Read-EbiJson -Path $Path     # kernel/Json.ps1: UTF-8, hashtables, failure as a record
    if (-not $read['ok']) {
        $result['failure'] = 'workflow_invalid'
        $result['message'] = ('cannot read workflow: {0}' -f $read['message'])
        Write-Host ('  [refused] {0}' -f $result['message']) -ForegroundColor Red
        return $result
    }
    $workflow = $read['value']
    if (-not ($workflow -is [hashtable])) {
        $result['failure'] = 'workflow_invalid'; $result['message'] = 'workflow JSON must be an object'
        Write-Host ('  [refused] {0}' -f $result['message']) -ForegroundColor Red
        return $result
    }
    $result['workflowId'] = [string]$workflow['id']

    $problems = @(Get-EbiWorkflowProblems -Workflow $workflow)
    if ($problems.Count -gt 0) {
        $result['failure'] = 'workflow_invalid'
        $result['message'] = ($problems -join '; ')
        Write-Host ('  [refused] {0}' -f $result['message']) -ForegroundColor Red
        return $result
    }

    # ---- context, scope parts ---------------------------------------------------
    $ctx = New-EbiContext -WorkDir $WorkDir -RunId $RunId -DryRun $dryRunFlag -Profile $Profile
    $result['session'] = $ctx['Session']
    $wfId = [string]$workflow['id']
    $wfVars = @{}
    if ($workflow.Contains('vars') -and $workflow['vars'] -is [System.Collections.IDictionary]) { foreach ($k in $workflow['vars'].Keys) { $wfVars[[string]$k] = $workflow['vars'][$k] } }
    foreach ($k in $Vars.Keys) { $wfVars[[string]$k] = $Vars[$k] }
    $pageName = if ($workflow.Contains('page') -and $null -ne $workflow['page']) { [string]$workflow['page'] } else { '' }
    $run = @{ runId = $RunId; startedAt = (Get-Date).ToString('o'); operator = $Operator; workDir = $WorkDir; timeWindow = $null }
    $source = if ($workflow.Contains('source') -and $null -ne $workflow['source']) { [hashtable]$workflow['source'] } else { $null }
    $keyColumns = @(Get-EbiWorklistKeyColumns -Profile $Profile)
    if ($null -ne $source -and $source.Contains('keyColumns') -and $null -ne $source['keyColumns']) {
        $keyColumns = @(Get-EbiManifestArray -Manifest $source -Key 'keyColumns')
    }
    $groupColumn = Get-EbiWorklistGroupColumn -Profile $Profile
    $groupBy = if ($null -ne $source -and $source.Contains('groupBy') -and $null -ne $source['groupBy']) { [string]$source['groupBy'] } else { '' }
    if ($groupColumn -eq '' -and $groupBy -ne '') { $groupColumn = $groupBy }

    $state = @{
        ctx = $ctx; registry = (New-EbiRegistry -ModulesRoot $ModulesRoot)
        workDir = $WorkDir; runId = $RunId; workflowId = $wfId
    }
    $tagsBase = @{ workflow = $wfId }

    Write-Host ('  run {0}  workflow {1}{2}' -f $RunId, $wfId, $(if ($dryRunFlag) { '  (dry run)' } else { '' })) -ForegroundColor Cyan
    Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'run' -Tags $tagsBase -Action 'run' -Status 'start' -Message $Path

    $stopped   = $false     # setup failed or the run was cancelled: no more each
    $cancelled = $false
    try {
        # ---- setup ------------------------------------------------------------------
        if ($workflow.Contains('setup') -and $null -ne $workflow['setup']) {
            $setupSteps = @{}
            $scope = New-EbiTemplateScope -Vars $wfVars -Profile $Profile -PageName $pageName -Run $run -Item $null -Steps $setupSteps -KeyColumns $keyColumns -GroupColumn $groupColumn
            foreach ($call in $workflow['setup']) {
                if ($stopped) { [void]$records.Add((New-EbiSkipRecord -Section 'setup' -Call $call -Key '' -Group '' -Why 'earlier step failed' -State $state)); continue }
                $rec = . Invoke-EbiStepCall -State $state -Section 'setup' -Call $call -Scope $scope -Steps $setupSteps
                [void]$records.Add($rec)
                if ($rec['status'] -eq 'fail') { $stopped = $true; if ($rec['failure'] -eq 'cancelled') { $cancelled = $true } }
            }
        }

        # ---- each -------------------------------------------------------------------
        if (-not $stopped -and $null -ne $source -and $workflow.Contains('each') -and $null -ne $workflow['each']) {
            $tableName = [string]$source['table']
            if (-not $ctx['Session'].Contains($tableName) -or [string]$ctx['Session'][$tableName]['kind'] -ne 'worklist') {
                $why = if ($ctx['Session'].Contains($tableName)) { ('source.table "{0}" is a "{1}", not a worklist' -f $tableName, [string]$ctx['Session'][$tableName]['kind']) } else { ('source.table "{0}" is not registered; a setup step must load the worklist with "as": "{0}"' -f $tableName) }
                [void]$records.Add(@{ section = 'each'; id = '(source)'; use = ''; item = ''; group = ''; status = 'fail'; failure = $(if ($ctx['Session'].Contains($tableName)) { 'session_kind_mismatch' } else { 'session_missing' }); message = $why; outputs = @{}; warnings = @() })
                Write-Host ('  [fail] each/(source)  ' + $why) -ForegroundColor Red
                Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'each' -Tags $tagsBase -Action 'source' -Status 'fail' -Message $why
                $stopped = $true
            } else {
                $table = $ctx['Session'][$tableName]['value']
                $rows = @()
                if ($null -ne $table -and ($table -is [System.Collections.IDictionary]) -and $table.Contains('rows') -and $null -ne $table['rows']) { $rows = @($table['rows']) }
                $sel = $source['select']
                $field = if ($sel.Contains('field')) { [string]$sel['field'] } else { '' }
                $spec = Get-EbiWorklistColumnSpec -Profile $Profile -Field $field
                $limitValue = if ($Limit -gt 0) { $Limit } elseif ($source.Contains('limit') -and $null -ne $source['limit']) { [int]$source['limit'] } else { 0 }
                $selected = Select-EbiWorklistRows -Rows $rows -Field $field -PendingWhen ([string]$sel['pendingWhen']) -Spec $spec -Only $Only -KeyColumns $keyColumns -Limit 0
                $result['total'] = $selected['total']
                $orderBy = if ($source.Contains('orderBy') -and $null -ne $source['orderBy']) { [string]$source['orderBy'] } else { '' }
                $ordered = @(Sort-EbiWorklistRows -Rows $selected['rows'] -GroupBy $groupBy -OrderBy $orderBy -KeyColumns $keyColumns)
                if ($limitValue -gt 0 -and $ordered.Count -gt $limitValue) { $ordered = @($ordered[0..($limitValue - 1)]) }
                $result['selected'] = $ordered.Count
                Write-Host ('  each: {0} of {1} row{2} selected' -f $ordered.Count, $selected['total'], $(if ($selected['total'] -eq 1) { '' } else { 's' })) -ForegroundColor Cyan
                Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'each' -Tags $tagsBase -Action 'source' -Status 'ok' -Data @{ total = $selected['total']; selected = $ordered.Count; groupBy = $groupBy }

                $groupOutputs = @{}    # group -> @{ stepId -> outputs } from once:group steps
                foreach ($row in $ordered) {
                    if ($stopped) { break }
                    $key   = Get-EbiKeyDisplay -Item $row -KeyColumns $keyColumns
                    $group = Get-EbiWorklistGroupOf -Row $row -GroupBy $groupBy
                    $steps = @{}
                    if ($groupBy -ne '' -and $groupOutputs.Contains($group)) { foreach ($k in $groupOutputs[$group].Keys) { $steps[$k] = $groupOutputs[$group][$k] } }
                    $scope = New-EbiTemplateScope -Vars $wfVars -Profile $Profile -PageName $pageName -Run $run -Item $row -Steps $steps -KeyColumns $keyColumns -GroupColumn $groupColumn
                    $itemRec = @{ key = $key; group = $group; status = 'ok'; failure = ''; message = '' }
                    Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'each' -Key $key -Tags $tagsBase -Action 'item' -Status 'start' -Data @{ group = $group }
                    $itemFailed = $false
                    foreach ($call in $workflow['each']) {
                        $once = if ($call.Contains('once') -and $null -ne $call['once']) { [string]$call['once'] } else { '' }
                        if ($itemFailed) {
                            [void]$records.Add((New-EbiSkipRecord -Section 'each' -Call $call -Key $key -Group $group -Why 'earlier step of this item failed' -State $state)); continue
                        }
                        if ($once -eq 'group') {
                            if (-not $groupOutputs.Contains($group)) { $groupOutputs[$group] = @{} }
                            if ($groupOutputs[$group].Contains([string]$call['id'])) {
                                # Already ran for this group: outputs are in $steps (replayed
                                # above); no new record, one trace line (6.3).
                                Write-EbiStepLine -Status 'skip' -Section 'each' -Id ([string]$call['id']) -Use ([string]$call['use']) -Key $key -Detail ('once: group -- replayed for group "' + $group + '"')
                                Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'each' -Key $key -Tags @{ workflow = $wfId; step = [string]$call['id']; use = [string]$call['use']; group = $group } -Action 'step' -Status 'skip' -Message ('once: group -- replayed for group "' + $group + '"')
                                continue
                            }
                        }
                        $rec = . Invoke-EbiStepCall -State $state -Section 'each' -Call $call -Scope $scope -Steps $steps -Key $key -Group $group
                        [void]$records.Add($rec)
                        if ($rec['status'] -eq 'fail') {
                            $itemFailed = $true
                            $itemRec['status'] = 'fail'; $itemRec['failure'] = $rec['failure']; $itemRec['message'] = ($rec['id'] + ': ' + $rec['message'])
                            if ($rec['failure'] -eq 'cancelled') { $itemRec['status'] = 'cancelled'; $cancelled = $true; $stopped = $true }
                            continue
                        }
                        if ($once -eq 'group') { $groupOutputs[$group][[string]$call['id']] = $rec['outputs'] }
                    }
                    [void]$itemRecords.Add($itemRec)
                    Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'each' -Key $key -Tags $tagsBase -Action 'item' -Status $itemRec['status'] -Message $itemRec['message'] -Data @{ group = $group; failure = $itemRec['failure'] }
                }
                if ($cancelled) {
                    # The rows never reached are listed, so the summary says
                    # what the cancel left undone.
                    $done = $itemRecords.Count
                    foreach ($row in $ordered) {
                        if ($done -gt 0) { $done--; continue }
                        [void]$itemRecords.Add(@{ key = (Get-EbiKeyDisplay -Item $row -KeyColumns $keyColumns); group = (Get-EbiWorklistGroupOf -Row $row -GroupBy $groupBy); status = 'skip'; failure = ''; message = 'run cancelled' })
                    }
                }
            }
        }
    } finally {
        # ---- teardown: every exit path above lands here (1.1) ------------------
        # An unexpected exception inside the loops above (a runner bug, not a
        # step failure -- those are caught per call) still reaches teardown;
        # the exception itself propagates after this block.
        if ($workflow.Contains('teardown') -and $null -ne $workflow['teardown']) {
            $teardownSteps = @{}
            $scope = New-EbiTemplateScope -Vars $wfVars -Profile $Profile -PageName $pageName -Run $run -Item $null -Steps $teardownSteps -KeyColumns $keyColumns -GroupColumn $groupColumn
            foreach ($call in $workflow['teardown']) {
                $rec = . Invoke-EbiStepCall -State $state -Section 'teardown' -Call $call -Scope $scope -Steps $teardownSteps
                [void]$records.Add($rec)
            }
        }

        $result['steps'] = $records.ToArray()
        $result['items'] = $itemRecords.ToArray()
        $failed = @($records.ToArray() | Where-Object { $_['status'] -eq 'fail' })
        $result['ok'] = ($failed.Count -eq 0)
        if ($failed.Count -gt 0) {
            $result['failure'] = [string]$failed[0]['failure']
            $where = if ([string]$failed[0]['item'] -ne '') { $failed[0]['section'] + '[' + $failed[0]['item'] + ']/' + $failed[0]['id'] } else { $failed[0]['section'] + '/' + $failed[0]['id'] }
            $result['message'] = ('{0}: {1}' -f $where, $failed[0]['message'])
        }
        if ($cancelled) { $result['failure'] = 'cancelled' }
        $left = @($ctx['Session'].Keys)
        $itemsFailed = @($itemRecords.ToArray() | Where-Object { $_['status'] -eq 'fail' }).Count
        Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'run' -Tags $tagsBase -Action 'run' -Status $(if ($result['ok']) { 'ok' } else { 'fail' }) -Message $result['message'] -Data @{ steps = $records.Count; failed = $failed.Count; items = $itemRecords.Count; itemsFailed = $itemsFailed; cancelled = $cancelled; sessionLeft = $left }
        Write-Host ('  run {0}  {1}  ({2} step{3}, {4} failed{5}{6})' -f $RunId, $(if ($result['ok']) { 'OK' } elseif ($cancelled) { 'CANCELLED' } else { 'FAIL' }), $records.Count, $(if ($records.Count -eq 1) { '' } else { 's' }), $failed.Count, $(if ($itemRecords.Count -gt 0) { ('; ' + $itemRecords.Count + ' item(s), ' + $itemsFailed + ' failed') } else { '' }), $(if ($left.Count -gt 0) { '; still registered: ' + ($left -join ', ') } else { '' })) -ForegroundColor $(if ($result['ok']) { 'Green' } else { 'Red' })
    }
    return $result
}
