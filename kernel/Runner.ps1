#Requires -Version 5.1
# ============================================================
#  kernel/Runner.ps1
#
#  The ebi-dance workflow runner. Dot-source only (no param() block,
#  ASCII source, no class -- CLAUDE.md conventions).
#
#  P1-03 main body + P1-04 fault tolerance and resume. What runs:
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
#      (6.3), no new record. "once": "groupEnd": runs once after the
#      group's last item (7.2), even when that item was skipped by a
#      policy -- a registered group resource must still be released
#    - onError (6): retry (only a failure the manifest marks transient;
#      backoff doubles; exhausted -> ask) / ask (r retry, s skip this
#      item, q cancel) / skip / fail (abort the run); byFailure overrides
#      per failure id; per-call onError overrides the top-level one;
#      default ask. A human step's operator_quit bypasses all of it and
#      becomes the reserved 'cancelled'. A destructive step gets a
#      confirm gate first unless the call says "confirm": false (6.2).
#      Asking goes through one handler (-AskHandler; kernel/Gate.ps1 is
#      P1-05); a dry run answers itself and never blocks
#    - the ledger (kernel/Ledger.ps1): every completed or when-skipped
#      "each" step is appended to run/<runId>/ledger.jsonl WITH its
#      outputs; -Resume replays those outputs instead of re-running the
#      step, except for provides / releases steps, which always run
#      again so Session resources come back (STEP-CONTRACT 6.2). setup
#      and teardown re-run on every resume. run/<runId>/run.json holds
#      the run.* scope and the arguments (P0-R16)
#    - warnings ride in the trace and are summarised at the end of the
#      run (3.1)
#    - the resource channel of STEP-CONTRACT.md 3.4 point 7 and the 3.1
#      return contract, through kernel/Registry.ps1
#    - one trace event per step (kernel/Trace.ps1), phase = section,
#      key = the item's key display
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
. (Join-Path $PSScriptRoot 'Ledger.ps1')
. (Join-Path $PSScriptRoot 'Gate.ps1')

function Get-EbiWorkflowSchemaVersion { return 1 }

# --- runner-level reserved failure ids (STEP-CONTRACT.md 3.1 table) ---
function Get-EbiRunnerFailureIds {
    return @('internal_error', 'contract_violation', 'step_not_found', 'input_invalid',
             'session_missing', 'session_kind_mismatch', 'session_name_taken',
             'cancelled', 'workflow_invalid')
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
        foreach ($pr in @(Test-EbiOnErrorShape -OnError $Workflow['onError'] -Where 'onError')) { [void]$problems.Add($pr) }
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
                foreach ($pr in @(Test-EbiOnErrorShape -OnError $call['onError'] -Where ($where + '.onError'))) { [void]$problems.Add($pr) }
            }
            if ($call.Contains('confirm') -and $null -ne $call['confirm'] -and -not ($call['confirm'] -is [bool])) {
                [void]$problems.Add(('{0}: "confirm" must be true or false' -f $where))
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
                }
            }
        }
    }
    return $problems.ToArray()
}

function Get-EbiErrorPolicies { return @('retry', 'ask', 'skip', 'fail') }

function Test-EbiOnErrorShape {
    # PURE. Problems with an onError object (WORKFLOW-SCHEMA 6): policy in
    # the four, times / backoffMs numbers, byFailure a map of the same.
    param($OnError, [string]$Where)
    $out = New-Object System.Collections.ArrayList
    if (-not ($OnError -is [System.Collections.IDictionary])) { [void]$out.Add($Where + ' must be an object'); return $out.ToArray() }
    if ($OnError.Contains('policy') -and $null -ne $OnError['policy'] -and (Get-EbiErrorPolicies) -notcontains [string]$OnError['policy']) {
        [void]$out.Add(('{0}.policy "{1}" is not one of retry, ask, skip, fail' -f $Where, [string]$OnError['policy']))
    }
    foreach ($num in @('times', 'backoffMs')) {
        if ($OnError.Contains($num) -and $null -ne $OnError[$num] -and -not (Test-EbiIsNumber $OnError[$num])) {
            [void]$out.Add(('{0}.{1} must be a number' -f $Where, $num))
        }
    }
    if ($OnError.Contains('byFailure') -and $null -ne $OnError['byFailure']) {
        if (-not ($OnError['byFailure'] -is [System.Collections.IDictionary])) { [void]$out.Add($Where + '.byFailure must be an object keyed by failure id') }
        else {
            foreach ($fid in $OnError['byFailure'].Keys) {
                foreach ($pr in @(Test-EbiOnErrorShape -OnError $OnError['byFailure'][$fid] -Where ($Where + '.byFailure.' + [string]$fid))) { [void]$out.Add($pr) }
            }
        }
    }
    return $out.ToArray()
}

function Test-EbiFailureTransient {
    # Is this failure id marked transient in the manifest? Reserved
    # runner ids are never transient (STEP-CONTRACT 3.1).
    param($Manifest, [string]$FailureId)
    if ((Get-EbiRunnerFailureIds) -contains $FailureId) { return $false }
    if ($null -eq $Manifest -or -not $Manifest.Contains('failures') -or $null -eq $Manifest['failures']) { return $false }
    foreach ($f in $Manifest['failures']) {
        if ($f -is [System.Collections.IDictionary] -and $f.Contains('id') -and [string]$f['id'] -eq $FailureId) {
            return ($f.Contains('transient') -and ($f['transient'] -is [bool]) -and $f['transient'])
        }
    }
    return $false
}

function Resolve-EbiErrorPolicy {
    <#
      PURE. The policy for one failure of one call (WORKFLOW-SCHEMA 6, 6.0):
      the call's onError wins over the workflow's; inside each, byFailure
      for this id wins over policy. Default ask. retry on a non-transient
      failure degrades to ask right away (P0-R5).
      Returns @{ policy; times; backoffMs; transient; source }.
    #>
    param($Workflow, $Call, [string]$FailureId, $Manifest)
    $picked = @{ policy = 'ask'; times = 3; backoffMs = 800; source = 'default' }
    foreach ($level in @(@{ obj = $(if ($null -ne $Workflow -and $Workflow.Contains('onError')) { $Workflow['onError'] } else { $null }); name = 'onError' },
                         @{ obj = $(if ($null -ne $Call -and $Call.Contains('onError')) { $Call['onError'] } else { $null }); name = 'step onError' })) {
        $o = $level['obj']
        if (-not ($o -is [System.Collections.IDictionary])) { continue }
        $chosen = $o
        $src = $level['name']
        if ($o.Contains('byFailure') -and ($o['byFailure'] -is [System.Collections.IDictionary]) -and $o['byFailure'].Contains($FailureId) -and ($o['byFailure'][$FailureId] -is [System.Collections.IDictionary])) {
            $chosen = $o['byFailure'][$FailureId]; $src = $level['name'] + '.byFailure.' + $FailureId
        }
        if ($chosen.Contains('policy') -and $null -ne $chosen['policy']) { $picked['policy'] = [string]$chosen['policy']; $picked['source'] = $src }
        if ($chosen.Contains('times') -and $null -ne $chosen['times']) { $picked['times'] = [int]$chosen['times'] }
        if ($chosen.Contains('backoffMs') -and $null -ne $chosen['backoffMs']) { $picked['backoffMs'] = [int]$chosen['backoffMs'] }
    }
    $picked['transient'] = Test-EbiFailureTransient -Manifest $Manifest -FailureId $FailureId
    if ($picked['policy'] -eq 'retry' -and -not $picked['transient']) { $picked['policy'] = 'ask'; $picked['source'] = $picked['source'] + ' (retry refused: "' + $FailureId + '" is not transient)' }
    return $picked
}

function Invoke-EbiDefaultAsk {
    <#
      The built-in answerer for -AskHandler when none is given: a plain
      console prompt (kernel/Gate.ps1, P1-05, replaces the rendering).
      It never blocks when nobody can answer: under DryRun, or when the
      console's input is redirected (CI, a scheduled run), a confirm is
      yes and an error is skip, and the line says so.
      Question shapes -- the contract P1-05 renders:
        @{ kind='error';   section; id; use; key; group; failure; message; attempt; transient; policy } -> 'r' | 's' | 'q'
        @{ kind='confirm'; section; id; use; key; group; effects; with }                                -> 'y' | 'n' | 'q'
    #>
    param([hashtable]$Question, [bool]$DryRun)
    # kernel/Gate.ps1 (P1-05) renders the panel and applies the same
    # never-block rule; the plain prompt below is only the fallback when
    # Gate.ps1 is not loaded (it always is by this file).
    if (Test-Path -LiteralPath 'function:Invoke-EbiGateAsk') { return (Invoke-EbiGateAsk -Question $Question -DryRun $DryRun) }
    $auto = $DryRun
    $autoWhy = 'dry run'
    if (-not $auto) {
        try { if ([Console]::IsInputRedirected) { $auto = $true; $autoWhy = 'no console to ask' } } catch { }
    }
    $where = if ([string]$Question['key'] -ne '') { $Question['section'] + '[' + $Question['key'] + ']/' + $Question['id'] } else { $Question['section'] + '/' + $Question['id'] }
    if ([string]$Question['kind'] -eq 'confirm') {
        Write-Host ''
        Write-Host ('  CONFIRM  {0}  ({1}) is destructive' -f $where, $Question['use']) -ForegroundColor Yellow
        if ($auto) { Write-Host ('  (' + $autoWhy + ': yes)') -ForegroundColor DarkGray; return 'y' }
        Write-Host '  y=do it / n=skip this item / q=cancel the run : ' -ForegroundColor Magenta -NoNewline
        $a = ([string](Read-Host)).Trim().ToLowerInvariant()
        if ($a -eq 'y') { return 'y' }; if ($a -eq 'q') { return 'q' }; return 'n'
    }
    Write-Host ''
    Write-Host ('  FAILED   {0}  ({1}): {2}: {3}' -f $where, $Question['use'], $Question['failure'], $Question['message']) -ForegroundColor Yellow
    if ($auto) { Write-Host ('  (' + $autoWhy + ': skip)') -ForegroundColor DarkGray; return 's' }
    Write-Host '  r=retry / s=skip this item / q=cancel the run : ' -ForegroundColor Magenta -NoNewline
    $a = ([string](Read-Host)).Trim().ToLowerInvariant()
    if ($a -eq 'r') { return 'r' }; if ($a -eq 'q') { return 'q' }; return 's'
}

function Invoke-EbiStepWithPolicy {
    <#
      One call under fault tolerance: ledger replay, the destructive
      confirm gate, then attempts under the resolved onError policy.
      MUST be dot-sourced (it dot-sources Invoke-EbiStepCall).

      Returns @{ outcome; records; last } where outcome is
        ok        ran (or replayed from the ledger) and succeeded
        skipped   when was false
        skip      given up on this ITEM: policy skip, operator s, confirm n
        fail      policy fail: abort the run
        cancelled operator q, or the step's operator_quit
      and records holds every attempt's record (plus a 'skip' record when
      the item was given up), last the final one.
    #>
    param([hashtable]$State, [string]$Section, [hashtable]$Call, [hashtable]$Scope, [hashtable]$Steps, [string]$Key = '', [string]$Group = '', [string]$LedgerGroup = '')

    $ebiPol = @{ records = (New-Object System.Collections.ArrayList); id = [string]$Call['id']; use = [string]$Call['use']; attempt = 0 }
    $ebiPol['ctx'] = $State['ctx']

    # -- ledger replay (STEP-CONTRACT 6.1 / 6.2) ----------------------------
    if ($Section -eq 'each' -and $State['ledgerDone'].Count -gt 0) {
        if ($null -eq (Get-EbiStep -Registry $State['registry'] -Use $ebiPol['use'])) { [void](. Import-EbiStep -Registry $State['registry'] -Use $ebiPol['use']) }
        $ebiPol['entry'] = Get-EbiStep -Registry $State['registry'] -Use $ebiPol['use']
        if ($null -ne $ebiPol['entry']) {
            $ebiPol['ledgerKey'] = Get-EbiLedgerKey -Item $Key -Group $LedgerGroup -Step $ebiPol['id']
            $ebiPol['resourceStep'] = ((@(Get-EbiManifestArray -Manifest $ebiPol['entry']['Manifest'] -Key 'provides')).Count -gt 0) -or ((@(Get-EbiManifestArray -Manifest $ebiPol['entry']['Manifest'] -Key 'releases')).Count -gt 0)
            if ($State['ledgerDone'].Contains($ebiPol['ledgerKey']) -and -not $ebiPol['resourceStep']) {
                $ebiPol['old'] = $State['ledgerDone'][$ebiPol['ledgerKey']]
                $ebiPol['outputs'] = $(if ($ebiPol['old'].Contains('outputs') -and ($ebiPol['old']['outputs'] -is [hashtable])) { $ebiPol['old']['outputs'] } else { @{} })
                $Steps[$ebiPol['id']] = $ebiPol['outputs']
                $ebiPol['rec'] = @{ section = $Section; id = $ebiPol['id']; use = $ebiPol['use']; item = $Key; group = $Group; status = 'replayed'; failure = ''; message = ('replayed from the ledger (' + [string]$ebiPol['old']['status'] + ')'); outputs = $ebiPol['outputs']; warnings = @() }
                [void]$ebiPol['records'].Add($ebiPol['rec'])
                Write-EbiStepLine -Status 'skip' -Section $Section -Id $ebiPol['id'] -Use $ebiPol['use'] -Key $Key -Detail $ebiPol['rec']['message']
                Write-TraceEvent -WorkDir $State['workDir'] -RunId $State['runId'] -Phase $Section -Key $Key -Tags @{ workflow = $State['workflowId']; step = $ebiPol['id']; use = $ebiPol['use'] } -Action 'step' -Status 'replayed' -Message $ebiPol['rec']['message'] -Data @{ outputs = $ebiPol['outputs'] }
                return @{ outcome = 'ok'; records = $ebiPol['records'].ToArray(); last = $ebiPol['rec']; replayed = $true }
            }
        }
    }

    # -- destructive: confirm first (WORKFLOW-SCHEMA 6.2) --------------------
    if ($null -eq (Get-EbiStep -Registry $State['registry'] -Use $ebiPol['use'])) { [void](. Import-EbiStep -Registry $State['registry'] -Use $ebiPol['use']) }
    $ebiPol['entry'] = Get-EbiStep -Registry $State['registry'] -Use $ebiPol['use']
    if ($null -ne $ebiPol['entry']) {
        $ebiPol['effects'] = $(if ($ebiPol['entry']['Manifest'].Contains('effects')) { [string]$ebiPol['entry']['Manifest']['effects'] } else { '' })
        $ebiPol['confirmOff'] = ($Call.Contains('confirm') -and ($Call['confirm'] -is [bool]) -and -not $Call['confirm'])
        if ($ebiPol['effects'] -eq 'destructive' -and -not $ebiPol['confirmOff']) {
            $ebiPol['shown'] = Expand-EbiTemplate -Value $(if ($Call.Contains('with') -and $null -ne $Call['with']) { $Call['with'] } else { @{} }) -Scope $Scope
            $ebiPol['answer'] = [string](& $State['ask'] @{ kind = 'confirm'; section = $Section; id = $ebiPol['id']; use = $ebiPol['use']; key = $Key; group = $Group; effects = 'destructive'; with = $(if ($ebiPol['shown']['ok']) { $ebiPol['shown']['value'] } else { $Call['with'] }) })
            Write-TraceEvent -WorkDir $State['workDir'] -RunId $State['runId'] -Phase $Section -Key $Key -Tags @{ workflow = $State['workflowId']; step = $ebiPol['id']; use = $ebiPol['use'] } -Action 'confirm' -Status $ebiPol['answer'] -Message 'destructive step: operator confirmation'
            if ($ebiPol['answer'] -eq 'q') {
                $ebiPol['rec'] = @{ section = $Section; id = $ebiPol['id']; use = $ebiPol['use']; item = $Key; group = $Group; status = 'fail'; failure = 'cancelled'; message = 'operator cancelled at the destructive confirm gate'; outputs = @{}; warnings = @() }
                [void]$ebiPol['records'].Add($ebiPol['rec'])
                Write-EbiStepLine -Status 'fail' -Section $Section -Id $ebiPol['id'] -Use $ebiPol['use'] -Key $Key -Detail 'cancelled: at the confirm gate'
                return @{ outcome = 'cancelled'; records = $ebiPol['records'].ToArray(); last = $ebiPol['rec']; replayed = $false }
            }
            if ($ebiPol['answer'] -ne 'y') {
                $ebiPol['rec'] = @{ section = $Section; id = $ebiPol['id']; use = $ebiPol['use']; item = $Key; group = $Group; status = 'skip'; failure = ''; message = 'operator declined the destructive step; item left pending'; outputs = @{}; warnings = @() }
                [void]$ebiPol['records'].Add($ebiPol['rec'])
                Write-EbiStepLine -Status 'skip' -Section $Section -Id $ebiPol['id'] -Use $ebiPol['use'] -Key $Key -Detail $ebiPol['rec']['message']
                Write-TraceEvent -WorkDir $State['workDir'] -RunId $State['runId'] -Phase $Section -Key $Key -Tags @{ workflow = $State['workflowId']; step = $ebiPol['id']; use = $ebiPol['use'] } -Action 'step' -Status 'skip' -Message $ebiPol['rec']['message']
                return @{ outcome = 'skip'; records = $ebiPol['records'].ToArray(); last = $ebiPol['rec']; replayed = $false }
            }
        }
    }

    # -- attempts -------------------------------------------------------------
    while ($true) {
        $ebiPol['attempt']++
        $ebiPol['rec'] = . Invoke-EbiStepCall -State $State -Section $Section -Call $Call -Scope $Scope -Steps $Steps -Key $Key -Group $Group
        $ebiPol['rec']['attempt'] = $ebiPol['attempt']
        [void]$ebiPol['records'].Add($ebiPol['rec'])
        if ($ebiPol['rec']['status'] -eq 'ok' -or $ebiPol['rec']['status'] -eq 'skipped') {
            if ($Section -eq 'each') {
                [void](Add-EbiLedgerRecord -Path $State['ledgerFile'] -Record (New-EbiLedgerRecord -RunId $State['runId'] -Item $Key -Group $LedgerGroup -Step $ebiPol['id'] -Status $ebiPol['rec']['status'] -Outputs $ebiPol['rec']['outputs']))
            }
            return @{ outcome = $ebiPol['rec']['status']; records = $ebiPol['records'].ToArray(); last = $ebiPol['rec']; replayed = $false }
        }
        if ($ebiPol['rec']['failure'] -eq 'cancelled') {
            return @{ outcome = 'cancelled'; records = $ebiPol['records'].ToArray(); last = $ebiPol['rec']; replayed = $false }
        }
        $ebiPol['manifest'] = $(if ($null -ne (Get-EbiStep -Registry $State['registry'] -Use $ebiPol['use'])) { (Get-EbiStep -Registry $State['registry'] -Use $ebiPol['use'])['Manifest'] } else { $null })
        $ebiPol['policy'] = Resolve-EbiErrorPolicy -Workflow $State['workflow'] -Call $Call -FailureId $ebiPol['rec']['failure'] -Manifest $ebiPol['manifest']
        $ebiPol['rec']['policy'] = $ebiPol['policy']['policy']
        Write-TraceEvent -WorkDir $State['workDir'] -RunId $State['runId'] -Phase $Section -Key $Key -Tags @{ workflow = $State['workflowId']; step = $ebiPol['id']; use = $ebiPol['use'] } -Action 'onError' -Status $ebiPol['policy']['policy'] -Message ('attempt ' + $ebiPol['attempt'] + ': ' + $ebiPol['rec']['failure'] + ' -> ' + $ebiPol['policy']['policy'] + ' (' + $ebiPol['policy']['source'] + ')') -Data @{ attempt = $ebiPol['attempt']; failure = $ebiPol['rec']['failure']; transient = $ebiPol['policy']['transient'] }

        $ebiPol['decision'] = $ebiPol['policy']['policy']
        if ($ebiPol['decision'] -eq 'retry') {
            if ($ebiPol['attempt'] -le [int]$ebiPol['policy']['times']) {
                $ebiPol['wait'] = [int]$ebiPol['policy']['backoffMs'] * [math]::Pow(2, $ebiPol['attempt'] - 1)
                Write-Host ('  [retry] {0}: attempt {1} of {2} failed ({3}); waiting {4} ms' -f $ebiPol['id'], $ebiPol['attempt'], ([int]$ebiPol['policy']['times'] + 1), $ebiPol['rec']['failure'], [int]$ebiPol['wait']) -ForegroundColor DarkYellow
                if ($ebiPol['wait'] -gt 0) { Start-Sleep -Milliseconds ([int]$ebiPol['wait']) }
                continue
            }
            $ebiPol['decision'] = 'ask'    # exhausted (WORKFLOW-SCHEMA 6)
        }
        if ($ebiPol['decision'] -eq 'ask') {
            $ebiPol['answer'] = [string](& $State['ask'] @{ kind = 'error'; section = $Section; id = $ebiPol['id']; use = $ebiPol['use']; key = $Key; group = $Group; failure = $ebiPol['rec']['failure']; message = $ebiPol['rec']['message']; attempt = $ebiPol['attempt']; transient = $ebiPol['policy']['transient']; policy = $ebiPol['policy']['policy']; evidence = $ebiPol['rec']['outputs'] })
            Write-TraceEvent -WorkDir $State['workDir'] -RunId $State['runId'] -Phase $Section -Key $Key -Tags @{ workflow = $State['workflowId']; step = $ebiPol['id']; use = $ebiPol['use'] } -Action 'ask' -Status $ebiPol['answer'] -Message ('operator answered ' + $ebiPol['answer'])
            if ($ebiPol['answer'] -eq 'r') { continue }
            if ($ebiPol['answer'] -eq 'q') {
                $ebiPol['rec']['failure'] = 'cancelled'; $ebiPol['rec']['message'] = ('operator cancelled after: ' + $ebiPol['rec']['message'])
                return @{ outcome = 'cancelled'; records = $ebiPol['records'].ToArray(); last = $ebiPol['rec']; replayed = $false }
            }
            $ebiPol['decision'] = 'skip'
        }
        if ($ebiPol['decision'] -eq 'skip') {
            $ebiPol['skipRec'] = @{ section = $Section; id = $ebiPol['id']; use = $ebiPol['use']; item = $Key; group = $Group; status = 'skip'; failure = ''; message = ('skipped after ' + $ebiPol['rec']['failure'] + ' (' + $ebiPol['policy']['source'] + '); item left pending'); outputs = @{}; warnings = @() }
            [void]$ebiPol['records'].Add($ebiPol['skipRec'])
            Write-EbiStepLine -Status 'skip' -Section $Section -Id $ebiPol['id'] -Use $ebiPol['use'] -Key $Key -Detail $ebiPol['skipRec']['message']
            Write-TraceEvent -WorkDir $State['workDir'] -RunId $State['runId'] -Phase $Section -Key $Key -Tags @{ workflow = $State['workflowId']; step = $ebiPol['id']; use = $ebiPol['use'] } -Action 'step' -Status 'skip' -Message $ebiPol['skipRec']['message']
            return @{ outcome = 'skip'; records = $ebiPol['records'].ToArray(); last = $ebiPol['rec']; replayed = $false }
        }
        # fail: abort the run
        return @{ outcome = 'fail'; records = $ebiPol['records'].ToArray(); last = $ebiPol['rec']; replayed = $false }
    }
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
        Item    = $null        # the current worklist row inside "each" (P1-28), $null in setup / teardown
        KeyColumns = @()       # the key columns the run iterates by (profile or source.keyColumns)
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
        Resume      continue run/<RunId>/: replay the ledger, restore
                    run.* from run.json (RunId is required with it)
        AskHandler  scriptblock answering the runner's questions (see
                    Invoke-EbiDefaultAsk for the two shapes); default: a
                    console prompt, or self-answering under DryRun

      Returns a hashtable:
        ok          no item failed, nothing cancelled, no unrecovered
                    setup / teardown failure (a retried-then-ok step is
                    fine; a skipped item is reported, not a failure)
        runId, workflowId
        failure     '' or the first unrecovered failure id
                    ('workflow_invalid' when refused; 'cancelled')
        message
        steps       one record per attempt, in execution order:
                    @{ section; id; use; item; group; status
                       ('ok'|'fail'|'skip'|'skipped'|'replayed');
                       failure; message; outputs; warnings; attempt?; policy? }
        items       one record per selected row: @{ key; group; status
                    ('ok'|'fail'|'skip'|'cancelled'); failure; message }
        selected / total, warnings (count), resumed, session
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
        [int]$Limit = 0,
        [switch]$Resume,
        [scriptblock]$AskHandler = $null
    )

    $dryRunFlag = [bool]$DryRun.IsPresent
    $resumeFlag = [bool]$Resume.IsPresent
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
        steps = @(); items = @(); selected = 0; total = 0; warnings = 0; resumed = $resumeFlag; session = @{}
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

    # ---- run.json + ledger (resume) ------------------------------------------------
    $wfId = [string]$workflow['id']
    $run = @{ runId = $RunId; startedAt = (Get-Date).ToString('o'); operator = $Operator; workDir = $WorkDir; timeWindow = $null }
    $ledgerDone = @{}
    if ($resumeFlag) {
        $prev = Read-EbiRunFile -WorkDir $WorkDir -RunId $RunId
        if (-not $prev['ok']) {
            $result['failure'] = 'workflow_invalid'; $result['message'] = ('cannot resume run {0}: {1}' -f $RunId, $prev['message'])
            Write-Host ('  [refused] {0}' -f $result['message']) -ForegroundColor Red
            return $result
        }
        $prevDoc = $prev['value']
        $prevWf = if ($prevDoc.Contains('workflow') -and ($prevDoc['workflow'] -is [hashtable]) -and $prevDoc['workflow'].Contains('id')) { [string]$prevDoc['workflow']['id'] } else { '' }
        if ($prevWf -ne $wfId) {
            $result['failure'] = 'workflow_invalid'; $result['message'] = ('cannot resume run {0}: it ran workflow "{1}", not "{2}"' -f $RunId, $prevWf, $wfId)
            Write-Host ('  [refused] {0}' -f $result['message']) -ForegroundColor Red
            return $result
        }
        foreach ($k in @('startedAt', 'operator', 'timeWindow')) { if ($prevDoc.Contains($k) -and $null -ne $prevDoc[$k]) { $run[$k] = $prevDoc[$k] } }
        $ledger = Read-EbiLedger -Path (Get-EbiLedgerFile -WorkDir $WorkDir -RunId $RunId)
        if (-not $ledger['ok']) {
            $result['failure'] = 'workflow_invalid'; $result['message'] = ('cannot resume run {0}: {1}' -f $RunId, $ledger['message'])
            Write-Host ('  [refused] {0}' -f $result['message']) -ForegroundColor Red
            return $result
        }
        $ledgerDone = $ledger['done']
        if (@($ledger['badLines']).Count -gt 0) { Write-Host ('  [ledger WARN] {0}' -f $ledger['message']) -ForegroundColor DarkYellow }
        Write-Host ('  resume {0}: {1} ledger record(s)' -f $RunId, $ledger['count']) -ForegroundColor Cyan
    }
    $workflow['_path'] = $Path
    $runArgs = @{ path = $Path; workDir = $WorkDir; dryRun = $dryRunFlag; only = $(if ($null -eq $Only) { @() } else { @($Only) }); limit = $Limit; vars = $Vars; resume = $resumeFlag }
    $wrote = Write-EbiRunFile -WorkDir $WorkDir -RunId $RunId -Run $run -Workflow $workflow -RunArgs $runArgs -Finished $false
    if (-not $wrote['ok']) { Write-Host ('  [run.json WARN] {0}' -f $wrote['message']) -ForegroundColor DarkYellow }

    # ---- context, scope parts ---------------------------------------------------
    $ctx = New-EbiContext -WorkDir $WorkDir -RunId $RunId -DryRun $dryRunFlag -Profile $Profile
    $result['session'] = $ctx['Session']
    $wfVars = @{}
    if ($workflow.Contains('vars') -and $workflow['vars'] -is [System.Collections.IDictionary]) { foreach ($k in $workflow['vars'].Keys) { $wfVars[[string]$k] = $workflow['vars'][$k] } }
    foreach ($k in $Vars.Keys) { $wfVars[[string]$k] = $Vars[$k] }
    $pageName = if ($workflow.Contains('page') -and $null -ne $workflow['page']) { [string]$workflow['page'] } else { '' }
    $source = if ($workflow.Contains('source') -and $null -ne $workflow['source']) { [hashtable]$workflow['source'] } else { $null }
    $keyColumns = @(Get-EbiWorklistKeyColumns -Profile $Profile)
    if ($null -ne $source -and $source.Contains('keyColumns') -and $null -ne $source['keyColumns']) {
        $keyColumns = @(Get-EbiManifestArray -Manifest $source -Key 'keyColumns')
    }
    $groupColumn = Get-EbiWorklistGroupColumn -Profile $Profile
    $groupBy = if ($null -ne $source -and $source.Contains('groupBy') -and $null -ne $source['groupBy']) { [string]$source['groupBy'] } else { '' }
    if ($groupColumn -eq '' -and $groupBy -ne '') { $groupColumn = $groupBy }

    $askDefault = { param($q) Invoke-EbiDefaultAsk -Question $q -DryRun $dryRunFlag }
    $state = @{
        ctx = $ctx; registry = (New-EbiRegistry -ModulesRoot $ModulesRoot)
        workDir = $WorkDir; runId = $RunId; workflowId = $wfId; workflow = $workflow
        ledgerDone = $ledgerDone; ledgerFile = (Get-EbiLedgerFile -WorkDir $WorkDir -RunId $RunId)
        ask = $(if ($null -ne $AskHandler) { $AskHandler } else { $askDefault })
    }
    $tagsBase = @{ workflow = $wfId }

    Write-Host ('  run {0}  workflow {1}{2}{3}' -f $RunId, $wfId, $(if ($dryRunFlag) { '  (dry run)' } else { '' }), $(if ($resumeFlag) { '  (resume)' } else { '' })) -ForegroundColor Cyan
    Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'run' -Tags $tagsBase -Action 'run' -Status 'start' -Message $Path -Data @{ resume = $resumeFlag; dryRun = $dryRunFlag }

    $stopped   = $false     # setup failed, policy fail, or cancelled: no more each
    $cancelled = $false
    $aborted   = $false     # policy fail
    $unrecovered = New-Object System.Collections.ArrayList   # records that ended a section / item in failure
    try {
        # ---- setup ------------------------------------------------------------------
        if ($workflow.Contains('setup') -and $null -ne $workflow['setup']) {
            $setupSteps = @{}
            $scope = New-EbiTemplateScope -Vars $wfVars -Profile $Profile -PageName $pageName -Run $run -Item $null -Steps $setupSteps -KeyColumns $keyColumns -GroupColumn $groupColumn
            foreach ($call in $workflow['setup']) {
                if ($stopped) { [void]$records.Add((New-EbiSkipRecord -Section 'setup' -Call $call -Key '' -Group '' -Why 'earlier step failed' -State $state)); continue }
                $done = . Invoke-EbiStepWithPolicy -State $state -Section 'setup' -Call $call -Scope $scope -Steps $setupSteps
                foreach ($r in $done['records']) { [void]$records.Add($r) }
                switch ([string]$done['outcome']) {
                    'cancelled' { $cancelled = $true; $stopped = $true; [void]$unrecovered.Add($done['last']) }
                    'fail'      { $aborted = $true; $stopped = $true; [void]$unrecovered.Add($done['last']) }
                    'skip'      { $stopped = $true; [void]$unrecovered.Add($done['last']) }   # a setup step given up on: nothing downstream is trustworthy
                }
            }
        }

        # ---- each -------------------------------------------------------------------
        if (-not $stopped -and $null -ne $source -and $workflow.Contains('each') -and $null -ne $workflow['each']) {
            $tableName = [string]$source['table']
            if (-not $ctx['Session'].Contains($tableName) -or [string]$ctx['Session'][$tableName]['kind'] -ne 'worklist') {
                $why = if ($ctx['Session'].Contains($tableName)) { ('source.table "{0}" is a "{1}", not a worklist' -f $tableName, [string]$ctx['Session'][$tableName]['kind']) } else { ('source.table "{0}" is not registered; a setup step must load the worklist with "as": "{0}"' -f $tableName) }
                $srcRec = @{ section = 'each'; id = '(source)'; use = ''; item = ''; group = ''; status = 'fail'; failure = $(if ($ctx['Session'].Contains($tableName)) { 'session_kind_mismatch' } else { 'session_missing' }); message = $why; outputs = @{}; warnings = @() }
                [void]$records.Add($srcRec); [void]$unrecovered.Add($srcRec)
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

                $groupOutputs = @{}    # group -> @{ stepId -> outputs } from once:group steps (this process)
                $groupEndCalls = @($workflow['each'] | Where-Object { $_.Contains('once') -and [string]$_['once'] -eq 'groupEnd' })
                $currentGroup = $null
                $lastRowOfGroup = $null

                # once:groupEnd for the group just finished (7.2): runs on the
                # group's last row, whatever happened to that row's own steps,
                # so a resource registered at the group head is released.
                function Invoke-EbiGroupEndFor {
                    param([string]$G, $Row)
                    if ($groupBy -eq '' -or $groupEndCalls.Count -eq 0 -or $null -eq $Row) { return }
                    $gKey = Get-EbiKeyDisplay -Item $Row -KeyColumns $keyColumns
                    $gSteps = @{}
                    if ($groupOutputs.Contains($G)) { foreach ($k in $groupOutputs[$G].Keys) { $gSteps[$k] = $groupOutputs[$G][$k] } }
                    $gScope = New-EbiTemplateScope -Vars $wfVars -Profile $Profile -PageName $pageName -Run $run -Item $Row -Steps $gSteps -KeyColumns $keyColumns -GroupColumn $groupColumn
                    $ctx['Item'] = $Row; $ctx['KeyColumns'] = $keyColumns
                    foreach ($gCall in $groupEndCalls) {
                        $gDone = . Invoke-EbiStepWithPolicy -State $state -Section 'each' -Call $gCall -Scope $gScope -Steps $gSteps -Key $gKey -Group $G -LedgerGroup $G
                        foreach ($r in $gDone['records']) { [void]$records.Add($r) }
                        if ($gDone['outcome'] -eq 'cancelled') { $script:ebiGroupEndOutcome = 'cancelled'; return }
                        if ($gDone['outcome'] -eq 'fail') { $script:ebiGroupEndOutcome = 'fail'; return }
                        if ($gDone['outcome'] -eq 'skip') { [void]$unrecovered.Add($gDone['last']) }
                    }
                }

                foreach ($row in $ordered) {
                    if ($stopped) { break }
                    $key   = Get-EbiKeyDisplay -Item $row -KeyColumns $keyColumns
                    $group = Get-EbiWorklistGroupOf -Row $row -GroupBy $groupBy
                    if ($groupBy -ne '' -and $null -ne $currentGroup -and $group -ne $currentGroup) {
                        $script:ebiGroupEndOutcome = ''
                        Invoke-EbiGroupEndFor -G $currentGroup -Row $lastRowOfGroup
                        if ($script:ebiGroupEndOutcome -eq 'cancelled') { $cancelled = $true; $stopped = $true; break }
                        if ($script:ebiGroupEndOutcome -eq 'fail') { $aborted = $true; $stopped = $true; break }
                    }
                    $currentGroup = $group; $lastRowOfGroup = $row
                    $steps = @{}
                    if ($groupBy -ne '' -and $groupOutputs.Contains($group)) { foreach ($k in $groupOutputs[$group].Keys) { $steps[$k] = $groupOutputs[$group][$k] } }
                    $scope = New-EbiTemplateScope -Vars $wfVars -Profile $Profile -PageName $pageName -Run $run -Item $row -Steps $steps -KeyColumns $keyColumns -GroupColumn $groupColumn
                    $ctx['Item'] = $row; $ctx['KeyColumns'] = $keyColumns   # STEP-CONTRACT 3.2: the current row (flow.checkpoint / progress.event default key)
                    $itemRec = @{ key = $key; group = $group; status = 'ok'; failure = ''; message = '' }
                    Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'each' -Key $key -Tags $tagsBase -Action 'item' -Status 'start' -Data @{ group = $group }
                    $itemDone = $false
                    foreach ($call in $workflow['each']) {
                        $once = if ($call.Contains('once') -and $null -ne $call['once']) { [string]$call['once'] } else { '' }
                        if ($once -eq 'groupEnd') { continue }
                        if ($itemDone) {
                            [void]$records.Add((New-EbiSkipRecord -Section 'each' -Call $call -Key $key -Group $group -Why ('earlier step of this item ' + $itemRec['status']) -State $state)); continue
                        }
                        if ($once -eq 'group') {
                            if (-not $groupOutputs.Contains($group)) { $groupOutputs[$group] = @{} }
                            if ($groupOutputs[$group].Contains([string]$call['id'])) {
                                # Already ran for this group in this process: outputs are
                                # in $steps (replayed above); no new record, one trace line (6.3).
                                Write-EbiStepLine -Status 'skip' -Section 'each' -Id ([string]$call['id']) -Use ([string]$call['use']) -Key $key -Detail ('once: group -- replayed for group "' + $group + '"')
                                Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'each' -Key $key -Tags @{ workflow = $wfId; step = [string]$call['id']; use = [string]$call['use']; group = $group } -Action 'step' -Status 'skip' -Message ('once: group -- replayed for group "' + $group + '"')
                                continue
                            }
                        }
                        $done = . Invoke-EbiStepWithPolicy -State $state -Section 'each' -Call $call -Scope $scope -Steps $steps -Key $key -Group $group -LedgerGroup $(if ($once -eq 'group') { $group } else { '' })
                        foreach ($r in $done['records']) { [void]$records.Add($r) }
                        switch ([string]$done['outcome']) {
                            'cancelled' { $itemRec['status'] = 'cancelled'; $itemRec['failure'] = 'cancelled'; $itemRec['message'] = ([string]$call['id'] + ': ' + $done['last']['message']); $cancelled = $true; $stopped = $true; $itemDone = $true; [void]$unrecovered.Add($done['last']) }
                            'fail'      { $itemRec['status'] = 'fail'; $itemRec['failure'] = $done['last']['failure']; $itemRec['message'] = ([string]$call['id'] + ': ' + $done['last']['message']); $aborted = $true; $stopped = $true; $itemDone = $true; [void]$unrecovered.Add($done['last']) }
                            'skip'      { $itemRec['status'] = 'skip'; $itemRec['failure'] = $done['last']['failure']; $itemRec['message'] = ([string]$call['id'] + ': ' + $done['last']['message']); $itemDone = $true }
                            default     { if ($once -eq 'group') { $groupOutputs[$group][[string]$call['id']] = $done['last']['outputs'] } }
                        }
                    }
                    [void]$itemRecords.Add($itemRec)
                    Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'each' -Key $key -Tags $tagsBase -Action 'item' -Status $itemRec['status'] -Message $itemRec['message'] -Data @{ group = $group; failure = $itemRec['failure'] }
                }
                if (-not $stopped -and $groupBy -ne '' -and $null -ne $currentGroup) {
                    $script:ebiGroupEndOutcome = ''
                    Invoke-EbiGroupEndFor -G $currentGroup -Row $lastRowOfGroup
                    if ($script:ebiGroupEndOutcome -eq 'cancelled') { $cancelled = $true; $stopped = $true }
                    if ($script:ebiGroupEndOutcome -eq 'fail') { $aborted = $true; $stopped = $true }
                }
                if ($stopped) {
                    # The rows never reached are listed, so the summary says
                    # what the cancel / abort left undone.
                    $done = $itemRecords.Count
                    foreach ($row in $ordered) {
                        if ($done -gt 0) { $done--; continue }
                        [void]$itemRecords.Add(@{ key = (Get-EbiKeyDisplay -Item $row -KeyColumns $keyColumns); group = (Get-EbiWorklistGroupOf -Row $row -GroupBy $groupBy); status = 'skip'; failure = ''; message = $(if ($cancelled) { 'run cancelled' } else { 'run aborted (onError: fail)' }) })
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
            $ctx['Item'] = $null; $ctx['KeyColumns'] = $keyColumns
            $scope = New-EbiTemplateScope -Vars $wfVars -Profile $Profile -PageName $pageName -Run $run -Item $null -Steps $teardownSteps -KeyColumns $keyColumns -GroupColumn $groupColumn
            foreach ($call in $workflow['teardown']) {
                $done = . Invoke-EbiStepWithPolicy -State $state -Section 'teardown' -Call $call -Scope $scope -Steps $teardownSteps
                foreach ($r in $done['records']) { [void]$records.Add($r) }
                if ($done['outcome'] -eq 'cancelled') { $cancelled = $true; [void]$unrecovered.Add($done['last']) }
                elseif ($done['outcome'] -eq 'fail' -or $done['outcome'] -eq 'skip') { [void]$unrecovered.Add($done['last']) }
            }
        }

        $result['steps'] = $records.ToArray()
        $result['items'] = $itemRecords.ToArray()
        $itemsFailed = @($itemRecords.ToArray() | Where-Object { $_['status'] -eq 'fail' -or $_['status'] -eq 'cancelled' }).Count
        $itemsSkipped = @($itemRecords.ToArray() | Where-Object { $_['status'] -eq 'skip' }).Count
        $result['ok'] = ($unrecovered.Count -eq 0 -and -not $cancelled -and -not $aborted -and $itemsFailed -eq 0)
        if ($unrecovered.Count -gt 0) {
            $first = $unrecovered[0]
            $result['failure'] = [string]$first['failure']
            $where = if ([string]$first['item'] -ne '') { $first['section'] + '[' + $first['item'] + ']/' + $first['id'] } else { $first['section'] + '/' + $first['id'] }
            $result['message'] = ('{0}: {1}' -f $where, $first['message'])
        }
        if ($cancelled) { $result['failure'] = 'cancelled' }

        # warnings: every one that rode in a record, said again at the end (3.1)
        $warnLines = New-Object System.Collections.ArrayList
        foreach ($r in $records) {
            foreach ($w in @($r['warnings'])) {
                $code = if ($w -is [System.Collections.IDictionary] -and $w.Contains('code')) { [string]$w['code'] } else { '' }
                $msg  = if ($w -is [System.Collections.IDictionary] -and $w.Contains('message')) { [string]$w['message'] } else { [string]$w }
                $where = if ([string]$r['item'] -ne '') { $r['section'] + '[' + $r['item'] + ']/' + $r['id'] } else { $r['section'] + '/' + $r['id'] }
                [void]$warnLines.Add(('{0}: {1} {2}' -f $where, $code, $msg).Trim())
            }
        }
        $result['warnings'] = $warnLines.Count
        if ($warnLines.Count -gt 0) {
            Write-Host ('  {0} warning{1}:' -f $warnLines.Count, $(if ($warnLines.Count -eq 1) { '' } else { 's' })) -ForegroundColor Yellow
            foreach ($line in $warnLines) { Write-Host ('    [warn ] ' + $line) -ForegroundColor Yellow }
        }

        $left = @($ctx['Session'].Keys)
        $finalStatus = if ($result['ok']) { 'ok' } else { 'fail' }
        Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'run' -Tags $tagsBase -Action 'run' -Status $finalStatus -Message $result['message'] -Data @{ steps = $records.Count; items = $itemRecords.Count; itemsFailed = $itemsFailed; itemsSkipped = $itemsSkipped; warnings = $warnLines.Count; cancelled = $cancelled; aborted = $aborted; sessionLeft = $left }
        [void](Write-EbiRunFile -WorkDir $WorkDir -RunId $RunId -Run $run -Workflow $workflow -RunArgs $runArgs -Finished ($result['ok']) -Result @{ ok = $result['ok']; failure = $result['failure']; message = $result['message']; items = $itemRecords.Count; itemsFailed = $itemsFailed; itemsSkipped = $itemsSkipped; warnings = $warnLines.Count })
        Write-Host ('  run {0}  {1}  ({2} step record{3}{4}{5}{6})' -f $RunId, $(if ($result['ok']) { 'OK' } elseif ($cancelled) { 'CANCELLED' } else { 'FAIL' }), $records.Count, $(if ($records.Count -eq 1) { '' } else { 's' }), $(if ($itemRecords.Count -gt 0) { ('; ' + $itemRecords.Count + ' item(s), ' + $itemsFailed + ' failed, ' + $itemsSkipped + ' skipped') } else { '' }), $(if ($warnLines.Count -gt 0) { '; ' + $warnLines.Count + ' warning(s)' } else { '' }), $(if ($left.Count -gt 0) { '; still registered: ' + ($left -join ', ') } else { '' })) -ForegroundColor $(if ($result['ok']) { 'Green' } else { 'Red' })
    }
    return $result
}
