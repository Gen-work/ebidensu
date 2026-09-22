#Requires -Version 5.1
# ============================================================
#  kernel/Runner.ps1
#
#  The ebi-dance workflow runner. Dot-source only (no param() block,
#  ASCII source, no class -- CLAUDE.md conventions).
#
#  P0-07 SPIKE. This file is the seed of the real runner (P1-03/P1-04),
#  not a throwaway: what is here stays, what is missing is listed so
#  nobody mistakes "not implemented" for "not needed".
#
#  What runs today:
#    - read a workflow JSON (spec/WORKFLOW-SCHEMA.md 1-2)
#    - load each step file named by "use" -- dot-source it in THIS
#      runspace and capture its Invoke-Step right away, because every
#      step defines the same function name (P1-02 will lift this into
#      kernel/Registry.ps1; the mechanism is decided here)
#    - run "setup", then "teardown" in a finally block (the 1.1 guarantee:
#      normal end, a step failure and an unexpected exception all reach
#      teardown; Ctrl+C is explicitly NOT covered)
#    - $Ctx per spec/STEP-CONTRACT.md 3.2, with $Ctx.Session present from
#      day one: a hashtable of name -> @{ kind; value; registeredBy }
#    - the resource channel of STEP-CONTRACT.md 3.4 point 7: a provides
#      step hands its handle back under the reserved return key
#      "resource"; "with.as" registers it; a consumer's type='session'
#      input is REPLACED by the registered value before Invoke-Step;
#      a releases step drops the name afterwards
#    - the 3.1 return-value contract: hashtable with ok, failure id
#      declared in the manifest, outputs JSON-serializable, exceptions
#      become internal_error
#    - one trace event per step (kernel/Trace.ps1); warnings ride in
#      "data" as 3.1 requires
#
#  What is refused, loudly, rather than half-done:
#    - "source" / "each"        -> P1-03 (foreach, once, groupBy)
#    - "{{...}}" templates       -> P1-01 (kernel/Context.ps1)
#    - "when" / "onError"        -> P1-03 / P1-04
#    - input schema validation   -> P1-02 (kernel/Registry.ps1)
#    - ledger / resume           -> P1-04
#  A workflow that uses any of these fails BEFORE the first step runs,
#  with failure 'unsupported_in_spike', so a literal "{{item.key}}" is
#  never handed to a step as if it were a value.
#
#  Hashtable access in this file is by index ($h['k']), never by dot:
#  under Set-StrictMode (which Tests/Run-Tests.ps1 turns on) a missing
#  key read with dot syntax throws.
# ============================================================

. (Join-Path $PSScriptRoot 'Trace.ps1')

# --- runner-level reserved failure ids (STEP-CONTRACT.md 3.4 point 7 table) ---
function Get-EbiRunnerFailureIds {
    return @('internal_error', 'contract_violation', 'step_not_found',
             'session_missing', 'session_kind_mismatch', 'session_name_taken',
             'unsupported_in_spike')
}

function ConvertTo-EbiHashtable {
    <#
      ConvertFrom-Json hands back PSCustomObjects; the runner wants plain
      hashtables (index access never throws under StrictMode, and the step
      contract is written in terms of hashtables). Recursive. Arrays come
      back as object[] built by an explicit loop -- the @() wrap over an
      indexed collection is the shape this repo bans (CLAUDE.md, R4).
    #>
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $Value.Keys) { $h[[string]$k] = ConvertTo-EbiHashtable $Value[$k] }
        return $h
    }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Collections.IList]) {
        $list = New-Object System.Collections.ArrayList
        foreach ($item in $Value) { [void]$list.Add((ConvertTo-EbiHashtable $item)) }
        # The unary comma matters: a one-element array returned bare is
        # unrolled into its element, and a "setup" with a single step call
        # would come back as that call instead of a list of one. Callers
        # assign the result; none wraps it in @().
        return ,$list.ToArray()
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = ConvertTo-EbiHashtable $p.Value }
        return $h
    }
    return $Value
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

function Get-EbiDefaultModulesRoot {
    # kernel/ and modules/ are siblings under the repo root.
    return (Join-Path (Split-Path $PSScriptRoot -Parent) 'modules')
}

function Get-EbiStepPath {
    # 'screen.capture_window' -> <ModulesRoot>/screen/screen.capture_window.ps1
    # '' when the id has no group part (a bare verb is not a step id).
    param([string]$ModulesRoot, [string]$Use)
    if ([string]::IsNullOrWhiteSpace($Use)) { return '' }
    $dot = $Use.IndexOf('.')
    if ($dot -le 0 -or $dot -eq ($Use.Length - 1)) { return '' }
    $group = $Use.Substring(0, $dot)
    return (Join-Path (Join-Path $ModulesRoot $group) ($Use + '.ps1'))
}

function Test-EbiValueHasTemplate {
    # True when a with-value (or anything nested in it) contains '{{'.
    # Templates are P1-01; the spike refuses them instead of passing the
    # literal text to a step.
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

function Get-EbiWorkflowSpikeProblems {
    <#
      PURE. Everything the spike cannot run, found before anything runs.
      Returns an array of one-line strings; empty means the workflow is
      within the spike's reach. Also the lint-lite that WORKFLOW-SCHEMA.md
      2 already requires of every runner: every call has a non-empty id
      and use, ids are unique within a section.
    #>
    param([hashtable]$Workflow)
    $problems = New-Object System.Collections.ArrayList

    if ($null -eq $Workflow) { [void]$problems.Add('workflow is empty'); return $problems.ToArray() }
    if (-not $Workflow.Contains('id') -or [string]::IsNullOrWhiteSpace([string]$Workflow['id'])) {
        [void]$problems.Add('top-level "id" is required')
    }
    foreach ($k in @('source', 'each')) {
        if ($Workflow.Contains($k) -and $null -ne $Workflow[$k]) {
            [void]$problems.Add(('"{0}" is not supported by the P0-07 spike (P1-03)' -f $k))
        }
    }
    if ($Workflow.Contains('onError') -and $null -ne $Workflow['onError']) {
        [void]$problems.Add('top-level "onError" is not supported by the P0-07 spike (P1-04); the spike stops at the first failure')
    }

    foreach ($section in @('setup', 'teardown')) {
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
            if ([string]::IsNullOrWhiteSpace($id))  { [void]$problems.Add(('{0}[{1}]: "id" is required' -f $section, $n)) }
            if ([string]::IsNullOrWhiteSpace($use)) { [void]$problems.Add(('{0}[{1}]: "use" is required' -f $section, $n)) }
            if ($id -ne '' -and $seen.Contains($id)) {
                [void]$problems.Add(('{0}: step id "{1}" is used twice' -f $section, $id))
            }
            $seen[$id] = $true
            foreach ($k in @('when', 'onError', 'once')) {
                if ($call.Contains($k) -and $null -ne $call[$k]) {
                    [void]$problems.Add(('{0}/{1}: "{2}" is not supported by the P0-07 spike' -f $section, $id, $k))
                }
            }
            if ($call.Contains('with') -and (Test-EbiValueHasTemplate $call['with'])) {
                [void]$problems.Add(('{0}/{1}: "with" contains a {{{{...}}}} template; templates are P1-01' -f $section, $id))
            }
        }
    }
    return $problems.ToArray()
}

function Test-EbiJsonSerializable {
    param($Value)
    try { [void]($Value | ConvertTo-Json -Compress -Depth 10 -ErrorAction Stop); return $true }
    catch { return $false }
}

function Get-EbiManifestArray {
    # provides / releases / needs may be missing, $null, a scalar or an
    # array. Always hand back string[] (explicit loop, see R4).
    param($Manifest, [string]$Key)
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $Manifest -or -not $Manifest.Contains($Key)) { return $out.ToArray() }
    $v = $Manifest[$Key]
    if ($null -eq $v) { return $out.ToArray() }
    if ($v -is [string]) { [void]$out.Add($v); return $out.ToArray() }
    if ($v -is [System.Collections.IEnumerable]) {
        foreach ($item in $v) { if ($null -ne $item) { [void]$out.Add([string]$item) } }
        return $out.ToArray()
    }
    [void]$out.Add([string]$v)
    return $out.ToArray()
}

function Get-EbiManifestFailureIds {
    param($Manifest)
    $ids = New-Object System.Collections.ArrayList
    if ($null -eq $Manifest -or -not $Manifest.Contains('failures')) { return $ids.ToArray() }
    $f = $Manifest['failures']
    if ($null -eq $f) { return $ids.ToArray() }
    foreach ($item in $f) {
        if ($item -is [System.Collections.IDictionary] -and $item.Contains('id')) { [void]$ids.Add([string]$item['id']) }
    }
    return $ids.ToArray()
}

function Get-EbiSessionInputs {
    # name -> sessionKind for every inputs entry with type = 'session'.
    param($Manifest)
    $map = @{}
    if ($null -eq $Manifest -or -not $Manifest.Contains('inputs')) { return $map }
    $inputs = $Manifest['inputs']
    if (-not ($inputs -is [System.Collections.IDictionary])) { return $map }
    foreach ($name in $inputs.Keys) {
        $spec = $inputs[$name]
        if (-not ($spec -is [System.Collections.IDictionary])) { continue }
        if (-not $spec.Contains('type') -or [string]$spec['type'] -ne 'session') { continue }
        $kind = if ($spec.Contains('sessionKind')) { [string]$spec['sessionKind'] } else { '' }
        $map[[string]$name] = $kind
    }
    return $map
}

function Resolve-EbiStepInputs {
    <#
      PURE (given Session). Turn a call's "with" into the $In a step
      receives (STEP-CONTRACT.md 3.4 points 3 and 7):
        - "as" is a runner field: taken out, never reaches $In
        - each type='session' input named in "with" is looked up in
          $Session and REPLACED by the registered value
      Returns @{ ok; In; As; failure; message }.
    #>
    param($Manifest, [hashtable]$With, [hashtable]$Session)

    $in = @{}
    $as = ''
    if ($null -ne $With) {
        foreach ($k in $With.Keys) {
            if ([string]$k -eq 'as') { $as = [string]$With[$k]; continue }
            $in[[string]$k] = $With[$k]
        }
    }

    $provides = @(Get-EbiManifestArray -Manifest $Manifest -Key 'provides')
    if ($as -ne '' -and $provides.Count -eq 0) {
        return @{ ok = $false; In = $in; As = $as; failure = 'contract_violation';
                  message = ('"as" given but step "{0}" provides nothing' -f [string]$Manifest['id']) }
    }
    if ($as -ne '' -and $Session.Contains($as)) {
        return @{ ok = $false; In = $in; As = $as; failure = 'session_name_taken';
                  message = ('session name "{0}" is already registered and not released' -f $as) }
    }

    $sessionInputs = Get-EbiSessionInputs -Manifest $Manifest
    foreach ($name in $sessionInputs.Keys) {
        if (-not $in.Contains($name)) { continue }   # required-ness is P1-02's check
        $wanted = [string]$in[$name]
        if (-not $Session.Contains($wanted)) {
            return @{ ok = $false; In = $in; As = $as; failure = 'session_missing';
                      message = ('input "{0}" names session resource "{1}", which is not registered' -f $name, $wanted) }
        }
        $entry = $Session[$wanted]
        $kind  = [string]$sessionInputs[$name]
        if ($kind -ne '' -and [string]$entry['kind'] -ne $kind) {
            return @{ ok = $false; In = $in; As = $as; failure = 'session_kind_mismatch';
                      message = ('input "{0}" wants kind "{1}" but "{2}" is a "{3}"' -f $name, $kind, $wanted, [string]$entry['kind']) }
        }
        $in[$name] = $entry['value']
    }
    return @{ ok = $true; In = $in; As = $as; failure = ''; message = '' }
}

function Test-EbiStepReturn {
    <#
      PURE. The 3.1 return-value contract plus 3.4 point 7's resource rule.
      Returns @{ ok; failure; message; Outputs; Warnings; Resource; HasResource }
      where ok=$false means a contract violation (never the step's own
      failure -- that is reported separately by the caller).
    #>
    param($Manifest, $Return, [bool]$WantsResource)

    if (-not ($Return -is [hashtable])) {
        return @{ ok = $false; failure = 'contract_violation';
                  message = 'Invoke-Step must return a [hashtable] (got ' + $(if ($null -eq $Return) { 'null' } else { $Return.GetType().Name }) + ')' }
    }
    if (-not $Return.Contains('ok')) {
        return @{ ok = $false; failure = 'contract_violation'; message = 'return value has no "ok" key' }
    }

    $reserved = @('ok', 'failure', 'message', 'warnings', 'resource')
    $outputs  = @{}
    foreach ($k in $Return.Keys) { if ($reserved -notcontains [string]$k) { $outputs[[string]$k] = $Return[$k] } }

    $stepOk = [bool]$Return['ok']
    if (-not $stepOk) {
        $fid = if ($Return.Contains('failure')) { [string]$Return['failure'] } else { '' }
        $declared = @(Get-EbiManifestFailureIds -Manifest $Manifest)
        if ($fid -eq '' -or $declared -notcontains $fid) {
            return @{ ok = $false; failure = 'contract_violation';
                      message = ('failure id "{0}" is not declared in the manifest of {1}' -f $fid, [string]$Manifest['id']) }
        }
    }

    $hasResource = $Return.Contains('resource')
    $provides = @(Get-EbiManifestArray -Manifest $Manifest -Key 'provides')
    if ($hasResource -and $provides.Count -eq 0) {
        return @{ ok = $false; failure = 'contract_violation';
                  message = ('{0} returned "resource" but provides nothing' -f [string]$Manifest['id']) }
    }
    if ($WantsResource -and $stepOk -and -not $hasResource) {
        return @{ ok = $false; failure = 'contract_violation';
                  message = ('{0} was called with "as" but returned no "resource" key (return resource = $null under DryRun)' -f [string]$Manifest['id']) }
    }

    if (-not (Test-EbiJsonSerializable $outputs)) {
        return @{ ok = $false; failure = 'contract_violation';
                  message = ('outputs of {0} are not JSON-serializable; handles belong in $Ctx.Session (3.4)' -f [string]$Manifest['id']) }
    }

    $warnings = if ($Return.Contains('warnings') -and $null -ne $Return['warnings']) { $Return['warnings'] } else { @() }
    $resource = if ($hasResource) { $Return['resource'] } else { $null }
    return @{ ok = $true; failure = ''; message = ''; Outputs = $outputs; Warnings = $warnings;
              Resource = $resource; HasResource = $hasResource }
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
    param([string]$WorkDir, [string]$RunId, [bool]$DryRun)
    return @{
        WorkDir = $WorkDir
        RunId   = $RunId
        Profile = @{}          # profile loading is P2-01; empty, present
        Log     = (New-EbiLog)
        DryRun  = $DryRun
        Session = @{}          # name -> @{ kind; value; registeredBy }
    }
}

function Write-EbiStepLine {
    param([string]$Status, [string]$Section, [string]$Id, [string]$Use, [string]$Detail)
    $color = switch ($Status) { 'ok' { 'Green' } 'fail' { 'Red' } 'skip' { 'DarkGray' } default { 'Gray' } }
    Write-Host ('  [{0,-4}] {1}/{2}  {3}  {4}' -f $Status, $Section, $Id, $Use, $Detail) -ForegroundColor $color
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

function Invoke-EbiWorkflow {
    <#
      Run a workflow JSON within the spike's reach (see the file header).

      Returns a hashtable:
        ok          every step call returned ok
        runId       run/<runId>/trace.jsonl holds the events
        workflowId
        failure     '' or the runner-level failure that stopped the run
        message
        steps       one record per step call, in execution order:
                    @{ section; id; use; status ('ok'|'fail'|'skip');
                       failure; message; outputs; warnings }
        session     the live $Ctx.Session (in-process only; a caller
                    that wants to inspect what was left registered)
    #>
    param(
        [string]$Path,
        [string]$WorkDir,
        [string]$ModulesRoot = '',
        [string]$RunId = '',
        [switch]$DryRun
    )

    $dryRunFlag = [bool]$DryRun.IsPresent
    if ([string]::IsNullOrWhiteSpace($ModulesRoot)) { $ModulesRoot = Get-EbiDefaultModulesRoot }
    if ([string]::IsNullOrWhiteSpace($RunId))       { $RunId = New-EbiRunId }
    if ([string]::IsNullOrWhiteSpace($WorkDir))     { $WorkDir = (Get-Location).ProviderPath }
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
        steps = @(); session = @{}
    }
    $records = New-Object System.Collections.ArrayList

    # ---- read + refuse what the spike cannot run --------------------------
    $workflow = $null
    try {
        if (-not (Test-Path -LiteralPath $Path)) { throw ('workflow file not found: {0}' -f $Path) }
        $raw = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding($false)))
        $workflow = ConvertTo-EbiHashtable ($raw | ConvertFrom-Json)
    } catch {
        $result['failure'] = 'unsupported_in_spike'
        $result['message'] = ('cannot read workflow: {0}' -f $_.Exception.Message)
        Write-Host ('  [refused] {0}' -f $result['message']) -ForegroundColor Red
        return $result
    }
    if (-not ($workflow -is [hashtable])) {
        $result['failure'] = 'unsupported_in_spike'; $result['message'] = 'workflow JSON must be an object'
        Write-Host ('  [refused] {0}' -f $result['message']) -ForegroundColor Red
        return $result
    }
    $result['workflowId'] = [string]$workflow['id']

    $problems = @(Get-EbiWorkflowSpikeProblems -Workflow $workflow)
    if ($problems.Count -gt 0) {
        $result['failure'] = 'unsupported_in_spike'
        $result['message'] = ($problems -join '; ')
        Write-Host ('  [refused] {0}' -f $result['message']) -ForegroundColor Red
        return $result
    }

    # ---- context ------------------------------------------------------------
    $ctx = New-EbiContext -WorkDir $WorkDir -RunId $RunId -DryRun $dryRunFlag
    $result['session'] = $ctx['Session']
    $registry = @{}     # use -> @{ Manifest; Invoke }
    $wfId = [string]$workflow['id']
    $tagsBase = @{ workflow = $wfId }

    Write-Host ('  run {0}  workflow {1}{2}' -f $RunId, $wfId, $(if ($dryRunFlag) { '  (dry run)' } else { '' })) -ForegroundColor Cyan
    Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'run' -Tags $tagsBase -Action 'run' -Status 'start' -Message $Path

    $stopped = $false
    try {
        foreach ($section in @('setup', 'teardown')) {
            if (-not $workflow.Contains($section) -or $null -eq $workflow[$section]) { continue }
            # teardown runs even after a failure (1.1); setup stops at the first one.
            if ($section -eq 'setup' -and $stopped) { continue }

            foreach ($call in $workflow[$section]) {
                $id  = [string]$call['id']
                $use = [string]$call['use']
                $with = if ($call.Contains('with') -and $null -ne $call['with']) { [hashtable]$call['with'] } else { @{} }
                $tags = @{ workflow = $wfId; step = $id; use = $use }

                if ($section -eq 'setup' -and $stopped) {
                    [void]$records.Add(@{ section = $section; id = $id; use = $use; status = 'skip'; failure = ''; message = 'earlier step failed'; outputs = @{}; warnings = @() })
                    Write-EbiStepLine -Status 'skip' -Section $section -Id $id -Use $use -Detail 'earlier step failed'
                    Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase $section -Tags $tags -Action 'step' -Status 'skip' -Message 'earlier step failed'
                    continue
                }

                Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase $section -Tags $tags -Action 'step' -Status 'start'

                # -- load (once per use) --------------------------------------
                if (-not $registry.Contains($use)) {
                    $stepPath = Get-EbiStepPath -ModulesRoot $ModulesRoot -Use $use
                    $loadErr = ''
                    if ($stepPath -eq '' -or -not (Test-Path -LiteralPath $stepPath)) {
                        $loadErr = ('no step file for "{0}" (looked at {1})' -f $use, $(if ($stepPath -eq '') { '<not a <group>.<verb> id>' } else { $stepPath }))
                    } else {
                        try {
                            $Manifest = $null
                            $captured = $null
                            # A step that forgets Invoke-Step must not inherit the previous
                            # step's, so the name is cleared before every load.
                            if (Test-Path -LiteralPath 'function:Invoke-Step') { Remove-Item -LiteralPath 'function:Invoke-Step' -ErrorAction SilentlyContinue }
                            # Dot-sourced HERE, in the runner's own scope, so the step's
                            # prefixed helper functions outlive this block; Invoke-Step is
                            # captured immediately because the next file overwrites it.
                            . $stepPath
                            $fnItem = Get-Item -LiteralPath 'function:Invoke-Step' -ErrorAction SilentlyContinue
                            if ($null -ne $fnItem) { $captured = $fnItem.ScriptBlock }
                            if ($null -eq $Manifest -or -not ($Manifest -is [hashtable])) { throw 'step defines no [hashtable] $Manifest' }
                            if ($null -eq $captured) { throw 'step defines no Invoke-Step' }
                            if (-not $Manifest.Contains('id') -or [string]$Manifest['id'] -ne $use) {
                                throw ('manifest id "{0}" does not match "{1}"' -f $(if ($Manifest.Contains('id')) { [string]$Manifest['id'] } else { '' }), $use)
                            }
                            $registry[$use] = @{ Manifest = $Manifest; Invoke = $captured }
                        } catch {
                            $loadErr = ('{0}: {1}' -f $stepPath, $_.Exception.Message)
                        }
                    }
                    if ($loadErr -ne '') {
                        [void]$records.Add(@{ section = $section; id = $id; use = $use; status = 'fail'; failure = 'step_not_found'; message = $loadErr; outputs = @{}; warnings = @() })
                        Write-EbiStepLine -Status 'fail' -Section $section -Id $id -Use $use -Detail ('step_not_found: ' + $loadErr)
                        Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase $section -Tags $tags -Action 'step' -Status 'fail' -Message $loadErr -Data @{ failure = 'step_not_found' }
                        if ($section -eq 'setup') { $stopped = $true }
                        continue
                    }
                }
                $entry    = $registry[$use]
                $manifest = $entry['Manifest']

                # -- inputs: strip "as", resolve session names ------------------
                $resolved = Resolve-EbiStepInputs -Manifest $manifest -With $with -Session $ctx['Session']
                if (-not $resolved['ok']) {
                    [void]$records.Add(@{ section = $section; id = $id; use = $use; status = 'fail'; failure = $resolved['failure']; message = $resolved['message']; outputs = @{}; warnings = @() })
                    Write-EbiStepLine -Status 'fail' -Section $section -Id $id -Use $use -Detail ($resolved['failure'] + ': ' + $resolved['message'])
                    Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase $section -Tags $tags -Action 'step' -Status 'fail' -Message $resolved['message'] -Data @{ failure = $resolved['failure'] }
                    if ($section -eq 'setup') { $stopped = $true }
                    continue
                }
                $in = $resolved['In']
                $as = $resolved['As']

                # -- call -------------------------------------------------------
                $ret = $null
                $threw = ''
                try {
                    $ret = & $entry['Invoke'] $in $ctx
                } catch {
                    $threw = $_.Exception.Message
                }

                $status = 'ok'; $failure = ''; $message = ''; $outputs = @{}; $warnings = @()
                if ($threw -ne '') {
                    $status = 'fail'; $failure = 'internal_error'; $message = $threw
                } else {
                    $check = Test-EbiStepReturn -Manifest $manifest -Return $ret -WantsResource ($as -ne '')
                    if (-not $check['ok']) {
                        $status = 'fail'; $failure = $check['failure']; $message = $check['message']
                    } else {
                        $outputs  = $check['Outputs']
                        $warnings = $check['Warnings']
                        if (-not [bool]$ret['ok']) {
                            $status  = 'fail'
                            $failure = [string]$ret['failure']
                            $message = if ($ret.Contains('message')) { [string]$ret['message'] } else { '' }
                        } else {
                            # 3.4 point 7: register / release. "resource" never reaches
                            # outputs, trace or the records below.
                            if ($as -ne '') {
                                $provides = @(Get-EbiManifestArray -Manifest $manifest -Key 'provides')
                                $ctx['Session'][$as] = @{ kind = $provides[0]; value = $check['Resource']; registeredBy = $id }
                            }
                            $releases = @(Get-EbiManifestArray -Manifest $manifest -Key 'releases')
                            if ($releases.Count -gt 0) {
                                $sessionInputs = Get-EbiSessionInputs -Manifest $manifest
                                foreach ($name in $sessionInputs.Keys) {
                                    if (-not $with.Contains($name)) { continue }
                                    if ($releases -contains [string]$sessionInputs[$name]) {
                                        $ctx['Session'].Remove([string]$with[$name])
                                    }
                                }
                            }
                        }
                    }
                }

                $rec = @{ section = $section; id = $id; use = $use; status = $status; failure = $failure; message = $message; outputs = $outputs; warnings = $warnings }
                [void]$records.Add($rec)
                $detail = if ($status -eq 'ok') { Format-EbiOutputs $outputs } else { $failure + ': ' + $message }
                $warnCount = @($warnings).Count
                if ($warnCount -gt 0) { $detail += ('  ({0} warning{1})' -f $warnCount, $(if ($warnCount -eq 1) { '' } else { 's' })) }
                Write-EbiStepLine -Status $status -Section $section -Id $id -Use $use -Detail $detail
                $data = @{ outputs = $outputs }
                if ($warnCount -gt 0) { $data['warnings'] = $warnings }
                if ($status -ne 'ok') { $data['failure'] = $failure }
                Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase $section -Tags $tags -Action 'step' -Status $status -Message $message -Data $data

                if ($status -ne 'ok' -and $section -eq 'setup') { $stopped = $true }
            }
        }
    } finally {
        # An unexpected exception inside the loop above (a runner bug, not a
        # step failure -- those are caught per call) still lands here, so a
        # teardown that was already reached keeps its 1.1 guarantee. The
        # exception itself propagates after this block.
        $result['steps'] = $records.ToArray()
        $failed = @($records.ToArray() | Where-Object { $_['status'] -eq 'fail' })
        $result['ok'] = ($failed.Count -eq 0)
        if ($failed.Count -gt 0) {
            $result['failure'] = [string]$failed[0]['failure']
            $result['message'] = ('{0}/{1}: {2}' -f $failed[0]['section'], $failed[0]['id'], $failed[0]['message'])
        }
        $left = @($ctx['Session'].Keys)
        Write-TraceEvent -WorkDir $WorkDir -RunId $RunId -Phase 'run' -Tags $tagsBase -Action 'run' -Status $(if ($result['ok']) { 'ok' } else { 'fail' }) -Message $result['message'] -Data @{ steps = $records.Count; failed = $failed.Count; sessionLeft = $left }
        Write-Host ('  run {0}  {1}  ({2} step{3}, {4} failed{5})' -f $RunId, $(if ($result['ok']) { 'OK' } else { 'FAIL' }), $records.Count, $(if ($records.Count -eq 1) { '' } else { 's' }), $failed.Count, $(if ($left.Count -gt 0) { '; still registered: ' + ($left -join ', ') } else { '' })) -ForegroundColor $(if ($result['ok']) { 'Green' } else { 'Red' })
    }
    return $result
}
