#Requires -Version 5.1
# ============================================================
#  kernel/Lint.ps1
#
#  `ebi lint` (P1-08): the static checks of WORKFLOW-SCHEMA.md section 9,
#  run over a workflow JSON, the step catalog and (when available) the
#  profile, without running anything. Dot-source only (no param() block,
#  ASCII source, no class).
#
#  Invoke-EbiLint is PURE given its inputs and returns
#      @{ ok; errors = @(@{ where; message }); warnings = @(...) }
#  ok is "no errors" -- warnings never fail a lint. Every finding names
#  where it is (section/id, or a top-level field). The shape checks the
#  runner already refuses at load time (Get-EbiWorkflowProblems) are the
#  first pass here, so lint and runner can never disagree about them.
#
#  What is checked (the section 9 list, plus the items the specs scatter
#  and section 9 collects):
#    schema / id / sections / unique ids / pendingWhen / when / once /
#    onError shapes                                (runner shape checks)
#    use exists in the catalog
#    with: undeclared parameters, missing required, literal values of the
#      wrong type (a templated value cannot be type-checked statically)
#    templates: every {{...}} resolves -- vars keys, run keys, item only in
#      each (and its columns when the profile declares them), page needs
#      the top-level binding, profile/page paths exist in the loaded
#      profile, steps.<id> is an earlier step of the same section
#    session resources: "as" only on a provides step, and REQUIRED on one
#      (P0-R17); a live name is not registered twice; every type='session'
#      input is a literal name registered earlier by a matching kind;
#      source.table is registered in setup by a provides-worklist call;
#      a name of a kind marked mustRelease in STEP-CONTRACT 3.4 point 6 is
#      released later
#    idempotent: every setup step; every provides / releases step
#    onError.byFailure: known failure id (manifest or reserved); retry only
#      on a transient one
#    warnings: tier fallback in use; destructive with confirm:false; needs
#      excel / browser / calibrated:*; profile not loaded; item column not
#      declared by the profile
#    profile: page exists; worklist role:key columns == key.columns
# ============================================================

. (Join-Path $PSScriptRoot 'Runner.ps1')

function Get-EbiMustReleaseKinds {
    <#
      The resource-kind table of STEP-CONTRACT.md 3.4 point 6, parsed out
      of the spec (the same rule Tests/StepContract.ps1 applies): kind ->
      [bool] mustRelease. Empty when the spec cannot be read -- callers
      treat that as "cannot check", not "nothing to release".
    #>
    param([string]$SpecPath = '')
    if ([string]::IsNullOrWhiteSpace($SpecPath)) { $SpecPath = Join-Path (Join-Path (Join-Path (Split-Path $PSScriptRoot -Parent) 'docs') 'ebi-dance') (Join-Path 'spec' 'STEP-CONTRACT.md') }
    $kinds = @{}
    if (-not (Test-Path -LiteralPath $SpecPath -PathType Leaf)) { return $kinds }
    $text = ''
    try { $text = [System.IO.File]::ReadAllText($SpecPath, (New-Object System.Text.UTF8Encoding($false))) } catch { return $kinds }
    $rowRx = [regex]'^\s*\|\s*`([A-Za-z][A-Za-z0-9_]*)`\s*\|\s*`\$(true|false)`'
    $inTable = $false
    foreach ($line in ($text -split "`r?`n")) {
        if (-not $inTable) {
            if ($line -match '`mustRelease`' -and $line.TrimStart().StartsWith('|')) { $inTable = $true }
            continue
        }
        $m = $rowRx.Match($line)
        if ($m.Success) { $kinds[$m.Groups[1].Value] = ($m.Groups[2].Value -eq 'true'); continue }
        if ($line.TrimStart().StartsWith('|')) { continue }
        if ($kinds.Count -gt 0) { break }
    }
    return $kinds
}

function ConvertTo-EbiLintCatalog {
    # Catalog entries (Get-EbiCatalogEntries) -> use -> manifest map.
    param($Entries)
    $map = @{}
    foreach ($e in @($Entries)) { if ($e['ok']) { $map[[string]$e['use']] = $e['Manifest'] } }
    return $map
}

function Invoke-EbiLint {
    <#
      PURE. See the file header. Catalog is a use -> manifest map
      (ConvertTo-EbiLintCatalog); Profile a hashtable or $null;
      MustRelease a kind -> bool map (Get-EbiMustReleaseKinds) or $null.
    #>
    param([hashtable]$Workflow, [hashtable]$Catalog, $Profile = $null, $MustRelease = $null)
    $errors   = New-Object System.Collections.ArrayList
    $warnings = New-Object System.Collections.ArrayList
    function Err  { param([string]$Where, [string]$Message) [void]$errors.Add(@{ where = $Where; message = $Message }) }
    function Warn { param([string]$Where, [string]$Message) [void]$warnings.Add(@{ where = $Where; message = $Message }) }
    if ($null -eq $Catalog) { $Catalog = @{} }
    $hasProfile = ($Profile -is [System.Collections.IDictionary]) -and $Profile.Count -gt 0

    # -- 1. the runner's shape checks ----------------------------------------
    foreach ($p in @(Get-EbiWorkflowProblems -Workflow $Workflow)) { Err -Where 'workflow' -Message $p }
    if ($null -eq $Workflow) { return @{ ok = $false; errors = $errors.ToArray(); warnings = $warnings.ToArray() } }

    $pageName = if ($Workflow.Contains('page') -and $null -ne $Workflow['page']) { [string]$Workflow['page'] } else { '' }
    $vars = if ($Workflow.Contains('vars') -and ($Workflow['vars'] -is [System.Collections.IDictionary])) { $Workflow['vars'] } else { @{} }
    $source = if ($Workflow.Contains('source') -and ($Workflow['source'] -is [System.Collections.IDictionary])) { $Workflow['source'] } else { $null }
    $groupBy = if ($null -ne $source -and $source.Contains('groupBy') -and $null -ne $source['groupBy']) { [string]$source['groupBy'] } else { '' }
    $runKeys = @('runId', 'startedAt', 'date', 'dateSlash', 'mmdd', 'toolDir', 'operator', 'workDir', 'timeWindow')

    # -- 2. profile-level facts -----------------------------------------------
    if (-not $hasProfile) {
        Warn -Where 'profile' -Message 'no profile loaded: profile.* / page.* references and worklist columns are not checked'
    } else {
        if ($pageName -ne '') {
            $pages = if ($Profile.Contains('pages') -and ($Profile['pages'] -is [System.Collections.IDictionary])) { $Profile['pages'] } else { @{} }
            if (-not $pages.Contains($pageName)) { Err -Where 'page' -Message ('page "' + $pageName + '" is not in the profile''s pages.json (' + $(if ($pages.Count) { ($pages.Keys | Sort-Object) -join ', ' } else { 'none' }) + ')') }
        }
        $wl = if ($Profile.Contains('worklist') -and ($Profile['worklist'] -is [System.Collections.IDictionary])) { $Profile['worklist'] } else { $null }
        if ($null -ne $wl) {
            $keyCols = @(Get-EbiWorklistKeyColumns -Profile $Profile)
            $roleKey = New-Object System.Collections.ArrayList
            if ($wl.Contains('columns') -and $null -ne $wl['columns']) {
                foreach ($c in $wl['columns']) { if ($c -is [System.Collections.IDictionary] -and $c.Contains('role') -and [string]$c['role'] -eq 'key' -and $c.Contains('name')) { [void]$roleKey.Add([string]$c['name']) } }
            }
            foreach ($k in $keyCols) { if (-not ($roleKey -contains $k)) { Err -Where 'profile.worklist' -Message ('key.columns names "' + $k + '" but no column declares role: key for it (PROFILE-SCHEMA 6.1)') } }
            foreach ($k in $roleKey) { if (-not ($keyCols -contains $k)) { Err -Where 'profile.worklist' -Message ('column "' + $k + '" has role: key but is not in key.columns (PROFILE-SCHEMA 6.1)') } }
        }
    }
    $columnNames = New-Object System.Collections.ArrayList
    if ($hasProfile -and $Profile.Contains('worklist') -and ($Profile['worklist'] -is [System.Collections.IDictionary]) -and $Profile['worklist'].Contains('columns') -and $null -ne $Profile['worklist']['columns']) {
        foreach ($c in $Profile['worklist']['columns']) { if ($c -is [System.Collections.IDictionary] -and $c.Contains('name')) { [void]$columnNames.Add([string]$c['name']) } }
    }
    $staticScope = New-EbiTemplateScope -Vars $vars -Profile $(if ($hasProfile) { $Profile } else { @{} }) -PageName $pageName -Run @{} -Item $null -Steps @{}

    # -- 3. walk the sections in execution order ------------------------------
    $registered = @{}      # name -> @{ kind; where; live }
    $usedManifests = New-Object System.Collections.ArrayList
    $fallbackUsed = New-Object System.Collections.ArrayList
    $needsSeen = New-Object System.Collections.ArrayList

    function Test-Ref {
        # One template / when reference path, in context.
        param([string]$Path, [string]$Where, [string]$Section, $StepsSoFar)
        $p = $Path.Trim()
        $dot = $p.IndexOf('.')
        $prefix = if ($dot -lt 0) { $p } else { $p.Substring(0, $dot) }
        $rest = if ($dot -lt 0) { '' } else { $p.Substring($dot + 1) }
        switch ($prefix) {
            'vars' {
                $first = ($rest -split '\.')[0]
                if ($rest -eq '' -or -not $vars.Contains($first)) { Err -Where $Where -Message ('{{' + $p + '}}: vars has no "' + $first + '" (declared: ' + $(if ($vars.Count) { ($vars.Keys | Sort-Object) -join ', ' } else { 'none' }) + ')') }
            }
            'run' {
                $first = ($rest -split '\.')[0]
                if (-not ($runKeys -contains $first)) { Err -Where $Where -Message ('{{' + $p + '}}: run has no "' + $first + '" (' + ($runKeys -join ', ') + ')') }
            }
            'item' {
                if ($Section -ne 'each') { Err -Where $Where -Message ('{{' + $p + '}}: item.* is only available in "each"') }
                else {
                    $first = ($rest -split '\.')[0]
                    if ($first -eq '') { Err -Where $Where -Message ('{{' + $p + '}}: item needs a column or key / keySafe / group') }
                    elseif ($first -ne 'key' -and $first -ne 'keySafe' -and $first -ne 'group' -and $columnNames.Count -gt 0 -and -not ($columnNames -contains $first)) {
                        Warn -Where $Where -Message ('{{' + $p + '}}: the profile''s worklist.json declares no column "' + $first + '"')
                    }
                }
            }
            'page' {
                if ($pageName -eq '') { Err -Where $Where -Message ('{{' + $p + '}}: the workflow has no top-level "page" binding (P0-R6)') }
                elseif ($hasProfile) {
                    $r = Resolve-EbiPath -Scope $staticScope -Path $p
                    if (-not $r['ok']) { Err -Where $Where -Message ('{{' + $p + '}}: ' + [string]$r['message']) }
                }
            }
            'profile' {
                if ($hasProfile) {
                    $r = Resolve-EbiPath -Scope $staticScope -Path $p
                    if (-not $r['ok']) { Err -Where $Where -Message ('{{' + $p + '}}: ' + [string]$r['message']) }
                }
            }
            'steps' {
                $b = $rest.IndexOf('.out.')
                if ($b -le 0) {
                    if ($rest.EndsWith('.out')) { Err -Where $Where -Message ('{{' + $p + '}}: name a field after .out.') }
                    else { Err -Where $Where -Message ('{{' + $p + '}}: a steps reference is steps.<id>.out.<field>') }
                } else {
                    $sid = $rest.Substring(0, $b)
                    if (-not ($StepsSoFar -contains $sid)) { Err -Where $Where -Message ('{{' + $p + '}}: step "' + $sid + '" is not an earlier step of this section') }
                }
            }
            default { Err -Where $Where -Message ('{{' + $p + '}}: unknown scope "' + $prefix + '" (vars, profile, page, run, item, steps)') }
        }
    }

    foreach ($section in @('setup', 'each', 'teardown')) {
        if (-not $Workflow.Contains($section) -or -not ($Workflow[$section] -is [System.Collections.IList])) { continue }
        $stepsSoFar = New-Object System.Collections.ArrayList
        foreach ($call in $Workflow[$section]) {
            if (-not ($call -is [System.Collections.IDictionary])) { continue }
            $id  = if ($call.Contains('id'))  { [string]$call['id'] }  else { '' }
            $use = if ($call.Contains('use')) { [string]$call['use'] } else { '' }
            $where = $section + '/' + $(if ($id -ne '') { $id } else { '?' })
            $with = if ($call.Contains('with') -and ($call['with'] -is [System.Collections.IDictionary])) { $call['with'] } else { @{} }
            $manifest = $null
            if ($use -ne '' -and $Catalog.Contains($use)) { $manifest = $Catalog[$use] }
            elseif ($use -ne '') { Err -Where $where -Message ('use "' + $use + '" is not in the catalog') }

            # templates and when, whatever the manifest
            foreach ($k in $with.Keys) {
                if ([string]$k -eq 'as') { continue }
                foreach ($ref in @(Get-EbiTemplateReferences -Value $with[$k])) { Test-Ref -Path $ref -Where ($where + '.with.' + [string]$k) -Section $section -StepsSoFar $stepsSoFar }
                if ($with[$k] -is [string]) { $t = Get-EbiTemplateTokens -Value $with[$k]; if (-not $t['ok']) { Err -Where ($where + '.with.' + [string]$k) -Message ([string]$t['message']) } }
            }
            if ($call.Contains('when') -and $null -ne $call['when']) {
                $w = ConvertFrom-EbiWhen -Text ([string]$call['when'])
                if ($w['ok']) { Test-Ref -Path $w['path'] -Where ($where + '.when') -Section $section -StepsSoFar $stepsSoFar }
            }

            if ($null -ne $manifest) {
                [void]$usedManifests.Add($manifest)
                $provides = @(Get-EbiManifestArray -Manifest $manifest -Key 'provides')
                $releases = @(Get-EbiManifestArray -Manifest $manifest -Key 'releases')
                $idem = ($manifest.Contains('idempotent') -and ($manifest['idempotent'] -is [bool]) -and $manifest['idempotent'])
                $tier = if ($manifest.Contains('tier')) { [string]$manifest['tier'] } else { '' }
                $effects = if ($manifest.Contains('effects')) { [string]$manifest['effects'] } else { '' }
                if ($tier -eq 'fallback') { [void]$fallbackUsed.Add($use) }
                if ($effects -eq 'destructive' -and $call.Contains('confirm') -and ($call['confirm'] -is [bool]) -and -not $call['confirm']) {
                    Warn -Where $where -Message ('destructive step "' + $use + '" runs WITHOUT its confirm gate (confirm: false)')
                }
                foreach ($n in @(Get-EbiManifestArray -Manifest $manifest -Key 'needs')) {
                    if ($n -eq 'excel' -or $n -eq 'browser' -or $n.StartsWith('calibrated:')) { if (-not ($needsSeen -contains $n)) { [void]$needsSeen.Add($n) } }
                }
                if ($section -eq 'setup' -and -not $idem) { Err -Where $where -Message ('"' + $use + '" is not idempotent; setup re-runs on every resume (STEP-CONTRACT 6.2)') }
                if (($provides.Count -gt 0 -or $releases.Count -gt 0) -and -not $idem) { Err -Where $where -Message ('"' + $use + '" provides / releases a resource but is not idempotent (STEP-CONTRACT 6.2)') }

                # "as" and registration
                $as = if ($with.Contains('as')) { [string]$with['as'] } else { '' }
                if ($as -ne '' -and $provides.Count -eq 0) { Err -Where $where -Message ('"as" given but "' + $use + '" provides nothing') }
                if ($as -eq '' -and $provides.Count -gt 0) { Err -Where $where -Message ('"' + $use + '" provides "' + $provides[0] + '" but the call has no "as": every provides call must register its resource (STEP-CONTRACT 3.4 point 7)') }
                if ($as -ne '' -and $provides.Count -gt 0) {
                    if ($registered.Contains($as) -and $registered[$as]['live']) { Err -Where $where -Message ('"as": "' + $as + '" is already registered by ' + $registered[$as]['where'] + ' and not released yet (STEP-CONTRACT 3.4 point 3)') }
                    $registered[$as] = @{ kind = $provides[0]; where = $where; live = $true; section = $section; once = $(if ($call.Contains('once')) { [string]$call['once'] } else { '' }) }
                }

                # inputs against the schema
                $inputs = if ($manifest.Contains('inputs') -and ($manifest['inputs'] -is [System.Collections.IDictionary])) { $manifest['inputs'] } else { @{} }
                foreach ($k in $with.Keys) {
                    if ([string]$k -eq 'as') { continue }
                    if (-not $inputs.Contains([string]$k)) { Err -Where $where -Message ('"' + $use + '" declares no input "' + [string]$k + '" (declared: ' + $(if ($inputs.Count) { ($inputs.Keys | Sort-Object) -join ', ' } else { 'none' }) + ')') }
                }
                foreach ($name in ($inputs.Keys | Sort-Object)) {
                    $spec = $inputs[$name]
                    if (-not ($spec -is [System.Collections.IDictionary])) { continue }
                    $type = if ($spec.Contains('type') -and $null -ne $spec['type']) { [string]$spec['type'] } else { 'any' }
                    $required = $spec.Contains('required') -and [bool]$spec['required']
                    $given = $with.Contains([string]$name) -and $null -ne $with[[string]$name]
                    if (-not $given) {
                        if ($required) { Err -Where $where -Message ('"' + $use + '" requires input "' + [string]$name + '"') }
                        continue
                    }
                    $v = $with[[string]$name]
                    if ($type -eq 'session') {
                        if (-not ($v -is [string]) -or (Test-EbiTemplateString $v)) { Err -Where $where -Message ('input "' + [string]$name + '" is a session name: a literal string, not a template') ; continue }
                        $kind = if ($spec.Contains('sessionKind')) { [string]$spec['sessionKind'] } else { '' }
                        if (-not $registered.Contains($v) -or -not $registered[$v]['live']) { Err -Where $where -Message ('input "' + [string]$name + '" names session resource "' + $v + '", which no earlier call registered with "as" (or it was released)') }
                        elseif ($kind -ne '' -and [string]$registered[$v]['kind'] -ne $kind) { Err -Where $where -Message ('input "' + [string]$name + '" wants kind "' + $kind + '" but "' + $v + '" is a "' + [string]$registered[$v]['kind'] + '" (registered by ' + $registered[$v]['where'] + ')') }
                        if ($releases -contains $kind -and $registered.Contains($v)) { $registered[$v]['live'] = $false }
                        continue
                    }
                    if (Test-EbiValueHasTemplate $v) { continue }   # cannot be typed statically
                    $c = ConvertTo-EbiInputValue -Type $type -Value $v
                    if (-not $c['ok']) { Err -Where $where -Message ('input "' + [string]$name + '" ' + [string]$c['message']) }
                    elseif ($spec.Contains('enum') -and $null -ne $spec['enum']) {
                        $allowed = @($spec['enum'] | ForEach-Object { [string]$_ })
                        if (-not ($allowed -contains [string]$c['value'])) { Err -Where $where -Message ('input "' + [string]$name + '" must be one of ' + ($allowed -join ', ') + '; got "' + [string]$v + '"') }
                    }
                }

                # per-call onError.byFailure against this manifest
                if ($call.Contains('onError') -and ($call['onError'] -is [System.Collections.IDictionary]) -and $call['onError'].Contains('byFailure') -and ($call['onError']['byFailure'] -is [System.Collections.IDictionary])) {
                    $declared = @(Get-EbiManifestFailureIds -Manifest $manifest)
                    foreach ($fid in $call['onError']['byFailure'].Keys) {
                        $f = [string]$fid
                        $known = ($declared -contains $f) -or ((Get-EbiRunnerFailureIds) -contains $f)
                        if (-not $known) { Err -Where ($where + '.onError.byFailure') -Message ('"' + $f + '" is not a failure "' + $use + '" declares, nor a reserved id') }
                        $pol = $call['onError']['byFailure'][$fid]
                        if ($pol -is [System.Collections.IDictionary] -and $pol.Contains('policy') -and [string]$pol['policy'] -eq 'retry' -and -not (Test-EbiFailureTransient -Manifest $manifest -FailureId $f)) {
                            Err -Where ($where + '.onError.byFailure.' + $f) -Message ('policy retry on "' + $f + '", which is not transient (P0-R5)')
                        }
                    }
                }
            }
            if ($id -ne '') { [void]$stepsSoFar.Add($id) }
        }
    }

    # -- 4. top-level onError.byFailure against every used manifest -----------
    if ($Workflow.Contains('onError') -and ($Workflow['onError'] -is [System.Collections.IDictionary]) -and $Workflow['onError'].Contains('byFailure') -and ($Workflow['onError']['byFailure'] -is [System.Collections.IDictionary])) {
        foreach ($fid in $Workflow['onError']['byFailure'].Keys) {
            $f = [string]$fid
            $known = (Get-EbiRunnerFailureIds) -contains $f
            $transientSomewhere = $false
            foreach ($m in $usedManifests) {
                if (@(Get-EbiManifestFailureIds -Manifest $m) -contains $f) { $known = $true }
                if (Test-EbiFailureTransient -Manifest $m -FailureId $f) { $transientSomewhere = $true }
            }
            if (-not $known) { Err -Where 'onError.byFailure' -Message ('"' + $f + '" is not a failure any used step declares, nor a reserved id') }
            $pol = $Workflow['onError']['byFailure'][$fid]
            if ($pol -is [System.Collections.IDictionary] -and $pol.Contains('policy') -and [string]$pol['policy'] -eq 'retry' -and -not $transientSomewhere) {
                Err -Where ('onError.byFailure.' + $f) -Message ('policy retry on "' + $f + '", which no used step marks transient (P0-R5)')
            }
        }
    }

    # -- 5. source.table and mustRelease ---------------------------------------
    if ($null -ne $source -and $source.Contains('table') -and -not [string]::IsNullOrWhiteSpace([string]$source['table'])) {
        $t = [string]$source['table']
        if (-not $registered.Contains($t)) { Err -Where 'source.table' -Message ('"' + $t + '" is not registered by any setup call with "as": "' + $t + '" (a provides-worklist step, P0-R11)') }
        elseif ([string]$registered[$t]['kind'] -ne 'worklist') { Err -Where 'source.table' -Message ('"' + $t + '" is a "' + [string]$registered[$t]['kind'] + '", not a worklist') }
        elseif ([string]$registered[$t]['section'] -ne 'setup') { Err -Where 'source.table' -Message ('"' + $t + '" must be registered in setup, not ' + [string]$registered[$t]['section']) }
    }
    if ($MustRelease -is [System.Collections.IDictionary] -and $MustRelease.Count -gt 0) {
        foreach ($name in ($registered.Keys | Sort-Object)) {
            $r = $registered[$name]
            $kind = [string]$r['kind']
            if ($MustRelease.Contains($kind) -and [bool]$MustRelease[$kind] -and $r['live']) {
                Err -Where $r['where'] -Message ('"' + $name + '" (' + $kind + ') is registered but never released; the kind is marked mustRelease in STEP-CONTRACT 3.4 point 6 -- add a releases step later in the same section (once:groupEnd for once:group) or in teardown')
            }
        }
    } elseif ($registered.Count -gt 0) {
        Warn -Where 'workflow' -Message 'the mustRelease kind table could not be read from STEP-CONTRACT.md; release checks skipped'
    }

    # -- 6. warnings about the machine ------------------------------------------
    if ($fallbackUsed.Count -gt 0) { Warn -Where 'workflow' -Message ('uses fallback-tier step(s): ' + (($fallbackUsed.ToArray() | Sort-Object -Unique) -join ', ') + ' -- run ebi calibrate first (STEP-CONTRACT 5)') }
    foreach ($n in $needsSeen) { Warn -Where 'workflow' -Message ('needs "' + $n + '" on this machine') }

    return @{ ok = ($errors.Count -eq 0); errors = $errors.ToArray(); warnings = $warnings.ToArray() }
}

function Format-EbiLintReport {
    # Lines for the console: errors then warnings, then a one-line total.
    param([hashtable]$Result, [string]$Path = '')
    $L = New-Object System.Collections.ArrayList
    if ($Path -ne '') { [void]$L.Add('lint ' + $Path) }
    foreach ($e in @($Result['errors']))   { [void]$L.Add('  [ERROR] ' + [string]$e['where'] + ': ' + [string]$e['message']) }
    foreach ($w in @($Result['warnings'])) { [void]$L.Add('  [WARN ] ' + [string]$w['where'] + ': ' + [string]$w['message']) }
    [void]$L.Add(('  {0} error(s), {1} warning(s) -- {2}' -f @($Result['errors']).Count, @($Result['warnings']).Count, $(if ($Result['ok']) { 'OK' } else { 'FAIL' })))
    return $L.ToArray()
}
