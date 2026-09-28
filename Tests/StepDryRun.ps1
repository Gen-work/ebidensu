# StepDryRun.ps1 -- the DryRun contract harness (P1-36). Dot-source only
# (no param() block, ASCII source). Test-StepDryRun.ps1 runs it over every
# shipped step; Test-Steps.ps1 reuses it.
#
# For one step: load it through the registry, take manifest.example.with,
# replace every {{template}} with a fixture of the declared type (a
# session input gets a fake resource of its kind registered under the
# example's name), pass the result through the P1-02 schema check, call
# Invoke-Step with $Ctx.DryRun = $true, and judge the return:
#   - a hashtable with ok = $true that passes Test-EbiStepReturn
#   - EVERY key of manifest.outputs present in it (ledger replay and
#     template resolution assume return keys == outputs)
#   - JSON-serializable without loss (kernel/Json.ps1)
#   - a ui / write / destructive step said what it would do ($Ctx.Log)
#   - nothing was written under the temp work dir
# Every problem is one line naming the step and the key.

function New-StepDryRunLog {
    $log = New-Object PSObject
    $log | Add-Member -MemberType NoteProperty -Name Lines -Value (New-Object System.Collections.ArrayList)
    $log | Add-Member -MemberType ScriptMethod -Name Info  -Value { param($m) [void]$this.Lines.Add('info:' + $m) }
    $log | Add-Member -MemberType ScriptMethod -Name Warn  -Value { param($m) [void]$this.Lines.Add('warn:' + $m) }
    $log | Add-Member -MemberType ScriptMethod -Name Debug -Value { param($m) }
    return $log
}

function New-StepDryRunWorklist {
    param([string]$TmpRoot)
    return @{ path = (Join-Path $TmpRoot 'wl.csv'); columns = @('Correl_ID_S', 'JOB_NAME', 'before_transferStatus', 'composed', 'note'); rows = @(@{ Correl_ID_S = 'ABC123'; JOB_NAME = 'J1'; before_transferStatus = ''; composed = '0'; note = '' }) }
}

function Get-StepDryRunNamedFixtures {
    # Inputs whose contract needs a particular shape -- what the workflow's
    # template would have produced -- keyed by input name.
    return @{
        value       = 'ok'                                            # {{steps.gate.out.code}}
        code        = 'unknown'                                       # so a gate really asks (and auto-answers)
        candidates  = @{ candidates = @(@{ id = 'c1'; candidate = 'fixture'; evidence = @{ source = 'test' } }); suggestion = @{ id = 'c1'; reason = 'fixture' }; doubts = '' }
        grammar     = @{ parser = 'regex'; pattern = '^(?<key>\S+)\s+(?<time>.+)$' }
        rules       = @{ rules = @(@{ field = 'key'; op = 'present'; else = 'unknown'; message = 'fixture' }); default = 'ok' }
        fingerprint = @{ ok = @('fixture') }
        records     = @(@{ key = 'fixture'; time = '2026/09/28 9:00:00' })
        record      = @{ key = 'fixture' }
        text        = 'fixture 2026/09/28 9:00:00'
        key         = 'fixture'
    }
}

function ConvertTo-StepDryRunWith {
    <#
      PURE. manifest -> @{ with; session } where with is example.with with
      every template replaced by a fixture, and session holds a fake
      resource for every session input the example names (unless the
      example carries 'as': a provides step registers its own).
    #>
    param($Manifest, [string]$TmpRoot)
    $with = @{}
    $ex = if ($Manifest.Contains('example') -and ($Manifest['example'] -is [System.Collections.IDictionary])) { $Manifest['example'] } else { @{} }
    if ($ex.Contains('with') -and ($ex['with'] -is [System.Collections.IDictionary])) { foreach ($k in $ex['with'].Keys) { $with[[string]$k] = $ex['with'][$k] } }
    $inputs = if ($Manifest.Contains('inputs') -and ($Manifest['inputs'] -is [System.Collections.IDictionary])) { $Manifest['inputs'] } else { @{} }
    $named = Get-StepDryRunNamedFixtures
    foreach ($k in @($with.Keys)) {
        $v = $with[$k]
        if (-not ($v -is [string]) -or -not (Test-EbiTemplateString $v)) { continue }
        $spec = if ($inputs.Contains($k) -and ($inputs[$k] -is [System.Collections.IDictionary])) { $inputs[$k] } else { @{} }
        $type = if ($spec.Contains('type')) { [string]$spec['type'] } else { 'string' }
        $enum = @(if ($spec.Contains('enum') -and $null -ne $spec['enum']) { $spec['enum'] })
        if ($enum.Count -gt 0) { $with[$k] = $enum[0]; continue }
        if ($named.Contains($k)) { $with[$k] = $named[$k]; continue }
        switch ($type) {
            'int'     { $with[$k] = 100 }
            'bool'    { $with[$k] = $false }
            'map'     { $with[$k] = @{ ok = @('fixture') } }
            'list'    { $with[$k] = @('fixture') }
            'rect'    { $with[$k] = @{ x = 0; y = 0; w = 1; h = 1 } }
            'path'    { $with[$k] = 'fixture/' + $k }
            'session' { $with[$k] = 'fixture' + $k }
            default   { $with[$k] = 'fixture' }
        }
    }
    $session = @{}
    if (-not $with.Contains('as')) {
        foreach ($k in $inputs.Keys) {
            $spec = $inputs[$k]
            if (-not ($spec -is [System.Collections.IDictionary]) -or [string]$spec['type'] -ne 'session' -or -not $with.Contains([string]$k)) { continue }
            $kind = [string]$spec['sessionKind']
            $value = switch ($kind) { 'window' { 4242 } 'worklist' { New-StepDryRunWorklist -TmpRoot $TmpRoot } default { @{ fake = $kind } } }
            $session[[string]$with[[string]$k]] = @{ kind = $kind; value = $value; registeredBy = 'dryrun-harness' }
        }
    }
    return @{ with = $with; session = $session }
}

function Invoke-StepDryRunCheck {
    <#
      Load and dry-run one step; -> @{ ok; problems (string[]); ret; log }.
      Loads into THIS function's scope (Import-EbiStep is dot-sourced here),
      so the step's helpers never leak into the caller.
    #>
    param([hashtable]$Registry, [string]$Use, [string]$TmpRoot)
    $problems = New-Object System.Collections.ArrayList
    $r = . Import-EbiStep -Registry $Registry -Use $Use
    if (-not $r['ok']) { [void]$problems.Add($Use + ': does not load: ' + $r['message']); return @{ ok = $false; problems = $problems.ToArray(); ret = $null; log = $null } }
    $m = $r['Entry']['Manifest']
    $fx = ConvertTo-StepDryRunWith -Manifest $m -TmpRoot $TmpRoot
    $log = New-StepDryRunLog
    $ctx = @{ WorkDir = $TmpRoot; RunId = 'dryrun'; Profile = @{}; Log = $log; DryRun = $true; Session = $fx['session']; Item = $null; KeyColumns = @('Correl_ID_S', 'JOB_NAME') }
    if ($ctx['Session'].Count -gt 0) { foreach ($n in $ctx['Session'].Keys) { if ([string]$ctx['Session'][$n]['kind'] -eq 'worklist') { $ctx['Item'] = $ctx['Session'][$n]['value']['rows'][0] } } }
    $res = Resolve-EbiStepInputs -Manifest $m -With $fx['with'] -Session $ctx['Session']
    if (-not $res['ok']) { [void]$problems.Add($Use + ': example inputs do not pass the schema: ' + $res['message']); return @{ ok = $false; problems = $problems.ToArray(); ret = $null; log = $log } }
    $before = @(Get-ChildItem -LiteralPath $TmpRoot -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
    $ret = $null
    try { $ret = & $r['Entry']['Invoke'] $res['In'] $ctx } catch { [void]$problems.Add($Use + ': Invoke-Step threw under DryRun: ' + $_.Exception.Message); return @{ ok = $false; problems = $problems.ToArray(); ret = $null; log = $log } }
    if (-not ($ret -is [hashtable])) { [void]$problems.Add($Use + ': DryRun returned ' + $(if ($null -eq $ret) { 'null' } else { $ret.GetType().Name }) + ', not a hashtable'); return @{ ok = $false; problems = $problems.ToArray(); ret = $ret; log = $log } }
    $chk = Test-EbiStepReturn -Manifest $m -Return $ret -WantsResource ($res['As'] -ne '')
    if (-not $chk['ok']) { [void]$problems.Add($Use + ': ' + $chk['message']) }
    if (-not ($ret.Contains('ok') -and $ret['ok'] -eq $true)) { [void]$problems.Add($Use + ': DryRun did not return ok=$true' + $(if ($ret.Contains('message')) { ' (' + [string]$ret['message'] + ')' } else { '' })) }
    $outputs = if ($m.Contains('outputs') -and ($m['outputs'] -is [System.Collections.IDictionary])) { $m['outputs'] } else { @{} }
    foreach ($k in $outputs.Keys) { if (-not $ret.Contains([string]$k)) { [void]$problems.Add($Use + ": DryRun return lacks output '" + $k + "' declared in the manifest") } }
    if (-not (Test-EbiJsonSerializable -Value $ret)) { [void]$problems.Add($Use + ': DryRun return is not JSON-serializable (a handle, a COM object, or nesting past depth 20)') }
    $effects = if ($m.Contains('effects')) { [string]$m['effects'] } else { '' }
    if (($effects -in @('ui', 'write', 'destructive')) -and $log.Lines.Count -eq 0) { [void]$problems.Add($Use + ': a ' + $effects + ' step said nothing on DryRun ($Ctx.Log)') }
    $after = @(Get-ChildItem -LiteralPath $TmpRoot -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
    if ($after.Count -gt $before.Count) { [void]$problems.Add($Use + ': DryRun wrote file(s): ' + (@($after | Where-Object { $before -notcontains $_ }) -join ', ')) }
    return @{ ok = ($problems.Count -eq 0); problems = $problems.ToArray(); ret = $ret; log = $log }
}
