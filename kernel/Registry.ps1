#Requires -Version 5.1
# ============================================================
#  kernel/Registry.ps1
#
#  The ebi-dance step registry (P1-02). Dot-source only (no param() block,
#  ASCII source, no class -- CLAUDE.md conventions).
#
#  What lives here, lifted out of the P0-07 runner spike and completed:
#    - where a step file is (Get-EbiStepPath) and which files under
#      modules/ are steps at all (Find-EbiStepFiles / Get-EbiStepCatalog,
#      the manifest-only scan that ebi lint / ebi help / kernel/Docs.ps1
#      read)
#    - loading a step for RUNNING (Import-EbiStep): dot-source the file,
#      capture its Invoke-Step into the registry table AT ONCE -- every
#      step defines the same function name, so the next file loaded would
#      overwrite it -- then remove the bare name, so the table is the only
#      way a step is ever called
#    - the "with" -> $In pipeline of STEP-CONTRACT.md 2.2 / 3.4 point 3
#      (Resolve-EbiStepInputs): "as" is lifted out, every input is checked
#      against the manifest's inputs schema (Test-EbiStepInputs: required,
#      type, enum, default, unknown parameters -- each problem names the
#      parameter), then type='session' names are replaced by the registered
#      instance
#    - the 3.1 return-value contract (Test-EbiStepReturn)
#
#  Scope rule for Import-EbiStep -- read this before calling it:
#    A step's prefixed helper functions (BrowserEnsure-FindProcess, ...) are
#    defined by dot-sourcing the file, and land in whatever scope does the
#    dot-sourcing. If that scope is a registry function's own, they vanish
#    when it returns and the captured Invoke-Step fails at call time with
#    "not recognized". So Import-EbiStep must itself be dot-sourced into
#    the scope that will run the steps:
#
#        $r = . Import-EbiStep -Registry $registry -Use 'browser.ensure'
#
#    It refuses (ok=$false, internal_error) when called any other way,
#    rather than loading a step whose helpers are already gone. Its own
#    locals therefore live in the caller's scope too; they are kept to
#    $Manifest (the contract requires that name) plus one prefixed
#    hashtable, both removed before it returns.
#
#  Hashtable access is by index ($h['k']), never by dot: under
#  Set-StrictMode a missing key read with dot syntax throws.
# ============================================================

# --- input types the runner can check (STEP-CONTRACT.md 2.2) ---------------

function Get-EbiInputTypes {
    return @('string', 'int', 'bool', 'path', 'rect', 'list', 'map', 'session', 'any')
}

function Test-EbiIsNumber {
    # A numeric scalar (not bool, not char, not a string).
    param($Value)
    if ($null -eq $Value) { return $false }
    foreach ($t in @([int], [long], [int16], [byte], [sbyte], [uint16], [uint32], [uint64], [double], [single], [decimal])) {
        if ($Value -is $t) { return $true }
    }
    return $false
}

function Get-EbiValueTypeName {
    # How a value is described in a validation message: what the workflow
    # author wrote, in their terms ('string "abc"', 'number 4.5', 'map').
    param($Value)
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { return ('bool ' + $Value.ToString().ToLowerInvariant()) }
    if ($Value -is [string]) {
        $s = $Value
        if ($s.Length -gt 30) { $s = $s.Substring(0, 27) + '...' }
        return ('string "' + $s + '"')
    }
    if (Test-EbiIsNumber $Value) { return ('number ' + [string]$Value) }
    if ($Value -is [System.Collections.IDictionary]) { return 'map' }
    if ($Value -is [System.Collections.IList]) { return 'list' }
    return $Value.GetType().Name
}

function ConvertTo-EbiInputValue {
    <#
      PURE. Check one value against one 2.2 input type and hand back the
      value as the step expects it. Returns @{ ok; value; message }; a
      message names what was expected and what arrived, never the parameter
      (the caller adds that).

      Leniency is deliberate and small: a whole-value template keeps the
      type of what it pulled out (WORKFLOW-SCHEMA.md 4.2), and worklist
      columns are strings, so an int input accepts "3" and a bool input
      accepts "true". Nothing else is coerced: a fractional number is not an
      int, a bare scalar is not a one-element list.
    #>
    param([string]$Type, $Value)

    switch ($Type) {
        'string' {
            if ($Value -is [string]) { return @{ ok = $true; value = $Value; message = '' } }
            if ($Value -is [bool] -or $Value -is [char] -or (Test-EbiIsNumber $Value)) {
                return @{ ok = $true; value = [string]$Value; message = '' }
            }
            return @{ ok = $false; value = $Value; message = ('expects string, got ' + (Get-EbiValueTypeName $Value)) }
        }
        'path' {
            if ($Value -is [string]) { return @{ ok = $true; value = $Value; message = '' } }
            if (Test-EbiIsNumber $Value) { return @{ ok = $true; value = [string]$Value; message = '' } }
            return @{ ok = $false; value = $Value; message = ('expects path (a string), got ' + (Get-EbiValueTypeName $Value)) }
        }
        'int' {
            if ($Value -is [bool]) {
                return @{ ok = $false; value = $Value; message = ('expects int, got ' + (Get-EbiValueTypeName $Value)) }
            }
            if (Test-EbiIsNumber $Value) {
                $d = [double]$Value
                if ([math]::Floor($d) -ne $d -or $d -gt [int]::MaxValue -or $d -lt [int]::MinValue) {
                    return @{ ok = $false; value = $Value; message = ('expects int, got ' + (Get-EbiValueTypeName $Value)) }
                }
                return @{ ok = $true; value = [int]$d; message = '' }
            }
            if ($Value -is [string]) {
                if ($Value -match '^\s*[-+]?\d{1,10}\s*$') {
                    try { return @{ ok = $true; value = [int]$Value.Trim(); message = '' } } catch { }
                }
                return @{ ok = $false; value = $Value; message = ('expects int, got ' + (Get-EbiValueTypeName $Value)) }
            }
            return @{ ok = $false; value = $Value; message = ('expects int, got ' + (Get-EbiValueTypeName $Value)) }
        }
        'bool' {
            if ($Value -is [bool]) { return @{ ok = $true; value = $Value; message = '' } }
            if ($Value -is [string]) {
                $t = $Value.Trim().ToLowerInvariant()
                if ($t -eq 'true')  { return @{ ok = $true; value = $true;  message = '' } }
                if ($t -eq 'false') { return @{ ok = $true; value = $false; message = '' } }
            }
            return @{ ok = $false; value = $Value; message = ('expects bool (true or false), got ' + (Get-EbiValueTypeName $Value)) }
        }
        'rect' {
            if (-not ($Value -is [System.Collections.IDictionary])) {
                return @{ ok = $false; value = $Value; message = ('expects rect {X,Y,W,H}, got ' + (Get-EbiValueTypeName $Value)) }
            }
            $rect = @{}
            foreach ($k in @('X', 'Y', 'W', 'H')) {
                if (-not $Value.Contains($k)) {
                    return @{ ok = $false; value = $Value; message = ('expects rect {X,Y,W,H}, got a map without ' + $k) }
                }
                $c = ConvertTo-EbiInputValue -Type 'int' -Value $Value[$k]
                if (-not $c['ok']) {
                    return @{ ok = $false; value = $Value; message = ('expects rect {X,Y,W,H}; ' + $k + ' ' + $c['message']) }
                }
                $rect[$k] = $c['value']
            }
            return @{ ok = $true; value = $rect; message = '' }
        }
        'list' {
            if ($Value -is [string] -or -not ($Value -is [System.Collections.IList])) {
                return @{ ok = $false; value = $Value; message = ('expects list, got ' + (Get-EbiValueTypeName $Value)) }
            }
            return @{ ok = $true; value = $Value; message = '' }
        }
        'map' {
            if (-not ($Value -is [System.Collections.IDictionary])) {
                return @{ ok = $false; value = $Value; message = ('expects map, got ' + (Get-EbiValueTypeName $Value)) }
            }
            return @{ ok = $true; value = $Value; message = '' }
        }
        'session' {
            if ($Value -is [string] -and -not [string]::IsNullOrWhiteSpace($Value)) {
                return @{ ok = $true; value = $Value; message = '' }
            }
            return @{ ok = $false; value = $Value; message = ('expects the name of a registered session resource (a non-empty string), got ' + (Get-EbiValueTypeName $Value)) }
        }
        'any' {
            return @{ ok = $true; value = $Value; message = '' }
        }
    }
    return @{ ok = $false; value = $Value; message = ('is declared with unknown type "' + $Type + '" (known: ' + ((Get-EbiInputTypes) -join ', ') + ')') }
}

function Test-EbiStepInputs {
    <#
      PURE. Check a call's parameters (with "as" already lifted out) against
      the manifest's inputs schema and fill defaults.

      Returns @{ ok; failure; message; In; problems } where
        In        the checked hashtable the step receives: defaults filled,
                  values coerced per ConvertTo-EbiInputValue, null values
                  treated as "not given" (a default fills them, a required
                  one is missing, an optional one is dropped)
        problems  @( @{ input; kind; message } ), every one naming the
                  parameter; kinds: missing_required, unknown_input,
                  type_mismatch, enum_mismatch, bad_manifest
        failure   'input_invalid' when the call is wrong,
                  'contract_violation' when the manifest itself is
                  (inputs not a hashtable, an input with an unknown type)
      Problems are reported all at once, declared inputs first (sorted by
      name), then unknown parameters (sorted), so one run shows every
      mistake in the call.
    #>
    param($Manifest, [hashtable]$In)

    $out = @{}
    $problems = New-Object System.Collections.ArrayList
    $manifestBad = $false
    $stepId = if ($null -ne $Manifest -and $Manifest.Contains('id')) { [string]$Manifest['id'] } else { '?' }

    $inputs = $null
    if ($null -ne $Manifest -and $Manifest.Contains('inputs')) { $inputs = $Manifest['inputs'] }
    if ($null -eq $inputs) { $inputs = @{} }
    if (-not ($inputs -is [System.Collections.IDictionary])) {
        [void]$problems.Add(@{ input = ''; kind = 'bad_manifest';
            message = ('manifest of ' + $stepId + ' has inputs that are not a hashtable, so no parameter can be checked') })
        return @{ ok = $false; failure = 'contract_violation'; message = $problems[0]['message']; In = $out; problems = $problems.ToArray() }
    }

    $given = @{}
    if ($null -ne $In) { foreach ($k in $In.Keys) { $given[[string]$k] = $In[$k] } }

    $declared = New-Object System.Collections.ArrayList
    foreach ($k in $inputs.Keys) { [void]$declared.Add([string]$k) }
    $declaredSorted = @($declared.ToArray() | Sort-Object)

    foreach ($name in $declaredSorted) {
        $spec = $inputs[$name]
        if (-not ($spec -is [System.Collections.IDictionary])) {
            [void]$problems.Add(@{ input = $name; kind = 'bad_manifest';
                message = ('manifest of ' + $stepId + ' declares input ''' + $name + ''' without a hashtable spec') })
            $manifestBad = $true
            continue
        }
        $type = if ($spec.Contains('type') -and $null -ne $spec['type']) { [string]$spec['type'] } else { '' }
        if ($type -eq '') { $type = 'any' }
        if ((Get-EbiInputTypes) -notcontains $type) {
            [void]$problems.Add(@{ input = $name; kind = 'bad_manifest';
                message = ('manifest of ' + $stepId + ' declares input ''' + $name + ''' with unknown type "' + $type + '"') })
            $manifestBad = $true
            continue
        }
        $required   = $spec.Contains('required') -and [bool]$spec['required']
        $hasDefault = $spec.Contains('default')
        $present    = $given.Contains($name) -and $null -ne $given[$name]

        if (-not $present) {
            if ($hasDefault) { $out[$name] = $spec['default']; continue }
            if ($required) {
                [void]$problems.Add(@{ input = $name; kind = 'missing_required'; message = ('missing required input ''' + $name + '''') })
            }
            continue
        }

        $conv = ConvertTo-EbiInputValue -Type $type -Value $given[$name]
        if (-not $conv['ok']) {
            [void]$problems.Add(@{ input = $name; kind = 'type_mismatch'; message = ('input ''' + $name + ''' ' + $conv['message']) })
            continue
        }
        $value = $conv['value']

        if ($spec.Contains('enum') -and $null -ne $spec['enum']) {
            $allowed = New-Object System.Collections.ArrayList
            foreach ($e in $spec['enum']) { [void]$allowed.Add([string]$e) }
            $hit = $false
            foreach ($a in $allowed) { if ([string]$value -ceq $a) { $hit = $true; break } }
            if (-not $hit) {
                [void]$problems.Add(@{ input = $name; kind = 'enum_mismatch';
                    message = ('input ''' + $name + ''' must be one of ' + ($allowed.ToArray() -join ', ') + '; got ' + (Get-EbiValueTypeName $value)) })
                continue
            }
        }
        $out[$name] = $value
    }

    $unknown = New-Object System.Collections.ArrayList
    foreach ($k in $given.Keys) { if ($declared -notcontains [string]$k) { [void]$unknown.Add([string]$k) } }
    foreach ($k in @($unknown.ToArray() | Sort-Object)) {
        $declaredText = if ($declaredSorted.Count -gt 0) { $declaredSorted -join ', ' } else { 'none' }
        [void]$problems.Add(@{ input = $k; kind = 'unknown_input'; message = ('unknown input ''' + $k + ''' (declared: ' + $declaredText + ')') })
    }

    if ($problems.Count -eq 0) {
        return @{ ok = $true; failure = ''; message = ''; In = $out; problems = @() }
    }
    $messages = New-Object System.Collections.ArrayList
    foreach ($p in $problems) { [void]$messages.Add([string]$p['message']) }
    return @{ ok = $false; failure = $(if ($manifestBad) { 'contract_violation' } else { 'input_invalid' });
              message = ($messages.ToArray() -join '; '); In = $out; problems = $problems.ToArray() }
}

# --- manifest readers (shared with the runner) -------------------------------

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

function Test-EbiJsonSerializable {
    param($Value)
    try { [void]($Value | ConvertTo-Json -Compress -Depth 10 -ErrorAction Stop); return $true }
    catch { return $false }
}

# --- "with" -> $In --------------------------------------------------------------

function Resolve-EbiStepInputs {
    <#
      PURE (given Session). Turn a call's "with" into the $In a step
      receives (STEP-CONTRACT.md 2.2 and 3.4 points 3 and 7), in this order:
        1. "as" is a runner field: taken out, never reaches $In; only a
           provides step may carry it (contract_violation otherwise)
        2. the rest is checked against the manifest's inputs schema
           (Test-EbiStepInputs): input_invalid / contract_violation, every
           problem naming its parameter
        3. the "as" name must not be live already (session_name_taken)
        4. each type='session' input is looked up in $Session and REPLACED
           by the registered value (session_missing / session_kind_mismatch)
      Returns @{ ok; In; As; failure; message; problems }.
    #>
    param($Manifest, [hashtable]$With, [hashtable]$Session)

    $raw = @{}
    $as = ''
    if ($null -ne $With) {
        foreach ($k in $With.Keys) {
            if ([string]$k -eq 'as') { $as = [string]$With[$k]; continue }
            $raw[[string]$k] = $With[$k]
        }
    }
    $stepId = if ($null -ne $Manifest -and $Manifest.Contains('id')) { [string]$Manifest['id'] } else { '?' }

    $provides = @(Get-EbiManifestArray -Manifest $Manifest -Key 'provides')
    if ($as -ne '' -and $provides.Count -eq 0) {
        return @{ ok = $false; In = $raw; As = $as; failure = 'contract_violation'; problems = @();
                  message = ('"as" given but step "{0}" provides nothing' -f $stepId) }
    }

    $checked = Test-EbiStepInputs -Manifest $Manifest -In $raw
    if (-not $checked['ok']) {
        return @{ ok = $false; In = $checked['In']; As = $as; failure = $checked['failure'];
                  message = $checked['message']; problems = $checked['problems'] }
    }
    $in = $checked['In']

    if ($as -ne '' -and $Session.Contains($as)) {
        return @{ ok = $false; In = $in; As = $as; failure = 'session_name_taken'; problems = @();
                  message = ('session name "{0}" is already registered and not released' -f $as) }
    }

    $sessionInputs = Get-EbiSessionInputs -Manifest $Manifest
    foreach ($name in $sessionInputs.Keys) {
        if (-not $in.Contains($name)) { continue }   # an optional session input that was not given
        $wanted = [string]$in[$name]
        if (-not $Session.Contains($wanted)) {
            return @{ ok = $false; In = $in; As = $as; failure = 'session_missing'; problems = @();
                      message = ('input "{0}" names session resource "{1}", which is not registered' -f $name, $wanted) }
        }
        $entry = $Session[$wanted]
        $kind  = [string]$sessionInputs[$name]
        if ($kind -ne '' -and [string]$entry['kind'] -ne $kind) {
            return @{ ok = $false; In = $in; As = $as; failure = 'session_kind_mismatch'; problems = @();
                      message = ('input "{0}" wants kind "{1}" but "{2}" is a "{3}"' -f $name, $kind, $wanted, [string]$entry['kind']) }
        }
        $in[$name] = $entry['value']
    }
    return @{ ok = $true; In = $in; As = $as; failure = ''; message = ''; problems = @() }
}

# --- return value -----------------------------------------------------------------

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

# --- where steps are ----------------------------------------------------------------

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

function Find-EbiStepFiles {
    <#
      Every step file under a modules root: <root>/<group>/<group>.<verb>.ps1,
      the group directory name matching the file's group part. Anything else
      under modules/ (the pre-conversion libraries in modules/verify, a
      README) is not a step and is not listed. Sorted by use. Each entry is
      @{ use; group; path }.
    #>
    param([string]$ModulesRoot)
    $found = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($ModulesRoot) -or -not (Test-Path -LiteralPath $ModulesRoot)) { return $found.ToArray() }
    $rx = [regex]'^([a-z][a-z0-9]*)\.([a-z][a-z0-9_]*)\.ps1$'
    foreach ($dir in @(Get-ChildItem -LiteralPath $ModulesRoot -Directory -ErrorAction SilentlyContinue | Sort-Object Name)) {
        foreach ($f in @(Get-ChildItem -LiteralPath $dir.FullName -Filter '*.ps1' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
            $m = $rx.Match($f.Name)
            if (-not $m.Success) { continue }
            if ($m.Groups[1].Value -ne $dir.Name) { continue }
            [void]$found.Add(@{ use = ($m.Groups[1].Value + '.' + $m.Groups[2].Value); group = $m.Groups[1].Value; path = $f.FullName })
        }
    }
    return $found.ToArray()
}

function Test-EbiLoadedStep {
    # What a loaded file must have shown: a [hashtable] $Manifest whose id is
    # the use, and an Invoke-Step defined by THAT file. '' when fine, else
    # the one-line reason (shared by the manifest-only scan and Import).
    param($Manifest, [string]$Use, [string]$InvokeFile, [string]$Path)
    if ($null -eq $Manifest -or -not ($Manifest -is [hashtable])) { return 'step defines no [hashtable] $Manifest' }
    if (-not $Manifest.Contains('id') -or [string]$Manifest['id'] -ne $Use) {
        return ('manifest id "{0}" does not match "{1}"' -f $(if ($Manifest.Contains('id')) { [string]$Manifest['id'] } else { '' }), $Use)
    }
    if ([string]::IsNullOrEmpty($InvokeFile)) { return 'step defines no Invoke-Step' }
    if ([System.IO.Path]::GetFullPath($InvokeFile) -ne [System.IO.Path]::GetFullPath($Path)) {
        return ('Invoke-Step in scope comes from {0}, not from this step' -f $InvokeFile)
    }
    return ''
}

function Read-EbiStepManifest {
    <#
      Manifest-only load, for ebi lint / ebi help / kernel/Docs.ps1: the file
      is dot-sourced in a throwaway child scope, so nothing it defines
      survives and no Invoke-Step is captured. Returns
      @{ ok; message; Manifest; use; path }.
    #>
    param([string]$Path, [string]$Use = '')
    if ([string]::IsNullOrWhiteSpace($Use)) {
        $leaf = Split-Path -Leaf $Path
        $Use = if ($leaf.EndsWith('.ps1')) { $leaf.Substring(0, $leaf.Length - 4) } else { $leaf }
    }
    $result = @{ ok = $false; message = ''; Manifest = $null; use = $Use; path = $Path }
    if (-not (Test-Path -LiteralPath $Path)) { $result['message'] = ('no step file at {0}' -f $Path); return $result }
    $loaded = $null
    try {
        $loaded = & {
            param($p)
            $Manifest = $null
            . $p
            $fn = Get-Item -LiteralPath 'function:Invoke-Step' -ErrorAction SilentlyContinue
            $file = if ($null -ne $fn -and $null -ne $fn.ScriptBlock) { [string]$fn.ScriptBlock.File } else { '' }
            return @{ Manifest = $Manifest; InvokeFile = $file }
        } $Path
    } catch {
        $result['message'] = ('{0}: {1}' -f $Path, $_.Exception.Message)
        return $result
    }
    $why = Test-EbiLoadedStep -Manifest $loaded['Manifest'] -Use $Use -InvokeFile $loaded['InvokeFile'] -Path $Path
    if ($why -ne '') { $result['message'] = ('{0}: {1}' -f $Path, $why); return $result }
    $result['ok'] = $true
    $result['Manifest'] = $loaded['Manifest']
    return $result
}

function Get-EbiStepCatalog {
    <#
      Scan a modules root and read every step's manifest (Read-EbiStepManifest,
      child scope, nothing captured). The list ebi lint / ebi help /
      kernel/Docs.ps1 start from: one entry per step file, sorted by use,
      @{ use; group; path; ok; message; Manifest } -- a broken step is an
      entry with ok=$false, reported next to the others, never an exception.
    #>
    param([string]$ModulesRoot = '')
    if ([string]::IsNullOrWhiteSpace($ModulesRoot)) { $ModulesRoot = Get-EbiDefaultModulesRoot }
    $out = New-Object System.Collections.ArrayList
    foreach ($f in @(Find-EbiStepFiles -ModulesRoot $ModulesRoot)) {
        $r = Read-EbiStepManifest -Path $f['path'] -Use $f['use']
        [void]$out.Add(@{ use = $f['use']; group = $f['group']; path = $f['path']; ok = $r['ok']; message = $r['message']; Manifest = $r['Manifest'] })
    }
    return $out.ToArray()
}

# --- the registry: loading steps for running --------------------------------------------

function New-EbiRegistry {
    # use -> @{ Manifest; Invoke; Path }, plus where to look.
    param([string]$ModulesRoot = '')
    if ([string]::IsNullOrWhiteSpace($ModulesRoot)) { $ModulesRoot = Get-EbiDefaultModulesRoot }
    return @{ ModulesRoot = $ModulesRoot; Steps = @{} }
}

function Get-EbiStep {
    # The loaded entry for a use, or $null when it has not been imported.
    param([hashtable]$Registry, [string]$Use)
    if ($null -eq $Registry -or [string]::IsNullOrWhiteSpace($Use)) { return $null }
    if ($Registry['Steps'].Contains($Use)) { return $Registry['Steps'][$Use] }
    return $null
}

function Import-EbiStep {
    <#
      Load one step for running and put it in the registry table.

      MUST be dot-sourced by the scope that will run the step:
          $r = . Import-EbiStep -Registry $registry -Use 'browser.ensure'
      (see the file header for why). Any other invocation is refused.

      Sequence: locate the file, clear any Invoke-Step already in scope,
      dot-source the file, capture Invoke-Step at once (the next file
      overwrites it), check $Manifest / id / that the Invoke-Step came from
      this file, store @{ Manifest; Invoke; Path } under the use, then
      remove the bare Invoke-Step so the table is the only way to call it.

      Returns @{ ok; failure; message; Entry }: failure is 'step_not_found'
      for a missing or unloadable file (the reason names the path), and
      'internal_error' when this function itself was called without the dot.
      A use already in the table is returned as-is, not loaded twice.
    #>
    param([hashtable]$Registry, [string]$Use)

    if ($MyInvocation.InvocationName -ne '.') {
        return @{ ok = $false; failure = 'internal_error'; Entry = $null;
                  message = 'Import-EbiStep must be dot-sourced (. Import-EbiStep -Registry ... -Use ...) so the step''s helper functions land in the caller''s scope; called as "' + $MyInvocation.InvocationName + '"' }
    }
    if ($null -ne (Get-EbiStep -Registry $Registry -Use $Use)) {
        return @{ ok = $true; failure = ''; message = ''; Entry = $Registry['Steps'][$Use] }
    }

    # Everything below lives in the caller's scope (dot-sourced), so state is
    # kept in one prefixed hashtable plus $Manifest, and both are removed
    # before returning.
    $ebiImportState = @{ path = (Get-EbiStepPath -ModulesRoot $Registry['ModulesRoot'] -Use $Use); err = ''; captured = $null; file = '' }
    $Manifest = $null

    if ($ebiImportState['path'] -eq '' -or -not (Test-Path -LiteralPath $ebiImportState['path'])) {
        $ebiImportState['err'] = ('no step file for "{0}" (looked at {1})' -f $Use, $(if ($ebiImportState['path'] -eq '') { '<not a <group>.<verb> id>' } else { $ebiImportState['path'] }))
    } else {
        try {
            # A step that forgets Invoke-Step must not inherit the previous
            # step's, so the name is cleared before every load.
            if (Test-Path -LiteralPath 'function:Invoke-Step') { Remove-Item -LiteralPath 'function:Invoke-Step' -ErrorAction SilentlyContinue }
            . ($ebiImportState['path'])
            $ebiImportState['fn'] = Get-Item -LiteralPath 'function:Invoke-Step' -ErrorAction SilentlyContinue
            if ($null -ne $ebiImportState['fn']) {
                $ebiImportState['captured'] = $ebiImportState['fn'].ScriptBlock
                $ebiImportState['file'] = [string]$ebiImportState['fn'].ScriptBlock.File
            }
            $ebiImportState['err'] = Test-EbiLoadedStep -Manifest $Manifest -Use $Use -InvokeFile $ebiImportState['file'] -Path $ebiImportState['path']
            if ($ebiImportState['err'] -ne '') { $ebiImportState['err'] = ('{0}: {1}' -f $ebiImportState['path'], $ebiImportState['err']) }
        } catch {
            $ebiImportState['err'] = ('{0}: {1}' -f $ebiImportState['path'], $_.Exception.Message)
        }
        # The bare name is never called: steps run from the table only.
        if (Test-Path -LiteralPath 'function:Invoke-Step') { Remove-Item -LiteralPath 'function:Invoke-Step' -ErrorAction SilentlyContinue }
    }

    if ($ebiImportState['err'] -eq '') {
        $Registry['Steps'][$Use] = @{ Manifest = $Manifest; Invoke = $ebiImportState['captured']; Path = $ebiImportState['path'] }
        $ebiImportState['result'] = @{ ok = $true; failure = ''; message = ''; Entry = $Registry['Steps'][$Use] }
    } else {
        $ebiImportState['result'] = @{ ok = $false; failure = 'step_not_found'; message = $ebiImportState['err']; Entry = $null }
    }
    $ebiImportResult = $ebiImportState['result']
    Remove-Variable -Name ebiImportState -ErrorAction SilentlyContinue
    Remove-Variable -Name Manifest -ErrorAction SilentlyContinue
    return $ebiImportResult
}
