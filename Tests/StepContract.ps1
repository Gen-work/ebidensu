#Requires -Version 5.1
# StepContract.ps1 -- the ebi-dance step contract checker.
#
# Dot-sourced by Tests\Test-StepContract.ps1. No param() block, ASCII source
# only (CLAUDE.md encoding rule), no class, no COM.
#
# The rule list this implements is spec\STEP-CONTRACT.md section 7. That list
# is the single place new manifest-side contract rules are collected: rules
# stated only in section 2 or section 3.4 are NOT picked up here automatically,
# which is exactly why section 7 carries the meta-rule "add the rule to this
# list in the same change".
#
# Split of concerns:
#   Get-StepContractFindings   PURE. Everything it judges is passed in, so
#                              every rule class is unit-testable from inline
#                              fixtures with no file and no dot-source.
#   Read-StepFile              the small impure half: parse + dot-source one
#                              file and hand the pure half what it needs.
#
# Every finding is @{ rule = '<id>'; message = '<one line>' }. The rule ids are
# stable -- Test-StepContract.ps1 asserts on them.

function Get-StepContractMustReleaseKinds {
    <#
      Parse the resource-kind table out of STEP-CONTRACT.md section 3.4 point 6.

      The table is deliberately hand-maintained prose, not derived from the
      manifests: 'provides'/'releases' say which step produces or frees a kind,
      they cannot say whether leaking that kind matters. Reading it from the
      spec keeps one copy -- a kind added to the table is picked up here with
      no second edit.

      Returns a hashtable kind -> [bool] mustRelease. An empty result means the
      table could not be found; callers must treat that as an error rather than
      as "no kinds are declared", or a reformat of the spec would silently turn
      the kind check off.
    #>
    param([string]$Text)

    $kinds = @{}
    if ([string]::IsNullOrEmpty($Text)) { return $kinds }

    $lines  = $Text -split "`r?`n"
    $rowRx  = [regex]'^\s*\|\s*`([A-Za-z][A-Za-z0-9_]*)`\s*\|\s*`\$(true|false)`'
    $inTable = $false

    foreach ($line in $lines) {
        if (-not $inTable) {
            # Anchor on the header cell, which is ASCII even though the rest of
            # the row is not.
            if ($line -match '`mustRelease`' -and $line.TrimStart().StartsWith('|')) {
                $inTable = $true
            }
            continue
        }
        $m = $rowRx.Match($line)
        if ($m.Success) {
            $kinds[$m.Groups[1].Value] = ($m.Groups[2].Value -eq 'true')
            continue
        }
        # The table ends at the first line that is not a table row. Separator
        # rows (|---|---|) are skipped rather than ending it.
        if ($line.TrimStart().StartsWith('|')) { continue }
        if ($kinds.Count -gt 0) { break }
    }
    return $kinds
}

function Get-StepIdFromFileName {
    # 'browser.find.ps1' -> 'browser.find'
    param([string]$FileName)
    if ([string]::IsNullOrEmpty($FileName)) { return '' }
    if ($FileName.EndsWith('.ps1')) { return $FileName.Substring(0, $FileName.Length - 4) }
    return $FileName
}

function Get-StepHelperPrefix {
    <#
      'browser.find'          -> 'BrowserFind-'
      'screen.capture_window' -> 'ScreenCaptureWindow-'

      Every step file is dot-sourced into the same runspace, one after another.
      Invoke-Step collides by design and the runner captures it per step
      (P1-02); a bare helper name like Get-Row would just overwrite the one a
      previously loaded step defined, silently and with no error anywhere.

      Both '.' and '_' are segment separators: a step id's verb may be
      snake_case, and 'ScreenCapture_window-' would be an odd thing to ask
      anyone to type.
    #>
    param([string]$StepId)
    if ([string]::IsNullOrEmpty($StepId)) { return '' }
    $parts = @($StepId -split '[._]' | Where-Object { $_ -ne '' })
    $sb = New-Object System.Text.StringBuilder
    foreach ($p in $parts) {
        [void]$sb.Append($p.Substring(0, 1).ToUpperInvariant())
        if ($p.Length -gt 1) { [void]$sb.Append($p.Substring(1)) }
    }
    return ($sb.ToString() + '-')
}

function Get-StepNonAsciiLines {
    # Returns the 1-based line numbers holding a character outside 0x00-0x7F.
    param([string]$Text)
    $hits = @()
    if ([string]::IsNullOrEmpty($Text)) { return $hits }
    $lines = $Text -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        foreach ($ch in $lines[$i].ToCharArray()) {
            if ([int]$ch -gt 127) { $hits += ($i + 1); break }
        }
    }
    return $hits
}

function Get-StepContractSerializableTypes {
    # Output types that survive ConvertTo-Json without losing anything. This is
    # the section 2.2 type table minus 'session': handles and COM objects go
    # through $Ctx.Session and never appear in outputs (P0-R2), which is also
    # what makes ledger replay possible at all (P0-R3).
    return @('string', 'int', 'bool', 'path', 'rect', 'list', 'map', 'any')
}

function ConvertTo-StepContractArray {
    <#
      Normalize a manifest field that should be an array.

      Written as an explicit loop rather than @($Value): the @() wrap over an
      indexed collection is the shape this repo has been bitten by twice on
      PS 5.1 ("argument types do not match" out of the binder), and it is a
      banned shape here.

      Returned unwrapped, so callers wrap the call in @() themselves. Do NOT
      add the usual ", $array" comma guard: it makes an empty array come back
      as a one-element array holding an empty array, which reads as a resource
      kind named '' and as a step with no Invoke-Step.
    #>
    param($Value)
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $Value) { return $out.ToArray() }
    if ($Value -is [string]) { [void]$out.Add($Value); return $out.ToArray() }
    if ($Value -is [System.Collections.IEnumerable]) {
        foreach ($item in $Value) { [void]$out.Add($item) }
        return $out.ToArray()
    }
    [void]$out.Add($Value)
    return $out.ToArray()
}

function Test-StepDictHasKey {
    <#
      Does this dictionary have that key?

      Goes through IDictionary.Contains rather than .ContainsKey, and that is
      the whole point: [ordered]@{} is an OrderedDictionary, which satisfies
      -is [IDictionary] and has Contains but NOT ContainsKey. Calling
      .ContainsKey on one throws MethodNotFound, and under the runner's
      $ErrorActionPreference = 'Stop' that ends the entire run. Reading every
      key through here makes "a malformed step cannot terminate the checker" a
      property of the code rather than something each call site has to
      remember.

      Note this is separate from the contract rule below, which still requires
      a plain [hashtable]: this function is about not crashing, that rule is
      about there being exactly one shape downstream tooling has to handle.
    #>
    param($Dict, [string]$Key)
    if ($null -eq $Dict) { return $false }
    if (-not ($Dict -is [System.Collections.IDictionary])) { return $false }
    return ([System.Collections.IDictionary]$Dict).Contains($Key)
}

function Get-StepContractFindings {
    <#
      Judge one step against the section 7 rule list. PURE -- no file access,
      no dot-source: the caller supplies the manifest, the source text and the
      facts read off the AST.

      Parameters:
        StepId         expected id, i.e. the file name without .ps1
        Manifest       the $Manifest hashtable, or $null when the file had none
        Text           the file's source, for the ASCII rule
        FunctionNames  every function the file defines
        HasParamBlock  whether the file has a file-level param() block
        LoadError      non-empty when the file could not be dot-sourced
        MustRelease    kind -> bool, from Get-StepContractMustReleaseKinds

      Returns an array of @{ rule; message }, empty when the step is clean.
    #>
    param(
        [string]$StepId,
        $Manifest,
        [string]$Text,
        $FunctionNames,
        [bool]$HasParamBlock,
        [string]$LoadError,
        $MustRelease
    )

    $findings = New-Object System.Collections.ArrayList
    function Add-Finding {
        param([string]$Rule, [string]$Message)
        [void]$findings.Add(@{ rule = $Rule; message = $Message })
    }

    # -- source-level rules (hold even when the manifest is unreadable) -------

    if (-not [string]::IsNullOrEmpty($LoadError)) {
        Add-Finding 'load' ('cannot be dot-sourced: ' + $LoadError)
    }

    if ($HasParamBlock) {
        Add-Finding 'param_block' 'step files are dot-sourced; a file-level param() block overwrites the caller''s variables'
    }

    $nonAscii = @(Get-StepNonAsciiLines -Text $Text)
    if ($nonAscii.Count -gt 0) {
        Add-Finding 'non_ascii' ('non-ASCII source on line(s) ' + ($nonAscii -join ', ') + '; build Japanese from [char]')
    }

    # $null FunctionNames means "not known" -- the file could not be parsed, so
    # there is no AST to read them off. An EMPTY array means "parsed fine, and
    # it really defines nothing", which is a missing Invoke-Step. Collapsing the
    # two is how a step with a manifest and no functions at all used to pass.
    if ($null -ne $FunctionNames) {
        $names = @(ConvertTo-StepContractArray -Value $FunctionNames)
        if (-not ($names -contains 'Invoke-Step')) {
            Add-Finding 'invoke_step_missing' 'no Invoke-Step function; a step file exports exactly $Manifest and Invoke-Step'
        }
        $prefix = Get-StepHelperPrefix -StepId $StepId
        foreach ($fn in $names) {
            $fnName = [string]$fn
            if ($fnName -eq 'Invoke-Step') { continue }
            if (-not $fnName.StartsWith($prefix)) {
                Add-Finding 'helper_prefix' ("helper '" + $fnName + "' must start with '" + $prefix + "'; every step is dot-sourced into one runspace and bare names overwrite each other")
            }
        }
    }

    if ($null -eq $Manifest) {
        Add-Finding 'manifest_missing' 'no $Manifest hashtable'
        return $findings.ToArray()
    }

    # A step file can assign $Manifest anything at all and still parse and
    # dot-source cleanly, so a wrong shape has to be a finding rather than an
    # exception: the checker's job is to report a malformed step next to the
    # others, not to take the whole test run down with it.
    #
    # The bar is a plain [hashtable], not merely IDictionary. [ordered]@{} would
    # satisfy IDictionary and work fine here, but the point of one documented
    # shape is that every consumer downstream -- runner, catalog generation
    # (P1-06), lint -- gets to assume it without each having to survive every
    # dictionary implementation .NET offers. Section 2 writes @{}; this rule is
    # what makes that a guarantee rather than an example.
    if (-not ($Manifest -is [hashtable])) {
        Add-Finding 'manifest_shape' ('$Manifest is ' + $Manifest.GetType().Name + ', not a hashtable; write it as @{} (an [ordered]@{} is not a hashtable)')
        return $findings.ToArray()
    }

    # -- identity ------------------------------------------------------------

    $declaredId = if ((Test-StepDictHasKey -Dict $Manifest -Key 'id')) { [string]$Manifest['id'] } else { '' }
    if ($declaredId -ne $StepId) {
        Add-Finding 'id_mismatch' ("manifest id '" + $declaredId + "' does not match the file name '" + $StepId + "'")
    }

    # -- inputs --------------------------------------------------------------

    $inputs = if ((Test-StepDictHasKey -Dict $Manifest -Key 'inputs')) { $Manifest['inputs'] } else { $null }
    $inputNames    = New-Object System.Collections.ArrayList
    $sessionKinds  = New-Object System.Collections.ArrayList
    if ($null -ne $inputs -and -not ($inputs -is [hashtable])) {
        # Present but not a map at all: inputs = 'bad', inputs = @(...). Every
        # per-input rule below is unreachable, so saying nothing would let the
        # step through clean -- the same silence the entry-level check fixed.
        Add-Finding 'field_container_shape' ('inputs is ' + $inputs.GetType().Name + ', not a hashtable; no input can be checked')
    }
    if ($inputs -is [hashtable]) {
        foreach ($key in $inputs.Keys) {
            [void]$inputNames.Add([string]$key)
            $spec = $inputs[$key]
            if (-not ($spec -is [hashtable])) {
                Add-Finding 'field_spec_shape' ("input '" + $key + "' is not a hashtable, so none of its rules can be checked")
                continue
            }

            $hasRequired = (Test-StepDictHasKey -Dict $spec -Key 'required') -and [bool]$spec['required']
            $hasDefault  = (Test-StepDictHasKey -Dict $spec -Key 'default')
            if ($hasRequired -and $hasDefault) {
                Add-Finding 'required_and_default' ("input '" + $key + "' declares both required and default")
            }

            $type = if ((Test-StepDictHasKey -Dict $spec -Key 'type')) { [string]$spec['type'] } else { '' }
            if ($type -eq 'session') {
                $kind = if ((Test-StepDictHasKey -Dict $spec -Key 'sessionKind')) { [string]$spec['sessionKind'] } else { '' }
                if ([string]::IsNullOrWhiteSpace($kind)) {
                    Add-Finding 'session_input_no_kind' ("input '" + $key + "' is type='session' but declares no sessionKind")
                } else {
                    [void]$sessionKinds.Add($kind)
                }
            }
        }
    }

    # -- outputs -------------------------------------------------------------

    $serializable = Get-StepContractSerializableTypes
    $outputs = if ((Test-StepDictHasKey -Dict $Manifest -Key 'outputs')) { $Manifest['outputs'] } else { $null }
    if ($null -ne $outputs -and -not ($outputs -is [hashtable])) {
        # A bad outputs container is the more dangerous of the two: with no
        # entries to walk, the JSON-serializable rule never fires and the step
        # otherwise reads as clean.
        Add-Finding 'field_container_shape' ('outputs is ' + $outputs.GetType().Name + ', not a hashtable; no output can be checked')
    }
    if ($outputs -is [hashtable]) {
        foreach ($key in $outputs.Keys) {
            $spec = $outputs[$key]
            if (-not ($spec -is [hashtable])) {
                Add-Finding 'field_spec_shape' ("output '" + $key + "' is not a hashtable, so none of its rules can be checked")
                continue
            }
            $type = if ((Test-StepDictHasKey -Dict $spec -Key 'type')) { [string]$spec['type'] } else { '' }
            if (-not ($serializable -contains $type)) {
                Add-Finding 'output_not_serializable' ("output '" + $key + "' has type '" + $type + "', which is not JSON-serializable; handles and COM objects go through " + '$Ctx.Session')
            }
        }
    }

    # -- failures ------------------------------------------------------------

    $failures = @(ConvertTo-StepContractArray -Value $(if ((Test-StepDictHasKey -Dict $Manifest -Key 'failures')) { $Manifest['failures'] } else { $null }))
    if ($failures.Count -eq 0) {
        Add-Finding 'failures_empty' 'failures must enumerate every failure id this step can return'
    }
    foreach ($f in $failures) {
        if (-not ($f -is [hashtable])) {
            Add-Finding 'failure_shape' 'each failures entry must be a hashtable with id and transient'
            continue
        }
        $fid = if ((Test-StepDictHasKey -Dict $f -Key 'id')) { [string]$f['id'] } else { '' }
        if ([string]::IsNullOrWhiteSpace($fid)) {
            Add-Finding 'failure_shape' 'a failures entry has no id'
        }
        if (-not (Test-StepDictHasKey -Dict $f -Key 'transient')) {
            Add-Finding 'failure_shape' ("failure '" + $fid + "' has no transient flag; retry policy needs it")
        } elseif (-not ($f['transient'] -is [bool])) {
            Add-Finding 'failure_shape' ("failure '" + $fid + "' has a non-boolean transient")
        }
    }

    # -- example -------------------------------------------------------------

    $example = if ((Test-StepDictHasKey -Dict $Manifest -Key 'example')) { $Manifest['example'] } else { $null }
    if ($example -is [hashtable] -and (Test-StepDictHasKey -Dict $example -Key 'with')) {
        $with = $example['with']
        if ($with -is [hashtable]) {
            foreach ($key in $with.Keys) {
                # 'as' is a runner-reserved field lifted out of 'with' before
                # the step ever sees it, so it is never declared in inputs.
                if ([string]$key -eq 'as') { continue }
                if (-not ($inputNames -contains [string]$key)) {
                    Add-Finding 'example_unknown_param' ("example passes '" + $key + "', which is not declared in inputs")
                }
            }
        }
    }

    # -- session resources ---------------------------------------------------

    $provides = @(ConvertTo-StepContractArray -Value $(if ((Test-StepDictHasKey -Dict $Manifest -Key 'provides')) { $Manifest['provides'] } else { $null }))
    $releases = @(ConvertTo-StepContractArray -Value $(if ((Test-StepDictHasKey -Dict $Manifest -Key 'releases')) { $Manifest['releases'] } else { $null }))

    if ($provides.Count -gt 1) {
        Add-Finding 'provides_multiple' ('provides has ' + $provides.Count + ' kinds; one call registers at most one resource, so split the step')
    }

    foreach ($kind in $releases) {
        $k = [string]$kind
        if (-not ($sessionKinds -contains $k)) {
            Add-Finding 'releases_unmatched' ("releases '" + $k + "' but no type='session' input declares that sessionKind, so there is no way to say which instance to free")
        }
    }

    $declaredKinds = New-Object System.Collections.ArrayList
    foreach ($kind in $provides) { [void]$declaredKinds.Add([string]$kind) }
    foreach ($kind in $releases) { [void]$declaredKinds.Add([string]$kind) }
    foreach ($k in $declaredKinds) {
        if (-not (Test-StepDictHasKey -Dict $MustRelease -Key $k)) {
            Add-Finding 'kind_undeclared' ("resource kind '" + $k + "' is not in the mustRelease table in STEP-CONTRACT.md section 3.4; declare whether leaking it matters")
        }
    }

    if ($declaredKinds.Count -gt 0) {
        $idempotent = (Test-StepDictHasKey -Dict $Manifest -Key 'idempotent') -and ($Manifest['idempotent'] -is [bool]) -and $Manifest['idempotent']
        if (-not $idempotent) {
            Add-Finding 'resource_not_idempotent' 'a step that provides or releases a Session resource must be idempotent; resume always really runs it'
        }
    }

    return $findings.ToArray()
}

function Read-StepFile {
    <#
      The impure half: parse one step file and dot-source it in a child scope
      so its $Manifest and its functions do not leak into the caller (every
      step defines Invoke-Step, so without isolation the last file loaded would
      win). Returns everything Get-StepContractFindings needs.
    #>
    param([string]$Path)

    $result = @{
        StepId        = Get-StepIdFromFileName -FileName (Split-Path -Leaf $Path)
        Manifest      = $null
        Text          = ''
        FunctionNames = $null   # $null = not known (parse failed); @() = really none
        HasParamBlock = $false
        LoadError     = ''
    }

    try {
        $result.Text = [System.IO.File]::ReadAllText($Path)
    } catch {
        $result.LoadError = $_.Exception.Message
        return $result
    }

    $tokens = $null; $errs = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errs)
    if ($errs -and $errs.Count -gt 0) {
        $result.LoadError = ('parse error on line {0}: {1}' -f $errs[0].Extent.StartLineNumber, $errs[0].Message)
        return $result
    }

    if ($null -ne $ast.ParamBlock) { $result.HasParamBlock = $true }

    $fnAsts = $ast.FindAll(
        { param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
    $names = New-Object System.Collections.ArrayList
    foreach ($fn in $fnAsts) { [void]$names.Add($fn.Name) }
    $result.FunctionNames = $names.ToArray()

    try {
        $result.Manifest = & {
            . $Path
            if (Get-Variable -Name Manifest -Scope Local -ErrorAction SilentlyContinue) { $Manifest } else { $null }
        }
    } catch {
        $result.LoadError = $_.Exception.Message
    }

    return $result
}

function Get-StepFiles {
    # Every .ps1 under modules/. Classification into steps and not-yet-steps is
    # Test-IsStepFile's job, because it needs the file's contents. Nothing under
    # legacy/ is ever considered: it is deliberately outside the catalog.
    param([string]$ModulesRoot)
    if (-not (Test-Path -LiteralPath $ModulesRoot)) { return @() }
    $found = @(Get-ChildItem -LiteralPath $ModulesRoot -Filter '*.ps1' -File -Recurse -ErrorAction SilentlyContinue |
        Sort-Object FullName)
    return $found
}

function Test-IsStepFile {
    <#
      Is this .ps1 under modules/ a step, and therefore bound by the contract?

      Two ways to be one, deliberately OR'd so there is no way to write a
      malformed step that the checker quietly skips:

      1. It is named for the convention in section 1, <group>.<verb>.ps1 with
         <group> equal to its own directory -- 'modules/screen/screen.crop.ps1'.
      2. It defines $Manifest or Invoke-Step. Defining either is opting in, so
         a step misfiled as 'helpers.ps1' is still checked; it just reports
         id_mismatch, which is the finding that tells the author to rename it.

      Everything else is a plain library. The refactor parks pre-conversion
      libraries under modules/ before they are rewritten as steps (P0-04 put
      five there), and those genuinely are not steps yet -- they have no
      manifest to check. They are reported by name so they stay visible rather
      than becoming permanent residents.
    #>
    param(
        [string]$FileName,
        [string]$GroupName,
        $Manifest,
        $FunctionNames
    )
    if ($null -ne $Manifest) { return $true }
    $names = @(ConvertTo-StepContractArray -Value $FunctionNames)
    if ($names -contains 'Invoke-Step') { return $true }

    $stem = Get-StepIdFromFileName -FileName $FileName
    $dot  = $stem.IndexOf('.')
    if ($dot -lt 1) { return $false }
    return ($stem.Substring(0, $dot) -eq $GroupName)
}

function Test-StepFileContract {
    # Convenience wrapper: read one file, judge it, return the findings.
    param([string]$Path, $MustRelease)
    $read = Read-StepFile -Path $Path
    return Get-StepContractFindings `
        -StepId        $read.StepId `
        -Manifest      $read.Manifest `
        -Text          $read.Text `
        -FunctionNames $read.FunctionNames `
        -HasParamBlock $read.HasParamBlock `
        -LoadError     $read.LoadError `
        -MustRelease   $MustRelease
}
