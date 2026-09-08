#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path $here 'StepContract.ps1')

Reset-Tests 'StepContract'

# Every rule class is asserted twice over: once against an inline manifest
# (fast, and it pins the exact rule id), and once against a real file written
# to a temp directory, because three of the rules -- load, param_block,
# non_ascii -- only exist at the file level and an inline fixture cannot reach
# them. The deliberately broken fixtures are written at run time and deleted
# afterwards rather than committed: a committed bad step would be picked up by
# the repo-wide parse check and by Check-Encoding.ps1, which is the same
# "a fixture that fights the real checks" trap the docs tests already avoid.

function Test-HasRule {
    param($Findings, [string]$Rule)
    foreach ($f in $Findings) { if ($f.rule -eq $Rule) { return $true } }
    return $false
}

function Test-RuleList {
    param($Findings)
    $names = New-Object System.Collections.ArrayList
    foreach ($f in $Findings) { [void]$names.Add($f.rule) }
    return (($names | Sort-Object) -join ',')
}

function New-CleanManifest {
    # The section 8 example, reduced to what the checker looks at.
    return @{
        id         = 'screen.crop'
        group      = 'screen'
        summary    = 'Crop a PNG by per-side pixel amounts and write the result'
        tier       = 'core'
        effects    = 'write'
        needs      = @()
        provides   = @()
        releases   = @()
        idempotent = $true
        inputs     = @{
            path = @{ type = 'path'; required = $true; desc = 'PNG to crop' }
            left = @{ type = 'int';  default  = 0 }
        }
        outputs    = @{
            path  = @{ type = 'path' }
            width = @{ type = 'int' }
        }
        failures   = @(
            @{ id = 'file_not_found';  transient = $false }
            @{ id = 'image_read_error'; transient = $true }
        )
        example    = @{ use = 'screen.crop'; with = @{ path = 'a.png'; left = 6 } }
    }
}

function Get-CleanFindings {
    param($Manifest, $MustRelease)
    return Get-StepContractFindings `
        -StepId        'screen.crop' `
        -Manifest      $Manifest `
        -Text          "# ascii only`nfunction Invoke-Step { param(`$In, `$Ctx) }`n" `
        -FunctionNames @('Invoke-Step') `
        -HasParamBlock $false `
        -LoadError     '' `
        -MustRelease   $MustRelease
}

# ---- the mustRelease kind table is read from the spec, not duplicated ------

$specPath = Join-Path $repoRoot 'docs/ebi-dance/spec/STEP-CONTRACT.md'
Assert-True (Test-Path -LiteralPath $specPath) 'STEP-CONTRACT.md is where the checker expects it'

$specText   = [System.IO.File]::ReadAllText($specPath)
$mustRelease = Get-StepContractMustReleaseKinds -Text $specText

Assert-True ($mustRelease.Count -ge 3) 'the mustRelease table is found in the real spec'
Assert-True ($mustRelease.ContainsKey('window'))   'spec table declares kind window'
Assert-True ($mustRelease.ContainsKey('workbook')) 'spec table declares kind workbook'
Assert-True ($mustRelease.ContainsKey('excelApp')) 'spec table declares kind excelApp'
Assert-Equal $false $mustRelease['window']   'window leaks harmlessly'
Assert-Equal $true  $mustRelease['workbook'] 'workbook must be released'
Assert-Equal $true  $mustRelease['excelApp'] 'excelApp must be released'

# A reformat of the spec must not silently disable the kind check: no table
# found has to be distinguishable from "the table declares nothing".
$noTable = Get-StepContractMustReleaseKinds -Text "# nothing here`nplain prose`n"
Assert-Equal 0 $noTable.Count 'no table found yields an empty map the caller can reject'

# ---- small pure helpers ---------------------------------------------------

Assert-Equal 'browser.find' (Get-StepIdFromFileName -FileName 'browser.find.ps1') 'step id is the file name without .ps1'
Assert-Equal 'browser.find' (Get-StepIdFromFileName -FileName 'browser.find')     'already-stripped name is left alone'

Assert-Equal 'BrowserFind-' (Get-StepHelperPrefix -StepId 'browser.find')    'helper prefix is the PascalCased id'
Assert-Equal 'ScreenCaptureWindow-' (Get-StepHelperPrefix -StepId 'screen.capture_window') 'underscores are segment separators too'
Assert-Equal 'Flow-' (Get-StepHelperPrefix -StepId 'flow') 'a single-segment id still yields a prefix'

$asciiHits = @(Get-StepNonAsciiLines -Text "ok`nalso ok`n")
Assert-Equal 0 $asciiHits.Count 'pure ASCII text reports no lines'
$jp = [string][char]0x65E5
$mixedHits = @(Get-StepNonAsciiLines -Text ("line one`nline " + $jp + " two`nline three`n"))
Assert-Equal 1 $mixedHits.Count 'one offending line is reported'
Assert-Equal 2 $mixedHits[0]    'the reported line number is 1-based'

# ---- a clean step produces no findings ------------------------------------

$clean = @(Get-CleanFindings -Manifest (New-CleanManifest) -MustRelease $mustRelease)
Assert-Equal 0 $clean.Count ('a clean step yields no findings (got: ' + (Test-RuleList -Findings $clean) + ')')

# ---- one rule class at a time ---------------------------------------------

$m = New-CleanManifest; $m['id'] = 'screen.crumb'
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'id_mismatch') 'manifest id must equal the file name'

$m = New-CleanManifest; $m['inputs']['left'] = @{ type = 'int'; required = $true; default = 0 }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'required_and_default') 'required and default are mutually exclusive'

$m = New-CleanManifest; $m['failures'] = @()
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'failures_empty') 'failures must not be empty'

$m = New-CleanManifest; $m['failures'] = @(@{ id = 'boom' })
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'failure_shape') 'a failure with no transient flag is reported'

$m = New-CleanManifest; $m['failures'] = @(@{ id = 'boom'; transient = 'yes' })
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'failure_shape') 'a non-boolean transient is reported'

$m = New-CleanManifest; $m['example'] = @{ use = 'screen.crop'; with = @{ path = 'a.png'; nope = 1 } }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'example_unknown_param') 'the example may only pass declared inputs'

$m = New-CleanManifest; $m['example'] = @{ use = 'screen.crop'; with = @{ path = 'a.png'; as = 'wb' } }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (-not (Test-HasRule -Findings $f -Rule 'example_unknown_param')) 'as is runner-reserved and never declared in inputs'

$m = New-CleanManifest; $m['outputs']['handle'] = @{ type = 'session' }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'output_not_serializable') 'a session output is not JSON-serializable'

$m = New-CleanManifest; $m['outputs']['app'] = @{ type = 'comobject' }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'output_not_serializable') 'an unknown output type is rejected'

$m = New-CleanManifest; $m['inputs']['window'] = @{ type = 'session'; required = $true }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'session_input_no_kind') 'a session input must declare its sessionKind'

$m = New-CleanManifest
$m['provides']   = @('workbook', 'excelApp')
$m['idempotent'] = $true
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'provides_multiple') 'one call registers at most one resource'

$m = New-CleanManifest
$m['releases']   = @('workbook')
$m['idempotent'] = $true
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'releases_unmatched') 'releasing a kind needs a session input of that kind'

$m = New-CleanManifest
$m['releases']       = @('workbook')
$m['idempotent']     = $true
$m['inputs']['book'] = @{ type = 'session'; sessionKind = 'workbook'; required = $true }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (-not (Test-HasRule -Findings $f -Rule 'releases_unmatched')) 'a matching session input satisfies releases'

$m = New-CleanManifest
$m['provides']   = @('teapot')
$m['idempotent'] = $true
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'kind_undeclared') 'a kind absent from the spec table is reported'

$m = New-CleanManifest
$m['provides']   = @('window')
$m['idempotent'] = $false
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'resource_not_idempotent') 'a resource step must be idempotent'

$m = New-CleanManifest
$m['provides']   = @('window')
$m['idempotent'] = $true
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-Equal 0 $f.Count ('a well-formed resource step is clean (got: ' + (Test-RuleList -Findings $f) + ')')

$f = @(Get-StepContractFindings -StepId 'screen.crop' -Manifest $null -Text '' `
        -FunctionNames @('Invoke-Step') -HasParamBlock $false -LoadError '' -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'manifest_missing') 'a file with no $Manifest is reported'

$f = @(Get-StepContractFindings -StepId 'screen.crop' -Manifest (New-CleanManifest) -Text '' `
        -FunctionNames @('Invoke-Step', 'Get-Row') -HasParamBlock $false -LoadError '' -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'helper_prefix') 'a bare helper name is reported'

$f = @(Get-StepContractFindings -StepId 'screen.crop' -Manifest (New-CleanManifest) -Text '' `
        -FunctionNames @('Invoke-Step', 'ScreenCrop-Row') -HasParamBlock $false -LoadError '' -MustRelease $mustRelease)
Assert-True (-not (Test-HasRule -Findings $f -Rule 'helper_prefix')) 'a prefixed helper is accepted'

$f = @(Get-StepContractFindings -StepId 'screen.crop' -Manifest (New-CleanManifest) -Text '' `
        -FunctionNames @('ScreenCrop-Row') -HasParamBlock $false -LoadError '' -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'invoke_step_missing') 'a file with no Invoke-Step is reported'

# A manifest with NO functions at all used to pass clean: the check was guarded
# on the list being non-empty, so "defines nothing" read as "nothing to judge".
$f = @(Get-StepContractFindings -StepId 'screen.crop' -Manifest (New-CleanManifest) -Text '' `
        -FunctionNames @() -HasParamBlock $false -LoadError '' -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'invoke_step_missing') 'a file defining no functions at all is reported'

# ...but an unparseable file genuinely has no AST to read names off, so claiming
# Invoke-Step is missing there would be a finding we cannot support. $null means
# not known, @() means really none.
$f = @(Get-StepContractFindings -StepId 'screen.crop' -Manifest $null -Text '' `
        -FunctionNames $null -HasParamBlock $false -LoadError 'parse error' -MustRelease $mustRelease)
Assert-True (-not (Test-HasRule -Findings $f -Rule 'invoke_step_missing')) 'unknown function names do not fabricate an invoke_step_missing'
Assert-True (Test-HasRule -Findings $f -Rule 'load') 'the unparseable file is still reported as a load failure'

# $Manifest = 'bad' parses and dot-sources fine. Every rule below reaches for
# .ContainsKey, so without a shape check this threw MethodNotFound and, under
# the runner's $ErrorActionPreference = 'Stop', took the whole run down.
$f = @(Get-StepContractFindings -StepId 'screen.crop' -Manifest 'bad' -Text '' `
        -FunctionNames @('Invoke-Step') -HasParamBlock $false -LoadError '' -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'manifest_shape') 'a non-hashtable $Manifest is reported, not thrown'

$f = @(Get-StepContractFindings -StepId 'screen.crop' -Manifest @(1, 2) -Text '' `
        -FunctionNames @('Invoke-Step') -HasParamBlock $false -LoadError '' -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'manifest_shape') 'an array $Manifest is reported too'

# Same class as the two above: a malformed inputs/outputs entry used to be
# skipped in silence, while a malformed failures entry was reported. Now all
# three say so.
$m = New-CleanManifest; $m['inputs']['broken'] = 'not a hashtable'
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'field_spec_shape') 'a non-hashtable input spec is reported, not skipped'

$m = New-CleanManifest; $m['outputs']['broken'] = 'not a hashtable'
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'field_spec_shape') 'a non-hashtable output spec is reported, not skipped'

# ---- no dictionary flavour may terminate the checker -----------------------

# [ordered]@{} satisfies -is [IDictionary] but has Contains, NOT ContainsKey.
# An IDictionary guard therefore ACCEPTS it and the first key read then throws
# MethodNotFound, which under the runner's 'Stop' preference ends the whole run.
# Every key is now read through Test-StepDictHasKey, and the contract rule asks
# for a plain [hashtable], so this is a finding instead.
Assert-True (Test-StepDictHasKey -Dict @{ a = 1 } -Key 'a')            'the key helper reads a hashtable'
Assert-True (Test-StepDictHasKey -Dict ([ordered]@{ a = 1 }) -Key 'a') 'the key helper reads an ordered dictionary, which has no ContainsKey'
Assert-True (-not (Test-StepDictHasKey -Dict @{ a = 1 } -Key 'b'))     'a missing key is false, not an error'
Assert-True (-not (Test-StepDictHasKey -Dict 'not a dict' -Key 'a'))   'a non-dictionary is false, not an error'
Assert-True (-not (Test-StepDictHasKey -Dict $null -Key 'a'))          '$null is false, not an error'

$ordered = [ordered]@{
    id         = 'screen.crop'
    idempotent = $true
    inputs     = @{}
    outputs    = @{}
    failures   = @(@{ id = 'x'; transient = $false })
}
$f = @(Get-CleanFindings -Manifest $ordered -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'manifest_shape') 'an [ordered] manifest is reported, not thrown'

$m = New-CleanManifest; $m['inputs']['ord'] = [ordered]@{ type = 'string' }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'field_spec_shape') 'an [ordered] input spec is reported, not thrown'

$m = New-CleanManifest; $m['outputs']['ord'] = [ordered]@{ type = 'string' }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'field_spec_shape') 'an [ordered] output spec is reported, not thrown'

$m = New-CleanManifest; $m['failures'] = @([ordered]@{ id = 'x'; transient = $false })
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'failure_shape') 'an [ordered] failures entry is reported, not thrown'

$m = New-CleanManifest; $m['example'] = [ordered]@{ use = 'screen.crop'; with = @{ path = 'a.png' } }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True ($f.Count -ge 0) 'an [ordered] example does not throw'

# ---- a present-but-non-map inputs/outputs container -------------------------

# Reported now: with no entries to walk, every per-field rule is unreachable and
# the step would otherwise read as clean. The outputs case is the worse of the
# two, since the JSON-serializable rule is the one that goes missing.
$m = New-CleanManifest; $m['inputs'] = 'bad'
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'field_container_shape') 'a string inputs container is reported'

$m = New-CleanManifest; $m['inputs'] = @('a', 'b')
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'field_container_shape') 'an array inputs container is reported'

$m = New-CleanManifest; $m['outputs'] = 'bad'
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'field_container_shape') 'a string outputs container is reported'

$m = New-CleanManifest; $m['inputs'] = [ordered]@{ path = @{ type = 'path' } }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (Test-HasRule -Findings $f -Rule 'field_container_shape') 'an [ordered] inputs container is reported, not thrown'

# A manifest missing inputs/outputs entirely is not a container problem -- there
# is nothing present to be the wrong shape.
$m = New-CleanManifest; $m.Remove('inputs'); $m.Remove('outputs')
$m['example'] = @{ use = 'screen.crop'; with = @{} }
$f = @(Get-CleanFindings -Manifest $m -MustRelease $mustRelease)
Assert-True (-not (Test-HasRule -Findings $f -Rule 'field_container_shape')) 'an absent inputs/outputs is not reported as a bad container'

# ---- the same rules, through a real file on disk ---------------------------

$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('stepcontract_' + [System.Guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($tmpRoot)
try {
    # A clean step file: proves the file pipeline agrees with the inline path.
    $goodPath = Join-Path $tmpRoot 'screen.crop.ps1'
    $good = @(
        '$Manifest = @{'
        "    id = 'screen.crop'"
        "    group = 'screen'"
        "    summary = 'Crop a PNG'"
        "    tier = 'core'"
        "    effects = 'write'"
        '    idempotent = $true'
        "    inputs = @{ path = @{ type='path'; required=`$true } }"
        "    outputs = @{ path = @{ type='path' } }"
        "    failures = @( @{ id='file_not_found'; transient=`$false } )"
        "    example = @{ use='screen.crop'; with=@{ path='a.png' } }"
        '}'
        'function ScreenCrop-Helper { return 1 }'
        'function Invoke-Step { param($In, $Ctx) return @{ ok = $true } }'
    ) -join [Environment]::NewLine
    [System.IO.File]::WriteAllText($goodPath, $good)

    $f = @(Test-StepFileContract -Path $goodPath -MustRelease $mustRelease)
    Assert-Equal 0 $f.Count ('a clean step file yields no findings (got: ' + (Test-RuleList -Findings $f) + ')')

    # The deliberately broken fixture the card asks for. One file, every
    # file-level rule at once.
    $badPath = Join-Path $tmpRoot 'screen.bad.ps1'
    $bad = @(
        'param([switch]$Force)'
        '$Manifest = @{'
        "    id = 'screen.wrongname'"
        '    idempotent = $false'
        "    provides = @('teapot')"
        "    inputs = @{ win = @{ type='session' }; both = @{ type='int'; required=`$true; default=1 } }"
        "    outputs = @{ app = @{ type='comobject' } }"
        '    failures = @()'
        "    example = @{ use='screen.bad'; with=@{ ghost='x' } }"
        '}'
        '# non-ASCII marker: ' + [string][char]0x65E5
        'function Get-Unprefixed { return 1 }'
    ) -join [Environment]::NewLine
    [System.IO.File]::WriteAllText($badPath, $bad)

    $f = @(Test-StepFileContract -Path $badPath -MustRelease $mustRelease)

    Assert-True (Test-HasRule -Findings $f -Rule 'param_block')             'fixture: file-level param() is reported'
    Assert-True (Test-HasRule -Findings $f -Rule 'non_ascii')               'fixture: non-ASCII source is reported'
    Assert-True (Test-HasRule -Findings $f -Rule 'invoke_step_missing')     'fixture: missing Invoke-Step is reported'
    Assert-True (Test-HasRule -Findings $f -Rule 'helper_prefix')           'fixture: unprefixed helper is reported'
    Assert-True (Test-HasRule -Findings $f -Rule 'id_mismatch')             'fixture: id/file-name mismatch is reported'
    Assert-True (Test-HasRule -Findings $f -Rule 'required_and_default')    'fixture: required+default is reported'
    Assert-True (Test-HasRule -Findings $f -Rule 'session_input_no_kind')   'fixture: session input without sessionKind is reported'
    Assert-True (Test-HasRule -Findings $f -Rule 'output_not_serializable') 'fixture: non-serializable output is reported'
    Assert-True (Test-HasRule -Findings $f -Rule 'failures_empty')          'fixture: empty failures is reported'
    Assert-True (Test-HasRule -Findings $f -Rule 'example_unknown_param')   'fixture: undeclared example param is reported'
    Assert-True (Test-HasRule -Findings $f -Rule 'kind_undeclared')         'fixture: undeclared resource kind is reported'
    Assert-True (Test-HasRule -Findings $f -Rule 'resource_not_idempotent') 'fixture: non-idempotent resource step is reported'

    # A file that cannot be parsed at all still produces a finding rather than
    # throwing out of the checker.
    $brokenPath = Join-Path $tmpRoot 'screen.broken.ps1'
    [System.IO.File]::WriteAllText($brokenPath, '$Manifest = @{ id = ')
    $f = @(Test-StepFileContract -Path $brokenPath -MustRelease $mustRelease)
    Assert-True (Test-HasRule -Findings $f -Rule 'load') 'fixture: an unparseable file is reported, not thrown'
    Assert-True (-not (Test-HasRule -Findings $f -Rule 'invoke_step_missing')) 'fixture: an unparseable file is not also accused of missing Invoke-Step'

    # A conventionally named file with a perfectly good manifest and no code at
    # all. It parses, it dot-sources, and it is useless -- the runner has nothing
    # to call.
    $emptyPath = Join-Path $tmpRoot 'screen.empty.ps1'
    $empty = @(
        '$Manifest = @{'
        "    id = 'screen.empty'"
        "    group = 'screen'"
        "    summary = 'Does nothing at all'"
        "    tier = 'core'"
        "    effects = 'pure'"
        '    idempotent = $true'
        '    inputs = @{}'
        '    outputs = @{}'
        "    failures = @( @{ id='never'; transient=`$false } )"
        "    example = @{ use='screen.empty'; with=@{} }"
        '}'
    ) -join [Environment]::NewLine
    [System.IO.File]::WriteAllText($emptyPath, $empty)
    $f = @(Test-StepFileContract -Path $emptyPath -MustRelease $mustRelease)
    Assert-True (Test-HasRule -Findings $f -Rule 'invoke_step_missing') 'fixture: a manifest-only file has no entry point and is reported'

    # A file whose $Manifest is not a hashtable. This one is the reason the
    # shape check exists: it must come back as a finding, and the run must
    # survive to check the files after it.
    $shapePath = Join-Path $tmpRoot 'screen.shape.ps1'
    $shape = @(
        "`$Manifest = 'this is not a manifest'"
        'function Invoke-Step { param($In, $Ctx) return @{ ok = $true } }'
    ) -join [Environment]::NewLine
    [System.IO.File]::WriteAllText($shapePath, $shape)
    $f = @(Test-StepFileContract -Path $shapePath -MustRelease $mustRelease)
    Assert-True (Test-HasRule -Findings $f -Rule 'manifest_shape') 'fixture: a malformed $Manifest is reported, not thrown'

    # A file whose manifest and nested specs are all [ordered]@{}. It parses, it
    # dot-sources, and before the key reads went through Test-StepDictHasKey it
    # took the run down at the first identity check.
    $ordPath = Join-Path $tmpRoot 'screen.ordered.ps1'
    $ord = @(
        '$Manifest = [ordered]@{'
        "    id = 'screen.ordered'"
        "    group = 'screen'"
        "    summary = 'Ordered manifest'"
        "    tier = 'core'"
        "    effects = 'pure'"
        '    idempotent = $true'
        "    inputs = [ordered]@{ p = [ordered]@{ type='path'; required=`$true } }"
        "    outputs = [ordered]@{ q = [ordered]@{ type='path' } }"
        "    failures = @( [ordered]@{ id='x'; transient=`$false } )"
        "    example = [ordered]@{ use='screen.ordered'; with=@{ p='a.png' } }"
        '}'
        'function Invoke-Step { param($In, $Ctx) return @{ ok = $true } }'
    ) -join [Environment]::NewLine
    [System.IO.File]::WriteAllText($ordPath, $ord)
    $f = @(Test-StepFileContract -Path $ordPath -MustRelease $mustRelease)
    Assert-True (Test-HasRule -Findings $f -Rule 'manifest_shape') 'fixture: an [ordered] manifest is reported, not thrown'

    # A file whose inputs/outputs are present but are not maps at all.
    $contPath = Join-Path $tmpRoot 'screen.container.ps1'
    $cont = @(
        '$Manifest = @{'
        "    id = 'screen.container'"
        "    group = 'screen'"
        "    summary = 'Bad containers'"
        "    tier = 'core'"
        "    effects = 'pure'"
        '    idempotent = $true'
        "    inputs = 'bad'"
        "    outputs = @('also bad')"
        "    failures = @( @{ id='x'; transient=`$false } )"
        "    example = @{ use='screen.container'; with=@{} }"
        '}'
        'function Invoke-Step { param($In, $Ctx) return @{ ok = $true } }'
    ) -join [Environment]::NewLine
    [System.IO.File]::WriteAllText($contPath, $cont)
    $f = @(Test-StepFileContract -Path $contPath -MustRelease $mustRelease)
    $containerHits = 0
    foreach ($finding in $f) { if ($finding.rule -eq 'field_container_shape') { $containerHits++ } }
    Assert-Equal 2 $containerHits 'fixture: both a bad inputs and a bad outputs container are reported'

    # And prove none of them aborted anything: a good file read straight after
    # the malformed ones still comes back clean.
    $f = @(Test-StepFileContract -Path $goodPath -MustRelease $mustRelease)
    Assert-Equal 0 $f.Count 'fixture: a malformed step does not poison the files checked after it'

    # Isolation: loading one step must not leak its Invoke-Step into the next.
    # Without the child scope in Read-StepFile the second read would see the
    # first file's manifest.
    $otherPath = Join-Path $tmpRoot 'screen.other.ps1'
    $other = $good.Replace("id = 'screen.crop'", "id = 'screen.other'").Replace('ScreenCrop-Helper', 'ScreenOther-Helper').Replace("use='screen.crop'", "use='screen.other'")
    [System.IO.File]::WriteAllText($otherPath, $other)
    $readA = Read-StepFile -Path $goodPath
    $readB = Read-StepFile -Path $otherPath
    Assert-Equal 'screen.crop'  $readA.Manifest['id'] 'first file keeps its own manifest'
    Assert-Equal 'screen.other' $readB.Manifest['id'] 'a second read is not contaminated by the first'
}
finally {
    if (Test-Path -LiteralPath $tmpRoot) {
        Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ---- step vs not-yet-step ---------------------------------------------------

Assert-True (Test-IsStepFile -FileName 'screen.crop.ps1' -GroupName 'screen' -Manifest $null -FunctionNames @()) `
    'a conventionally named file is a step'
Assert-True (-not (Test-IsStepFile -FileName 'screen.crop.ps1' -GroupName 'verify' -Manifest $null -FunctionNames @())) `
    'the name only counts when the prefix is its own group'
Assert-True (-not (Test-IsStepFile -FileName 'SnapVerify.ps1' -GroupName 'verify' -Manifest $null -FunctionNames @('Test-HmAbend'))) `
    'a plain library parked under modules/ is not a step'
Assert-True (Test-IsStepFile -FileName 'helpers.ps1' -GroupName 'verify' -Manifest @{ id = 'x' } -FunctionNames @()) `
    'defining $Manifest opts a misnamed file in, so it cannot hide from the checker'
Assert-True (Test-IsStepFile -FileName 'helpers.ps1' -GroupName 'verify' -Manifest $null -FunctionNames @('Invoke-Step')) `
    'defining Invoke-Step opts a misnamed file in too'

# ---- and finally the real tree --------------------------------------------

$modulesRoot = Join-Path $repoRoot 'modules'
$moduleFiles = @(Get-StepFiles -ModulesRoot $modulesRoot)
$badSteps    = New-Object System.Collections.ArrayList
$notYetSteps = New-Object System.Collections.ArrayList
$checked     = 0

foreach ($sf in $moduleFiles) {
    $read  = Read-StepFile -Path $sf.FullName
    $group = Split-Path -Leaf (Split-Path -Parent $sf.FullName)
    $isStep = Test-IsStepFile -FileName $sf.Name -GroupName $group `
        -Manifest $read.Manifest -FunctionNames $read.FunctionNames
    if (-not $isStep) {
        [void]$notYetSteps.Add(('{0}/{1}' -f $group, $sf.Name))
        continue
    }
    $checked++
    $findings = @(Get-StepContractFindings `
        -StepId $read.StepId -Manifest $read.Manifest -Text $read.Text `
        -FunctionNames $read.FunctionNames -HasParamBlock $read.HasParamBlock `
        -LoadError $read.LoadError -MustRelease $mustRelease)
    foreach ($finding in $findings) {
        [void]$badSteps.Add(('{0}: [{1}] {2}' -f $sf.Name, $finding.rule, $finding.message))
    }
}

foreach ($line in $notYetSteps) {
    Write-Host ('  [note] ' + $line + ' is a library, not a step yet -- exempt until it is rewritten to the contract') -ForegroundColor DarkYellow
}
foreach ($line in $badSteps) { Write-Host ('  [step] ' + $line) -ForegroundColor Red }
Assert-Equal 0 $badSteps.Count ('every step under modules/ satisfies the contract (' + $checked + ' step file(s) checked, ' + $notYetSteps.Count + ' library file(s) exempt)')

exit (Complete-Tests)
