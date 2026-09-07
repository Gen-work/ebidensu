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
