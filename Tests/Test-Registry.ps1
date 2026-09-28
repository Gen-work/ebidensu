#Requires -Version 5.1
# Test-Registry.ps1 -- kernel/Registry.ps1 (P1-02): step discovery, loading
# with Invoke-Step captured per step, and the inputs schema check.
#
# The card's completion criteria: a missing required input, a type mismatch
# and an undeclared extra parameter are each reported BY PARAMETER NAME; two
# steps loaded one after the other still each call their own Invoke-Step.
# Fixture steps are written to a temp modules root at run time (same approach
# as Test-Runner.ps1), so nothing under modules/ is touched.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Registry.ps1')

Reset-Tests 'Registry'

function Get-ProblemInputs {
    param($Result)
    $names = New-Object System.Collections.ArrayList
    foreach ($p in @($Result['problems'])) { [void]$names.Add([string]$p['input']) }
    return ($names.ToArray() -join ',')
}
function Get-ProblemKinds {
    param($Result)
    $kinds = New-Object System.Collections.ArrayList
    foreach ($p in @($Result['problems'])) { [void]$kinds.Add([string]$p['kind']) }
    return ($kinds.ToArray() -join ',')
}

# ---------------------------------------------------------------- ConvertTo-EbiInputValue
$c = ConvertTo-EbiInputValue -Type 'string' -Value 'abc'
Assert-True ($c['ok'] -and $c['value'] -eq 'abc') 'string: a string passes unchanged'
$c = ConvertTo-EbiInputValue -Type 'string' -Value 42
Assert-True ($c['ok'] -and $c['value'] -is [string] -and $c['value'] -eq '42') 'string: a number is coerced to its text (whole-value templates keep types)'
$c = ConvertTo-EbiInputValue -Type 'string' -Value $true
Assert-True ($c['ok'] -and $c['value'] -is [string]) 'string: a bool is coerced to text'
$c = ConvertTo-EbiInputValue -Type 'string' -Value @{ a = 1 }
Assert-True (-not $c['ok'] -and $c['message'] -like 'expects string, got map') 'string: a map is a mismatch naming both sides'
$c = ConvertTo-EbiInputValue -Type 'string' -Value @(1, 2)
Assert-True (-not $c['ok'] -and $c['message'] -like '*got list') 'string: a list is a mismatch'

$c = ConvertTo-EbiInputValue -Type 'int' -Value 7
Assert-True ($c['ok'] -and $c['value'] -is [int] -and $c['value'] -eq 7) 'int: an int passes'
$c = ConvertTo-EbiInputValue -Type 'int' -Value ([long]12)
Assert-True ($c['ok'] -and $c['value'] -is [int]) 'int: a long (what ConvertFrom-Json gives for big numbers) becomes int'
$c = ConvertTo-EbiInputValue -Type 'int' -Value ([double]4.0)
Assert-True ($c['ok'] -and $c['value'] -is [int] -and $c['value'] -eq 4) 'int: a whole double becomes int'
$c = ConvertTo-EbiInputValue -Type 'int' -Value ([double]4.5)
Assert-True (-not $c['ok'] -and $c['message'] -like 'expects int, got number 4.5') 'int: a fractional number is a mismatch'
$c = ConvertTo-EbiInputValue -Type 'int' -Value ' -3 '
Assert-True ($c['ok'] -and $c['value'] -eq -3) 'int: a numeric string (a worklist cell) is accepted'
$c = ConvertTo-EbiInputValue -Type 'int' -Value 'abc'
Assert-True (-not $c['ok'] -and $c['message'] -like 'expects int, got string "abc"') 'int: a non-numeric string is a mismatch quoting the text'
$c = ConvertTo-EbiInputValue -Type 'int' -Value '99999999999'
Assert-True (-not $c['ok']) 'int: a string beyond int range is a mismatch, not an overflow exception'
$c = ConvertTo-EbiInputValue -Type 'int' -Value $true
Assert-True (-not $c['ok'] -and $c['message'] -like '*got bool true') 'int: a bool is a mismatch'
$c = ConvertTo-EbiInputValue -Type 'int' -Value $null
Assert-True (-not $c['ok'] -and $c['message'] -like '*got null') 'int: null is a mismatch'

$c = ConvertTo-EbiInputValue -Type 'bool' -Value $false
Assert-True ($c['ok'] -and $c['value'] -eq $false) 'bool: a bool passes'
$c = ConvertTo-EbiInputValue -Type 'bool' -Value 'TRUE'
Assert-True ($c['ok'] -and $c['value'] -eq $true) 'bool: "true" in any case is accepted'
$c = ConvertTo-EbiInputValue -Type 'bool' -Value 1
Assert-True (-not $c['ok'] -and $c['message'] -like 'expects bool*got number 1') 'bool: 1 is a mismatch (not a bool)'
$c = ConvertTo-EbiInputValue -Type 'bool' -Value 'yes'
Assert-True (-not $c['ok']) 'bool: "yes" is a mismatch'

$c = ConvertTo-EbiInputValue -Type 'path' -Value 'capture/x.png'
Assert-True ($c['ok'] -and $c['value'] -eq 'capture/x.png') 'path: a string passes unchanged (where it resolves is the step''s call for now)'
$c = ConvertTo-EbiInputValue -Type 'path' -Value @{}
Assert-True (-not $c['ok'] -and $c['message'] -like 'expects path*') 'path: a map is a mismatch'

$c = ConvertTo-EbiInputValue -Type 'rect' -Value @{ X = 1; Y = '2'; W = 30; H = 40.0 }
Assert-True ($c['ok'] -and $c['value']['Y'] -eq 2 -and $c['value']['H'] -is [int]) 'rect: a map with X,Y,W,H passes with int-coerced sides'
$c = ConvertTo-EbiInputValue -Type 'rect' -Value @{ X = 1; Y = 2; W = 30 }
Assert-True (-not $c['ok'] -and $c['message'] -like '*without H') 'rect: a missing side is named'
$c = ConvertTo-EbiInputValue -Type 'rect' -Value @{ X = 1; Y = 2; W = 'wide'; H = 4 }
Assert-True (-not $c['ok'] -and $c['message'] -like '*W expects int*') 'rect: a non-int side is named'
$c = ConvertTo-EbiInputValue -Type 'rect' -Value '1,2,3,4'
Assert-True (-not $c['ok'] -and $c['message'] -like 'expects rect {X,Y,W,H}, got string*') 'rect: a string is a mismatch'

$c = ConvertTo-EbiInputValue -Type 'list' -Value @('a', 'b')
Assert-True ($c['ok'] -and @($c['value']).Count -eq 2) 'list: an array passes'
$c = ConvertTo-EbiInputValue -Type 'list' -Value 'a'
Assert-True (-not $c['ok'] -and $c['message'] -like 'expects list, got string*') 'list: a bare scalar is a mismatch, not wrapped'
$c = ConvertTo-EbiInputValue -Type 'map' -Value @{ k = 1 }
Assert-True ($c['ok']) 'map: a hashtable passes'
$c = ConvertTo-EbiInputValue -Type 'map' -Value @(1)
Assert-True (-not $c['ok'] -and $c['message'] -like 'expects map, got list') 'map: a list is a mismatch'

$c = ConvertTo-EbiInputValue -Type 'session' -Value 'mainWindow'
Assert-True ($c['ok'] -and $c['value'] -eq 'mainWindow') 'session: a name passes (replacement is Resolve-EbiStepInputs'' job)'
$c = ConvertTo-EbiInputValue -Type 'session' -Value ''
Assert-True (-not $c['ok'] -and $c['message'] -like 'expects the name of a registered session resource*') 'session: an empty name is a mismatch'
$c = ConvertTo-EbiInputValue -Type 'session' -Value 4242
Assert-True (-not $c['ok']) 'session: a handle-looking number is a mismatch -- names only'

$c = ConvertTo-EbiInputValue -Type 'any' -Value @{ deep = @(1) }
Assert-True ($c['ok']) 'any: anything passes'
$c = ConvertTo-EbiInputValue -Type 'enum' -Value 'x'
Assert-True (-not $c['ok'] -and $c['message'] -like '*unknown type "enum"*') 'an unknown type is reported as such (it is the manifest that is wrong)'

# ---------------------------------------------------------------- Test-EbiStepInputs
$m = @{
    id = 'fake.probe'
    inputs = @{
        message  = @{ type = 'string'; required = $true }
        settleMs = @{ type = 'int';    default  = 400 }
        mode     = @{ type = 'string'; default  = 'Both'; enum = @('Ocr', 'Write', 'Both') }
        window   = @{ type = 'session'; sessionKind = 'window' }
        flag     = @{ type = 'bool' }
    }
}

$r = Test-EbiStepInputs -Manifest $m -In @{ message = 'hi' }
Assert-True ($r['ok']) 'inputs: required given, the rest optional -> ok'
Assert-Equal 400 $r['In']['settleMs'] 'inputs: a default is filled when the parameter is absent'
Assert-Equal 'Both' $r['In']['mode'] 'inputs: an enum parameter takes its default'
Assert-True (-not $r['In'].Contains('flag')) 'inputs: an optional parameter without a default is simply absent'
Assert-True (-not $r['In'].Contains('window')) 'inputs: an optional session input not given is absent (not null)'

$r = Test-EbiStepInputs -Manifest $m -In @{}
Assert-True (-not $r['ok']) 'inputs: missing required -> not ok'
Assert-Equal 'input_invalid' $r['failure'] 'inputs: the failure id is input_invalid'
Assert-Equal 'message' (Get-ProblemInputs $r) 'inputs: the missing required parameter is named'
Assert-Equal 'missing_required' (Get-ProblemKinds $r) 'inputs: ... with kind missing_required'
Assert-True ($r['message'] -like "missing required input 'message'") 'inputs: the one-line message names it too'

$r = Test-EbiStepInputs -Manifest $m -In @{ message = 'hi'; settleMs = 'soon' }
Assert-Equal 'settleMs' (Get-ProblemInputs $r) 'inputs: a type mismatch names the parameter'
Assert-Equal 'type_mismatch' (Get-ProblemKinds $r) 'inputs: ... with kind type_mismatch'
Assert-True ($r['message'] -like "input 'settleMs' expects int, got string ""soon""") 'inputs: the message says what was expected and what came'

$r = Test-EbiStepInputs -Manifest $m -In @{ message = 'hi'; colour = 'red' }
Assert-Equal 'colour' (Get-ProblemInputs $r) 'inputs: an undeclared extra parameter is named'
Assert-Equal 'unknown_input' (Get-ProblemKinds $r) 'inputs: ... with kind unknown_input'
Assert-True ($r['message'] -like "unknown input 'colour' (declared: flag, message, mode, settleMs, window)") 'inputs: the message lists what IS declared'

$r = Test-EbiStepInputs -Manifest $m -In @{ message = 'hi'; mode = 'both' }
Assert-Equal 'enum_mismatch' (Get-ProblemKinds $r) 'inputs: an enum is matched case-sensitively'
Assert-True ($r['message'] -like "input 'mode' must be one of Ocr, Write, Both; got string ""both""") 'inputs: the enum message lists the allowed values'
$r = Test-EbiStepInputs -Manifest $m -In @{ message = 'hi'; mode = 'Ocr' }
Assert-True ($r['ok'] -and $r['In']['mode'] -eq 'Ocr') 'inputs: an allowed enum value passes'

$r = Test-EbiStepInputs -Manifest $m -In @{ settleMs = 'x'; zzz = 1; aaa = 2 }
Assert-Equal 'message,settleMs,aaa,zzz' (Get-ProblemInputs $r) 'inputs: every problem is reported at once -- declared (sorted) then unknown (sorted)'
Assert-Equal 4 ($r['message'] -split '; ').Count 'inputs: the one-line message carries all four'

$r = Test-EbiStepInputs -Manifest $m -In @{ message = $null }
Assert-Equal 'missing_required' (Get-ProblemKinds $r) 'inputs: an explicit null counts as not given for a required parameter'
$r = Test-EbiStepInputs -Manifest $m -In @{ message = 'hi'; settleMs = $null }
Assert-Equal 400 $r['In']['settleMs'] 'inputs: an explicit null takes the default'
$r = Test-EbiStepInputs -Manifest $m -In @{ message = 'hi'; settleMs = '250' }
Assert-True ($r['In']['settleMs'] -is [int] -and $r['In']['settleMs'] -eq 250) 'inputs: the coerced value is what the step receives'
$r = Test-EbiStepInputs -Manifest $m -In @{ message = 'hi'; flag = 'false' }
Assert-True ($r['ok'] -and $r['In']['flag'] -eq $false) 'inputs: a bool from text arrives as a real bool'

$r = Test-EbiStepInputs -Manifest @{ id = 'fake.none' } -In @{ x = 1 }
Assert-Equal 'unknown_input' (Get-ProblemKinds $r) 'inputs: a manifest without inputs declares none, so anything given is unknown'
Assert-True ($r['message'] -like '*(declared: none)') 'inputs: ... and the message says none are declared'
$r = Test-EbiStepInputs -Manifest @{ id = 'fake.none'; inputs = @{} } -In @{}
Assert-True ($r['ok'] -and $r['In'].Count -eq 0) 'inputs: nothing declared, nothing given -> ok, empty In'
$r = Test-EbiStepInputs -Manifest @{ id = 'fake.none'; inputs = $null } -In $null
Assert-True ($r['ok']) 'inputs: null inputs and null In are both treated as empty'

$r = Test-EbiStepInputs -Manifest @{ id = 'fake.bad'; inputs = 'nope' } -In @{}
Assert-Equal 'contract_violation' $r['failure'] 'inputs: a manifest whose inputs is not a hashtable is the step''s fault (contract_violation)'
$r = Test-EbiStepInputs -Manifest @{ id = 'fake.bad'; inputs = @{ q = @{ type = 'enum' } } } -In @{ q = 'x' }
Assert-Equal 'contract_violation' $r['failure'] 'inputs: an input declared with an unknown type is contract_violation'
Assert-Equal 'q' (Get-ProblemInputs $r) 'inputs: ... naming the input'
Assert-True ($r['message'] -like '*fake.bad*unknown type "enum"*') 'inputs: ... and the step and the type'
$r = Test-EbiStepInputs -Manifest @{ id = 'fake.bad'; inputs = @{ q = 'string' } } -In @{}
Assert-Equal 'bad_manifest' (Get-ProblemKinds $r) 'inputs: an input spec that is not a hashtable is bad_manifest'
$r = Test-EbiStepInputs -Manifest @{ id = 'fake.bad'; inputs = @{ q = @{ type = 'enum' }; r = @{ type = 'int'; required = $true } } } -In @{}
Assert-Equal 'contract_violation' $r['failure'] 'inputs: a manifest problem outranks a call problem in the failure id'
Assert-Equal 'q,r' (Get-ProblemInputs $r) 'inputs: ... but both are still listed'
$r = Test-EbiStepInputs -Manifest @{ id = 'fake.untyped'; inputs = @{ q = @{ required = $true } } } -In @{ q = @{ x = 1 } }
Assert-True ($r['ok']) 'inputs: an input declared without a type is any'

# ---------------------------------------------------------------- Resolve-EbiStepInputs with the check
$cap = @{ id = 'fake.capture'; provides = @()
          inputs = @{ window = @{ type = 'session'; sessionKind = 'window'; required = $true }
                      saveAs = @{ type = 'path'; required = $true } } }
$session = @{ w = @{ kind = 'window'; value = 4242; registeredBy = 'e' } }
$r = Resolve-EbiStepInputs -Manifest $cap -With @{ window = 'w'; saveAs = 'x.png' } -Session $session
Assert-True ($r['ok'] -and $r['In']['window'] -eq 4242) 'resolve: after the check, the session name is replaced by the instance'
$r = Resolve-EbiStepInputs -Manifest $cap -With @{ window = 'w' } -Session $session
Assert-Equal 'input_invalid' $r['failure'] 'resolve: a missing required input is input_invalid'
Assert-Equal 'saveAs' (Get-ProblemInputs $r) 'resolve: ... naming it'
$r = Resolve-EbiStepInputs -Manifest $cap -With @{ window = 4242; saveAs = 'x.png' } -Session $session
Assert-Equal 'input_invalid' $r['failure'] 'resolve: a session input that is not a name is input_invalid before any lookup'
Assert-Equal 'window' (Get-ProblemInputs $r) 'resolve: ... naming the session input'
$r = Resolve-EbiStepInputs -Manifest $cap -With @{ window = 'nope'; saveAs = 'x.png' } -Session $session
Assert-Equal 'session_missing' $r['failure'] 'resolve: a well-formed but unregistered name is still session_missing'
$r = Resolve-EbiStepInputs -Manifest $cap -With @{ window = 'w'; saveAs = 'x.png'; extra = 1 } -Session $session
Assert-Equal 'extra' (Get-ProblemInputs $r) 'resolve: an undeclared parameter is reported through resolve too'
$r = Resolve-EbiStepInputs -Manifest @{ id = 'fake.pure'; provides = @(); inputs = @{} } -With @{ as = 'w' } -Session @{}
Assert-Equal 'contract_violation' $r['failure'] 'resolve: as on a non-provides step is still a contract violation'
Assert-True (-not $r['In'].Contains('as')) 'resolve: as never reaches In'
$ens = @{ id = 'fake.ensure'; provides = @('window'); inputs = @{ title = @{ type = 'string'; default = 'Edge' } } }
$r = Resolve-EbiStepInputs -Manifest $ens -With @{ as = 'w2' } -Session $session
Assert-True ($r['ok'] -and $r['As'] -eq 'w2' -and $r['In']['title'] -eq 'Edge') 'resolve: as is lifted out, then the schema fills the default'
$r = Resolve-EbiStepInputs -Manifest $ens -With @{ as = 'w'; title = 5 } -Session $session
Assert-Equal 'session_name_taken' $r['failure'] 'resolve: a live name is refused after the inputs check passes'
$r = Resolve-EbiStepInputs -Manifest $ens -With @{ as = 'w'; title = @{} } -Session $session
Assert-Equal 'input_invalid' $r['failure'] 'resolve: ... and the inputs check comes first'

# ---------------------------------------------------------------- discovery + loading (temp modules root)
$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ebidensu-registry-' + [guid]::NewGuid().ToString('N'))
$modules = Join-Path $tmpRoot 'modules'
foreach ($g in @('fake', 'other')) { New-Item -ItemType Directory -Path (Join-Path $modules $g) -Force | Out-Null }

function Write-Fixture {
    param([string]$Group, [string]$Name, [string]$Body)
    $path = Join-Path (Join-Path $modules $Group) $Name
    [System.IO.File]::WriteAllText($path, $Body, (New-Object System.Text.UTF8Encoding($false)))
}

Write-Fixture 'fake' 'fake.alpha.ps1' @'
$Manifest = @{ id = 'fake.alpha'; group = 'fake'; summary = 'fixture alpha'; tier = 'core'; effects = 'pure'
  needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs = @{ n = @{ type = 'int'; default = 1 } }; outputs = @{ who = @{ type = 'string' } }
  failures = @( @{ id = 'never'; transient = $false } ); example = @{ use = 'fake.alpha'; with = @{ n = 2 } } }
function FakeAlpha-Who { return 'alpha' }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; who = (FakeAlpha-Who) + [string]$In['n'] } }
'@
Write-Fixture 'fake' 'fake.beta.ps1' @'
$Manifest = @{ id = 'fake.beta'; group = 'fake'; summary = 'fixture beta'; tier = 'core'; effects = 'pure'
  needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs = @{}; outputs = @{ who = @{ type = 'string' } }
  failures = @( @{ id = 'never'; transient = $false } ); example = @{ use = 'fake.beta'; with = @{} } }
function FakeBeta-Who { return 'beta' }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; who = (FakeBeta-Who) } }
'@
Write-Fixture 'fake' 'fake.wrongid.ps1' @'
$Manifest = @{ id = 'fake.other'; inputs = @{}; outputs = @{}; failures = @( @{ id = 'x'; transient = $false } ) }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true } }
'@
Write-Fixture 'fake' 'fake.noinvoke.ps1' @'
$Manifest = @{ id = 'fake.noinvoke'; inputs = @{}; outputs = @{}; failures = @( @{ id = 'x'; transient = $false } ) }
'@
Write-Fixture 'fake' 'fake.throws.ps1' @'
$Manifest = @{ id = 'fake.throws' }
throw 'fixture explodes on load'
'@
Write-Fixture 'fake' 'FakeLibrary.ps1' @'
function Get-FakeLibraryThing { return 1 }
'@
Write-Fixture 'fake' 'other.misplaced.ps1' @'
$Manifest = @{ id = 'other.misplaced' }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true } }
'@
Write-Fixture 'other' 'other.gamma.ps1' @'
$Manifest = @{ id = 'other.gamma'; inputs = @{}; outputs = @{}; failures = @( @{ id = 'x'; transient = $false } ) }
function OtherGamma-Helper { return 'gamma' }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; who = (OtherGamma-Helper) } }
'@

try {
    # -- paths and discovery
    Assert-Equal '' (Get-EbiStepPath -ModulesRoot $modules -Use 'alpha') 'path: a bare verb is not a step id'
    Assert-Equal '' (Get-EbiStepPath -ModulesRoot $modules -Use '.alpha') 'path: a leading dot is not a step id'
    Assert-True ((Get-EbiStepPath -ModulesRoot $modules -Use 'fake.alpha').EndsWith('fake' + [IO.Path]::DirectorySeparatorChar + 'fake.alpha.ps1')) 'path: <root>/<group>/<use>.ps1'
    Assert-True ((Get-EbiDefaultModulesRoot).EndsWith('modules')) 'default modules root is the repo modules/ dir'

    $files = @(Find-EbiStepFiles -ModulesRoot $modules)
    $uses = @($files | ForEach-Object { $_['use'] })
    Assert-Equal 'fake.alpha,fake.beta,fake.noinvoke,fake.throws,fake.wrongid,other.gamma' ($uses -join ',') 'find: every <group>/<group>.<verb>.ps1, sorted by use'
    Assert-True (-not ($uses -contains 'other.misplaced')) 'find: a step file in the wrong group directory is not a step'
    Assert-True (-not (($files | ForEach-Object { $_['path'] }) -like '*FakeLibrary*')) 'find: a library file is not listed'
    Assert-Equal 'fake' $files[0]['group'] 'find: entries carry the group'
    Assert-Equal 0 @(Find-EbiStepFiles -ModulesRoot (Join-Path $tmpRoot 'absent')).Count 'find: a missing root lists nothing (no throw)'
    Assert-Equal 0 @(Find-EbiStepFiles -ModulesRoot '').Count 'find: an empty root lists nothing'

    # -- manifest-only reads
    $rm = Read-EbiStepManifest -Path (Join-Path (Join-Path $modules 'fake') 'fake.alpha.ps1')
    Assert-True ($rm['ok'] -and $rm['Manifest']['id'] -eq 'fake.alpha') 'read: a good step yields its manifest'
    Assert-Equal 'fake.alpha' $rm['use'] 'read: the use is derived from the file name when not given'
    Assert-True (-not (Test-Path -LiteralPath 'function:FakeAlpha-Who')) 'read: nothing the file defines leaks into the caller'
    Assert-True (-not (Test-Path -LiteralPath 'function:Invoke-Step')) 'read: no Invoke-Step is left behind'
    $rm = Read-EbiStepManifest -Path (Join-Path (Join-Path $modules 'fake') 'fake.wrongid.ps1')
    Assert-True (-not $rm['ok'] -and $rm['message'] -like '*manifest id "fake.other" does not match "fake.wrongid"*') 'read: an id/file-name mismatch is reported with both names'
    $rm = Read-EbiStepManifest -Path (Join-Path (Join-Path $modules 'fake') 'fake.noinvoke.ps1')
    Assert-True (-not $rm['ok'] -and $rm['message'] -like '*defines no Invoke-Step*') 'read: a file without Invoke-Step is reported'
    $rm = Read-EbiStepManifest -Path (Join-Path (Join-Path $modules 'fake') 'fake.throws.ps1')
    Assert-True (-not $rm['ok'] -and $rm['message'] -like '*fixture explodes on load*') 'read: a file that throws on load is reported, not thrown'
    $rm = Read-EbiStepManifest -Path (Join-Path (Join-Path $modules 'fake') 'fake.absent.ps1')
    Assert-True (-not $rm['ok'] -and $rm['message'] -like 'no step file at *') 'read: a missing file is reported'

    $cat = @(Get-EbiStepCatalog -ModulesRoot $modules)
    Assert-Equal 6 $cat.Count 'catalog: one entry per step file, broken ones included'
    $okUses = @($cat | Where-Object { $_['ok'] } | ForEach-Object { $_['use'] })
    Assert-Equal 'fake.alpha,fake.beta,other.gamma' ($okUses -join ',') 'catalog: the loadable steps read ok'
    $bad = @($cat | Where-Object { -not $_['ok'] })
    Assert-Equal 3 $bad.Count 'catalog: the three broken fixtures are entries with ok=false'
    Assert-True (($bad | ForEach-Object { $_['message'] }) -join ' ' -like '*fake.throws*') 'catalog: a broken entry says which file and why'
    Assert-True (-not (Test-Path -LiteralPath 'function:Invoke-Step')) 'catalog: no Invoke-Step left behind after a scan'

    # -- the registry proper
    $reg = New-EbiRegistry -ModulesRoot $modules
    Assert-Equal $modules $reg['ModulesRoot'] 'registry: remembers its modules root'
    Assert-True ($null -eq (Get-EbiStep -Registry $reg -Use 'fake.alpha')) 'registry: nothing is loaded until imported'

    # the card's second completion criterion: two steps loaded, each still calls its own code
    $ia = . Import-EbiStep -Registry $reg -Use 'fake.alpha'
    Assert-True ($ia['ok']) 'import: fake.alpha loads'
    $ib = . Import-EbiStep -Registry $reg -Use 'fake.beta'
    Assert-True ($ib['ok']) 'import: fake.beta loads after it'
    $ig = . Import-EbiStep -Registry $reg -Use 'other.gamma'
    Assert-True ($ig['ok']) 'import: other.gamma loads from its own group directory'
    $ctx = @{ DryRun = $false; Session = @{} }
    $ra = & (Get-EbiStep -Registry $reg -Use 'fake.alpha')['Invoke'] @{ n = 3 } $ctx
    $rb = & (Get-EbiStep -Registry $reg -Use 'fake.beta')['Invoke'] @{} $ctx
    $rg = & (Get-EbiStep -Registry $reg -Use 'other.gamma')['Invoke'] @{} $ctx
    Assert-Equal 'alpha3' $ra['who'] 'import: alpha''s captured Invoke-Step runs alpha''s code (and its helper) after beta was loaded'
    Assert-Equal 'beta' $rb['who'] 'import: beta''s captured Invoke-Step runs beta''s code'
    Assert-Equal 'gamma' $rg['who'] 'import: gamma''s captured Invoke-Step runs gamma''s code'
    Assert-True (-not (Test-Path -LiteralPath 'function:Invoke-Step')) 'import: the bare Invoke-Step is removed after capture -- the table is the only way in'
    Assert-True (Test-Path -LiteralPath 'function:FakeAlpha-Who') 'import: the step''s prefixed helper lives in the importing scope'
    Assert-True (-not (Test-Path -LiteralPath 'variable:Manifest')) 'import: $Manifest is cleaned up after the load'
    Assert-True (-not (Test-Path -LiteralPath 'variable:ebiImportState')) 'import: the import''s working state is cleaned up'
    Assert-True ((Get-EbiStep -Registry $reg -Use 'fake.alpha')['Path'].EndsWith('fake.alpha.ps1')) 'import: the entry remembers the file'
    Assert-Equal 'fake.alpha' (Get-EbiStep -Registry $reg -Use 'fake.alpha')['Manifest']['id'] 'import: the entry holds the manifest'

    $again = . Import-EbiStep -Registry $reg -Use 'fake.alpha'
    Assert-True ($again['ok'] -and [object]::ReferenceEquals($again['Entry'], (Get-EbiStep -Registry $reg -Use 'fake.alpha'))) 'import: importing a loaded use returns the same entry, no second dot-source'

    $wrong = Import-EbiStep -Registry $reg -Use 'fake.beta2'
    Assert-Equal 'internal_error' $wrong['failure'] 'import: called WITHOUT the dot is refused (helpers would be lost)'
    Assert-True ($wrong['message'] -like '*must be dot-sourced*') 'import: ... and the message says how to call it'
    $wrong = & Import-EbiStep -Registry $reg -Use 'fake.beta2'
    Assert-Equal 'internal_error' $wrong['failure'] 'import: called with & is refused too'

    $r = . Import-EbiStep -Registry $reg -Use 'fake.absent'
    Assert-Equal 'step_not_found' $r['failure'] 'import: a use with no file is step_not_found'
    Assert-True ($r['message'] -like '*fake.absent.ps1*') 'import: ... naming the path it looked at'
    $r = . Import-EbiStep -Registry $reg -Use 'absent'
    Assert-True ($r['message'] -like '*not a <group>.<verb> id*') 'import: a bare verb says why it cannot be a step'
    $r = . Import-EbiStep -Registry $reg -Use 'fake.wrongid'
    Assert-Equal 'step_not_found' $r['failure'] 'import: an id mismatch is step_not_found'
    $r = . Import-EbiStep -Registry $reg -Use 'fake.noinvoke'
    Assert-True ($r['failure'] -eq 'step_not_found' -and $r['message'] -like '*defines no Invoke-Step*') 'import: no Invoke-Step is step_not_found, not the previously loaded step''s entry point'
    Assert-True (-not (Test-Path -LiteralPath 'function:Invoke-Step')) 'import: ... and nothing is left in scope after a failed load either'
    $r = . Import-EbiStep -Registry $reg -Use 'fake.throws'
    Assert-True ($r['failure'] -eq 'step_not_found' -and $r['message'] -like '*fixture explodes on load*') 'import: a throw during load is step_not_found with the message'
    Assert-True ($null -eq (Get-EbiStep -Registry $reg -Use 'fake.throws')) 'import: a failed load is not cached'
    Assert-Equal 3 $reg['Steps'].Count 'import: only the three good steps are in the table'

    # -- the real modules tree
    $real = @(Get-EbiStepCatalog)
    $realUses = @($real | ForEach-Object { $_['use'] })
    foreach ($u in @('browser.ensure', 'human.prepare', 'screen.capture_window')) {
        Assert-True ($realUses -contains $u) ('real: ' + $u + ' is in the catalog')
    }
    Assert-True (-not ($realUses -contains 'verify.SnapVerify')) 'real: modules/verify libraries are not steps'
    $realBad = @($real | Where-Object { -not $_['ok'] })
    Assert-Equal 0 $realBad.Count ('real: every shipped step reads ok' + $(if ($realBad.Count -gt 0) { ': ' + (($realBad | ForEach-Object { $_['message'] }) -join '; ') } else { '' }))
    foreach ($e in $real) {
        $chk = Test-EbiStepInputs -Manifest $e['Manifest'] -In $(if ($e['Manifest']['example'].Contains('with')) { $e['Manifest']['example']['with'] } else { @{} })
        # the example may carry 'as' (a runner field), which the resolver strips
        # before this check, and a typed input may hold a {{template}} in the
        # example -- the runner expands it before the check runs (P1-03)
        $exWith = $(if ($e['Manifest']['example'].Contains('with')) { $e['Manifest']['example']['with'] } else { @{} })
        $probs = @($chk['problems'] | Where-Object { $_['input'] -ne 'as' -and -not ($exWith.Contains($_['input']) -and ($exWith[$_['input']] -is [string]) -and $exWith[$_['input']] -match '\{\{') })
        Assert-Equal 0 $probs.Count ('real: the example of ' + $e['use'] + ' passes its own inputs schema' + $(if ($probs.Count -gt 0) { ': ' + $chk['message'] } else { '' }))
    }
} finally {
    if (Test-Path -LiteralPath $tmpRoot) { Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

$rc = Complete-Tests
exit $rc
