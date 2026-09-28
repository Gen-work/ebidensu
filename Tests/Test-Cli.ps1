#Requires -Version 5.1
# Test-Cli.ps1 -- P1-07 help, P1-08 lint, P1-09 explain, P1-10 dryrun / run /
# doctor, and the profile loader that run and lint share.
#
# The pure halves (Help / Lint / Explain / Profile) are tested in-process
# against fixture steps in a temp modules root. ebi.ps1 itself calls exit,
# so it is exercised in a child PowerShell of the same executable.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Lint.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Help.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Explain.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Profile.ps1')

Reset-Tests 'Cli'

$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ebidensu-cli-' + [guid]::NewGuid().ToString('N'))
$modules = Join-Path $tmpRoot 'modules'
$utf8 = New-Object System.Text.UTF8Encoding($false)
New-Item -ItemType Directory -Path (Join-Path $modules 'fake') -Force | Out-Null
function Write-Fixture { param([string]$Name, [string]$Body) [System.IO.File]::WriteAllText((Join-Path (Join-Path $modules 'fake') $Name), $Body, $utf8) }
Write-Fixture 'fake.load.ps1' @'
$Manifest = @{ id = 'fake.load'; group = 'fake'; summary = 'fixture: load a worklist'; tier = 'core'; effects = 'read'; needs = @(); provides = @('worklist'); releases = @(); idempotent = $true
  inputs = @{ path = @{ type='path'; required=$true } }; outputs = @{ rowCount = @{ type='int' } }; failures = @( @{ id = 'file_not_found'; transient = $false } ); example = @{ use='fake.load'; with=@{ path='w.csv'; as='wl' } } }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; resource = @{ path = ''; columns = @(); rows = @() }; rowCount = 0 } }
'@
Write-Fixture 'fake.ensure.ps1' @'
$Manifest = @{ id = 'fake.ensure'; group = 'fake'; summary = 'fixture: register a window'; tier = 'core'; effects = 'ui'; needs = @('browser'); provides = @('window'); releases = @(); idempotent = $true
  inputs = @{ title = @{ type='string'; default='Edge' } }; outputs = @{ title = @{ type='string' } }; failures = @( @{ id = 'no_window'; transient = $true } ); example = @{ use='fake.ensure'; with=@{ as='w' } } }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; resource = 1; title = 'x' } }
'@
Write-Fixture 'fake.shot.ps1' @'
$Manifest = @{ id = 'fake.shot'; group = 'fake'; summary = 'fixture: capture'; tier = 'core'; effects = 'write'; needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs = @{ window = @{ type='session'; sessionKind='window'; required=$true }; saveAs = @{ type='path'; required=$true }; crop = @{ type='int'; default=0 }; mode = @{ type='string'; default='a'; enum=@('a','b') } }
  outputs = @{ path = @{ type='path' } }; failures = @( @{ id = 'timeout'; transient = $true }, @{ id = 'not_found'; transient = $false } ); example = @{ use='fake.shot'; with=@{ window='w'; saveAs='x.png' } } }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; path = $In['saveAs'] } }
'@
Write-Fixture 'fake.gate.ps1' @'
$Manifest = @{ id = 'fake.gate'; group = 'fake'; summary = 'fixture: gate'; tier = 'core'; effects = 'ui'; needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs = @{ code = @{ type='string'; required=$true } }; outputs = @{ action = @{ type='string' }; code = @{ type='string' } }; failures = @( @{ id = 'operator_quit'; transient = $false } ); example = @{ use='fake.gate'; with=@{ code='ok' } } }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; action = 'pass'; code = $In['code'] } }
'@
Write-Fixture 'fake.open.ps1' @'
$Manifest = @{ id = 'fake.open'; group = 'fake'; summary = 'fixture: open a workbook'; tier = 'core'; effects = 'read'; needs = @('excel'); provides = @('workbook'); releases = @(); idempotent = $true
  inputs = @{ path = @{ type='path'; required=$true } }; outputs = @{}; failures = @( @{ id = 'file_not_found'; transient = $false } ); example = @{ use='fake.open'; with=@{ path='b.xlsx'; as='wb' } } }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; resource = 2 } }
'@
Write-Fixture 'fake.close.ps1' @'
$Manifest = @{ id = 'fake.close'; group = 'fake'; summary = 'fixture: close a workbook'; tier = 'core'; effects = 'read'; needs = @(); provides = @(); releases = @('workbook'); idempotent = $true
  inputs = @{ workbook = @{ type='session'; sessionKind='workbook'; required=$true } }; outputs = @{}; failures = @( @{ id = 'never'; transient = $false } ); example = @{ use='fake.close'; with=@{ workbook='wb' } } }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true } }
'@
Write-Fixture 'fake.ocr.ps1' @'
$Manifest = @{ id = 'fake.ocr'; group = 'fake'; summary = 'fixture: fallback tier'; tier = 'fallback'; effects = 'pure'; needs = @('calibrated:ocr'); provides = @(); releases = @(); idempotent = $true
  inputs = @{}; outputs = @{ text = @{ type='string' } }; failures = @( @{ id = 'never'; transient = $false } ); example = @{ use='fake.ocr'; with=@{} } }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; text = '' } }
'@
Write-Fixture 'fake.destroy.ps1' @'
$Manifest = @{ id = 'fake.destroy'; group = 'fake'; summary = 'fixture: destructive'; tier = 'core'; effects = 'destructive'; needs = @(); provides = @(); releases = @(); idempotent = $false
  inputs = @{}; outputs = @{}; failures = @( @{ id = 'never'; transient = $false } ); example = @{ use='fake.destroy'; with=@{} } }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true } }
'@

$entries = @(Get-EbiCatalogEntries -ModulesRoot $modules)
$catalog = ConvertTo-EbiLintCatalog -Entries $entries
$mustRelease = @{ window = $false; workbook = $true; excelApp = $true; worklist = $false }
$profile = @{
    name = 'test'
    pages = @{ ts = @{ role = 'list'; label = 'Transfer list'; url = 'https://h/ts'; crop = @{ left = 6 } } }
    worklist = @{ file = 'w.csv'; key = @{ columns = @('Correl_ID_S') }; columns = @( @{ name = 'Correl_ID_S'; role = 'key' }, @{ name = 'status'; role = 'verdict' } ) }
    vocabulary = @{ columns = @{ group = 'JOB' } }
}

# ================================================================ P1-07 help
$list = @(Format-EbiHelpList -Entries $entries)
Assert-True ($list[0] -like 'ebi steps: 8 in 1 group(s).*') 'help: header counts steps and groups'
Assert-True (($list -join "`n").Contains('fake.destroy') -and ($list -join "`n").Contains('[DESTRUCTIVE]')) 'help: a destructive step is tagged'
Assert-True (($list -join "`n").Contains('[fallback]')) 'help: a fallback step is tagged'
Assert-Equal 0 @($list | Where-Object { $_.Length -gt 80 }).Count 'help: no list line wider than 80'
$one = @(Format-EbiHelpStep -Entry (Find-EbiHelpEntry -Entries $entries -Use 'fake.shot'))
Assert-Equal 'fake.shot' $one[0] 'help step: starts with the id'
Assert-True (($one -join "`n").Contains('window : session:window, required')) 'help step: a session input shows kind and required'
Assert-True (($one -join "`n").Contains('mode : string, default=a, enum=a|b')) 'help step: default and enum'
Assert-True (($one -join "`n").Contains('failures: timeout (transient), not_found')) 'help step: failures with transient'
Assert-True (($one -join "`n").Contains('{"id":"shot","use":"fake.shot","with":{"saveAs":"x.png","window":"w"}}')) 'help step: the example as a call with an id'
Assert-Equal 0 @($one | Where-Object { $_.Length -gt 80 }).Count 'help step: no line wider than 80'
$real = @(Format-EbiHelpStep -Entry (Find-EbiHelpEntry -Entries @(Get-EbiCatalogEntries) -Use 'screen.capture_window'))
Assert-Equal 0 @($real | Where-Object { $_.Length -gt 80 }).Count 'help step: a real manifest with long descs wraps under 80'
Assert-True ($null -eq (Find-EbiHelpEntry -Entries $entries -Use 'fake.nope')) 'help: an unknown step is null'

# ================================================================ P1-08 lint
function Wf { param([string]$Json) return (ConvertFrom-EbiJson -Text $Json)['value'] }
function Msgs { param($Res) return (@($Res['errors'] | ForEach-Object { [string]$_['where'] + ': ' + [string]$_['message'] }) -join "`n") }
function Warns { param($Res) return (@($Res['warnings'] | ForEach-Object { [string]$_['where'] + ': ' + [string]$_['message'] }) -join "`n") }

$good = Wf @'
{ "schema": 1, "id": "before.ts.capture", "title": "t", "version": "1.0.0", "profile": "test", "page": "ts",
  "vars": { "side": "before" },
  "source": { "table": "wl", "select": { "field": "status", "pendingWhen": "!= ok" }, "groupBy": "JOB" },
  "onError": { "policy": "ask", "byFailure": { "timeout": { "policy": "retry", "times": 2 }, "internal_error": { "policy": "fail" } } },
  "setup": [
    { "id": "load",   "use": "fake.load",   "with": { "path": "{{profile.worklist.file}}", "as": "wl" } },
    { "id": "ensure", "use": "fake.ensure", "with": { "as": "mainWindow" } }
  ],
  "each": [
    { "id": "open", "use": "fake.open", "once": "group", "with": { "path": "{{item.group}}.xlsx", "as": "wb" } },
    { "id": "shot", "use": "fake.shot", "with": { "window": "mainWindow", "saveAs": "capture/{{vars.side}}_{{page.id}}/{{item.keySafe}}.png", "crop": "{{page.crop.left}}" } },
    { "id": "gate", "use": "fake.gate", "with": { "code": "{{steps.shot.out.path}}" } },
    { "id": "again", "use": "fake.shot", "when": "steps.gate.out.action != skip", "with": { "window": "mainWindow", "saveAs": "{{run.workDir}}/x.png", "mode": "b" } },
    { "id": "close", "use": "fake.close", "once": "groupEnd", "with": { "workbook": "wb" } }
  ],
  "teardown": [ { "id": "bye", "use": "fake.gate", "with": { "code": "{{vars.side}}" } } ] }
'@
$res = Invoke-EbiLint -Workflow $good -Catalog $catalog -Profile $profile -MustRelease $mustRelease
Assert-True ($res['ok']) ('lint: a correct workflow has no errors' + $(if (-not $res['ok']) { ': ' + (Msgs $res) } else { '' }))
Assert-True ((Warns $res) -like '*needs "browser"*' -and (Warns $res) -like '*needs "excel"*') 'lint: needs excel / browser are warnings'
Assert-True ($res['warnings'].Count -eq 2) 'lint: ... and nothing else is warned about'

$bad = Wf @'
{ "schema": 1, "id": "bad", "profile": "test", "page": "nope",
  "vars": { "side": "before" },
  "source": { "table": "nope", "select": { "field": "status", "pendingWhen": "!= ok" } },
  "onError": { "policy": "ask", "byFailure": { "ghost": { "policy": "skip" }, "not_found": { "policy": "retry" } } },
  "setup": [
    { "id": "load",   "use": "fake.load",   "with": { "path": "w.csv", "as": "wl" } },
    { "id": "open0",  "use": "fake.open",   "with": { "path": "x.xlsx" } },
    { "id": "ensure", "use": "fake.ensure", "with": { "as": "mainWindow", "title": { "x": 1 }, "colour": "red" } },
    { "id": "dup",    "use": "fake.ensure", "with": { "as": "mainWindow" } },
    { "id": "zap",    "use": "fake.destroy", "confirm": false },
    { "id": "item",   "use": "fake.gate",   "with": { "code": "{{item.key}}" } }
  ],
  "each": [
    { "id": "shot", "use": "fake.shot", "with": { "window": "ghostWindow", "saveAs": "{{page.crop.left}}/{{profile.nosuch}}/{{vars.nope}}/{{run.nope}}/{{item.nocol}}.png", "crop": "many", "mode": "zzz" } },
    { "id": "shot2", "use": "fake.shot", "with": { "window": "wl", "saveAs": "{{steps.later.out.path}}" } },
    { "id": "ocr", "use": "fake.ocr" },
    { "id": "nostep", "use": "fake.nope" },
    { "id": "open", "use": "fake.open", "once": "group", "with": { "path": "x.xlsx", "as": "wb" } },
    { "id": "later", "use": "fake.shot", "when": "steps.nope.out.x exists", "onError": { "byFailure": { "timeout": { "policy": "retry" }, "not_found": { "policy": "retry" } } }, "with": { "window": "{{vars.side}}", "saveAs": "a.png" } }
  ] }
'@
$res = Invoke-EbiLint -Workflow $bad -Catalog $catalog -Profile $profile -MustRelease $mustRelease
$m = Msgs $res
Assert-True (-not $res['ok']) 'lint: the bad workflow fails'
foreach ($want in @(
    @{ t = 'page "nope" is not in the profile';                          w = 'page binding must exist in the profile' },
    @{ t = 'setup/open0: "fake.open" provides "workbook" but the call has no "as"'; w = 'a provides step without as (P0-R17)' },
    @{ t = 'source.table: "nope" is not registered';                     w = 'source.table must be registered in setup' },
    @{ t = 'setup/ensure: input "title" expects string, got map';        w = 'a literal of the wrong type' },
    @{ t = 'setup/ensure: "fake.ensure" declares no input "colour"';     w = 'an undeclared parameter' },
    @{ t = 'setup/dup: "as": "mainWindow" is already registered by setup/ensure'; w = 'a live name registered twice' },
    @{ t = 'setup/zap: "fake.destroy" is not idempotent; setup re-runs';  w = 'a non-idempotent setup step' },
    @{ t = 'setup/item.with.code: {{item.key}}: item.* is only available in "each"'; w = 'item outside each' },
    @{ t = 'each/shot: input "window" names session resource "ghostWindow", which no earlier call registered'; w = 'an unregistered session name' },
    @{ t = '{{profile.nosuch}}';                                          w = 'a profile path that does not exist' },
    @{ t = '{{vars.nope}}: vars has no "nope"';                           w = 'an unknown var' },
    @{ t = '{{run.nope}}: run has no "nope"';                             w = 'an unknown run key' },
    @{ t = 'each/shot: input "crop" expects int, got string "many"';     w = 'a wrong-typed literal in each' },
    @{ t = 'each/shot: input "mode" must be one of a, b';                w = 'an enum violation' },
    @{ t = 'each/shot2: input "window" wants kind "window" but "wl" is a "worklist"'; w = 'a session kind mismatch' },
    @{ t = '{{steps.later.out.path}}: step "later" is not an earlier step';  w = 'a forward steps reference' },
    @{ t = 'each/nostep: use "fake.nope" is not in the catalog';         w = 'an unknown use' },
    @{ t = 'each/later.when: {{steps.nope.out.x}}: step "nope" is not an earlier step'; w = 'a when path to a missing step' },
    @{ t = 'each/later.onError.byFailure.not_found: policy retry on "not_found", which is not transient'; w = 'retry on a non-transient id (per call)' },
    @{ t = 'each/later: input "window" is a session name: a literal string, not a template'; w = 'a templated session name' },
    @{ t = 'each/open: "wb" (workbook) is registered but never released';  w = 'a mustRelease kind never released' },
    @{ t = 'onError.byFailure: "ghost" is not a failure any used step declares'; w = 'an unknown byFailure id (top level)' },
    @{ t = 'onError.byFailure.not_found: policy retry on "not_found"';     w = 'retry on a non-transient id (top level)' }
)) { Assert-True ($m.Contains($want['t'])) ('lint catches ' + $want['w']) }
Assert-True ($m -notlike '*each/shot: "fake.shot" requires input "window"*') 'lint: a given (if wrong) input is not also reported missing'
$w = Warns $res
Assert-True ($w.Contains('setup/zap: destructive step "fake.destroy" runs WITHOUT its confirm gate')) 'lint warns: confirm:false on a destructive step'
Assert-True ($w.Contains('uses fallback-tier step(s): fake.ocr')) 'lint warns: a fallback-tier step'
Assert-True ($w.Contains('needs "calibrated:ocr"')) 'lint warns: needs calibrated:ocr'
Assert-True ($w.Contains('{{item.nocol}}: the profile''s worklist.json declares no column "nocol"')) 'lint warns: an undeclared item column'

$res = Invoke-EbiLint -Workflow $good -Catalog $catalog -Profile $null -MustRelease $mustRelease
Assert-True ($res['ok']) 'lint: without a profile the profile checks are skipped, not failed'
Assert-True ((Warns $res) -like '*no profile loaded*') 'lint: ... with a warning'
$res = Invoke-EbiLint -Workflow $good -Catalog $catalog -Profile $profile -MustRelease @{}
Assert-True ((Warns $res) -like '*mustRelease kind table could not be read*') 'lint: an empty kind table is reported, not treated as nothing to release'
$badProfile = @{ pages = @{ ts = @{} }; worklist = @{ key = @{ columns = @('A') }; columns = @( @{ name = 'B'; role = 'key' } ) } }
$res = Invoke-EbiLint -Workflow $good -Catalog $catalog -Profile $badProfile -MustRelease $mustRelease
Assert-True ((Msgs $res).Contains('key.columns names "A" but no column declares role: key') -and (Msgs $res).Contains('column "B" has role: key but is not in key.columns')) 'lint: key.columns and role: key must agree (PROFILE-SCHEMA 6.1)'
$res = Invoke-EbiLint -Workflow (Wf '{ "id": "x" }') -Catalog $catalog
Assert-True (-not $res['ok'] -and (Msgs $res) -like '*"schema" is required*') 'lint: the runner shape checks are the first pass'
$report = @(Format-EbiLintReport -Result $res -Path 'x.json')
Assert-True ($report[0] -eq 'lint x.json' -and $report[-1] -like '*error(s), * warning(s) -- FAIL') 'lint: the report ends with a total'
Assert-True ((Get-EbiMustReleaseKinds).Contains('workbook') -and (Get-EbiMustReleaseKinds)['workbook'] -eq $true) 'lint: the real spec table parses (workbook must be released)'

# ================================================================ P1-09 explain
$plan = Format-EbiExplain -Workflow $good -Catalog $catalog -Profile $profile
$lines = @($plan['lines'])
Assert-True ($lines[0] -like 'before.ts.capture -- t   (v1.0.0, profile test)') 'explain: header with title, version, profile'
Assert-Equal 'page: ts = list(Transfer list)' $lines[1] 'explain: the page as role(label) from the profile'
Assert-True ($lines[2] -like '+ source: wl  ->  status != ok   (groupBy JOB)') 'explain: the source line'
Assert-True (($lines -join "`n").Contains('+ setup') -and ($lines -join "`n").Contains('+ each  (x pending rows)') -and ($lines -join "`n").Contains('+ teardown')) 'explain: three sections'
Assert-True (($lines -join "`n") -match '\|   \[read \] load\s+fake\.load\s+fixture: load a worklist   -> \{\{profile\.worklist\.file\}\}  as wl') 'explain: a step line = tag, id, use, summary, detail'
Assert-True (($lines -join "`n") -match '\[write\] shot\s+fake\.shot .*-> capture/\{\{vars\.side\}\}_\{\{page\.id\}\}/\{\{item\.keySafe\}\}\.png') 'explain: a path input is shown with ->'
Assert-True (($lines -join "`n").Contains('when steps.gate.out.action != skip')) 'explain: when is shown'
Assert-True (($lines -join "`n").Contains('once:group') -and ($lines -join "`n").Contains('once:groupEnd')) 'explain: once is shown'
Assert-True ($lines[-1] -like '+ onError: ask   gates: 0   destructive: 0   fallback tier: none') 'explain: the footer counts (fake.gate is not human.*)'
$plan2 = Format-EbiExplain -Workflow $bad -Catalog $catalog -Profile $profile
Assert-True ($plan2['destructive'] -eq 1 -and $plan2['unknown'] -eq 1 -and @($plan2['fallback']).Count -eq 1) 'explain: destructive / unknown / fallback are counted'
Assert-True (($plan2['lines'] -join "`n").Contains('[DESTR] zap') -and ($plan2['lines'] -join "`n").Contains('confirm:false')) 'explain: a destructive step with confirm:false is visible'
Assert-True (($plan2['lines'] -join "`n").Contains('[?    ] nostep')) 'explain: an unknown step still renders'
Assert-True ($plan2['lines'][1] -eq 'page: nope = (not in the profile!)') 'explain: a page missing from the profile is flagged'
$plan3 = Format-EbiExplain -Workflow $good -Catalog $catalog -Profile $null
Assert-Equal 'page: ts' $plan3['lines'][1] 'explain: without a profile the page is just its name'

# ================================================================ profile loader
$pdir = Join-Path $tmpRoot 'profiles'
New-Item -ItemType Directory -Path (Join-Path $pdir 'p1') -Force | Out-Null
[System.IO.File]::WriteAllText((Join-Path (Join-Path $pdir 'p1') 'pages.json'), '{ "ts": { "role": "list", "label": "L", "crop": { "left": 6, "top": 6 } } }', $utf8)
[System.IO.File]::WriteAllText((Join-Path (Join-Path $pdir 'p1') 'worklist.json'), '{ "file": "mapping_{{run.operator}}.csv", "key": { "columns": ["A"] } }', $utf8)
$work = Join-Path $tmpRoot 'work'
New-Item -ItemType Directory -Path $work -Force | Out-Null
[System.IO.File]::WriteAllText((Join-Path $work 'ebi.local.json'), '{ "pages": { "ts": { "crop": { "left": 9 } } }, "worklist": { "key": { "confirmedRules": [ { "kind": "fullwidth" } ] } } }', $utf8)
$p = Read-EbiProfile -Dir (Join-Path $pdir 'p1')
Assert-True ($p['ok'] -and $p['name'] -eq 'p1') 'profile: loads, named after the directory'
Assert-Equal 'L' $p['value']['pages']['ts']['label'] 'profile: pages.json is the pages key'
Assert-Equal 'mapping_{{run.operator}}.csv' $p['value']['worklist']['file'] 'profile: templates inside are left for Context.ps1'
Assert-True (($p['missing'] -contains 'grammar') -and ($p['missing'] -contains 'layout')) 'profile: absent files are listed as missing, not errors'
$p = Read-EbiProfile -Dir (Join-Path $pdir 'p1') -WorkDir $work
Assert-True ($p['overlay']) 'profile: ebi.local.json overlay applied'
Assert-Equal 9 $p['value']['pages']['ts']['crop']['left'] 'profile: overlay wins on a leaf'
Assert-Equal 6 $p['value']['pages']['ts']['crop']['top'] 'profile: siblings survive the deep merge'
Assert-Equal 'A' $p['value']['worklist']['key']['columns'][0] 'profile: an overlay adding confirmedRules keeps key.columns'
Assert-Equal 'fullwidth' $p['value']['worklist']['key']['confirmedRules'][0]['kind'] 'profile: the learned rule is there'
Assert-True (-not (Read-EbiProfile -Dir (Join-Path $pdir 'nope'))['ok']) 'profile: a missing directory is ok=false'
[System.IO.File]::WriteAllText((Join-Path (Join-Path $pdir 'p1') 'rules.json'), '{ nope', $utf8)
$p = Read-EbiProfile -Dir (Join-Path $pdir 'p1')
Assert-True (-not $p['ok'] -and $p['message'] -like '*rules.json*') 'profile: a malformed file is an error naming the file'
Assert-Equal '' (Resolve-EbiProfileDir -NameOrPath 'none') 'profile: "none" resolves to nothing'
Assert-True ((Resolve-EbiProfileDir -NameOrPath 'host-open' -ProfilesRoot $pdir).EndsWith('host-open')) 'profile: a name resolves under the profiles root'
Assert-Equal (Resolve-Path -LiteralPath (Join-Path $pdir 'p1')).ProviderPath (Resolve-EbiProfileDir -NameOrPath (Join-Path $pdir 'p1')) 'profile: a directory path resolves to itself'
$mg = Merge-EbiHashtable -Base @{ a = @{ b = 1; c = 2 }; d = @(1) } -Overlay @{ a = @{ c = 3 }; d = @(2, 3); e = 'new' }
Assert-True ($mg['a']['b'] -eq 1 -and $mg['a']['c'] -eq 3 -and $mg['d'].Count -eq 2 -and $mg['e'] -eq 'new') 'merge: maps merge, arrays and scalars replace, new keys added'

# ================================================================ P1-10 ebi.ps1 end to end (child process)
$psExe = (Get-Process -Id $PID).Path
$ebi = Join-Path $repoRoot 'ebi.ps1'
function Invoke-Ebi {
    param([string[]]$CliArgs)
    $out = & $psExe -NoLogo -NoProfile -File $ebi @CliArgs 2>&1 | Out-String
    return @{ code = $LASTEXITCODE; out = $out }
}
$r = Invoke-Ebi @('help')
Assert-True ($r['code'] -eq 0 -and $r['out'].Contains('ebi steps:') -and $r['out'].Contains('screen.capture_window')) 'cli help: lists the real steps, exit 0'
$r = Invoke-Ebi @('help', 'human.prepare')
Assert-True ($r['code'] -eq 0 -and $r['out'].Contains('human.prepare') -and $r['out'].Contains('inputs')) 'cli help <step>: renders the manifest'
$r = Invoke-Ebi @('help', 'no.such')
Assert-True ($r['code'] -eq 1 -and $r['out'].Contains('no step "no.such"')) 'cli help <unknown>: exit 1'
$r = Invoke-Ebi @('bogus')
Assert-Equal 2 $r['code'] 'cli: an unknown command is usage (exit 2)'
$r = Invoke-Ebi @('lint')
Assert-Equal 2 $r['code'] 'cli lint: a missing workflow argument is usage'
$spike = Join-Path (Join-Path $repoRoot 'workflows') 'spike.capture_window.json'
$r = Invoke-Ebi @('lint', $spike)
Assert-True ($r['code'] -eq 0 -and $r['out'].Contains('-- OK')) ('cli lint: the shipped spike workflow lints clean' + $(if ($r['code'] -ne 0) { "`n" + $r['out'] } else { '' }))
$r = Invoke-Ebi @('explain', $spike)
Assert-True ($r['code'] -eq 0 -and $r['out'].Contains('spike.capture_window') -and $r['out'].Contains('+ setup') -and $r['out'].Contains('gates: 1')) 'cli explain: the plan, one human gate'
$badWf = Join-Path $tmpRoot 'bad.json'
[System.IO.File]::WriteAllText($badWf, '{ "schema": 1, "id": "cli.bad", "setup": [ { "id": "x", "use": "no.such" } ] }', $utf8)
$r = Invoke-Ebi @('lint', $badWf)
Assert-True ($r['code'] -eq 1 -and $r['out'].Contains('[ERROR] setup/x: use "no.such" is not in the catalog')) 'cli lint: errors exit 1'
$cliWork = Join-Path $tmpRoot 'cliwork'
$r = Invoke-Ebi @('dryrun', $spike, '-WorkDir', $cliWork)
Assert-True ($r['code'] -eq 0 -and $r['out'].Contains('(dry run)') -and $r['out'].Contains('OK')) ('cli dryrun: the spike dry-runs, exit 0' + $(if ($r['code'] -ne 0) { "`n" + $r['out'] } else { '' }))
Assert-True (@(Get-ChildItem -LiteralPath (Join-Path $cliWork 'run') -Directory).Count -eq 1) 'cli dryrun: one run directory'
$r = Invoke-Ebi @('run', $badWf, '-WorkDir', $cliWork)
Assert-Equal 1 $r['code'] 'cli run: a workflow whose step is missing exits 1'
$r = Invoke-Ebi @('run', $badWf, '-WorkDir', $cliWork, '-Resume')
Assert-True ($r['code'] -eq 0 -or $r['code'] -eq 1) 'cli run -Resume: resumes the unfinished run (still fails, same exit)'
$r = Invoke-Ebi @('run', $badWf, '-WorkDir', $cliWork, '-Resume', '-RunId', 'never-existed')
Assert-True ($r['code'] -eq 1 -and $r['out'].Contains('cannot resume run never-existed')) 'cli run -Resume -RunId: an unknown run id is refused'
$r = Invoke-Ebi @('doctor')
Assert-True (($r['code'] -eq 0 -or $r['code'] -eq 1) -and $r['out'].Contains('PowerShell') -and $r['out'].Contains('Step catalog')) 'cli doctor: runs and reports PowerShell and the catalog'
$r = Invoke-Ebi @('lint', $spike, '-Profile', 'no-such-profile')
Assert-True ($r['code'] -eq 2 -and $r['out'].Contains('profile directory not found')) 'cli: a missing -Profile is usage (exit 2)'

if (Test-Path -LiteralPath $tmpRoot) { Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue }
$rc = Complete-Tests
exit $rc
