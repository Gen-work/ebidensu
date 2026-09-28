# Test-StepDryRun.ps1 -- P1-36: every shipped step really runs once in CI,
# under DryRun, from its own manifest example (Tests/StepDryRun.ps1 is the
# harness). A fixture step that forgets an output in its DryRun branch must
# be caught by name; one whose return JSON cannot carry must be caught too.
# ASCII source; no param() block.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_TestCommon.ps1')
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'kernel/Registry.ps1')
. (Join-Path $repoRoot 'kernel/Json.ps1')
. (Join-Path $repoRoot 'kernel/Context.ps1')
. (Join-Path $PSScriptRoot 'StepDryRun.ps1')

Reset-Tests 'StepDryRun'
$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-dryrun-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null

# --- 1. the fixture builder ---------------------------------------------------------
Write-Host '  -- fixtures'
$m = @{ id = 'x.y'; inputs = @{ a = @{ type = 'int' }; b = @{ type = 'bool' }; c = @{ type = 'path' }; d = @{ type = 'string'; enum = @('one', 'two') }; w = @{ type = 'session'; sessionKind = 'window' }; wl = @{ type = 'session'; sessionKind = 'worklist' }; lit = @{ type = 'string' } }
        example = @{ with = @{ a = '{{page.n}}'; b = '{{vars.flag}}'; c = 'capture/{{item.keySafe}}.png'; d = '{{page.mode}}'; w = 'mainWindow'; wl = 'wl'; lit = 'kept' } } }
$fx = ConvertTo-StepDryRunWith -Manifest $m -TmpRoot $tmpRoot
Assert-True ($fx['with']['a'] -eq 100 -and $fx['with']['b'] -eq $false -and $fx['with']['c'] -eq 'fixture/c' -and $fx['with']['d'] -eq 'one' -and $fx['with']['lit'] -eq 'kept') 'fixtures: int / bool / path / enum by type; literals kept'
Assert-True ($fx['session'].Contains('mainWindow') -and $fx['session']['mainWindow']['kind'] -eq 'window' -and $fx['session'].Contains('wl') -and $fx['session']['wl']['kind'] -eq 'worklist') 'fixtures: a fake resource per session input, under the example name'
$fx = ConvertTo-StepDryRunWith -Manifest @{ id = 'x.y'; provides = @('window'); inputs = @{}; example = @{ with = @{ as = 'mainWindow' } } } -TmpRoot $tmpRoot
Assert-True ($fx['session'].Count -eq 0 -and $fx['with']['as'] -eq 'mainWindow') 'fixtures: a provides example registers its own, the session starts empty'

# --- 2. negative fixtures: the harness catches the omissions it exists for -----------
Write-Host '  -- negative fixtures'
$fixRoot = Join-Path $tmpRoot 'modules'
New-Item -ItemType Directory -Path (Join-Path $fixRoot 'fake') -Force | Out-Null
$common = @'
$Manifest = @{
  id = 'fake.__ID__'; group = 'fake'; summary = 'fixture'; tier = 'core'; effects = '__EFFECTS__'
  needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs = @{ saveAs = @{ type = 'path'; required = $true } }
  outputs = @{ path = @{ type = 'path' }; width = @{ type = 'int' } }
  failures = @( @{ id = 'save_failed'; transient = $true } )
  example = @{ use = 'fake.__ID__'; with = @{ saveAs = 'capture/{{item.keySafe}}.png' } }
}
function Invoke-Step {
    param($In, $Ctx)
    if ($Ctx['DryRun']) { __DRY__ }
    return @{ ok = $true; path = $In['saveAs']; width = 1 }
}
'@
function Write-Fake { param([string]$Id, [string]$Effects, [string]$Dry) [System.IO.File]::WriteAllText((Join-Path (Join-Path $fixRoot 'fake') ('fake.' + $Id + '.ps1')), ($common.Replace('__ID__', $Id).Replace('__EFFECTS__', $Effects).Replace('__DRY__', $Dry)), (New-Object System.Text.UTF8Encoding($false))) }
Write-Fake 'good'    'write' "`$Ctx.Log.Info('would save'); return @{ ok = `$true; path = `$In['saveAs']; width = 0 }"
Write-Fake 'nopath'  'write' "`$Ctx.Log.Info('would save'); return @{ ok = `$true; width = 0 }"
Write-Fake 'silent'  'write' "return @{ ok = `$true; path = `$In['saveAs']; width = 0 }"
Write-Fake 'deep'    'write' "`$Ctx.Log.Info('x'); `$d = @{ leaf = 1 }; 1..25 | ForEach-Object { `$d = @{ n = `$d } }; return @{ ok = `$true; path = `$In['saveAs']; width = `$d }"
Write-Fake 'notok'   'write' "`$Ctx.Log.Info('x'); return @{ ok = `$false; failure = 'save_failed'; message = 'dry run failed'; path = ''; width = 0 }"
Write-Fake 'writes'  'write' "`$Ctx.Log.Info('x'); [IO.File]::WriteAllText((Join-Path `$Ctx['WorkDir'] 'oops.txt'), 'x'); return @{ ok = `$true; path = 'p'; width = 0 }"
Write-Fake 'throws'  'write' "throw 'boom'"
Write-Fake 'scalar'  'write' "`$Ctx.Log.Info('x'); return 'done'"
Write-Fake 'pureok'  'pure'  "return @{ ok = `$true; path = 'p'; width = 0 }"
$freg = New-EbiRegistry -ModulesRoot $fixRoot
$c = Invoke-StepDryRunCheck -Registry $freg -Use 'fake.good' -TmpRoot $tmpRoot
Assert-True ($c['ok'] -and @($c['problems']).Count -eq 0) 'a correct DryRun branch passes'
$c = Invoke-StepDryRunCheck -Registry $freg -Use 'fake.nopath' -TmpRoot $tmpRoot
Assert-True (-not $c['ok'] -and @($c['problems'])[0] -like "*lacks output 'path'*") 'a DryRun branch that forgot path is reported BY NAME'
$c = Invoke-StepDryRunCheck -Registry $freg -Use 'fake.silent' -TmpRoot $tmpRoot
Assert-True (-not $c['ok'] -and @($c['problems'])[0] -like '*said nothing on DryRun*') 'a write step that says nothing on DryRun is reported'
$c = Invoke-StepDryRunCheck -Registry $freg -Use 'fake.deep' -TmpRoot $tmpRoot
Assert-True (-not $c['ok'] -and (@($c['problems']) -join ' ') -like '*not JSON-serializable*') 'a return kernel/Json.ps1 cannot carry (nesting past depth 20) is reported'
$c = Invoke-StepDryRunCheck -Registry $freg -Use 'fake.notok' -TmpRoot $tmpRoot
Assert-True (-not $c['ok'] -and (@($c['problems']) -join ' ') -like '*did not return ok=$true*') 'a DryRun that fails is reported'
$c = Invoke-StepDryRunCheck -Registry $freg -Use 'fake.writes' -TmpRoot $tmpRoot
Assert-True (-not $c['ok'] -and (@($c['problems']) -join ' ') -like '*wrote file(s)*oops.txt*') 'a DryRun that writes a file is reported'
Remove-Item -LiteralPath (Join-Path $tmpRoot 'oops.txt') -Force -ErrorAction SilentlyContinue
$c = Invoke-StepDryRunCheck -Registry $freg -Use 'fake.throws' -TmpRoot $tmpRoot
Assert-True (-not $c['ok'] -and @($c['problems'])[0] -like '*threw under DryRun*boom*') 'a throwing DryRun is reported, not propagated'
$c = Invoke-StepDryRunCheck -Registry $freg -Use 'fake.scalar' -TmpRoot $tmpRoot
Assert-True (-not $c['ok'] -and @($c['problems'])[0] -like '*not a hashtable*') 'a non-hashtable return is reported'
$c = Invoke-StepDryRunCheck -Registry $freg -Use 'fake.pureok' -TmpRoot $tmpRoot
Assert-True $c['ok'] 'a pure step need not log, and runs the same way'
Assert-True (-not (Test-Path -LiteralPath 'function:Invoke-Step')) 'no bare Invoke-Step leaks out of the harness'

# --- 3. every shipped step ---------------------------------------------------------------
Write-Host '  -- every shipped step dry-runs from its manifest example'
$modulesRoot = Join-Path $repoRoot 'modules'
$reg = New-EbiRegistry -ModulesRoot $modulesRoot
$catalog = @(Get-EbiStepCatalog -ModulesRoot $modulesRoot)
$uses = @(@($catalog) | Where-Object { $_['ok'] } | ForEach-Object { $_['use'] })
Assert-True ($uses.Count -ge 35) ('the catalog has the P1 steps (' + $uses.Count + ')')
foreach ($u in $uses) {
    $c = Invoke-StepDryRunCheck -Registry $reg -Use $u -TmpRoot $tmpRoot
    Assert-True $c['ok'] ('dry-run contract: ' + $u + $(if (-not $c['ok']) { ' -- ' + (@($c['problems']) -join ' | ') } else { '' }))
}
Assert-True ((Get-ChildItem -LiteralPath $tmpRoot -Recurse -File | Where-Object { $_.FullName -notlike '*modules*' } | Measure-Object).Count -eq 0) 'the dry runs wrote nothing under the work dir'

Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
$failed = Complete-Tests
exit $failed
