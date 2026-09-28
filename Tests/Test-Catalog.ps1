#Requires -Version 5.1
# Test-Catalog.ps1 -- kernel/Docs.ps1 (P1-06): manifests -> CATALOG.md +
# catalog.json. Fixtures for the shape, then the real modules/ tree: the
# committed docs/ebi-dance/CATALOG.md and catalog.json must equal a fresh
# generation, so they can never drift from the code.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Docs.ps1')

Reset-Tests 'Catalog'

$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ebidensu-catalog-' + [guid]::NewGuid().ToString('N'))
$modules = Join-Path $tmpRoot 'modules'
foreach ($g in @('fake', 'verify')) { New-Item -ItemType Directory -Path (Join-Path $modules $g) -Force | Out-Null }
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path (Join-Path $modules 'verify') 'verify.zeta.ps1'), @'
$Manifest = @{
  id = 'verify.zeta'; group = 'verify'; summary = 'fixture: zeta'; tier = 'core'; effects = 'pure'
  needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs = @{ text = @{ type='string'; required=$true; desc='the text | with a pipe' }; mode = @{ type='string'; default='a'; enum=@('a','b') }; win = @{ type='session'; sessionKind='window' } }
  outputs = @{ records = @{ type='list'; desc='rows' } }
  failures = @( @{ id = 'parse_error'; transient = $false }, @{ id = 'timeout'; transient = $true } )
  example = @{ use='verify.zeta'; with=@{ text='{{steps.wait.out.text}}'; mode='b' } }
  notes = 'two
lines'
}
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; records = @() } }
'@, $utf8)
[System.IO.File]::WriteAllText((Join-Path (Join-Path $modules 'fake') 'fake.alpha.ps1'), @'
$Manifest = @{
  id = 'fake.alpha'; group = 'fake'; summary = 'fixture: alpha'; tier = 'fallback'; effects = 'ui'
  needs = @('foreground'); provides = @('window'); releases = @(); idempotent = $true
  inputs = @{}; outputs = @{}
  failures = @( @{ id = 'x'; transient = $false } )
  example = @{ use='fake.alpha'; with=@{ as='w' } }
}
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; resource = $null } }
'@, $utf8)
[System.IO.File]::WriteAllText((Join-Path (Join-Path $modules 'fake') 'fake.broken.ps1'), @'
$Manifest = @{ id = 'fake.other' }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true } }
'@, $utf8)

try {
    $entries = @(Get-EbiCatalogEntries -ModulesRoot $modules)
    Assert-Equal 'verify.zeta,fake.alpha,fake.broken' (@($entries | ForEach-Object { $_['use'] }) -join ',') 'entries: group order (verify before an unknown group), then use'

    $data = ConvertTo-EbiCatalogData -Entries $entries -ModulesRoot $modules
    Assert-Equal 2 $data['steps'].Count 'json: two loadable steps'
    Assert-Equal 1 $data['broken'].Count 'json: the broken one is listed under broken, not dropped'
    Assert-True ($data['broken']['fake.broken'] -like '*does not match*') 'json: ... with its reason'
    Assert-Equal 'modules/verify/verify.zeta.ps1' $data['steps']['verify.zeta']['file'] 'json: the file path is relative to the repo, forward slashes'
    Assert-Equal 'window' $data['steps']['verify.zeta']['inputs']['win']['sessionKind'] 'json: the manifest is carried verbatim'
    $text1 = ConvertTo-EbiCatalogJsonText -Data $data
    $text2 = ConvertTo-EbiCatalogJsonText -Data (ConvertTo-EbiCatalogData -Entries @(Get-EbiCatalogEntries -ModulesRoot $modules) -ModulesRoot $modules)
    Assert-Equal $text1 $text2 'json: regenerating is byte-identical (sorted keys, no timestamp)'
    $back = ConvertFrom-EbiJson -Text $text1
    Assert-True ($back['ok'] -and $back['value']['steps']['verify.zeta']['summary'] -eq 'fixture: zeta') 'json: reads back'
    Assert-True ($text1.Contains('"transient": true') -or $text1.Contains('"transient":true')) 'json: booleans are booleans'
    Assert-True (-not ($text1 -match '\\u[0-9a-fA-F]{4}')) 'json: no \uXXXX'

    $md = Format-EbiCatalogMarkdown -Entries $entries
    Assert-True ($md.StartsWith('# CATALOG')) 'md: title'
    Assert-True ($md.Contains("`n## verify`n")) 'md: a section per group'
    Assert-True ($md.Contains('### `verify.zeta`')) 'md: an entry per step'
    Assert-True ($md.Contains('| `text` | string | yes |  |  | the text \| with a pipe |')) 'md: inputs table row (pipe escaped, required marked)'
    Assert-True ($md.Contains('| `mode` | string |  | a | a, b |  |')) 'md: default and enum'
    Assert-True ($md.Contains('| `win` | session:window |')) 'md: a session input shows its kind'
    Assert-True ($md.Contains('| `records` | list | rows |')) 'md: outputs table'
    Assert-True ($md.Contains('failures: `parse_error` (not transient), `timeout` (transient)')) 'md: failures with transient flags'
    Assert-True ($md.Contains('{"id":"zeta","use":"verify.zeta","with":{')) 'md: the example is a JSON call with an id (docs check rule)'
    Assert-True ($md.Contains('Notes: two lines')) 'md: notes on one line'
    Assert-True ($md.Contains('- needs: `foreground` / provides: `window` / releases: -')) 'md: needs / provides / releases'
    Assert-True ($md.Contains('effects: `ui` / tier: `fallback`')) 'md: effects and tier'
    Assert-True ($md.Contains("`n## broken`n") -and $md.Contains('- `fake.broken`:')) 'md: broken steps get their own section'
    Assert-True ($md.Contains('| flow | (none yet) |')) 'md: an empty group still shows in the overview'
    Assert-True ($md.Contains('| verify | `verify.zeta` |')) 'md: the overview lists each group''s steps'

    $docsDir = Join-Path $tmpRoot 'docs'
    $w = Write-EbiCatalog -ModulesRoot $modules -DocsDir $docsDir
    Assert-True ($w['ok'] -and $w['steps'] -eq 2 -and $w['broken'] -eq 1) 'write: both files, counts reported'
    Assert-True ((Test-Path -LiteralPath $w['markdownPath']) -and (Test-Path -LiteralPath $w['jsonPath'])) 'write: files exist'
    $bytes = [System.IO.File]::ReadAllBytes($w['markdownPath'])
    Assert-True (-not ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB)) 'write: no BOM'
} finally {
    if (Test-Path -LiteralPath $tmpRoot) { Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue }
}

# ---------------------------------------------------------------- the real tree: committed catalog == generated
$real = Get-EbiCatalogText
$docsDir = Join-Path (Join-Path $repoRoot 'docs') 'ebi-dance'
$mdPath = Join-Path $docsDir 'CATALOG.md'
$jsPath = Join-Path $docsDir 'catalog.json'
Assert-True (Test-Path -LiteralPath $mdPath) 'real: docs/ebi-dance/CATALOG.md is committed'
Assert-True (Test-Path -LiteralPath $jsPath) 'real: docs/ebi-dance/catalog.json is committed'
$utf8 = New-Object System.Text.UTF8Encoding($false)
$mdNow = if (Test-Path -LiteralPath $mdPath) { [System.IO.File]::ReadAllText($mdPath, $utf8) } else { '' }
$jsNow = if (Test-Path -LiteralPath $jsPath) { [System.IO.File]::ReadAllText($jsPath, $utf8) } else { '' }
$mdA = ($mdNow -replace "`r`n", "`n") -split "`n"; $mdB = ($real['markdown'] -replace "`r`n", "`n") -split "`n"
$firstDiff = ''
for ($i = 0; $i -lt [Math]::Max($mdA.Count, $mdB.Count); $i++) { $x = if ($i -lt $mdA.Count) { $mdA[$i] } else { '<eof>' }; $y = if ($i -lt $mdB.Count) { $mdB[$i] } else { '<eof>' }; if ($x -ne $y) { $firstDiff = (' (first difference at line ' + ($i + 1) + ': file "' + $x + '" vs generated "' + $y + '")'); break } }
Assert-True ($firstDiff -eq '') ('real: CATALOG.md is up to date (regenerate: . kernel/Docs.ps1; Write-EbiCatalog)' + $firstDiff)
Assert-True (($jsNow -replace "`r`n", "`n") -eq ($real['json'] -replace "`r`n", "`n")) 'real: catalog.json is up to date (regenerate: . kernel/Docs.ps1; Write-EbiCatalog)'
$broken = @($real['entries'] | Where-Object { -not $_['ok'] })
Assert-Equal 0 $broken.Count 'real: no shipped step is broken'
foreach ($u in @('browser.ensure', 'human.prepare', 'screen.capture_window')) { Assert-True ($real['markdown'].Contains('### `' + $u + '`')) ('real: ' + $u + ' is in the catalog') }

$rc = Complete-Tests
exit $rc
