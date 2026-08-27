#Requires -Version 5.1
# Test-Docs.ps1 -- makes the ebi-dance documentation self-checking.
#
# Five review rounds on PR #141 kept finding the same three failure classes by
# hand: a definition changed and a reference to it did not, the card counts in
# BACKLOG / Plan / README drifted apart, and workflow examples referenced steps
# or sections that no longer existed. All three are mechanical. This turns them
# into a test so they never reach a human reviewer again.
#
# Part 1 unit-tests the pure helpers against inline fixtures (so a broken
# checker fails loudly instead of silently passing everything). Part 2 runs
# them over the real docs tree.
#
# Patterns containing Chinese live in Tests\docs-checks.json; this file and
# DocsCheck.ps1 stay pure ASCII (CLAUDE.md encoding rule).

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path $here 'DocsCheck.ps1')

Reset-Tests 'Docs'

$mark = Get-DocsSectionMark

# ---------------------------------------------------------------- part 1 ----
# Unit tests for the pure helpers. Fixtures only -- no repo files touched.

$fixBacklog = @'
### [x] P0-00 note
### [x] P0-R1 [BLOCK] spec fix
### [ ] P0-01 tag
### [ ] P1-01 kernel
### [ ] P1-02 [BLOCK] runner
### [ ] P4-01 excel
'@
$stats = Get-DocsCardStats -Text $fixBacklog -BlockMarker '[BLOCK]'
Assert-Equal 2 $stats.Total['P0']  'P0-00 excluded from the phase count'
Assert-Equal 1 $stats.Blocks['P0'] 'block marker counted per phase'
Assert-Equal 2 $stats.Total['P1']  'P1 counted'
Assert-Equal 4 $stats.P0P3         'P0-P3 total derived'
Assert-Equal 1 $stats.P4P5         'P4-P5 total derived'
Assert-Equal 5 $stats.All          'grand total derived'
Assert-Equal 2 $stats.AllBlocks    'block total derived'
Assert-Equal 3 $stats.Rest         'rest = all - blocks'

Assert-Equal 4 (Get-DocsExpectedCount -Stats $stats -Kind 'p0p3')      'kind p0p3'
Assert-Equal 5 (Get-DocsExpectedCount -Stats $stats -Kind 'p0p3plus1') 'kind p0p3plus1 (grep counts P0-00 too)'
Assert-Equal 2 (Get-DocsExpectedCount -Stats $stats -Kind 'total:P0')  'kind total:<phase>'
Assert-Equal 1 (Get-DocsExpectedCount -Stats $stats -Kind 'blocks:P1') 'kind blocks:<phase>'
Assert-Equal 0 (Get-DocsExpectedCount -Stats $stats -Kind 'total:P9')  'unknown phase resolves to 0'

$hits = @(Find-DocsTextHit -Text "alpha`nbeta needle`ngamma`nneedle" -Needle 'needle')
Assert-Equal 2 $hits.Count 'dead-name hits counted'
Assert-Equal 2 $hits[0]    'dead-name hit reports its line number'

$ids = @(Get-DocsSectionId -Text "## 3. head`ntext`n### 6.1 sub`n#### 6.1.2 deep`nnot ## 9. inline")
Assert-True ($ids -contains '3')     'section id from ##'
Assert-True ($ids -contains '6.1')   'section id from ###'
Assert-True ($ids -contains '6.1.2') 'three-level section id'
Assert-Equal 3 $ids.Count            'inline text is not a heading'

$refs = @(Get-DocsSectionRef -Line ('see `A.md` ' + $mark + '1.1 and `B.md` ' + $mark + '2'))
Assert-Equal 2   $refs.Count  'both refs resolved'
Assert-Equal 'A.md' $refs[0].File 'first ref binds to the nearest preceding file'
Assert-Equal '1.1'  $refs[0].Ref  'first ref number'
Assert-Equal 'B.md' $refs[1].File 'second ref binds to the later file'

$refs2 = @(Get-DocsSectionRef -Line ($mark + '6.2,' + $mark + '7;`B.md` ' + $mark + '9'))
Assert-Equal 1 $refs2.Count 'refs with no filename before them are skipped, not guessed'
Assert-Equal 'B.md' $refs2[0].File 'only the resolvable ref survives'

$refs3 = @(Get-DocsSectionRef -Line ('`spec/C.md` ' + $mark + '4'))
Assert-Equal 'C.md' $refs3[0].File 'path-qualified filename resolves to its leaf'

$fixBlocks = @'
prose
```jsonc
{ "id": "a", "use": "x.y" }
```
more
```text
{ "use": "not-json" }
```
'@
$blocks = @(Get-DocsJsonBlock -Text $fixBlocks)
Assert-Equal 1 $blocks.Count      'only json/jsonc fences are collected'
Assert-Equal 3 $blocks[0].Lines[0].Line 'block lines carry their file line number'

$goodBlock = @(Get-DocsJsonBlock -Text ("``````jsonc`n" + '"each": [' + "`n" + '{ "id": "shot", "use": "s.c" },' + "`n" + '{ "id": "crop", "use": "s.p", "with": { "p": "{{steps.shot.out.path}}" } }' + "`n]`n``````"))
Assert-Equal 0 @(Test-DocsExampleBlock -Block $goodBlock[0]).Count 'a well-formed workflow example passes'

$noId = @(Get-DocsJsonBlock -Text ("``````jsonc`n" + '"each": [ { "use": "s.c" } ]' + "`n``````"))
$iss = @(Test-DocsExampleBlock -Block $noId[0])
Assert-Equal 1        $iss.Count 'a step call without an id is reported'
Assert-Equal 'no-id'  $iss[0].Kind 'reported as no-id'

$undef = @(Get-DocsJsonBlock -Text ("``````jsonc`n" + '"each": [ { "id": "crop", "use": "s.p", "with": { "p": "{{steps.ghost.out.path}}" } } ]' + "`n``````"))
$iss2 = @(Test-DocsExampleBlock -Block $undef[0])
Assert-Equal 1            $iss2.Count 'an unresolvable {{steps.X}} is reported'
Assert-Equal 'undef-step' $iss2[0].Kind 'reported as undef-step'
Assert-Equal 'ghost'      $iss2[0].Name 'reports which step id is missing'

$frag = @(Get-DocsJsonBlock -Text ("``````jsonc`n" + '{ "id": "cp", "use": "flow.checkpoint", "with": { "v": "{{steps.verdict.out.code}}" } }' + "`n``````"))
Assert-Equal 0 @(Test-DocsExampleBlock -Block $frag[0]).Count 'a fragment may reference a step defined elsewhere'

# ---------------------------------------------------------------- part 2 ----
# The same helpers, pointed at the real tree.

$cfgPath = Join-Path $here 'docs-checks.json'
Assert-True (Test-Path -LiteralPath $cfgPath) 'docs-checks.json present'
$cfg = Get-Content -LiteralPath $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json

$docFiles = New-Object System.Collections.Generic.List[object]
foreach ($entry in $cfg.scan) {
    $full = Join-Path $repoRoot $entry
    if (Test-Path -LiteralPath $full -PathType Container) {
        foreach ($f in @(Get-ChildItem -LiteralPath $full -Filter '*.md' -File -Recurse)) {
            [void]$docFiles.Add($f.FullName)
        }
    } elseif (Test-Path -LiteralPath $full) {
        [void]$docFiles.Add((Resolve-Path -LiteralPath $full).Path)
    }
}
Assert-True ($docFiles.Count -gt 0) ('docs tree found ({0} markdown files)' -f $docFiles.Count)

$body = @{}
foreach ($f in $docFiles) {
    $body[$f] = Get-Content -LiteralPath $f -Raw -Encoding UTF8
}

# --- check 1: dead names ---
$deadHits = New-Object System.Collections.Generic.List[string]
foreach ($entry in $cfg.deadNames) {
    foreach ($f in $docFiles) {
        foreach ($ln in @(Find-DocsTextHit -Text $body[$f] -Needle $entry.text)) {
            [void]$deadHits.Add(('{0}:{1} {2}' -f (Split-Path -Leaf $f), $ln, $entry.note))
        }
    }
}
foreach ($h in $deadHits) { Write-Host ('      ' + $h) -ForegroundColor Yellow }
Assert-Equal 0 $deadHits.Count 'no retired spelling survives anywhere in docs'

# --- check 2: card counts agree across BACKLOG / Plan / README ---
$backlogPath = $null
foreach ($f in $docFiles) { if ((Split-Path -Leaf $f) -eq 'BACKLOG.md') { $backlogPath = $f } }
Assert-True ($null -ne $backlogPath) 'BACKLOG.md located'
$stats2 = Get-DocsCardStats -Text $body[$backlogPath] -BlockMarker $cfg.blockMarker

$countBad = New-Object System.Collections.Generic.List[string]
foreach ($chk in $cfg.countChecks) {
    $target = $null
    foreach ($f in $docFiles) { if ((Split-Path -Leaf $f) -eq $chk.file) { $target = $f } }
    if ($null -eq $target) {
        [void]$countBad.Add(('{0}: file not found' -f $chk.file)); continue
    }
    $m = [regex]::Match($body[$target], $chk.pattern)
    if (-not $m.Success) {
        [void]$countBad.Add(('{0}: anchor not found -- did the wording change? ({1})' -f $chk.file, ($chk.kinds -join ',')))
        continue
    }
    for ($k = 0; $k -lt $chk.kinds.Count; $k++) {
        $expect = Get-DocsExpectedCount -Stats $stats2 -Kind $chk.kinds[$k]
        $actual = [int]$m.Groups[$k + 1].Value
        if ($expect -ne $actual) {
            [void]$countBad.Add(('{0}: {1} says {2}, cards say {3}' -f $chk.file, $chk.kinds[$k], $actual, $expect))
        }
    }
}
foreach ($h in $countBad) { Write-Host ('      ' + $h) -ForegroundColor Yellow }
Assert-Equal 0 $countBad.Count ('card counts agree everywhere (P0-P3={0}, all={1}, blocks={2})' -f $stats2.P0P3, $stats2.All, $stats2.AllBlocks)

# --- check 3: cross-file section references resolve ---
$sectionIds = @{}
foreach ($f in $docFiles) {
    $sectionIds[(Split-Path -Leaf $f)] = @(Get-DocsSectionId -Text $body[$f])
}
$refBad = New-Object System.Collections.Generic.List[string]
$refSeen = 0
foreach ($f in $docFiles) {
    $lines = $body[$f] -split "`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        foreach ($r in @(Get-DocsSectionRef -Line $lines[$i])) {
            if (-not $sectionIds.ContainsKey($r.File)) { continue }
            $refSeen++
            if (-not ($sectionIds[$r.File] -contains $r.Ref)) {
                [void]$refBad.Add(('{0}:{1} -> {2} {3}{4} does not exist' -f (Split-Path -Leaf $f), ($i + 1), $r.File, $mark, $r.Ref))
            }
        }
    }
}
foreach ($h in $refBad) { Write-Host ('      ' + $h) -ForegroundColor Yellow }
Assert-Equal 0 $refBad.Count ('every resolvable section reference exists ({0} checked)' -f $refSeen)

# --- check 4: workflow examples are internally consistent ---
$exBad = New-Object System.Collections.Generic.List[string]
foreach ($f in $docFiles) {
    foreach ($blk in @(Get-DocsJsonBlock -Text $body[$f])) {
        foreach ($iss in @(Test-DocsExampleBlock -Block $blk)) {
            if ($iss.Kind -eq 'no-id') {
                [void]$exBad.Add(('{0}:{1} step call has no "id" (ledger key, P0-R3)' -f (Split-Path -Leaf $f), $iss.Line))
            } else {
                [void]$exBad.Add(('{0}:{1} {{{{steps.{2}}}}} is not defined in this example' -f (Split-Path -Leaf $f), $iss.Line, $iss.Name))
            }
        }
    }
}
foreach ($h in $exBad) { Write-Host ('      ' + $h) -ForegroundColor Yellow }
Assert-Equal 0 $exBad.Count 'workflow examples define every step they reference'

exit (Complete-Tests)
