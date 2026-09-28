#Requires -Version 5.1
# ============================================================
#  kernel/Docs.ps1
#
#  Manifest -> catalog (P1-06). Dot-source only (no param() block, ASCII
#  source, no class).
#
#  docs/ebi-dance/catalog.json  what an Agent reads to assemble a
#                               workflow: every step's manifest, verbatim,
#                               plus a generated header (Plan.md 8.3)
#  docs/ebi-dance/CATALOG.md    the same for people: one section per
#                               group, one entry per step -- summary,
#                               effects / tier / needs / provides /
#                               releases, inputs and outputs tables,
#                               failures with their transient flag, the
#                               example as a JSON block, notes
#
#  Both are generated from Get-EbiStepCatalog (kernel/Registry.ps1), so
#  neither can drift from the code; Tests/Test-Docs.ps1 regenerates them
#  and fails when the committed copies differ. Never hand-edit them.
#
#  A step whose manifest cannot be read is listed too (under "broken"),
#  never dropped: a catalog that silently omits a step would hide exactly
#  the file that needs attention.
# ============================================================

. (Join-Path $PSScriptRoot 'Json.ps1')
. (Join-Path $PSScriptRoot 'Registry.ps1')

function Get-EbiCatalogGroups {
    # STEP-CONTRACT 2.1: the nine groups, in the order the catalog lists them.
    return @('browser', 'screen', 'file', 'excel', 'table', 'verify', 'human', 'progress', 'flow')
}

function Get-EbiCatalogEntries {
    # Get-EbiStepCatalog, sorted by group order then use. Broken steps kept.
    param([string]$ModulesRoot = '')
    $entries = @(Get-EbiStepCatalog -ModulesRoot $ModulesRoot)
    $order = @{}
    $i = 0
    foreach ($g in (Get-EbiCatalogGroups)) { $order[$g] = $i; $i++ }
    $keyed = New-Object System.Collections.ArrayList
    foreach ($e in $entries) {
        $g = [string]$e['group']
        $rank = if ($order.Contains($g)) { $order[$g] } else { 99 }
        [void]$keyed.Add(@{ rank = $rank; use = [string]$e['use']; entry = $e })
    }
    $sorted = @($keyed.ToArray() | Sort-Object -Property @{ Expression = { $_['rank'] } }, @{ Expression = { $_['use'] } })
    $out = New-Object System.Collections.ArrayList
    foreach ($k in $sorted) { [void]$out.Add($k['entry']) }
    return $out.ToArray()
}

function ConvertTo-EbiCatalogData {
    <#
      PURE (given entries). The catalog.json document: a hashtable with
        generated  'by kernel/Docs.ps1' (no timestamp -- the file is
                   committed and must not change when nothing did)
        groups     the nine group names
        steps      use -> manifest (hashtable, as the step declares it,
                   with the file's path relative to the modules root)
        broken     use -> message, for step files that could not be read
    #>
    param($Entries, [string]$ModulesRoot = '')
    $steps = @{}
    $broken = @{}
    foreach ($e in @($Entries)) {
        $use = [string]$e['use']
        if (-not $e['ok']) { $broken[$use] = [string]$e['message']; continue }
        $m = @{}
        foreach ($k in $e['Manifest'].Keys) { $m[[string]$k] = $e['Manifest'][$k] }
        $rel = [string]$e['path']
        if ($ModulesRoot -ne '' -and $rel.StartsWith($ModulesRoot)) { $rel = $rel.Substring($ModulesRoot.Length).TrimStart('\', '/') }
        $m['file'] = ('modules/' + ($rel -replace '\\', '/'))
        $steps[$use] = $m
    }
    return @{ generated = 'kernel/Docs.ps1 (P1-06) -- do not edit; run Tests/Run-Tests.ps1 to see drift'; groups = (Get-EbiCatalogGroups); steps = $steps; broken = $broken }
}

function ConvertTo-EbiSortedKeys {
    # PURE. The same value with every dictionary's keys in sorted order
    # ([ordered]), so serializing it is stable across processes -- a plain
    # hashtable iterates in a different order from one PowerShell to the
    # next, and a generated file must not change when nothing did.
    param($Value)
    if ($Value -is [System.Collections.IDictionary]) {
        $o = [ordered]@{}
        foreach ($k in ($Value.Keys | Sort-Object)) { $o[[string]$k] = ConvertTo-EbiSortedKeys $Value[$k] }
        return $o
    }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Collections.IList]) {
        $l = New-Object System.Collections.ArrayList
        foreach ($i in $Value) { [void]$l.Add((ConvertTo-EbiSortedKeys $i)) }
        return ,$l.ToArray()
    }
    return $Value
}

function ConvertTo-EbiCatalogJsonText {
    # catalog.json text with keys in a stable order at every level.
    param([hashtable]$Data)
    return (ConvertTo-EbiJson -Value (ConvertTo-EbiSortedKeys $Data))
}

function ConvertTo-EbiCatalogCell {
    # One markdown table cell: no pipes, no line breaks.
    param($Value)
    if ($null -eq $Value) { return '' }
    $t = if ($Value -is [bool]) { $Value.ToString().ToLowerInvariant() }
         elseif ($Value -is [System.Collections.IDictionary] -or ($Value -is [System.Collections.IList])) { ConvertTo-EbiJson -Value $Value -Compress }
         else { [string]$Value }
    return ($t -replace '\|', '\|' -replace "`r?`n", ' ')
}

function Format-EbiCatalogStep {
    # PURE. The CATALOG.md section for one loadable step.
    param([string]$Use, [hashtable]$Manifest, [string]$File)
    $L = New-Object System.Collections.ArrayList
    function Val { param([string]$K) if ($Manifest.Contains($K) -and $null -ne $Manifest[$K]) { return $Manifest[$K] }; return $null }
    function Arr { param([string]$K) return @(Get-EbiManifestArray -Manifest $Manifest -Key $K) }
    [void]$L.Add('### `' + $Use + '`')
    [void]$L.Add('')
    [void]$L.Add([string](Val 'summary'))
    [void]$L.Add('')
    $needs = @(Arr 'needs'); $prov = @(Arr 'provides'); $rel = @(Arr 'releases')
    [void]$L.Add(('- file: `{0}`' -f $File))
    [void]$L.Add(('- effects: `{0}` / tier: `{1}` / idempotent: `{2}`' -f [string](Val 'effects'), [string](Val 'tier'), (ConvertTo-EbiCatalogCell (Val 'idempotent'))))
    [void]$L.Add(('- needs: {0} / provides: {1} / releases: {2}' -f $(if ($needs.Count) { '`' + ($needs -join '`, `') + '`' } else { '-' }), $(if ($prov.Count) { '`' + ($prov -join '`, `') + '`' } else { '-' }), $(if ($rel.Count) { '`' + ($rel -join '`, `') + '`' } else { '-' })))
    [void]$L.Add('')
    $inputs = Val 'inputs'
    if ($inputs -is [System.Collections.IDictionary] -and $inputs.Count -gt 0) {
        [void]$L.Add('| input | type | required | default | enum | desc |')
        [void]$L.Add('|-------|------|----------|---------|------|------|')
        foreach ($k in ($inputs.Keys | Sort-Object)) {
            $sp = $inputs[$k]
            if (-not ($sp -is [System.Collections.IDictionary])) { [void]$L.Add(('| `{0}` | ? | | | | (not a hashtable) |' -f $k)); continue }
            $type = if ($sp.Contains('type')) { [string]$sp['type'] } else { 'any' }
            if ($type -eq 'session' -and $sp.Contains('sessionKind')) { $type = 'session:' + [string]$sp['sessionKind'] }
            $req = if ($sp.Contains('required') -and [bool]$sp['required']) { 'yes' } else { '' }
            $def = if ($sp.Contains('default')) { ConvertTo-EbiCatalogCell $sp['default'] } else { '' }
            if ($sp.Contains('default') -and $def -eq '') { $def = '(empty)' }
            $enum = if ($sp.Contains('enum') -and $null -ne $sp['enum']) { (@($sp['enum']) -join ', ') } else { '' }
            $desc = if ($sp.Contains('desc')) { ConvertTo-EbiCatalogCell $sp['desc'] } else { '' }
            [void]$L.Add(('| `{0}` | {1} | {2} | {3} | {4} | {5} |' -f $k, $type, $req, $def, $enum, $desc))
        }
    } else { [void]$L.Add('inputs: none') }
    [void]$L.Add('')
    $outputs = Val 'outputs'
    if ($outputs -is [System.Collections.IDictionary] -and $outputs.Count -gt 0) {
        [void]$L.Add('| output | type | desc |')
        [void]$L.Add('|--------|------|------|')
        foreach ($k in ($outputs.Keys | Sort-Object)) {
            $sp = $outputs[$k]
            $type = if ($sp -is [System.Collections.IDictionary] -and $sp.Contains('type')) { [string]$sp['type'] } else { '?' }
            $desc = if ($sp -is [System.Collections.IDictionary] -and $sp.Contains('desc')) { ConvertTo-EbiCatalogCell $sp['desc'] } else { '' }
            [void]$L.Add(('| `{0}` | {1} | {2} |' -f $k, $type, $desc))
        }
    } else { [void]$L.Add('outputs: none') }
    [void]$L.Add('')
    $fails = New-Object System.Collections.ArrayList
    $fl = Val 'failures'
    if ($null -ne $fl) {
        foreach ($f in $fl) {
            if ($f -is [System.Collections.IDictionary] -and $f.Contains('id')) {
                $tr = if ($f.Contains('transient') -and ($f['transient'] -is [bool]) -and $f['transient']) { 'transient' } else { 'not transient' }
                [void]$fails.Add('`' + [string]$f['id'] + '` (' + $tr + ')')
            }
        }
    }
    [void]$L.Add('failures: ' + $(if ($fails.Count) { $fails.ToArray() -join ', ' } else { 'none declared' }))
    [void]$L.Add('')
    $ex = Val 'example'
    if ($ex -is [System.Collections.IDictionary]) {
        $call = [ordered]@{ id = ($Use -replace '^[^.]+\.', '') }
        $call['use'] = $(if ($ex.Contains('use')) { [string]$ex['use'] } else { $Use })
        if ($ex.Contains('with') -and $null -ne $ex['with']) { $call['with'] = ConvertTo-EbiSortedKeys $ex['with'] }
        [void]$L.Add('```json')
        [void]$L.Add((ConvertTo-EbiJson -Value $call -Compress))
        [void]$L.Add('```')
        [void]$L.Add('')
    }
    $notes = Val 'notes'
    if ($null -ne $notes -and [string]$notes -ne '') {
        [void]$L.Add('Notes: ' + ([string]$notes -replace "`r?`n", ' '))
        [void]$L.Add('')
    }
    return $L.ToArray()
}

function Format-EbiCatalogMarkdown {
    # PURE (given entries). The whole CATALOG.md.
    param($Entries)
    $L = New-Object System.Collections.ArrayList
    [void]$L.Add('# CATALOG -- ebi-dance steps')
    [void]$L.Add('')
    [void]$L.Add('> Generated by `kernel/Docs.ps1` (P1-06) from every `$Manifest` under `modules/`.')
    [void]$L.Add('> **Do not edit.** `Tests/Test-Docs.ps1` regenerates it and fails on drift; the')
    [void]$L.Add('> Agent-facing twin is `catalog.json`. Contract: `spec/STEP-CONTRACT.md`.')
    [void]$L.Add('')
    $byGroup = @{}
    $broken = New-Object System.Collections.ArrayList
    foreach ($e in @($Entries)) {
        if (-not $e['ok']) { [void]$broken.Add($e); continue }
        $g = [string]$e['group']
        if (-not $byGroup.Contains($g)) { $byGroup[$g] = New-Object System.Collections.ArrayList }
        [void]$byGroup[$g].Add($e)
    }
    $total = 0
    foreach ($g in $byGroup.Keys) { $total += $byGroup[$g].Count }
    [void]$L.Add(('{0} step(s) in {1} group(s).' -f $total, $byGroup.Count) + $(if ($broken.Count) { (' **{0} step file(s) could not be read -- see the end.**' -f $broken.Count) } else { '' }))
    [void]$L.Add('')
    [void]$L.Add('| group | steps |')
    [void]$L.Add('|-------|-------|')
    foreach ($g in (Get-EbiCatalogGroups)) {
        $names = if ($byGroup.Contains($g)) { @($byGroup[$g] | ForEach-Object { '`' + [string]$_['use'] + '`' }) -join ', ' } else { '(none yet)' }
        [void]$L.Add(('| {0} | {1} |' -f $g, $names))
    }
    [void]$L.Add('')
    # The nine groups in order, then any group the contract does not know
    # (a misfiled step is still shown, never dropped; the contract checker
    # is what complains about the group name).
    $groupOrder = New-Object System.Collections.ArrayList
    foreach ($g in (Get-EbiCatalogGroups)) { [void]$groupOrder.Add($g) }
    foreach ($g in ($byGroup.Keys | Sort-Object)) { if (-not ($groupOrder -contains $g)) { [void]$groupOrder.Add($g) } }
    foreach ($g in $groupOrder) {
        if (-not $byGroup.Contains($g)) { continue }
        [void]$L.Add('## ' + $g)
        [void]$L.Add('')
        foreach ($e in $byGroup[$g]) {
            $rel = 'modules/' + (([string]$e['path'] -replace '\\', '/') -replace '^.*?/modules/', '')
            foreach ($line in @(Format-EbiCatalogStep -Use ([string]$e['use']) -Manifest $e['Manifest'] -File $rel)) { [void]$L.Add($line) }
        }
    }
    if ($broken.Count -gt 0) {
        [void]$L.Add('## broken')
        [void]$L.Add('')
        [void]$L.Add('Step files that could not be read. They are listed so nobody mistakes an absent entry for an absent step.')
        [void]$L.Add('')
        foreach ($e in $broken) { [void]$L.Add(('- `{0}`: {1}' -f [string]$e['use'], ([string]$e['message'] -replace "`r?`n", ' '))) }
        [void]$L.Add('')
    }
    return (($L.ToArray() -join "`n").TrimEnd() + "`n")
}

function Get-EbiCatalogText {
    # Both documents as text: @{ markdown; json }.
    param([string]$ModulesRoot = '')
    if ([string]::IsNullOrWhiteSpace($ModulesRoot)) { $ModulesRoot = Get-EbiDefaultModulesRoot }
    $entries = @(Get-EbiCatalogEntries -ModulesRoot $ModulesRoot)
    return @{ markdown = (Format-EbiCatalogMarkdown -Entries $entries); json = (ConvertTo-EbiCatalogJsonText -Data (ConvertTo-EbiCatalogData -Entries $entries -ModulesRoot $ModulesRoot)) + "`n"; entries = $entries }
}

function Write-EbiCatalog {
    # Generate CATALOG.md + catalog.json into DocsDir (default docs/ebi-dance).
    # Returns @{ ok; markdownPath; jsonPath; steps; broken; message }.
    param([string]$ModulesRoot = '', [string]$DocsDir = '')
    if ([string]::IsNullOrWhiteSpace($DocsDir)) { $DocsDir = Join-Path (Join-Path (Split-Path $PSScriptRoot -Parent) 'docs') 'ebi-dance' }
    $t = Get-EbiCatalogText -ModulesRoot $ModulesRoot
    $md = Join-Path $DocsDir 'CATALOG.md'
    $js = Join-Path $DocsDir 'catalog.json'
    try {
        if (-not (Test-Path -LiteralPath $DocsDir)) { New-Item -ItemType Directory -Path $DocsDir -Force | Out-Null }
        [System.IO.File]::WriteAllText($md, $t['markdown'], (Get-EbiJsonEncoding))
        [System.IO.File]::WriteAllText($js, $t['json'], (Get-EbiJsonEncoding))
    } catch {
        return @{ ok = $false; markdownPath = $md; jsonPath = $js; steps = 0; broken = 0; message = $_.Exception.Message }
    }
    $ok = @($t['entries'] | Where-Object { $_['ok'] }).Count
    $bad = @($t['entries']).Count - $ok
    return @{ ok = $true; markdownPath = $md; jsonPath = $js; steps = $ok; broken = $bad; message = '' }
}
