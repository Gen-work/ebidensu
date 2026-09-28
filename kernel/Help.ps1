#Requires -Version 5.1
# ============================================================
#  kernel/Help.ps1
#
#  `ebi help` (P1-07). Dot-source only (no param() block, ASCII source,
#  no class). Pure text functions over the catalog entries
#  (kernel/Docs.ps1's Get-EbiCatalogEntries); the CLI prints the lines.
#  Every line is at most 80 characters wide.
# ============================================================

. (Join-Path $PSScriptRoot 'Docs.ps1')
. (Join-Path $PSScriptRoot 'Gate.ps1')     # ConvertTo-EbiGateWrapped

function Get-EbiHelpWidth { return 80 }

function ConvertTo-EbiHelpFit {
    # One line, cut to the width with '...'.
    param([string]$Text, [int]$Width = 0)
    if ($Width -le 0) { $Width = Get-EbiHelpWidth }
    if ($null -eq $Text) { return '' }
    if ($Text.Length -le $Width) { return $Text }
    return ($Text.Substring(0, $Width - 3) + '...')
}

function Format-EbiHelpList {
    # PURE. `ebi help`: every step by group, one line each.
    param($Entries)
    $L = New-Object System.Collections.ArrayList
    $byGroup = @{}
    $broken = New-Object System.Collections.ArrayList
    foreach ($e in @($Entries)) {
        if (-not $e['ok']) { [void]$broken.Add($e); continue }
        $g = [string]$e['group']
        if (-not $byGroup.Contains($g)) { $byGroup[$g] = New-Object System.Collections.ArrayList }
        [void]$byGroup[$g].Add($e)
    }
    $groups = New-Object System.Collections.ArrayList
    foreach ($g in (Get-EbiCatalogGroups)) { if ($byGroup.Contains($g)) { [void]$groups.Add($g) } }
    foreach ($g in ($byGroup.Keys | Sort-Object)) { if (-not ($groups -contains $g)) { [void]$groups.Add($g) } }
    $total = 0
    foreach ($g in $groups) { $total += $byGroup[$g].Count }
    [void]$L.Add(('ebi steps: {0} in {1} group(s).  `ebi help <step>` shows one in full.' -f $total, $groups.Count))
    $width = 0
    foreach ($e in @($Entries)) { if ($e['ok'] -and ([string]$e['use']).Length -gt $width) { $width = ([string]$e['use']).Length } }
    foreach ($g in $groups) {
        [void]$L.Add('')
        [void]$L.Add($g)
        foreach ($e in $byGroup[$g]) {
            $m = $e['Manifest']
            $summary = if ($m.Contains('summary')) { [string]$m['summary'] } else { '' }
            $tag = ''
            if ($m.Contains('tier') -and [string]$m['tier'] -eq 'fallback') { $tag = ' [fallback]' }
            if ($m.Contains('effects') -and [string]$m['effects'] -eq 'destructive') { $tag += ' [DESTRUCTIVE]' }
            [void]$L.Add((ConvertTo-EbiHelpFit -Text ('  ' + ([string]$e['use']).PadRight($width) + '  ' + $summary + $tag)))
        }
    }
    if ($broken.Count -gt 0) {
        [void]$L.Add('')
        [void]$L.Add('broken (not loadable):')
        foreach ($e in $broken) { [void]$L.Add((ConvertTo-EbiHelpFit -Text ('  ' + [string]$e['use'] + '  ' + ([string]$e['message'] -replace "`r?`n", ' ')))) }
    }
    return $L.ToArray()
}

function Format-EbiHelpStep {
    # PURE. `ebi help <step>`: one manifest rendered in full, wrapped to 80.
    param($Entry)
    $L = New-Object System.Collections.ArrayList
    $w = Get-EbiHelpWidth
    if (-not $Entry['ok']) {
        [void]$L.Add([string]$Entry['use'] + ': cannot be loaded')
        foreach ($x in @(ConvertTo-EbiGateWrapped -Text ([string]$Entry['message']) -Width ($w - 2))) { [void]$L.Add('  ' + $x) }
        return $L.ToArray()
    }
    $m = $Entry['Manifest']
    function Val { param([string]$K) if ($m.Contains($K) -and $null -ne $m[$K]) { return $m[$K] }; return $null }
    function Wrap { param([string]$Text, [int]$Indent = 2) foreach ($x in @(ConvertTo-EbiGateWrapped -Text $Text -Width ($w - $Indent))) { [void]$L.Add((' ' * $Indent) + $x) } }
    [void]$L.Add([string]$Entry['use'])
    Wrap -Text ([string](Val 'summary'))
    [void]$L.Add('')
    $needs = @(Get-EbiManifestArray -Manifest $m -Key 'needs'); $prov = @(Get-EbiManifestArray -Manifest $m -Key 'provides'); $rel = @(Get-EbiManifestArray -Manifest $m -Key 'releases')
    Wrap -Text ('effects: ' + [string](Val 'effects') + '   tier: ' + [string](Val 'tier') + '   idempotent: ' + $(if ((Val 'idempotent') -eq $true) { 'true' } else { 'false' }))
    Wrap -Text ('needs: ' + $(if ($needs.Count) { $needs -join ', ' } else { '-' }) + '   provides: ' + $(if ($prov.Count) { $prov -join ', ' } else { '-' }) + '   releases: ' + $(if ($rel.Count) { $rel -join ', ' } else { '-' }))
    Wrap -Text ('file: ' + 'modules/' + (([string]$Entry['path'] -replace '\\', '/') -replace '^.*?/modules/', ''))
    [void]$L.Add('')
    [void]$L.Add('inputs')
    $inputs = Val 'inputs'
    if ($inputs -is [System.Collections.IDictionary] -and $inputs.Count -gt 0) {
        foreach ($k in ($inputs.Keys | Sort-Object)) {
            $sp = $inputs[$k]
            if (-not ($sp -is [System.Collections.IDictionary])) { Wrap -Text ([string]$k + ' : (not a hashtable)'); continue }
            $type = if ($sp.Contains('type')) { [string]$sp['type'] } else { 'any' }
            if ($type -eq 'session' -and $sp.Contains('sessionKind')) { $type = 'session:' + [string]$sp['sessionKind'] }
            $bits = New-Object System.Collections.ArrayList
            [void]$bits.Add($type)
            if ($sp.Contains('required') -and [bool]$sp['required']) { [void]$bits.Add('required') }
            if ($sp.Contains('default')) { $d = $sp['default']; [void]$bits.Add('default=' + $(if ($null -eq $d) { 'null' } elseif ($d -is [string] -and $d -eq '') { '""' } elseif ($d -is [System.Collections.IDictionary] -or $d -is [System.Collections.IList]) { ConvertTo-EbiJson -Value $d -Compress } else { [string]$d })) }
            if ($sp.Contains('enum') -and $null -ne $sp['enum']) { [void]$bits.Add('enum=' + (@($sp['enum']) -join '|')) }
            Wrap -Text ([string]$k + ' : ' + ($bits.ToArray() -join ', '))
            if ($sp.Contains('desc') -and -not [string]::IsNullOrWhiteSpace([string]$sp['desc'])) { Wrap -Text ([string]$sp['desc']) -Indent 6 }
        }
    } else { [void]$L.Add('  (none)') }
    [void]$L.Add('')
    [void]$L.Add('outputs')
    $outputs = Val 'outputs'
    if ($outputs -is [System.Collections.IDictionary] -and $outputs.Count -gt 0) {
        foreach ($k in ($outputs.Keys | Sort-Object)) {
            $sp = $outputs[$k]
            $type = if ($sp -is [System.Collections.IDictionary] -and $sp.Contains('type')) { [string]$sp['type'] } else { '?' }
            $desc = if ($sp -is [System.Collections.IDictionary] -and $sp.Contains('desc')) { '  ' + [string]$sp['desc'] } else { '' }
            Wrap -Text ([string]$k + ' : ' + $type + $desc)
        }
    } else { [void]$L.Add('  (none)') }
    [void]$L.Add('')
    $fails = New-Object System.Collections.ArrayList
    $fl = Val 'failures'
    if ($null -ne $fl) { foreach ($f in $fl) { if ($f -is [System.Collections.IDictionary] -and $f.Contains('id')) { [void]$fails.Add([string]$f['id'] + $(if ($f.Contains('transient') -and ($f['transient'] -is [bool]) -and $f['transient']) { ' (transient)' } else { '' })) } } }
    Wrap -Text ('failures: ' + $(if ($fails.Count) { $fails.ToArray() -join ', ' } else { '(none declared)' })) -Indent 0
    $ex = Val 'example'
    if ($ex -is [System.Collections.IDictionary]) {
        [void]$L.Add('')
        [void]$L.Add('example')
        $call = [ordered]@{ id = ([string]$Entry['use'] -replace '^[^.]+\.', ''); use = $(if ($ex.Contains('use')) { [string]$ex['use'] } else { [string]$Entry['use'] }) }
        if ($ex.Contains('with') -and $null -ne $ex['with']) { $call['with'] = ConvertTo-EbiSortedKeys $ex['with'] }
        Wrap -Text (ConvertTo-EbiJson -Value $call -Compress)
    }
    $notes = Val 'notes'
    if ($null -ne $notes -and [string]$notes -ne '') {
        [void]$L.Add('')
        [void]$L.Add('notes')
        Wrap -Text ([string]$notes -replace "`r?`n", ' ')
    }
    return $L.ToArray()
}

function Find-EbiHelpEntry {
    param($Entries, [string]$Use)
    foreach ($e in @($Entries)) { if ([string]$e['use'] -eq $Use) { return $e } }
    return $null
}
