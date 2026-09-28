#Requires -Version 5.1
# ============================================================
#  kernel/Explain.ps1
#
#  `ebi explain` (P1-09): a workflow JSON rendered as the execution plan
#  a person (or an Agent) reviews before running it -- the shape of
#  Plan.md section 9, in ASCII. Dot-source only (no param() block, ASCII
#  source, no class). Pure text over the workflow, the catalog and the
#  profile; nothing runs.
#
#      before.transferStatus.capture -- <title>   (v1.0.0, profile host-open)
#      page: transferStatus = list(<label>)
#      + source: wl  ->  before_transferStatus != ok   (groupBy JOB_NAME)
#      |
#      + setup
#      |   [human ] load     table.load          Load the worklist CSV
#      |   [ui    ] ensure   browser.ensure      ... -> as mainWindow
#      + each  (x pending rows)
#      |   [write ] shot     screen.capture_window  -> capture/.../{{item.keySafe}}.png
#      |   [pure  ] verdict  verify.assert
#      |   [human ] gate     human.gate
#      |   [write ] checkpoint flow.checkpoint  when steps.gate.out.action != skip
#      + teardown
#      |   ...
#      + onError: ask   gates: 1   destructive: 0   fallback tier: none
#
#  The tag is the step's effects (human.* steps show as human); a step
#  not in the catalog shows as [?     ] so the plan is still readable
#  while lint says what is wrong.
# ============================================================

. (Join-Path $PSScriptRoot 'Context.ps1')
. (Join-Path $PSScriptRoot 'Registry.ps1')

function Get-EbiExplainTag {
    param([string]$Use, $Manifest)
    if ($Use.StartsWith('human.')) { return 'human' }
    if ($null -eq $Manifest) { return '?' }
    $e = if ($Manifest.Contains('effects')) { [string]$Manifest['effects'] } else { '?' }
    if ($e -eq 'destructive') { return 'DESTR' }
    return $e
}

function Get-EbiExplainDetail {
    # What is worth showing after the step name: a path-like input, a
    # session name, a template-bearing text input, "as", "when", "once".
    param($Call, $Manifest)
    $parts = New-Object System.Collections.ArrayList
    $with = if ($Call.Contains('with') -and ($Call['with'] -is [System.Collections.IDictionary])) { $Call['with'] } else { @{} }
    $inputs = if ($null -ne $Manifest -and $Manifest.Contains('inputs') -and ($Manifest['inputs'] -is [System.Collections.IDictionary])) { $Manifest['inputs'] } else { @{} }
    foreach ($k in ($with.Keys | Sort-Object)) {
        $name = [string]$k
        if ($name -eq 'as') { continue }
        $v = $with[$k]
        $type = if ($inputs.Contains($name) -and ($inputs[$name] -is [System.Collections.IDictionary]) -and $inputs[$name].Contains('type')) { [string]$inputs[$name]['type'] } else { '' }
        if ($type -eq 'path') { [void]$parts.Add('-> ' + [string]$v); continue }
        if ($type -eq 'session') { [void]$parts.Add($name + '=' + [string]$v); continue }
        if ($v -is [string] -and (Test-EbiTemplateString $v) -and $v.Length -le 48) { [void]$parts.Add($name + '=' + $v); continue }
    }
    if ($with.Contains('as')) { [void]$parts.Add('as ' + [string]$with['as']) }
    if ($Call.Contains('when') -and $null -ne $Call['when']) { [void]$parts.Add('when ' + [string]$Call['when']) }
    if ($Call.Contains('once') -and $null -ne $Call['once']) { [void]$parts.Add('once:' + [string]$Call['once']) }
    if ($Call.Contains('confirm') -and ($Call['confirm'] -is [bool]) -and -not $Call['confirm']) { [void]$parts.Add('confirm:false') }
    return ($parts.ToArray() -join '  ')
}

function Format-EbiExplain {
    <#
      PURE. Workflow (hashtable), Catalog (use -> manifest map), Profile
      (hashtable or $null) -> lines. Returns @{ lines; gates; destructive;
      fallback; unknown } so the CLI can also print the counts.
    #>
    param([hashtable]$Workflow, [hashtable]$Catalog, $Profile = $null)
    if ($null -eq $Catalog) { $Catalog = @{} }
    $L = New-Object System.Collections.ArrayList
    $hasProfile = ($Profile -is [System.Collections.IDictionary]) -and $Profile.Count -gt 0
    $id = if ($Workflow.Contains('id')) { [string]$Workflow['id'] } else { '?' }
    $title = if ($Workflow.Contains('title')) { [string]$Workflow['title'] } else { '' }
    $ver = if ($Workflow.Contains('version')) { [string]$Workflow['version'] } else { '' }
    $prof = if ($Workflow.Contains('profile')) { [string]$Workflow['profile'] } else { '' }
    [void]$L.Add(($id + $(if ($title -ne '') { ' -- ' + $title } else { '' }) + '   (' + $(if ($ver -ne '') { 'v' + $ver + ', ' } else { '' }) + 'profile ' + $(if ($prof -ne '') { $prof } else { 'none' }) + ')'))
    $pageName = if ($Workflow.Contains('page') -and $null -ne $Workflow['page']) { [string]$Workflow['page'] } else { '' }
    if ($pageName -ne '') {
        $desc = $pageName
        if ($hasProfile -and $Profile.Contains('pages') -and ($Profile['pages'] -is [System.Collections.IDictionary]) -and $Profile['pages'].Contains($pageName) -and ($Profile['pages'][$pageName] -is [System.Collections.IDictionary])) {
            $pg = $Profile['pages'][$pageName]
            $role = if ($pg.Contains('role')) { [string]$pg['role'] } else { '?' }
            $label = if ($pg.Contains('label')) { [string]$pg['label'] } else { '' }
            $desc = $pageName + ' = ' + $role + '(' + $(if ($label -ne '') { $label } else { $pageName }) + ')'
        } elseif ($hasProfile) { $desc = $pageName + ' = (not in the profile!)' }
        [void]$L.Add('page: ' + $desc)
    }
    $source = if ($Workflow.Contains('source') -and ($Workflow['source'] -is [System.Collections.IDictionary])) { $Workflow['source'] } else { $null }
    if ($null -ne $source) {
        $sel = if ($source.Contains('select') -and ($source['select'] -is [System.Collections.IDictionary])) { $source['select'] } else { @{} }
        $field = if ($sel.Contains('field')) { [string]$sel['field'] } else { '' }
        $pw = if ($sel.Contains('pendingWhen')) { [string]$sel['pendingWhen'] } else { '?' }
        $extra = New-Object System.Collections.ArrayList
        if ($source.Contains('groupBy') -and $null -ne $source['groupBy']) { [void]$extra.Add('groupBy ' + [string]$source['groupBy']) }
        if ($source.Contains('orderBy') -and $null -ne $source['orderBy']) { [void]$extra.Add('orderBy ' + [string]$source['orderBy']) }
        if ($source.Contains('limit') -and $null -ne $source['limit'] -and [int]$source['limit'] -gt 0) { [void]$extra.Add('limit ' + [string]$source['limit']) }
        [void]$L.Add('+ source: ' + [string]$source['table'] + '  ->  ' + $(if ($field -ne '') { $field + ' ' } else { '' }) + $pw + $(if ($extra.Count) { '   (' + ($extra.ToArray() -join ', ') + ')' } else { '' }))
        [void]$L.Add('|')
    }
    $gates = 0; $destructive = 0; $unknown = 0
    $fallback = New-Object System.Collections.ArrayList
    $idWidth = 8; $useWidth = 12
    foreach ($section in @('setup', 'each', 'teardown')) {
        if (-not $Workflow.Contains($section) -or -not ($Workflow[$section] -is [System.Collections.IList])) { continue }
        foreach ($c in $Workflow[$section]) { if ($c -is [System.Collections.IDictionary]) { if ($c.Contains('id') -and ([string]$c['id']).Length -gt $idWidth) { $idWidth = ([string]$c['id']).Length }; if ($c.Contains('use') -and ([string]$c['use']).Length -gt $useWidth) { $useWidth = ([string]$c['use']).Length } } }
    }
    foreach ($section in @('setup', 'each', 'teardown')) {
        if (-not $Workflow.Contains($section) -or -not ($Workflow[$section] -is [System.Collections.IList])) { continue }
        [void]$L.Add('+ ' + $section + $(if ($section -eq 'each') { '  (x pending rows)' } else { '' }))
        foreach ($call in $Workflow[$section]) {
            if (-not ($call -is [System.Collections.IDictionary])) { continue }
            $cid = if ($call.Contains('id')) { [string]$call['id'] } else { '?' }
            $use = if ($call.Contains('use')) { [string]$call['use'] } else { '?' }
            $m = if ($Catalog.Contains($use)) { $Catalog[$use] } else { $null }
            if ($null -eq $m) { $unknown++ }
            $tag = Get-EbiExplainTag -Use $use -Manifest $m
            if ($tag -eq 'human') { $gates++ }
            if ($tag -eq 'DESTR') { $destructive++ }
            if ($null -ne $m -and $m.Contains('tier') -and [string]$m['tier'] -eq 'fallback' -and -not ($fallback -contains $use)) { [void]$fallback.Add($use) }
            $label = if ($call.Contains('label') -and $null -ne $call['label']) { [string]$call['label'] } elseif ($null -ne $m -and $m.Contains('summary')) { [string]$m['summary'] } else { '' }
            $detail = Get-EbiExplainDetail -Call $call -Manifest $m
            $line = '|   [' + $tag.PadRight(5) + '] ' + $cid.PadRight($idWidth) + ' ' + $use.PadRight($useWidth) + '  ' + $label
            if ($detail -ne '') { $line += '   ' + $detail }
            [void]$L.Add($line)
        }
    }
    $onErr = 'ask'
    if ($Workflow.Contains('onError') -and ($Workflow['onError'] -is [System.Collections.IDictionary]) -and $Workflow['onError'].Contains('policy')) { $onErr = [string]$Workflow['onError']['policy'] }
    [void]$L.Add(('+ onError: {0}   gates: {1}   destructive: {2}   fallback tier: {3}{4}' -f $onErr, $gates, $destructive, $(if ($fallback.Count) { $fallback.ToArray() -join ', ' } else { 'none' }), $(if ($unknown -gt 0) { '   UNKNOWN STEPS: ' + $unknown } else { '' })))
    return @{ lines = $L.ToArray(); gates = $gates; destructive = $destructive; fallback = $fallback.ToArray(); unknown = $unknown }
}
