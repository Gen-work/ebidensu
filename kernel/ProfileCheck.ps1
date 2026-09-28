#Requires -Version 5.1
# ============================================================
#  kernel/ProfileCheck.ps1
#
#  `ebi profile check` (P2-09) and the fixture runner it is built on
#  (PROFILE-SCHEMA.md 9): every profiles/<name>/fixtures/<page>/*.txt is
#  pushed through the same pure functions the workflow uses --
#  browser.assert_page's classifier, kernel/Parse.ps1's grammar,
#  kernel/Key.ps1's record match, verify.assert's rule table -- and the
#  result is compared with expected.json. Dot-source only (no param(),
#  ASCII source, no class). Pure over loaded data plus file reads of the
#  fixture directory.
#
#  Test-EbiProfileSchema  -> @{ ok; errors; warnings }   the five JSON files
#  Invoke-EbiFixtureCheck -> @{ ok; results; passed; failed; skipped }
#  Compare-EbiProfile     -> @{ lines; missing; extra; same; changed }   (diff)
#  New-EbiProfileSkeleton -> writes a commented skeleton directory      (new)
# ============================================================

. (Join-Path $PSScriptRoot 'Json.ps1')
. (Join-Path $PSScriptRoot 'Parse.ps1')
. (Join-Path $PSScriptRoot 'Key.ps1')
. (Join-Path $PSScriptRoot 'Worklist.ps1')
. (Join-Path $PSScriptRoot 'Context.ps1')
. (Join-Path $PSScriptRoot 'Rules.ps1')

function Get-EbiProfileRoles { return @('entry', 'form', 'record', 'list', 'document') }
function Get-EbiProfileColumnRoles { return @('key', 'group', 'owner', 'deliverable', 'verdict', 'bitmask', 'text', 'time') }
function Get-EbiProfileVerdicts { return @('ok', 'ng', 'unknown', 'pending') }

function Test-EbiProfileSchema {
    <#
      PURE. The loaded profile (hashtable: vocabulary / pages / grammar /
      rules / worklist / ...) -> @{ ok; errors; warnings } (string[]).
      Every message names the file and the key.
    #>
    param($Profile, $Missing = @())
    $E = New-Object System.Collections.ArrayList
    $W = New-Object System.Collections.ArrayList
    if ($null -eq $Profile -or -not ($Profile -is [System.Collections.IDictionary])) { [void]$E.Add('profile is not loaded'); return @{ ok = $false; errors = $E.ToArray(); warnings = $W.ToArray() } }
    foreach ($f in @('vocabulary', 'pages', 'worklist')) { if (@($Missing) -contains $f) { [void]$E.Add($f + '.json: missing (required)') } }
    foreach ($f in @('grammar', 'rules')) { if (@($Missing) -contains $f) { [void]$W.Add($f + '.json: missing; no page can be parsed or judged') } }
    # vocabulary
    if ($Profile.Contains('vocabulary') -and ($Profile['vocabulary'] -is [System.Collections.IDictionary])) {
        $v = $Profile['vocabulary']
        $roles = if ($v.Contains('roles') -and ($v['roles'] -is [System.Collections.IDictionary])) { $v['roles'] } else { @{} }
        $want = Get-EbiProfileRoles
        foreach ($r in $want) { if (-not $roles.Contains($r)) { [void]$E.Add('vocabulary.json: roles lacks "' + $r + '" (exactly the five: ' + ($want -join ', ') + ')') } }
        foreach ($r in $roles.Keys) { if (-not ($want -contains [string]$r)) { [void]$E.Add('vocabulary.json: roles has "' + $r + '", which is not one of the five') } }
        if ($v.Contains('columns') -and ($v['columns'] -is [System.Collections.IDictionary]) -and $v['columns'].Contains('key')) { [void]$E.Add('vocabulary.json: columns.key is not allowed; the key is declared once, in worklist.json key.columns (P0-R4)') }
        if (-not $v.Contains('sides') -or -not ($v['sides'] -is [System.Collections.IDictionary]) -or $v['sides'].Count -eq 0) { [void]$W.Add('vocabulary.json: no sides declared') }
    }
    # pages
    $pages = if ($Profile.Contains('pages') -and ($Profile['pages'] -is [System.Collections.IDictionary])) { $Profile['pages'] } else { @{} }
    if ($pages.Count -eq 0) { [void]$W.Add('pages.json: no page declared') }
    foreach ($name in $pages.Keys) {
        $p = $pages[$name]
        if (-not ($p -is [System.Collections.IDictionary])) { [void]$E.Add('pages.json: "' + $name + '" is not an object'); continue }
        $role = if ($p.Contains('role')) { [string]$p['role'] } else { '' }
        if (-not ((Get-EbiProfileRoles) -contains $role)) { [void]$E.Add('pages.json: "' + $name + '" role "' + $role + '" is not one of ' + ((Get-EbiProfileRoles) -join ', ')) }
        foreach ($reserved in @('grammar', 'rules', 'id')) { if ($p.Contains($reserved)) { [void]$E.Add('pages.json: "' + $name + '" has a field named "' + $reserved + '", which {{page.' + $reserved + '}} shadows (PROFILE-SCHEMA 3)') } }
        if (-not $p.Contains('fingerprint') -or -not ($p['fingerprint'] -is [System.Collections.IDictionary])) { [void]$E.Add('pages.json: "' + $name + '" has no fingerprint (3.1)') }
        else {
            $fp = $p['fingerprint']
            $okList = @(if ($fp.Contains('ok') -and $null -ne $fp['ok']) { $fp['ok'] })
            if ($okList.Count -eq 0) { [void]$W.Add('pages.json: "' + $name + '" fingerprint.ok is empty; browser.assert_page can never say ok on it') }
            foreach ($k in $fp.Keys) { if (-not ([string]$k -in @('ok', 'loading', 'empty', 'expired'))) { [void]$E.Add('pages.json: "' + $name + '" fingerprint has unknown kind "' + $k + '"') } }
            if (-not $fp.Contains('expired') -or @($fp['expired']).Count -eq 0) { [void]$W.Add('pages.json: "' + $name + '" fingerprint.expired is empty; a session timeout will read as unknown page') }
        }
        if (-not $p.Contains('label') -or [string]::IsNullOrWhiteSpace([string]$p['label'])) { [void]$W.Add('pages.json: "' + $name + '" has no label') }
    }
    # grammar
    $grammar = if ($Profile.Contains('grammar') -and ($Profile['grammar'] -is [System.Collections.IDictionary])) { $Profile['grammar'] } else { @{} }
    foreach ($name in $grammar.Keys) {
        $g = $grammar[$name]
        if (-not $pages.Contains($name)) { [void]$E.Add('grammar.json: "' + $name + '" is not a page in pages.json (keys are page names, P0-R1)') }
        if (-not ($g -is [System.Collections.IDictionary])) { [void]$E.Add('grammar.json: "' + $name + '" is not an object'); continue }
        $parser = if ($g.Contains('parser')) { [string]$g['parser'] } else { '' }
        if (-not ($parser -in @('delimited', 'labeled', 'columns', 'regex'))) { [void]$E.Add('grammar.json: "' + $name + '" parser "' + $parser + '" is not delimited | labeled | columns | regex') }
        $probe = ConvertFrom-EbiGrammar -Text 'x' -Grammar $g
        if (-not $probe['ok'] -and $probe['message'] -notlike '*header line not found*') { [void]$E.Add('grammar.json: "' + $name + '": ' + $probe['message']) }
        foreach ($rx in @(if ($g.Contains('ignore') -and $null -ne $g['ignore']) { $g['ignore'] })) { try { [void][regex]::new([string]$rx) } catch { [void]$E.Add('grammar.json: "' + $name + '" ignore pattern "' + $rx + '" is not a valid regex') } }
    }
    # rules
    $rules = if ($Profile.Contains('rules') -and ($Profile['rules'] -is [System.Collections.IDictionary])) { $Profile['rules'] } else { @{} }
    foreach ($name in $rules.Keys) {
        $r = $rules[$name]
        if (-not $pages.Contains($name)) { [void]$E.Add('rules.json: "' + $name + '" is not a page in pages.json (keys are page names, P0-R1)') }
        if (-not ($r -is [System.Collections.IDictionary])) { [void]$E.Add('rules.json: "' + $name + '" is not an object'); continue }
        $default = if ($r.Contains('default')) { [string]$r['default'] } else { 'ok' }
        if (-not ($default -in @('ok', 'ng', 'unknown'))) { [void]$E.Add('rules.json: "' + $name + '" default "' + $default + '" must be ok, ng or unknown') }
        $i = 0
        foreach ($rule in @(if ($r.Contains('rules') -and $null -ne $r['rules']) { $r['rules'] })) {
            $i++
            if (-not ($rule -is [System.Collections.IDictionary])) { [void]$E.Add('rules.json: "' + $name + '" rule ' + $i + ' is not an object'); continue }
            $op = if ($rule.Contains('op')) { [string]$rule['op'] } else { '' }
            if (-not ((Get-EbiRuleOps) -contains $op)) { [void]$E.Add('rules.json: "' + $name + '" rule ' + $i + ' op "' + $op + '" is not one of the twelve') }
            $else = if ($rule.Contains('else')) { [string]$rule['else'] } else { '' }
            if (-not ($else -in @('ng', 'unknown'))) { [void]$E.Add('rules.json: "' + $name + '" rule ' + $i + ' else "' + $else + '" -- must be ng or unknown, never ok (5.2)') }
            if (-not $rule.Contains('field') -or [string]::IsNullOrWhiteSpace([string]$rule['field'])) { [void]$E.Add('rules.json: "' + $name + '" rule ' + $i + ' has no field') }
            elseif ($grammar.Contains($name) -and ($grammar[$name] -is [System.Collections.IDictionary]) -and $grammar[$name].Contains('fields')) {
                $fields = @($grammar[$name]['fields'] | ForEach-Object { [string]$_ })
                if ($fields.Count -gt 0 -and -not ($fields -contains [string]$rule['field'])) { [void]$W.Add('rules.json: "' + $name + '" rule ' + $i + ' field "' + [string]$rule['field'] + '" is not one of the grammar''s fields (' + ($fields -join ', ') + ')') }
            }
            if (-not $rule.Contains('message') -or [string]::IsNullOrWhiteSpace([string]$rule['message'])) { [void]$W.Add('rules.json: "' + $name + '" rule ' + $i + ' has no message for the operator (5.3)') }
            if ($op -in @('equals', 'notEquals', 'in', 'notIn', 'matches', 'within', 'gt', 'lt', 'gte', 'lte') -and -not $rule.Contains('value')) { [void]$E.Add('rules.json: "' + $name + '" rule ' + $i + ' (' + $op + ') needs a value') }
        }
    }
    foreach ($name in $pages.Keys) {
        if (-not $grammar.Contains($name)) { [void]$W.Add('pages.json: "' + $name + '" has no grammar.json entry') }
        if (-not $rules.Contains($name)) { [void]$W.Add('pages.json: "' + $name + '" has no rules.json entry') }
    }
    # worklist
    if ($Profile.Contains('worklist') -and ($Profile['worklist'] -is [System.Collections.IDictionary])) {
        $wl = $Profile['worklist']
        if (-not $wl.Contains('file') -or [string]::IsNullOrWhiteSpace([string]$wl['file'])) { [void]$E.Add('worklist.json: no file') }
        $keyCols = @(Get-EbiWorklistKeyColumns -Profile $Profile)
        if ($keyCols.Count -eq 0) { [void]$E.Add('worklist.json: key.columns is empty (at least one key column, 6.1)') }
        $cols = @(if ($wl.Contains('columns') -and $null -ne $wl['columns']) { $wl['columns'] })
        $names = New-Object System.Collections.ArrayList
        $roleKey = New-Object System.Collections.ArrayList
        $i = 0
        foreach ($c in $cols) {
            $i++
            if (-not ($c -is [System.Collections.IDictionary]) -or -not $c.Contains('name')) { [void]$E.Add('worklist.json: column ' + $i + ' has no name'); continue }
            $n = [string]$c['name']
            if ($names -contains $n) { [void]$E.Add('worklist.json: column "' + $n + '" is declared twice') }
            [void]$names.Add($n)
            if ($n -in @('key', 'keySafe', 'group')) { [void]$W.Add('worklist.json: column "' + $n + '" collides with the reserved {{item.' + $n + '}} (WORKFLOW-SCHEMA 4.1)') }
            $role = if ($c.Contains('role')) { [string]$c['role'] } else { '' }
            if (-not ((Get-EbiProfileColumnRoles) -contains $role)) { [void]$E.Add('worklist.json: column "' + $n + '" role "' + $role + '" is not one of ' + ((Get-EbiProfileColumnRoles) -join ', ')) }
            if ($role -eq 'key') { [void]$roleKey.Add($n) }
            if ($c.Contains('values')) {
                if ($role -ne 'verdict') { [void]$E.Add('worklist.json: column "' + $n + '" has values but is not a verdict column') }
                elseif ($c['values'] -is [System.Collections.IDictionary]) { foreach ($k in $c['values'].Keys) { if (-not ((Get-EbiProfileVerdicts) -contains [string]$k)) { [void]$E.Add('worklist.json: column "' + $n + '" values has "' + $k + '", not a verdict (ok / ng / unknown / pending)') } } }
                else { [void]$E.Add('worklist.json: column "' + $n + '" values is not a map') }
            }
            if ($role -eq 'bitmask' -and (-not $c.Contains('bits') -or -not ($c['bits'] -is [System.Collections.IDictionary]) -or $c['bits'].Count -eq 0)) { [void]$E.Add('worklist.json: bitmask column "' + $n + '" declares no bits') }
        }
        foreach ($k in $keyCols) { if (-not ($roleKey -contains $k)) { [void]$E.Add('worklist.json: key.columns names "' + $k + '" but no column declares role: key for it (6.1)') } }
        foreach ($k in $roleKey) { if (-not ($keyCols -contains $k)) { [void]$E.Add('worklist.json: column "' + $k + '" has role: key but is not in key.columns (6.1)') } }
        foreach ($rule in @(Get-EbiKeyRules -Profile $Profile)) {
            $kind = [string]$rule['kind']
            if (-not ($kind -in @('suffix', 'prefix', 'fullwidth', 'case-insensitive'))) { [void]$E.Add('worklist.json: confirmedRules kind "' + $kind + '" is unknown') }
            if ($kind -in @('suffix', 'prefix')) { if (-not $rule.Contains('pattern')) { [void]$E.Add('worklist.json: confirmedRules ' + $kind + ' rule has no pattern') } else { try { [void][regex]::new([string]$rule['pattern']) } catch { [void]$E.Add('worklist.json: confirmedRules pattern "' + [string]$rule['pattern'] + '" is not a valid regex') } } }
        }
        $groupCol = Get-EbiWorklistGroupColumn -Profile $Profile
        if ($groupCol -ne '' -and $names.Count -gt 0 -and -not ($names -contains $groupCol)) { [void]$W.Add('vocabulary.json: columns.group "' + $groupCol + '" is not a worklist column') }
    }
    return @{ ok = ($E.Count -eq 0); errors = $E.ToArray(); warnings = $W.ToArray() }
}

function Invoke-EbiFixtureCase {
    <#
      PURE. One fixture text through the page pipeline:
        fingerprint -> kind; if not ok: @{ page = kind }
        grammar -> records; key -> record (tieBreak newest by the page's
        timeField, inside the window first); rules -> verdict.
      -> @{ page; verdict; message; matchedRow; records; unrecognized;
            failure }  (failure: not_found | ambiguous | grammar_invalid ...)
    #>
    param([string]$Text, $Page, $Grammar, $Rules, [string]$Key, $TimeWindow = $null)
    $fp = if ($Page -is [System.Collections.IDictionary] -and $Page.Contains('fingerprint')) { $Page['fingerprint'] } else { @{} }
    $kind = (Get-EbiPageKind -Text $Text -Fingerprint $fp)['kind']
    $out = @{ page = $kind; verdict = ''; message = ''; matchedRow = 0; records = 0; unrecognized = 0; failure = '' }
    if ($kind -ne 'ok') { return $out }
    if ($null -eq $Grammar) { $out['failure'] = 'no_grammar'; return $out }
    $parsed = ConvertFrom-EbiGrammar -Text $Text -Grammar $Grammar
    if (-not $parsed['ok']) { $out['failure'] = 'grammar_invalid'; $out['message'] = $parsed['message']; return $out }
    $records = @($parsed['records']); $out['records'] = $records.Count; $out['unrecognized'] = @($parsed['unrecognized']).Count
    if ($records.Count -eq 0) { $out['failure'] = 'no_records'; return $out }
    $field = 'key'
    $values = @(foreach ($r in $records) { if ($r.Contains($field)) { [string]$r[$field] } else { '' } })
    $m = Find-EbiKeyMatches -Records $values -Key $Key -Rules (Get-EbiKeyRules -Profile @{})
    $idx = @($m['indexes'])
    if ($idx.Count -eq 0) { $out['failure'] = 'not_found'; return $out }
    $timeField = if ($Page -is [System.Collections.IDictionary] -and $Page.Contains('timeField')) { [string]$Page['timeField'] } else { 'time' }
    $hits = @(foreach ($i in $idx) { @{ index = ($i + 1); record = $records[$i] } })
    $pick = Select-EbiNewestRecord -Hits $hits -TieBreak 'newest' -TimeField $timeField -Window $TimeWindow
    $out['matchedRow'] = $pick['index']
    if ($null -eq $Rules) { $out['failure'] = 'no_rules'; return $out }
    $scope = New-EbiTemplateScope -Vars @{} -Profile @{} -PageName '' -Run @{ timeWindow = $TimeWindow } -Item $null -Steps @{}
    $rulesEval = Expand-EbiTemplate -Value $Rules -Scope $scope
    $r2 = if ($rulesEval['ok']) { $rulesEval['value'] } else { $Rules }
    $v = Test-EbiRuleTable -Rules $r2
    if (-not $v['ok']) { $out['failure'] = 'rules_invalid'; $out['message'] = $v['message']; return $out }
    $j = Invoke-EbiRuleTable -Record $pick['record'] -Rules $v['rules'] -Default $v['default']
    $out['verdict'] = $j['code']; $out['message'] = $j['message']
    return $out
}

function Invoke-EbiFixtureCheck {
    <#
      Every fixtures/<page>/expected.json entry -> a result line.
      Entry keys are file names; a key with '#suffix' reuses the file named
      by its 'file' field under another key / window. Fields: key, verdict,
      message, matchedRow, records, page, timeWindow.
      -> @{ ok; results = @(@{ page; name; ok; message }); passed; failed; skipped }
    #>
    param([string]$Dir, $Profile)
    $results = New-Object System.Collections.ArrayList
    $fxRoot = Join-Path $Dir 'fixtures'
    $passed = 0; $failed = 0; $skipped = 0
    if (-not (Test-Path -LiteralPath $fxRoot -PathType Container)) { return @{ ok = $true; results = @(); passed = 0; failed = 0; skipped = 0 } }
    $pages = if ($Profile.Contains('pages') -and ($Profile['pages'] -is [System.Collections.IDictionary])) { $Profile['pages'] } else { @{} }
    $grammar = if ($Profile.Contains('grammar') -and ($Profile['grammar'] -is [System.Collections.IDictionary])) { $Profile['grammar'] } else { @{} }
    $rules = if ($Profile.Contains('rules') -and ($Profile['rules'] -is [System.Collections.IDictionary])) { $Profile['rules'] } else { @{} }
    foreach ($pageDir in @(Get-ChildItem -LiteralPath $fxRoot -Directory | Sort-Object Name)) {
        $pageName = $pageDir.Name
        $expPath = Join-Path $pageDir.FullName 'expected.json'
        if (-not (Test-Path -LiteralPath $expPath)) { [void]$results.Add(@{ page = $pageName; name = 'expected.json'; ok = $false; message = 'missing' }); $failed++; continue }
        $exp = Read-EbiJson -Path $expPath
        if (-not $exp['ok']) { [void]$results.Add(@{ page = $pageName; name = 'expected.json'; ok = $false; message = $exp['message'] }); $failed++; continue }
        if (-not $pages.Contains($pageName)) { [void]$results.Add(@{ page = $pageName; name = '*'; ok = $false; message = 'no such page in pages.json' }); $failed++; continue }
        $listed = New-Object System.Collections.ArrayList
        foreach ($name in ($exp['value'].Keys | Sort-Object)) {
            $e = $exp['value'][$name]
            $file = if ($e -is [System.Collections.IDictionary] -and $e.Contains('file')) { [string]$e['file'] } else { ([string]$name -split '#')[0] }
            [void]$listed.Add($file)
            $path = Join-Path $pageDir.FullName $file
            if (-not (Test-Path -LiteralPath $path)) { [void]$results.Add(@{ page = $pageName; name = [string]$name; ok = $false; message = ('fixture file ' + $file + ' is missing') }); $failed++; continue }
            $text = [System.IO.File]::ReadAllText($path, (New-Object System.Text.UTF8Encoding($false)))
            $key = if ($e.Contains('key')) { [string]$e['key'] } else { '' }
            $win = if ($e.Contains('timeWindow')) { $e['timeWindow'] } else { $null }
            $got = Invoke-EbiFixtureCase -Text $text -Page $pages[$pageName] -Grammar $(if ($grammar.Contains($pageName)) { $grammar[$pageName] } else { $null }) -Rules $(if ($rules.Contains($pageName)) { $rules[$pageName] } else { $null }) -Key $key -TimeWindow $win
            $diffs = New-Object System.Collections.ArrayList
            if ($e.Contains('page')) { if ([string]$got['page'] -ne [string]$e['page']) { [void]$diffs.Add('page: expected ' + $e['page'] + ', got ' + $got['page']) } }
            elseif ($got['page'] -ne 'ok') { [void]$diffs.Add('page classified as ' + $got['page'] + ', so nothing was judged') }
            if ($e.Contains('verdict')) {
                $want = [string]$e['verdict']
                if ($want -eq 'not_found') { if ($got['failure'] -ne 'not_found') { [void]$diffs.Add('expected not_found, got ' + $(if ($got['failure'] -ne '') { $got['failure'] } else { 'verdict ' + $got['verdict'] })) } }
                elseif ([string]$got['verdict'] -ne $want) { [void]$diffs.Add('verdict: expected ' + $want + ', got ' + $(if ($got['failure'] -ne '') { $got['failure'] + ' ' + $got['message'] } else { $got['verdict'] + ' ' + $got['message'] })) }
            }
            if ($e.Contains('message') -and [string]$got['message'] -ne [string]$e['message']) { [void]$diffs.Add('message: expected "' + $e['message'] + '", got "' + $got['message'] + '"') }
            if ($e.Contains('matchedRow') -and [int]$got['matchedRow'] -ne [int]$e['matchedRow']) { [void]$diffs.Add('matchedRow: expected ' + $e['matchedRow'] + ', got ' + $got['matchedRow']) }
            if ($e.Contains('records') -and [int]$got['records'] -ne [int]$e['records']) { [void]$diffs.Add('records: expected ' + $e['records'] + ', got ' + $got['records']) }
            if ($diffs.Count -eq 0) { $passed++; [void]$results.Add(@{ page = $pageName; name = [string]$name; ok = $true; message = ('page=' + $got['page'] + $(if ($got['verdict'] -ne '') { ' verdict=' + $got['verdict'] + ' row=' + $got['matchedRow'] } else { '' }) + $(if ([int]$got['unrecognized'] -gt 0) { ' unrecognized=' + $got['unrecognized'] } else { '' })) }) }
            else { $failed++; [void]$results.Add(@{ page = $pageName; name = [string]$name; ok = $false; message = ($diffs.ToArray() -join '; ') }) }
        }
        foreach ($f in @(Get-ChildItem -LiteralPath $pageDir.FullName -Filter '*.txt' -File)) { if (-not ($listed -contains $f.Name)) { $skipped++; [void]$results.Add(@{ page = $pageName; name = $f.Name; ok = $true; message = 'not in expected.json (skipped)' }) } }
    }
    return @{ ok = ($failed -eq 0); results = $results.ToArray(); passed = $passed; failed = $failed; skipped = $skipped }
}

function Get-EbiProfileLeaves {
    # PURE. Nested hashtable -> map 'a.b.c' -> value (arrays as JSON text).
    param($Value, [string]$Prefix = '', $Into = $null)
    if ($null -eq $Into) { $Into = @{} }
    if ($Value -is [System.Collections.IDictionary]) {
        if ($Value.Count -eq 0) { $Into[$Prefix] = '{}' }
        foreach ($k in $Value.Keys) { [void](Get-EbiProfileLeaves -Value $Value[$k] -Prefix $(if ($Prefix -eq '') { [string]$k } else { $Prefix + '.' + [string]$k }) -Into $Into) }
    } else {
        $Into[$Prefix] = $(if ($null -eq $Value) { 'null' } elseif ($Value -is [System.Collections.IList]) { ConvertTo-EbiJson -Value $Value -Compress } else { [string]$Value })
    }
    return $Into
}

function Compare-EbiProfile {
    # PURE. Two loaded profiles -> @{ lines; same; changed; missing; extra }.
    # 'missing' = in A not B (still to write), 'extra' = in B not A.
    param($A, $B, [string]$NameA = 'a', [string]$NameB = 'b')
    $la = Get-EbiProfileLeaves -Value $A; $lb = Get-EbiProfileLeaves -Value $B
    $L = New-Object System.Collections.ArrayList
    $same = 0; $changed = New-Object System.Collections.ArrayList; $missing = New-Object System.Collections.ArrayList; $extra = New-Object System.Collections.ArrayList
    foreach ($k in ($la.Keys | Sort-Object)) {
        if (-not $lb.Contains($k)) { [void]$missing.Add($k); [void]$L.Add('- ' + $k + ' = ' + $la[$k] + '   (only in ' + $NameA + ')'); continue }
        if ([string]$la[$k] -eq [string]$lb[$k]) { $same++; continue }
        [void]$changed.Add($k); [void]$L.Add('~ ' + $k + ' : ' + $la[$k] + ' -> ' + $lb[$k])
    }
    foreach ($k in ($lb.Keys | Sort-Object)) { if (-not $la.Contains($k)) { [void]$extra.Add($k); [void]$L.Add('+ ' + $k + ' = ' + $lb[$k] + '   (only in ' + $NameB + ')') } }
    [void]$L.Add(('{0} same, {1} changed, {2} only in {3}, {4} only in {5}' -f $same, $changed.Count, $missing.Count, $NameA, $extra.Count, $NameB))
    return @{ lines = $L.ToArray(); same = $same; changed = $changed.ToArray(); missing = $missing.ToArray(); extra = $extra.ToArray() }
}

function New-EbiProfileSkeleton {
    # Write profiles/<name>/ with every file's shape and a _doc key per field.
    param([string]$Dir, [string]$Name)
    if (Test-Path -LiteralPath $Dir) { return @{ ok = $false; message = ('already exists: ' + $Dir) } }
    New-Item -ItemType Directory -Path (Join-Path $Dir 'fixtures') -Force | Out-Null
    $files = @{
        'vocabulary.json' = @{ _doc = 'display names only; roles are exactly the five; columns maps neutral names to CSV column names (no key here, PROFILE-SCHEMA 2)'; sides = @{ before = 'before'; after = 'after' }; roles = @{ entry = 'entry'; form = 'form'; record = 'record'; list = 'list'; document = 'document' }; columns = @{ group = 'GROUP_COLUMN'; owner = 'OWNER_COLUMN'; deliverable = 'DELIVERABLE_COLUMN' } }
        'pages.json' = @{ _doc = 'one entry per page (a named page instance), keyed by page name, never by role (PROFILE-SCHEMA 3)'; examplePage = @{ role = 'list'; label = 'shown to people'; url = ''; openHint = 'what to open by hand'; fingerprint = @{ ok = @('a string only this page has'); loading = @(); empty = @(); expired = @('login') }; timeoutSec = 12; pollMs = 800; crop = @{ left = 6; top = 6; right = 6; bottom = 6 }; timeField = 'time' } }
        'grammar.json' = @{ _doc = 'page name -> parser (delimited | labeled | columns | regex); tune with ebi grammar tune (PROFILE-SCHEMA 4)'; examplePage = @{ parser = 'delimited'; delimiter = "`t"; rowWhen = @{ field = 0; matches = '^\d+$' }; fields = @('no', 'key', 'status', 'time'); ignore = @() } }
        'rules.json' = @{ _doc = 'page name -> { rules: [ { field, op, value, else (ng|unknown, never ok), message } ], default }'; examplePage = @{ rules = @(@{ field = 'status'; op = 'equals'; value = 'OK'; else = 'ng'; message = 'status is not OK' }); default = 'ok' } }
        'worklist.json' = @{ _doc = 'the CSV schema: key.columns (composite allowed), confirmedRules, columns with roles; verdict columns may declare values (PROFILE-SCHEMA 6)'; file = 'worklist.csv'; encoding = 'utf8-bom'; key = @{ columns = @('KEY_COLUMN'); confirmedRules = @(@{ kind = 'fullwidth' }, @{ kind = 'case-insensitive' }); ambiguityPolicy = 'listAndAsk' }; columns = @(@{ name = 'KEY_COLUMN'; role = 'key' }, @{ name = 'GROUP_COLUMN'; role = 'group' }, @{ name = 'before_examplePage'; role = 'verdict'; default = '' }) }
        'layout.json' = @{ _doc = 'only for compose / annotate workflows; numbers come from ebi probe on the office PC (PROFILE-SCHEMA 7)'; sheets = @{ before = 'before'; after = 'after' }; anchor = @{ column = 'A'; matchesKey = $true }; pictures = @{}; boxes = @{} }
    }
    foreach ($f in $files.Keys) { $w = Write-EbiJson -Path (Join-Path $Dir $f) -Value $files[$f]; if (-not $w['ok']) { return @{ ok = $false; message = $w['message'] } } }
    [System.IO.File]::WriteAllText((Join-Path $Dir 'README.md'), ('# profiles/' + $Name + "`n`nSkeleton from `ebi profile new`. Every JSON file carries a _doc key explaining its shape; delete the example entries once real ones exist, then `ebi profile check " + $Name + "`.`n"), (New-Object System.Text.UTF8Encoding($false)))
    return @{ ok = $true; message = ''; dir = $Dir }
}
