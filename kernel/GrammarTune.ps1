#Requires -Version 5.1
# ============================================================
#  kernel/GrammarTune.ps1
#
#  `ebi grammar tune` (P2-02, PROFILE-SCHEMA.md 4.1): the interactive
#  parser debugger. Feed it one page text; it parses with the page's
#  grammar, renders the records as an ASCII table and LISTS the lines it
#  did not recognise; d / r / c / p / i change one parameter and re-parse
#  at once; s writes the grammar into the profile's grammar.json AND saves
#  the text as fixtures/<page>/<name>.txt with an expected.json entry
#  (after the mask gate), so every tuning session leaves a regression test.
#
#  Dot-source only (no param(), ASCII source). The rendering and the
#  editing steps are pure (Format-EbiTuneView, Edit-EbiTuneGrammar); the
#  loop (Invoke-EbiGrammarTune) reads answers through kernel/Gate.ps1's
#  reader so a test can script it.
# ============================================================

. (Join-Path $PSScriptRoot 'Json.ps1')
. (Join-Path $PSScriptRoot 'Parse.ps1')
. (Join-Path $PSScriptRoot 'Gate.ps1')
. (Join-Path $PSScriptRoot 'Mask.ps1')
. (Join-Path $PSScriptRoot 'Profile.ps1')

function ConvertTo-EbiTuneCell {
    param([string]$Text, [int]$Width)
    $t = if ($null -eq $Text) { '' } else { $Text -replace '[\r\n\t]', ' ' }
    if ($t.Length -gt $Width) { $t = $t.Substring(0, [Math]::Max(0, $Width - 2)) + '..' }
    return $t.PadRight($Width)
}

function Format-EbiTuneView {
    <#
      PURE. Grammar + parse result -> lines: the grammar summary, up to
      Rows records as a table (fields in grammar order, else sorted), and
      the unrecognised lines, every one.
    #>
    param($Grammar, $Parsed, [int]$Rows = 5, [int]$Width = 100)
    $L = New-Object System.Collections.ArrayList
    $parser = if ($Grammar.Contains('parser')) { [string]$Grammar['parser'] } else { '?' }
    $summary = New-Object System.Collections.ArrayList
    foreach ($k in @('delimiter', 'pattern', 'lastNonEmpty')) { if ($Grammar.Contains($k) -and $null -ne $Grammar[$k]) { [void]$summary.Add($k + '=' + (ConvertTo-EbiJson -Value ([string]$Grammar[$k]) -Compress)) } }
    if ($Grammar.Contains('rowWhen') -and ($Grammar['rowWhen'] -is [System.Collections.IDictionary])) { [void]$summary.Add('rowWhen=field[' + [string]$Grammar['rowWhen']['field'] + '] matches ' + [string]$Grammar['rowWhen']['matches']) }
    if ($Grammar.Contains('headerLine') -and ($Grammar['headerLine'] -is [System.Collections.IDictionary]) -and $Grammar['headerLine'].Contains('contains')) { [void]$summary.Add('header contains ' + (@($Grammar['headerLine']['contains']) -join ' & ')) }
    if ($Grammar.Contains('ignore') -and @($Grammar['ignore']).Count -gt 0) { [void]$summary.Add('ignore=' + @($Grammar['ignore']).Count + ' pattern(s)') }
    [void]$L.Add('grammar: ' + $parser + $(if ($summary.Count) { ', ' + ($summary.ToArray() -join ', ') } else { '' }))
    if (-not $Parsed['ok']) { [void]$L.Add('  ! ' + $Parsed['message']); return $L.ToArray() }
    $records = @($Parsed['records'])
    $fields = New-Object System.Collections.ArrayList
    if ($Grammar.Contains('fields') -and $null -ne $Grammar['fields']) { foreach ($f in @($Grammar['fields'])) { [void]$fields.Add([string]$f) } }
    elseif ($Grammar.Contains('columns') -and ($Grammar['columns'] -is [System.Collections.IDictionary])) { foreach ($f in ($Grammar['columns'].Keys | Sort-Object)) { [void]$fields.Add([string]$f) } }
    elseif ($Grammar.Contains('pairs') -and ($Grammar['pairs'] -is [System.Collections.IDictionary])) { foreach ($f in ($Grammar['pairs'].Keys | Sort-Object)) { [void]$fields.Add([string]$f) } }
    elseif ($records.Count -gt 0) { foreach ($f in ($records[0].Keys | Sort-Object)) { if ([string]$f -ne '_line') { [void]$fields.Add([string]$f) } } }
    [void]$L.Add(('records: {0} shown of {1};  unrecognised: {2}' -f [Math]::Min($Rows, $records.Count), $records.Count, @($Parsed['unrecognized']).Count))
    if ($fields.Count -gt 0) {
        $w = @{}
        foreach ($f in $fields) { $w[$f] = [Math]::Max(4, $f.Length) }
        $shown = @($records | Select-Object -First $Rows)
        foreach ($r in $shown) { foreach ($f in $fields) { $v = if ($r.Contains($f)) { [string]$r[$f] } else { '' }; if ($v.Length -gt $w[$f]) { $w[$f] = [Math]::Min(28, $v.Length) } } }
        $line = '  | ' + (($fields | ForEach-Object { ConvertTo-EbiTuneCell -Text $_ -Width $w[$_] }) -join ' | ') + ' |'
        [void]$L.Add($line)
        [void]$L.Add('  |-' + (($fields | ForEach-Object { '-' * $w[$_] }) -join '-|-') + '-|')
        foreach ($r in $shown) { [void]$L.Add('  | ' + (($fields | ForEach-Object { ConvertTo-EbiTuneCell -Text $(if ($r.Contains($_)) { [string]$r[$_] } else { '' }) -Width $w[$_] }) -join ' | ') + ' |') }
    }
    $un = @($Parsed['unrecognized'])
    if ($un.Count -gt 0) {
        [void]$L.Add('  unrecognised lines (every one -- nothing is dropped in silence):')
        foreach ($u in $un) { [void]$L.Add(('    {0,4}: {1}' -f $u['line'], (ConvertTo-EbiTuneCell -Text ([string]$u['text']) -Width ($Width - 12)).TrimEnd())) }
    }
    if (@($Parsed['missing']).Count -gt 0) { [void]$L.Add('  labels not found: ' + (@($Parsed['missing']) -join ', ')) }
    return $L.ToArray()
}

function Edit-EbiTuneGrammar {
    <#
      PURE. One edit command on a grammar -> the new grammar (a copy).
        d <delimiter>       \t or ws or a literal
        r <field> <regex>   rowWhen (delimited / columns)
        c a,b,c             field names (delimited) / rename columns keys
        p <parser>          switch parser (keeps what carries over)
        i <regex>           add an ignore pattern
        x <regex>           regex parser's pattern
        l <field>           lastNonEmpty
        h a & b             header line must contain (columns parser)
    #>
    param($Grammar, [string]$Command)
    $g = @{}
    foreach ($k in $Grammar.Keys) { $g[$k] = $Grammar[$k] }
    $m = [regex]::Match($Command.Trim(), '^([a-z])\s*(.*)$')
    if (-not $m.Success) { return @{ ok = $false; grammar = $g; message = 'not an edit' } }
    $cmd = $m.Groups[1].Value; $arg = $m.Groups[2].Value.Trim()
    switch ($cmd) {
        'd' { $g['delimiter'] = $(if ($arg -eq '\t') { "`t" } else { $arg }); if (-not $g.Contains('parser')) { $g['parser'] = 'delimited' } }
        'r' { $mm = [regex]::Match($arg, '^(\S+)\s+(.+)$'); if (-not $mm.Success) { return @{ ok = $false; grammar = $g; message = 'r <field> <regex>' } }; $f = $mm.Groups[1].Value; $g['rowWhen'] = @{ field = $(if ($f -match '^\d+$') { [int]$f } else { $f }); matches = $mm.Groups[2].Value } }
        'c' { $g['fields'] = @($arg -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) }
        'p' { if (-not ($arg -in @('delimited', 'labeled', 'columns', 'regex'))) { return @{ ok = $false; grammar = $g; message = 'p delimited|labeled|columns|regex' } }; $g['parser'] = $arg }
        'i' { $g['ignore'] = @(@(if ($g.Contains('ignore') -and $null -ne $g['ignore']) { $g['ignore'] }) + @($arg)) }
        'x' { $g['pattern'] = $arg; if (-not $g.Contains('parser')) { $g['parser'] = 'regex' } }
        'l' { $g['lastNonEmpty'] = $arg }
        'h' { $g['headerLine'] = @{ contains = @($arg -split '&' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) } }
        default { return @{ ok = $false; grammar = $g; message = ('unknown edit ' + $cmd) } }
    }
    try { [void](ConvertFrom-EbiGrammar -Text 'probe' -Grammar $g) } catch { return @{ ok = $false; grammar = $Grammar; message = $_.Exception.Message } }
    return @{ ok = $true; grammar = $g; message = '' }
}

function Save-EbiTuneResult {
    <#
      s: grammar.json[<page>] = grammar, text -> fixtures/<page>/<name>.txt,
      expected.json[<name>.txt] = { records = <count> } (a verdict is added
      by hand later). The text passes the mask gate first, or nothing is
      written. -> @{ ok; message; grammarPath; fixturePath }
    #>
    param([string]$ProfileDir, [string]$Page, $Grammar, [string]$Text, [string]$FixtureName, [int]$RecordCount)
    $hits = @(Find-EbiMaskHitsInText -Text $Text)
    if ($hits.Count -gt 0) { return @{ ok = $false; message = ('the text has ' + $hits.Count + ' sensitive item(s) (' + (($hits | ForEach-Object { $_['rule'] } | Sort-Object -Unique) -join ', ') + '); mask them before saving a fixture'); grammarPath = ''; fixturePath = '' } }
    $gp = Join-Path $ProfileDir 'grammar.json'
    $all = @{}
    if (Test-Path -LiteralPath $gp) { $r = Read-EbiJson -Path $gp; if ($r['ok'] -and ($r['value'] -is [System.Collections.IDictionary])) { $all = $r['value'] } }
    $all[$Page] = $Grammar
    $w = Write-EbiJson -Path $gp -Value $all
    if (-not $w['ok']) { return @{ ok = $false; message = $w['message']; grammarPath = $gp; fixturePath = '' } }
    $fxDir = Join-Path (Join-Path $ProfileDir 'fixtures') $Page
    if (-not (Test-Path -LiteralPath $fxDir)) { New-Item -ItemType Directory -Path $fxDir -Force | Out-Null }
    $name = if ([string]::IsNullOrWhiteSpace($FixtureName)) { 'tuned-' + (Get-Date).ToString('yyyyMMdd-HHmmss') + '.txt' } elseif ($FixtureName.EndsWith('.txt')) { $FixtureName } else { $FixtureName + '.txt' }
    $fp = Join-Path $fxDir $name
    [System.IO.File]::WriteAllText($fp, $Text, (New-Object System.Text.UTF8Encoding($false)))
    $ep = Join-Path $fxDir 'expected.json'
    $exp = @{}
    if (Test-Path -LiteralPath $ep) { $r = Read-EbiJson -Path $ep; if ($r['ok'] -and ($r['value'] -is [System.Collections.IDictionary])) { $exp = $r['value'] } }
    if (-not $exp.Contains($name)) { $exp[$name] = @{ records = $RecordCount; _todo = 'add key + verdict (ok / ng / unknown) once judged by hand' } }
    $w = Write-EbiJson -Path $ep -Value $exp
    if (-not $w['ok']) { return @{ ok = $false; message = $w['message']; grammarPath = $gp; fixturePath = $fp } }
    return @{ ok = $true; message = ''; grammarPath = $gp; fixturePath = $fp }
}

function Invoke-EbiGrammarTune {
    <#
      The loop. Reader is a scriptblock returning the next typed line
      (tests); default Read-Host. -> @{ ok; grammar; saved; quit }
    #>
    param([string]$Text, [string]$ProfileDir, [string]$Page, $Grammar = $null, [scriptblock]$Reader = $null, [int]$Rows = 5)
    $g = if ($null -ne $Grammar -and ($Grammar -is [System.Collections.IDictionary])) { $Grammar } else { @{ parser = 'delimited'; delimiter = "`t"; fields = @('key') } }
    $read = if ($null -ne $Reader) { $Reader } else { { Read-Host } }
    $saved = $false
    $noConsole = $false
    if ($null -eq $Reader) { try { if ([Console]::IsInputRedirected) { $noConsole = $true } } catch { } }
    while ($true) {
        $parsed = ConvertFrom-EbiGrammar -Text $Text -Grammar $g
        Write-Host ''
        foreach ($l in @(Format-EbiTuneView -Grammar $g -Parsed $parsed -Rows $Rows)) { Write-Host ('  ' + $l) -ForegroundColor $(if ($l -like '*unrecognised lines*' -or $l -like '  ! *') { 'Yellow' } else { 'Gray' }) }
        Write-Host '  d=delimiter  r=<field> <regex> row rule  c=a,b,c fields  p=parser  i=ignore regex  x=regex pattern  l=lastNonEmpty  h=header  u=show all unrecognised  s=save (profile + fixture)  q=quit' -ForegroundColor DarkGray
        if ($noConsole) { Write-Host '  (no console: nothing changed)' -ForegroundColor DarkGray; return @{ ok = $true; grammar = $g; saved = $false; quit = $true } }
        Write-Host '  > ' -ForegroundColor Magenta -NoNewline
        $typed = [string](& $read)
        $t = if ($null -eq $typed) { '' } else { $typed.Trim() }
        if ($t -eq '' ) { continue }
        if ($t -eq 'q') { return @{ ok = $true; grammar = $g; saved = $saved; quit = $true } }
        if ($t -eq 'u') { foreach ($u in @($parsed['unrecognized'])) { Write-Host ('    {0,4}: {1}' -f $u['line'], $u['text']) -ForegroundColor Yellow }; continue }
        if ($t -match '^s(?:\s+(\S+))?$') {
            $name = if ($Matches.Count -gt 1 -and $null -ne $Matches[1]) { $Matches[1] } else { '' }
            $s = Save-EbiTuneResult -ProfileDir $ProfileDir -Page $Page -Grammar $g -Text $Text -FixtureName $name -RecordCount @($parsed['records']).Count
            if ($s['ok']) { $saved = $true; Write-Host ('  saved ' + $s['grammarPath'] + ' and ' + $s['fixturePath']) -ForegroundColor Green } else { Write-Host ('  not saved: ' + $s['message']) -ForegroundColor Red }
            continue
        }
        $e = Edit-EbiTuneGrammar -Grammar $g -Command $t
        if ($e['ok']) { $g = $e['grammar'] } else { Write-Host ('  ' + $e['message']) -ForegroundColor DarkYellow }
    }
}
