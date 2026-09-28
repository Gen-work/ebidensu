#Requires -Version 5.1
# ============================================================
#  kernel/Parse.ps1
#
#  Page text -> records, the four grammars of PROFILE-SCHEMA.md section 4
#  (P1-30 / P1-31), plus the time parsing every consumer shares. Dot-source
#  only (no param(), ASCII source, no class). Pure: no I/O.
#
#    delimited  tab / whitespace separated table; a data row is one whose
#               rowWhen field matches a regex (ConvertFrom-GfixJobListText)
#    labeled    "label: value" detail page, one record (ConvertFrom-HmPageText)
#    columns    fixed-width table under a header line (ConvertFrom-JenkinsListText)
#    regex      one pattern per line, named groups are the fields
#
#  The one rule that is not negotiable (section 4.1 point 2): a line the
#  grammar did not recognise is REPORTED, never dropped in silence. Every
#  parser returns @{ ok; records; unrecognized = @(@{ line; text }) } and
#  the step turns those into the standard warnings channel. A grammar may
#  declare `ignore` (regex list) for lines that are expected noise --
#  headers, footers, blank rulers -- so the report stays about the lines
#  that matter.
#
#  Time: ConvertTo-EbiDateTime accepts single-digit hours (H:mm:ss). The
#  old parsers required \d{2}: and silently lost every row before 10:00
#  (section 4.2). The default format list is the one place that knowledge
#  lives.
# ============================================================

function Get-EbiDateTimeFormats {
    # Every shape the project's pages have shown so far, single-digit
    # hour first. A grammar may add its own via 'timeFormats'.
    return @(
        'yyyy/M/d H:mm:ss', 'yyyy/MM/dd H:mm:ss', 'yyyy/M/d H:mm', 'yyyy/MM/dd H:mm',
        'yyyy-M-d H:mm:ss', 'yyyy-MM-dd H:mm:ss', 'yyyy-M-d H:mm', 'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-ddTHH:mm:sszzz',
        'yyyy/M/d', 'yyyy-M-d', 'yyyyMMddHHmmss', 'yyyyMMdd', 'H:mm:ss', 'H:mm'
    )
}

function ConvertTo-EbiDateTime {
    <#
      PURE. Text -> @{ ok; value ([datetime]); format }. Whitespace inside
      is collapsed to one space first (OCR and copy-paste both inject it).
      A time without a date is placed on Date (today by default) so it can
      be compared with a window on the same day.
    #>
    param([string]$Text, $Formats = $null, [datetime]$Date = [datetime]::Today)
    $t = if ($null -eq $Text) { '' } else { ([regex]::Replace($Text.Trim(), '\s+', ' ')) }
    if ($t -eq '') { return @{ ok = $false; value = $null; format = '' } }
    $fmts = @(if ($null -ne $Formats) { @($Formats) }) + @(Get-EbiDateTimeFormats)
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    foreach ($f in $fmts) {
        $fs = [string]$f
        if ($fs -eq '') { continue }
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact($t, $fs, $culture, [System.Globalization.DateTimeStyles]::AllowWhiteSpaces, [ref]$parsed)) {
            if ($fs -match '^H' ) { $parsed = $Date.Date.Add($parsed.TimeOfDay) }
            return @{ ok = $true; value = $parsed; format = $fs }
        }
    }
    return @{ ok = $false; value = $null; format = '' }
}

function Get-EbiGrammarLines {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return @() }
    return @([regex]::Split($Text, "\r?\n"))
}

function Test-EbiGrammarIgnored {
    param([string]$Line, $Grammar)
    if ($null -eq $Grammar -or -not ($Grammar -is [System.Collections.IDictionary]) -or -not $Grammar.Contains('ignore') -or $null -eq $Grammar['ignore']) { return $false }
    foreach ($p in @($Grammar['ignore'])) { if ([string]$p -ne '' -and $Line -match [string]$p) { return $true } }
    return $false
}

function ConvertFrom-EbiDelimitedText {
    param([string]$Text, $Grammar)
    $records = New-Object System.Collections.ArrayList
    $unrec = New-Object System.Collections.ArrayList
    $delim = if ($Grammar.Contains('delimiter') -and $null -ne $Grammar['delimiter']) { [string]$Grammar['delimiter'] } else { "`t" }
    $fields = @(if ($Grammar.Contains('fields')) { @($Grammar['fields']) | ForEach-Object { [string]$_ } })
    if ($fields.Count -eq 0) { return @{ ok = $false; message = 'delimited grammar needs "fields"'; records = @(); unrecognized = @() } }
    $rowWhen = if ($Grammar.Contains('rowWhen') -and ($Grammar['rowWhen'] -is [System.Collections.IDictionary])) { $Grammar['rowWhen'] } else { $null }
    $rwField = 0; $rwRegex = ''
    if ($null -ne $rowWhen) {
        if ($rowWhen.Contains('field')) { $f = $rowWhen['field']; if ($f -is [string] -and $f -notmatch '^\d+$') { $rwField = [array]::IndexOf($fields, [string]$f) } else { $rwField = [int]$f } }
        if ($rowWhen.Contains('matches')) { $rwRegex = [string]$rowWhen['matches'] }
    }
    $n = 0
    foreach ($line in (Get-EbiGrammarLines -Text $Text)) {
        $n++
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if (Test-EbiGrammarIgnored -Line $line -Grammar $Grammar) { continue }
        $parts = @(if ($delim -eq 'ws' -or $delim -eq ' ') { [regex]::Split($line.Trim(), '\s+') } else { $line -split [regex]::Escape($delim) })   # @(if) -- a one-field line must stay an array
        $ok = $true
        if ($rwRegex -ne '') {
            $probe = if ($rwField -ge 0 -and $rwField -lt $parts.Count) { ([string]$parts[$rwField]).Trim() } else { '' }
            if ($probe -notmatch $rwRegex) { $ok = $false }
        }
        if ($ok -and $parts.Count -lt $fields.Count) { $ok = $false }
        if (-not $ok) { [void]$unrec.Add(@{ line = $n; text = $line }); continue }
        $rec = @{}
        for ($i = 0; $i -lt $fields.Count; $i++) { $rec[$fields[$i]] = ([string]$parts[$i]).Trim() }
        # lastNonEmpty: a page whose rows sometimes carry an extra empty cell
        # (the HM abend row) keeps its key as the LAST non-empty cell, the
        # way ConvertFrom-HmPageText read it; the field named here takes it.
        if ($Grammar.Contains('lastNonEmpty') -and -not [string]::IsNullOrWhiteSpace([string]$Grammar['lastNonEmpty'])) {
            for ($i = $parts.Count - 1; $i -ge 0; $i--) { if (([string]$parts[$i]).Trim() -ne '') { $rec[[string]$Grammar['lastNonEmpty']] = ([string]$parts[$i]).Trim(); break } }
        }
        $rec['_line'] = $n
        [void]$records.Add($rec)
    }
    return @{ ok = $true; message = ''; records = $records.ToArray(); unrecognized = $unrec.ToArray() }
}

function ConvertFrom-EbiLabeledText {
    # pairs: name -> @{ after = '<label>'; take = 'line' | 'token' }. One record.
    param([string]$Text, $Grammar)
    $unrec = New-Object System.Collections.ArrayList
    $pairs = if ($Grammar.Contains('pairs') -and ($Grammar['pairs'] -is [System.Collections.IDictionary])) { $Grammar['pairs'] } else { $null }
    if ($null -eq $pairs -or $pairs.Count -eq 0) { return @{ ok = $false; message = 'labeled grammar needs "pairs"'; records = @(); unrecognized = @() } }
    $lines = @(Get-EbiGrammarLines -Text $Text)
    $rec = @{}
    $missing = New-Object System.Collections.ArrayList
    $usedLines = @{}
    foreach ($name in ($pairs.Keys | Sort-Object)) {
        $spec = $pairs[$name]
        $after = if ($spec -is [System.Collections.IDictionary] -and $spec.Contains('after')) { [string]$spec['after'] } else { '' }
        $take = if ($spec -is [System.Collections.IDictionary] -and $spec.Contains('take')) { [string]$spec['take'] } else { 'line' }
        $found = $false
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            $pos = if ($after -ne '') { $line.IndexOf($after, [System.StringComparison]::Ordinal) } else { -1 }
            if ($pos -lt 0) { continue }
            $rest = $line.Substring($pos + $after.Length).TrimStart(':', ' ', "`t", [char]0xFF1A, [char]0x3000)
            if ($rest -eq '' -and ($i + 1) -lt $lines.Count -and $take -ne 'token') { $rest = $lines[$i + 1].Trim(); $usedLines[$i + 1] = $true }
            $val = if ($take -eq 'token') { (@([regex]::Split($rest.Trim(), '\s+'))[0]) } else { $rest.Trim() }
            $rec[[string]$name] = $val
            $usedLines[$i] = $true
            $found = $true
            break
        }
        if (-not $found) { $rec[[string]$name] = ''; [void]$missing.Add([string]$name) }
    }
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($usedLines.Contains($i) -or [string]::IsNullOrWhiteSpace($lines[$i]) -or (Test-EbiGrammarIgnored -Line $lines[$i] -Grammar $Grammar)) { continue }
        [void]$unrec.Add(@{ line = ($i + 1); text = $lines[$i] })
    }
    $rec['_line'] = 1
    return @{ ok = $true; message = ''; records = @($rec); unrecognized = $unrec.ToArray(); missing = $missing.ToArray() }
}

function ConvertFrom-EbiColumnsText {
    # headerLine: @{ contains = @(...) }; columns: name -> @(start, end) character offsets.
    param([string]$Text, $Grammar)
    $records = New-Object System.Collections.ArrayList
    $unrec = New-Object System.Collections.ArrayList
    $cols = if ($Grammar.Contains('columns') -and ($Grammar['columns'] -is [System.Collections.IDictionary])) { $Grammar['columns'] } else { $null }
    if ($null -eq $cols -or $cols.Count -eq 0) { return @{ ok = $false; message = 'columns grammar needs "columns"'; records = @(); unrecognized = @() } }
    $must = @(if ($Grammar.Contains('headerLine') -and ($Grammar['headerLine'] -is [System.Collections.IDictionary]) -and $Grammar['headerLine'].Contains('contains')) { @($Grammar['headerLine']['contains']) | ForEach-Object { [string]$_ } })
    $lines = @(Get-EbiGrammarLines -Text $Text)
    $start = 0
    if ($must.Count -gt 0) {
        $start = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $all = $true
            foreach ($m in $must) { if ($lines[$i].IndexOf($m, [System.StringComparison]::Ordinal) -lt 0) { $all = $false; break } }
            if ($all) { $start = $i + 1; break }
        }
        if ($start -lt 0) { return @{ ok = $false; message = ('header line not found (must contain: ' + ($must -join ', ') + ')'); records = @(); unrecognized = @() } }
    }
    $rowWhen = if ($Grammar.Contains('rowWhen') -and ($Grammar['rowWhen'] -is [System.Collections.IDictionary]) -and $Grammar['rowWhen'].Contains('matches')) { [string]$Grammar['rowWhen']['matches'] } else { '' }
    $minLen = 0
    foreach ($k in $cols.Keys) { $r = @($cols[$k]); if ($r.Count -ge 1 -and [int]$r[0] -gt $minLen) { $minLen = [int]$r[0] } }
    for ($i = $start; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if (Test-EbiGrammarIgnored -Line $line -Grammar $Grammar) { continue }
        if ($line.Length -le $minLen -or ($rowWhen -ne '' -and $line -notmatch $rowWhen)) { [void]$unrec.Add(@{ line = ($i + 1); text = $line }); continue }
        $rec = @{}
        foreach ($k in $cols.Keys) {
            $r = @($cols[$k]); $s = [int]$r[0]; $e = if ($r.Count -gt 1 -and $null -ne $r[1]) { [int]$r[1] } else { $line.Length }
            if ($s -ge $line.Length) { $rec[[string]$k] = ''; continue }
            if ($e -gt $line.Length) { $e = $line.Length }
            $rec[[string]$k] = $line.Substring($s, $e - $s).Trim()
        }
        $rec['_line'] = $i + 1
        [void]$records.Add($rec)
    }
    return @{ ok = $true; message = ''; records = $records.ToArray(); unrecognized = $unrec.ToArray() }
}

function ConvertFrom-EbiRegexText {
    # pattern with named groups; every line that matches is a record.
    param([string]$Text, $Grammar)
    $records = New-Object System.Collections.ArrayList
    $unrec = New-Object System.Collections.ArrayList
    $pattern = if ($Grammar.Contains('pattern')) { [string]$Grammar['pattern'] } else { '' }
    if ($pattern -eq '') { return @{ ok = $false; message = 'regex grammar needs "pattern"'; records = @(); unrecognized = @() } }
    $rx = $null
    try { $rx = New-Object System.Text.RegularExpressions.Regex($pattern) } catch { return @{ ok = $false; message = ('bad pattern: ' + $_.Exception.Message); records = @(); unrecognized = @() } }
    $names = @($rx.GetGroupNames() | Where-Object { $_ -notmatch '^\d+$' })
    $n = 0
    foreach ($line in (Get-EbiGrammarLines -Text $Text)) {
        $n++
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if (Test-EbiGrammarIgnored -Line $line -Grammar $Grammar) { continue }
        $m = $rx.Match($line)
        if (-not $m.Success) { [void]$unrec.Add(@{ line = $n; text = $line }); continue }
        $rec = @{}
        foreach ($g in $names) { $rec[$g] = $m.Groups[$g].Value.Trim() }
        $rec['_line'] = $n
        [void]$records.Add($rec)
    }
    return @{ ok = $true; message = ''; records = $records.ToArray(); unrecognized = $unrec.ToArray() }
}

function ConvertFrom-EbiGrammar {
    <#
      PURE. Text + grammar -> @{ ok; message; parser; records; unrecognized;
      missing }. Records are hashtables field -> string plus '_line'.
    #>
    param([string]$Text, $Grammar)
    if ($null -eq $Grammar -or -not ($Grammar -is [System.Collections.IDictionary])) { return @{ ok = $false; message = 'grammar must be a map'; parser = ''; records = @(); unrecognized = @(); missing = @() } }
    $parser = if ($Grammar.Contains('parser')) { [string]$Grammar['parser'] } else { '' }
    $r = $null
    switch ($parser) {
        'delimited' { $r = ConvertFrom-EbiDelimitedText -Text $Text -Grammar $Grammar }
        'labeled'   { $r = ConvertFrom-EbiLabeledText -Text $Text -Grammar $Grammar }
        'columns'   { $r = ConvertFrom-EbiColumnsText -Text $Text -Grammar $Grammar }
        'regex'     { $r = ConvertFrom-EbiRegexText -Text $Text -Grammar $Grammar }
        default     { return @{ ok = $false; message = ('unknown parser "' + $parser + '" (delimited | labeled | columns | regex)'); parser = $parser; records = @(); unrecognized = @(); missing = @() } }
    }
    $missing = @(if ($r.Contains('missing')) { $r['missing'] })
    return @{ ok = $r['ok']; message = $r['message']; parser = $parser; records = @($r['records']); unrecognized = @($r['unrecognized']); missing = $missing }
}

# ============================================================
#  Page fingerprint (PROFILE-SCHEMA 3.1) -- browser.assert_page's classifier,
#  here so `ebi profile check` judges fixtures with the very same function.
# ============================================================

function Get-EbiFingerprintStrings {
    # PURE. The fingerprint entry for a kind as string[] (missing -> empty).
    param($Fingerprint, [string]$Kind)
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $Fingerprint -or -not ($Fingerprint -is [System.Collections.IDictionary]) -or -not $Fingerprint.Contains($Kind) -or $null -eq $Fingerprint[$Kind]) { return $out.ToArray() }
    $v = $Fingerprint[$Kind]
    if ($v -is [string]) { [void]$out.Add($v); return $out.ToArray() }
    foreach ($s in $v) { if ($null -ne $s -and [string]$s -ne '') { [void]$out.Add([string]$s) } }
    return $out.ToArray()
}

function Get-EbiPageKind {
    <#
      PURE. -> @{ kind; matched }. (P1-15; moved here in P2-09.)
        blank text                     -> loading (nothing arrived yet)
        any 'expired' string present   -> expired
        any 'empty' string present     -> empty
        any 'loading' string present   -> loading
        ALL 'ok' strings present       -> ok   (an empty ok list never matches)
        otherwise                      -> unknown
    #>
    param([string]$Text, $Fingerprint)
    $matched = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($Text)) { return @{ kind = 'loading'; matched = $matched.ToArray() } }
    foreach ($kind in @('expired', 'empty', 'loading')) {
        foreach ($s in @(Get-EbiFingerprintStrings -Fingerprint $Fingerprint -Kind $kind)) {
            if ($Text.IndexOf($s, [System.StringComparison]::Ordinal) -ge 0) { [void]$matched.Add($s); return @{ kind = $kind; matched = $matched.ToArray() } }
        }
    }
    $okList = @(Get-EbiFingerprintStrings -Fingerprint $Fingerprint -Kind 'ok')
    if ($okList.Count -gt 0) {
        $all = $true
        foreach ($s in $okList) { if ($Text.IndexOf($s, [System.StringComparison]::Ordinal) -ge 0) { [void]$matched.Add($s) } else { $all = $false } }
        if ($all) { return @{ kind = 'ok'; matched = $matched.ToArray() } }
    }
    return @{ kind = 'unknown'; matched = $matched.ToArray() }
}


# ============================================================
#  Tie-break among several records for one key (verify.match_record,
#  P1-32; here since P2-09 so the fixture runner picks the same row).
# ============================================================

function Select-EbiNewestRecord {
    <#
      PURE. Matched records (with their 1-based index) -> the chosen one
      under a tie-break. @{ index; record; reason }.
    #>
    param($Hits, [string]$TieBreak, [string]$TimeField, $Window)
    $hits = @($Hits)
    if ($hits.Count -eq 1) { return @{ index = $hits[0]['index']; record = $hits[0]['record']; reason = 'the only match' } }
    if ($TieBreak -eq 'first') { return @{ index = $hits[0]['index']; record = $hits[0]['record']; reason = 'first listed' } }
    $dated = New-Object System.Collections.ArrayList
    foreach ($h in $hits) {
        $rec = $h['record']
        $t = if ($rec.Contains($TimeField)) { ConvertTo-EbiDateTime -Text ([string]$rec[$TimeField]) } else { @{ ok = $false } }
        if ($t['ok']) { [void]$dated.Add(@{ index = $h['index']; record = $rec; time = $t['value'] }) }
    }
    if ($dated.Count -eq 0) { return @{ index = $hits[0]['index']; record = $hits[0]['record']; reason = 'no usable time on any match; first listed' } }
    $from = $null; $to = $null
    if ($null -ne $Window -and ($Window -is [System.Collections.IDictionary])) {
        if ($Window.Contains('from')) { $p = ConvertTo-EbiDateTime -Text ([string]$Window['from']); if ($p['ok']) { $from = $p['value'] } }
        if ($Window.Contains('to')) { $p = ConvertTo-EbiDateTime -Text ([string]$Window['to']); if ($p['ok']) { $to = $p['value'] } }
    }
    $pool = @($dated.ToArray())
    $reason = 'newest by ' + $TimeField
    if ($null -ne $from -or $null -ne $to) {
        $inside = @($pool | Where-Object { ($null -eq $from -or $_['time'] -ge $from) -and ($null -eq $to -or $_['time'] -le $to) })
        if ($inside.Count -gt 0) { $pool = $inside; $reason = 'newest inside the run window' }
    }
    $best = $pool[0]
    foreach ($d in $pool) { if ($d['time'] -gt $best['time']) { $best = $d } }
    return @{ index = $best['index']; record = $best['record']; reason = $reason }
}

