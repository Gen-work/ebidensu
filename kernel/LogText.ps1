#Requires -Version 5.1
# ============================================================
#  kernel/LogText.ps1
#
#  Pure text helpers for log-style evidence: the ONE place these live, so
#  the file.* / verify.* / excel.* steps that need them agree. Dot-source
#  only (no param(), ASCII source). Everything here is PURE except
#  Get-EbiCodePage (it registers the .NET Core code-page provider once).
#
#    Get-EbiCodePage            an Encoding by code page (932 works on
#                               PS 5.1 and 7)
#    ConvertFrom-EbiMixedBytes  bytes that are UTF-8 with stray CP932
#                               pairs (a Java log writing a full-width colon
#                               in SJIS) -> string, nothing dropped
#    Get-EbiTextLines           text -> lines (CRLF / LF), trailing empty
#                               line removed
#    Get-EbiLogBlocks           START <key> ... END <key> blocks, keyed
#    Find-EbiLineHits           which lines (0-based) match which patterns
#    Get-EbiTextCellUnits       display width: half-width 1, full-width 2
#    Get-EbiHighlightEndColumn  how many grid columns a line of text spans
#    ConvertTo-EbiClockText     an Excel time (fraction or "11:44:59.99")
#                               -> "HH:mm:ss", float artefacts rounded
#    Get-EbiTimeAround          a date + clock + minutes before/after ->
#                               @{ from; to } ISO strings
#    Select-EbiFilePairs        left files vs right files by line count and
#                               arrival order
#    Compare-EbiTextLines       two texts line by line (line ending blind)
#    ConvertTo-EbiColumnNumber / ConvertTo-EbiColumnLetter   A <-> 1
# ============================================================

function Get-EbiCodePage {
    # An Encoding by code page number. .NET Core (pwsh 7) needs the
    # code-pages provider registered before 932 exists; .NET Framework has
    # it built in, and the type below does not exist there (hence the try).
    param([int]$CodePage)
    try {
        $t = 'System.Text.CodePagesEncodingProvider' -as [type]
        if ($null -ne $t) { [System.Text.Encoding]::RegisterProvider($t::Instance) }
    } catch { }
    return [System.Text.Encoding]::GetEncoding($CodePage)
}

function Test-EbiCp932Lead {
    param([int]$B)
    return (($B -ge 0x81 -and $B -le 0x9F) -or ($B -ge 0xE0 -and $B -le 0xFC))
}

function Test-EbiCp932Trail {
    param([int]$B)
    return (($B -ge 0x40 -and $B -le 0x7E) -or ($B -ge 0x80 -and $B -le 0xFC))
}

function Get-EbiUtf8SequenceLength {
    # Length of a VALID UTF-8 sequence starting at $i, or 0 when the bytes
    # there are not one (overlong / truncated / bad continuation).
    param([byte[]]$Bytes, [int]$i)
    $b0 = [int]$Bytes[$i]
    if ($b0 -lt 0x80) { return 1 }
    $n = 0; $min = 0
    if ($b0 -ge 0xC2 -and $b0 -le 0xDF) { $n = 2 }
    elseif ($b0 -ge 0xE0 -and $b0 -le 0xEF) { $n = 3 }
    elseif ($b0 -ge 0xF0 -and $b0 -le 0xF4) { $n = 4 }
    else { return 0 }
    if ($i + $n -gt $Bytes.Length) { return 0 }
    for ($k = 1; $k -lt $n; $k++) {
        $c = [int]$Bytes[$i + $k]
        if ($c -lt 0x80 -or $c -gt 0xBF) { return 0 }
    }
    if ($n -eq 3) {
        $b1 = [int]$Bytes[$i + 1]
        if ($b0 -eq 0xE0 -and $b1 -lt 0xA0) { return 0 }
        if ($b0 -eq 0xED -and $b1 -gt 0x9F) { return 0 }   # surrogates
    }
    if ($n -eq 4) {
        $b1 = [int]$Bytes[$i + 1]
        if ($b0 -eq 0xF0 -and $b1 -lt 0x90) { return 0 }
        if ($b0 -eq 0xF4 -and $b1 -gt 0x8F) { return 0 }
    }
    return $n
}

function ConvertFrom-EbiMixedSegment {
    # The slow, byte-by-byte half of ConvertFrom-EbiMixedBytes for one range
    # of bytes that strict UTF-8 refused. @{ text; foreignPairs; badBytes }.
    param([byte[]]$Bytes, [int]$Start, [int]$End, $Utf8, $Sjis)
    $sb = New-Object System.Text.StringBuilder
    $i = $Start; $runStart = $Start
    $pairs = 0; $bad = 0
    while ($i -lt $End) {
        $n = Get-EbiUtf8SequenceLength -Bytes $Bytes -i $i
        if ($n -gt 0 -and ($i + $n) -le $End) { $i += $n; continue }
        if ($i -gt $runStart) { [void]$sb.Append($Utf8.GetString($Bytes, $runStart, $i - $runStart)) }
        $b0 = [int]$Bytes[$i]
        if ((Test-EbiCp932Lead $b0) -and ($i + 1) -lt $End -and (Test-EbiCp932Trail ([int]$Bytes[$i + 1]))) {
            [void]$sb.Append($Sjis.GetString($Bytes, $i, 2))
            $pairs++
            $i += 2
        } else {
            [void]$sb.Append([char]0xFFFD)
            $bad++
            $i += 1
        }
        $runStart = $i
    }
    if ($i -gt $runStart) { [void]$sb.Append($Utf8.GetString($Bytes, $runStart, $i - $runStart)) }
    return @{ text = $sb.ToString(); foreignPairs = $pairs; badBytes = $bad }
}

function ConvertFrom-EbiMixedBytes {
    <#
      PURE (but for Get-EbiCodePage). Decode bytes that are UTF-8 except for
      the odd CP932 pair -- the batch logs write their full-width colon
      ("update count<colon>1") in SJIS inside an otherwise UTF-8 file, so a
      plain UTF-8 read turns it into U+FFFD and Excel shows "_xDC81_F".
      Valid UTF-8 runs decode as UTF-8; an invalid byte that starts a valid
      CP932 pair decodes as that CP932 character; anything else becomes
      U+FFFD (never silently dropped). A UTF-8 BOM is skipped.
      Fast: the whole file is tried as strict UTF-8 first (one .NET call),
      then line by line, and only a line strict UTF-8 refuses is walked
      byte by byte (a 600 KB log has a few dozen such lines).
      Returns @{ text; foreignPairs; badBytes }.
    #>
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return @{ text = ''; foreignPairs = 0; badBytes = 0 } }
    $start = 0
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { $start = 3 }
    $strict = New-Object System.Text.UTF8Encoding($false, $true)
    try { return @{ text = $strict.GetString($Bytes, $start, $Bytes.Length - $start); foreignPairs = 0; badBytes = 0 } } catch { }
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $sjis = Get-EbiCodePage -CodePage 932
    $sb = New-Object System.Text.StringBuilder
    $pairs = 0; $bad = 0
    $pos = $start
    while ($pos -lt $Bytes.Length) {
        $nl = [Array]::IndexOf($Bytes, [byte]10, $pos)
        $end = if ($nl -lt 0) { $Bytes.Length } else { $nl + 1 }
        $ok = $true
        try { [void]$sb.Append($strict.GetString($Bytes, $pos, $end - $pos)) } catch { $ok = $false }
        if (-not $ok) {
            $seg = ConvertFrom-EbiMixedSegment -Bytes $Bytes -Start $pos -End $end -Utf8 $utf8 -Sjis $sjis
            [void]$sb.Append($seg['text'])
            $pairs += [int]$seg['foreignPairs']; $bad += [int]$seg['badBytes']
        }
        $pos = $end
    }
    return @{ text = $sb.ToString(); foreignPairs = $pairs; badBytes = $bad }
}

function Get-EbiTextLines {
    # PURE. Text -> string[]; CRLF and LF both split; a single trailing
    # empty line (the file's final newline) is not a line.
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return @() }
    $lines = [regex]::Split($Text, "\r?\n")
    if ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') {
        if ($lines.Count -eq 1) { return @() }
        $lines = $lines[0..($lines.Count - 2)]
    }
    return @($lines)
}

function Get-EbiLogBlocks {
    <#
      PURE. Cut a log into START/END blocks.
        StartPattern / EndPattern: regexes with a named group 'key'
          (e.g. '^(?<tag>\S+) START (?<key>\S+)\s*$').
        Keys: keep only blocks whose key is one of these (empty = all).
        AfterEnd: lines after the END line that still belong to the block
          (the date stamp the shell prints after END).
      A START with no END before the next START (or the end of the log)
      closes there and is flagged complete=$false -- reported, not dropped.
      Returns @{ blocks = @( @{ key; start; end; complete; lines } ...);
      missing = keys asked for but not found } (start/end are 0-based).
    #>
    param([string[]]$Lines, [string]$StartPattern, [string]$EndPattern, [string[]]$Keys = @(), [int]$AfterEnd = 0)
    $blocks = New-Object System.Collections.ArrayList
    $want = @{}
    foreach ($k in @($Keys)) { if (-not [string]::IsNullOrWhiteSpace($k)) { $want[[string]$k] = $true } }
    $rxS = New-Object System.Text.RegularExpressions.Regex($StartPattern)
    $rxE = New-Object System.Text.RegularExpressions.Regex($EndPattern)
    $n = @($Lines).Count
    $i = 0
    while ($i -lt $n) {
        $m = $rxS.Match([string]$Lines[$i])
        if (-not $m.Success) { $i++; continue }
        $key = $m.Groups['key'].Value
        $startAt = $i
        $endAt = -1
        $j = $i + 1
        while ($j -lt $n) {
            $me = $rxE.Match([string]$Lines[$j])
            if ($me.Success -and $me.Groups['key'].Value -eq $key) { $endAt = $j; break }
            if ($rxS.Match([string]$Lines[$j]).Success) { break }
            $j++
        }
        $complete = ($endAt -ge 0)
        if ($complete) {
            $last = [Math]::Min($n - 1, $endAt + [Math]::Max(0, $AfterEnd))
        } else {
            $last = $j - 1
        }
        if ($want.Count -eq 0 -or $want.Contains($key)) {
            $slice = @($Lines[$startAt..$last])
            [void]$blocks.Add(@{ key = $key; start = $startAt; end = $last; complete = $complete; lines = $slice })
        }
        $i = $last + 1
    }
    $found = @{}
    foreach ($b in $blocks) { $found[[string]$b['key']] = $true }
    $missing = @(foreach ($k in $want.Keys) { if (-not $found.Contains($k)) { $k } })
    return @{ blocks = $blocks.ToArray(); missing = $missing }
}

function Find-EbiLineHits {
    <#
      PURE. Which lines match which patterns. Patterns are regexes; each hit
      is @{ index (0-based); pattern (the pattern's position, 0-based);
      text }. In line order. A line matching two patterns is listed once,
      under the first pattern it matches.
    #>
    param([string[]]$Lines, [string[]]$Patterns)
    $hits = New-Object System.Collections.ArrayList
    $rx = @(foreach ($p in @($Patterns)) { New-Object System.Text.RegularExpressions.Regex([string]$p) })
    $i = 0
    foreach ($line in @($Lines)) {
        for ($p = 0; $p -lt $rx.Count; $p++) {
            if ($rx[$p].IsMatch([string]$line)) { [void]$hits.Add(@{ index = $i; pattern = $p; text = [string]$line }); break }
        }
        $i++
    }
    return $hits.ToArray()
}

function Get-EbiTextCellUnits {
    # PURE. Display width in half-width cells: ASCII / Latin-1 / half-width
    # katakana count 1, everything else (kanji, kana, full-width forms) 2.
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    $u = 0
    foreach ($ch in $Text.ToCharArray()) {
        $c = [int]$ch
        if ($c -lt 0x0100 -or ($c -ge 0xFF61 -and $c -le 0xFF9F)) { $u += 1 } else { $u += 2 }
    }
    return $u
}

function Get-EbiHighlightEndColumn {
    <#
      PURE. The last column a highlight starting at StartColumn must reach
      to cover Text, on a grid whose columns each hold UnitsPerColumn
      half-width characters (MS Gothic 10pt on the evidence sheets' 2.625
      wide columns holds 3: the operator's own highlights in the sample
      workbook end exactly at ceil(units / 3)). PadColumns widens it; never
      narrower than one column, never past MaxColumn (0 = no cap).
    #>
    param([string]$Text, [int]$StartColumn, [double]$UnitsPerColumn = 3, [int]$PadColumns = 0, [int]$MaxColumn = 0)
    if ($UnitsPerColumn -le 0) { $UnitsPerColumn = 3 }
    $units = Get-EbiTextCellUnits -Text ($(if ($null -eq $Text) { '' } else { $Text.TrimEnd() }))
    $cols = [int][Math]::Ceiling($units / $UnitsPerColumn)
    if ($cols -lt 1) { $cols = 1 }
    $end = $StartColumn + $cols - 1 + [Math]::Max(0, $PadColumns)
    if ($MaxColumn -gt 0 -and $end -gt $MaxColumn) { $end = $MaxColumn }
    if ($end -lt $StartColumn) { $end = $StartColumn }
    return $end
}

function ConvertTo-EbiClockText {
    <#
      PURE. An Excel time as the cell gives it -- a day fraction (0.4895833),
      a full serial (46304.4895833) or text with float artefacts
      ("11:44:59.9999999999984025", "13:15:00.0000000000031950") -> "HH:mm:ss",
      rounded to the nearest second. '' when unreadable ('-', blank).
    #>
    param($Value)
    if ($null -eq $Value) { return '' }
    $secs = $null
    if ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal] -or $Value -is [int] -or $Value -is [long]) {
        $f = [double]$Value
        $f = $f - [Math]::Floor($f)
        $secs = [Math]::Round($f * 86400.0)
    } else {
        $t = ([string]$Value).Trim()
        $m = [regex]::Match($t, '^(\d{1,2}):(\d{2})(?::(\d{2})(\.\d+)?)?$')
        if ($m.Success) {
            $s = 0.0
            if ($m.Groups[3].Success) { $s = [double]$m.Groups[3].Value }
            if ($m.Groups[4].Success) { $s += [double]::Parse('0' + $m.Groups[4].Value, [System.Globalization.CultureInfo]::InvariantCulture) }
            $secs = [Math]::Round([int]$m.Groups[1].Value * 3600 + [int]$m.Groups[2].Value * 60 + $s)
        } else {
            $d = 0.0
            if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
                $d = $d - [Math]::Floor($d)
                $secs = [Math]::Round($d * 86400.0)
            }
        }
    }
    if ($null -eq $secs) { return '' }
    $secs = [int]$secs % 86400
    return ('{0:00}:{1:00}:{2:00}' -f [int][Math]::Floor($secs / 3600), [int][Math]::Floor(($secs % 3600) / 60), [int]($secs % 60))
}

function Get-EbiTimeAround {
    <#
      PURE. @{ ok; from; to; at } for a window of BeforeMinutes before and
      AfterMinutes after Date + Clock. Date: 'yyyy-MM-dd' / 'yyyy/MM/dd'
      (or anything ConvertTo-EbiDateTime would read -- only the date part is
      used). Clock: whatever ConvertTo-EbiClockText reads. ISO strings
      without offset ('yyyy-MM-ddTHH:mm:ss'), the shape run.timeWindow and
      the 'within' rule op use.
    #>
    param([string]$Date, $Clock, [double]$BeforeMinutes = 0, [double]$AfterMinutes = 0)
    $c = ConvertTo-EbiClockText -Value $Clock
    if ($c -eq '') { return @{ ok = $false; message = ('unreadable time "' + [string]$Clock + '"'); from = ''; to = ''; at = '' } }
    $dm = [regex]::Match(([string]$Date).Trim(), '^(\d{4})[-/](\d{1,2})[-/](\d{1,2})')
    if (-not $dm.Success) { return @{ ok = $false; message = ('unreadable date "' + $Date + '"'); from = ''; to = ''; at = '' } }
    $parts = $c.Split(':')
    $at = New-Object DateTime ([int]$dm.Groups[1].Value), ([int]$dm.Groups[2].Value), ([int]$dm.Groups[3].Value), ([int]$parts[0]), ([int]$parts[1]), ([int]$parts[2])
    $fmt = 'yyyy-MM-ddTHH:mm:ss'
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    return @{ ok = $true; message = ''; at = $at.ToString($fmt, $inv); from = $at.AddMinutes(-1 * $BeforeMinutes).ToString($fmt, $inv); to = $at.AddMinutes($AfterMinutes).ToString($fmt, $inv) }
}

function Select-EbiFilePairs {
    <#
      PURE. Pair left files with right files (GIFT side vs GFIX side of one
      transfer). Each file is @{ name; lines; order } -- order is the arrival
      rank (time, or the name when the names carry a sequence number).
      Rule, in the operator's words: the row counts must be the same, and
      several files go in the order they arrived.
        1. same count on both sides and every in-order pair has equal line
           counts -> 'ok', by order
        2. otherwise pair by line count where a count is unique on both
           sides; anything left over -> 'unknown' with the leftovers listed
      Never guesses between two files of the same line count out of order.
      Returns @{ code = ok|unknown; pairs = @(@{ left; right; lines });
      leftOver = @(names); rightOver = @(names); reason }.
    #>
    param([object[]]$Left, [object[]]$Right)
    $l = @(@($Left) | Where-Object { $null -ne $_ } | Sort-Object -Property @{ Expression = { [string]$_['order'] } }, @{ Expression = { [string]$_['name'] } })
    $r = @(@($Right) | Where-Object { $null -ne $_ } | Sort-Object -Property @{ Expression = { [string]$_['order'] } }, @{ Expression = { [string]$_['name'] } })
    $pairs = New-Object System.Collections.ArrayList
    if ($l.Count -eq 0 -or $r.Count -eq 0) {
        return @{ code = 'unknown'; pairs = @(); leftOver = @($l | ForEach-Object { [string]$_['name'] }); rightOver = @($r | ForEach-Object { [string]$_['name'] }); reason = ('no files on the ' + $(if ($l.Count -eq 0) { 'left' } else { 'right' }) + ' side') }
    }
    if ($l.Count -eq $r.Count) {
        $allEqual = $true
        for ($i = 0; $i -lt $l.Count; $i++) { if ([int]$l[$i]['lines'] -ne [int]$r[$i]['lines']) { $allEqual = $false; break } }
        if ($allEqual) {
            for ($i = 0; $i -lt $l.Count; $i++) { [void]$pairs.Add(@{ left = [string]$l[$i]['name']; right = [string]$r[$i]['name']; lines = [int]$l[$i]['lines'] }) }
            return @{ code = 'ok'; pairs = $pairs.ToArray(); leftOver = @(); rightOver = @(); reason = 'paired by arrival order, line counts equal' }
        }
    }
    $lc = @{}; $rc = @{}
    foreach ($f in $l) { $k = [string][int]$f['lines']; if ($lc.Contains($k)) { $lc[$k]++ } else { $lc[$k] = 1 } }
    foreach ($f in $r) { $k = [string][int]$f['lines']; if ($rc.Contains($k)) { $rc[$k]++ } else { $rc[$k] = 1 } }
    $usedL = @{}; $usedR = @{}
    foreach ($f in $l) {
        $k = [string][int]$f['lines']
        if ($lc[$k] -eq 1 -and $rc.Contains($k) -and $rc[$k] -eq 1) {
            $g = @($r | Where-Object { [string][int]$_['lines'] -eq $k })[0]
            [void]$pairs.Add(@{ left = [string]$f['name']; right = [string]$g['name']; lines = [int]$f['lines'] })
            $usedL[[string]$f['name']] = $true; $usedR[[string]$g['name']] = $true
        }
    }
    $lo = @(foreach ($f in $l) { if (-not $usedL.Contains([string]$f['name'])) { [string]$f['name'] } })
    $ro = @(foreach ($f in $r) { if (-not $usedR.Contains([string]$f['name'])) { [string]$f['name'] } })
    $code = if ($lo.Count -eq 0 -and $ro.Count -eq 0) { 'ok' } else { 'unknown' }
    $why = if ($code -eq 'ok') { 'paired by unique line counts (arrival order disagreed)' } else { ('' + $lo.Count + ' left / ' + $ro.Count + ' right file(s) could not be paired by line count') }
    return @{ code = $code; pairs = $pairs.ToArray(); leftOver = $lo; rightOver = $ro; reason = $why }
}

function ConvertFrom-EbiTextBytes {
    <#
      PURE. A whole data file -> @{ text; encoding = utf8 | cp932 }. The
      file is ONE encoding: valid UTF-8 throughout (BOM skipped) is UTF-8,
      anything else is CP932. Unlike ConvertFrom-EbiMixedBytes (a log with
      stray bytes), no line is decoded on its own -- a Shift_JIS line can
      contain byte runs that are also valid UTF-8, and mixing would then
      read the same content two ways. GFIX files are UTF-8 where the GIFT
      side is Shift_JIS; a diff tool reads both as the same text.
    #>
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return @{ text = ''; encoding = 'utf8' } }
    $start = 0
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { $start = 3 }
    try {
        $t = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($Bytes, $start, $Bytes.Length - $start)
        return @{ text = $t; encoding = 'utf8' }
    } catch {
        return @{ text = (Get-EbiCodePage -CodePage 932).GetString($Bytes); encoding = 'cp932' }
    }
}

function ConvertTo-EbiJisNeutral {
    <#
      PURE. Fold the characters Shift_JIS <-> Unicode converters disagree on
      (Windows CP932 vs JIS-mapped Java/ICU: U+FF0D vs U+2212 minus, U+FF5E
      vs U+301C wave dash, U+2225 vs U+2016, U+FFE0/1/2 vs U+00A2/3/AC,
      U+2015 vs U+2014) to one form, so a UTF-8 file converted on the GFIX
      side reads as the same text as the GIFT side's Shift_JIS -- which is
      what the diff tool, comparing in Shift_JIS, reports.
    #>
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $map = @{ 0x2212 = 0xFF0D; 0x301C = 0xFF5E; 0x2016 = 0x2225; 0x00A2 = 0xFFE0; 0x00A3 = 0xFFE1; 0x00AC = 0xFFE2; 0x2014 = 0x2015 }
    $sb = New-Object System.Text.StringBuilder($Text.Length)
    foreach ($ch in $Text.ToCharArray()) { $c = [int]$ch; if ($map.ContainsKey($c)) { [void]$sb.Append([char]$map[$c]) } else { [void]$sb.Append($ch) } }
    return $sb.ToString()
}

function Get-EbiFirstCharDiff {
    # PURE. Where two lines first differ: @{ col (1-based, 0 = same); left; right } as U+XXXX ('' past the end).
    param([string]$Left, [string]$Right)
    $a = if ($null -eq $Left) { '' } else { $Left }; $b = if ($null -eq $Right) { '' } else { $Right }
    $n = [Math]::Max($a.Length, $b.Length)
    for ($i = 0; $i -lt $n; $i++) {
        $ca = if ($i -lt $a.Length) { 'U+{0:X4}' -f [int]$a[$i] } else { '' }
        $cb = if ($i -lt $b.Length) { 'U+{0:X4}' -f [int]$b[$i] } else { '' }
        if ($ca -ne $cb) { return @{ col = ($i + 1); left = $ca; right = $cb } }
    }
    return @{ col = 0; left = ''; right = '' }
}

function Compare-EbiTextLines {
    <#
      PURE. Two texts line by line, blind to CRLF vs LF and to one final
      newline. @{ identical; leftLines; rightLines; firstDiff (1-based line,
      0 when identical) }.
    #>
    param([string]$Left, [string]$Right)
    $a = @(Get-EbiTextLines -Text $Left)
    $b = @(Get-EbiTextLines -Text $Right)
    $n = [Math]::Min($a.Count, $b.Count)
    for ($i = 0; $i -lt $n; $i++) {
        if (-not [string]::Equals([string]$a[$i], [string]$b[$i], [System.StringComparison]::Ordinal)) { return @{ identical = $false; leftLines = $a.Count; rightLines = $b.Count; firstDiff = ($i + 1) } }
    }
    if ($a.Count -ne $b.Count) { return @{ identical = $false; leftLines = $a.Count; rightLines = $b.Count; firstDiff = ($n + 1) } }
    return @{ identical = $true; leftLines = $a.Count; rightLines = $b.Count; firstDiff = 0 }
}

function ConvertTo-EbiColumnNumber {
    # PURE. 'A' -> 1, 'CB' -> 80; a number passes through; bad -> 0.
    param($Column)
    if ($null -eq $Column) { return 0 }
    if ($Column -is [int] -or $Column -is [long] -or $Column -is [double]) { return [int]$Column }
    $s = ([string]$Column).Trim().ToUpperInvariant()
    if ($s -match '^\d+$') { return [int]$s }
    if ($s -notmatch '^[A-Z]{1,3}$') { return 0 }
    $n = 0
    foreach ($ch in $s.ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    return $n
}

function ConvertTo-EbiColumnLetter {
    # PURE. 1 -> 'A', 80 -> 'CB'.
    param([int]$Number)
    if ($Number -lt 1) { return '' }
    $s = ''
    $n = $Number
    while ($n -gt 0) {
        $r = ($n - 1) % 26
        $s = [string][char](65 + $r) + $s
        $n = [int][Math]::Floor(($n - 1) / 26)
    }
    return $s
}
