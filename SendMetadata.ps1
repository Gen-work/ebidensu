# ============================================================
#  SendMetadata.ps1
#
#  PURE send-side evidence metadata helpers for the SendVsGift
#  Stage 2 OCR flow -- NO Excel COM, NO WinRT, NO file I/O.
#  Dot-source only (no param() block). Unit-tested by
#  Tests\Test-SendMetadata.ps1.
#
#  Pipeline position:
#    EvidenceImageExport.ps1  -> PNGs of the send-sheet screenshots
#    OcrWindows.ps1           -> OCR lines (Text + word boxes)
#    SendMetadata.ps1 (here)  -> parse lines into a record parallel to
#                                gift_metadata.csv and compare the two
#
#  NOTE on spacing: the Japanese Windows OCR recognizer drops spaces
#  between tokens. Get-SendLineTextFromWords rebuilds them from the
#  word bounding boxes (X/Width) so fixed-position record text can be
#  approximated and first/last tokens stay separable.
# ============================================================

function Get-SendFirstToken {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $parts = @($Text.Trim() -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts.Count -eq 0) { return '' }
    return $parts[0]
}

# Rebuilds one text line from OCR word objects (Text/X/Width), inserting
# spaces sized from the running average character width. A gap wider than
# GapRatio * charWidth becomes round(gap / charWidth) spaces (min 1).
function Get-SendLineTextFromWords {
    param($Words, [double]$GapRatio = 0.5)
    $ws = @(@($Words) | Where-Object { $null -ne $_ } | Sort-Object { [double]$_.X })
    if ($ws.Count -eq 0) { return '' }
    $sb = New-Object System.Text.StringBuilder
    $prevEnd = $null
    $charW = 0.0
    foreach ($w in $ws) {
        $text = [string]$w.Text
        $x = 0.0
        $width = 0.0
        try { $x = [double]$w.X } catch {}
        try { $width = [double]$w.Width } catch {}
        $len = [Math]::Max(1, $text.Length)
        if ($width -gt 0) { $charW = $width / $len }
        if ($null -ne $prevEnd -and $charW -gt 0) {
            $gap = $x - $prevEnd
            if ($gap -ge ($GapRatio * $charW)) {
                $n = [int][Math]::Round($gap / $charW)
                if ($n -lt 1) { $n = 1 }
                [void]$sb.Append(' ' * $n)
            }
        }
        [void]$sb.Append($text)
        $prevEnd = $x + $width
    }
    return $sb.ToString()
}

# Normalizes OCR output to plain trimmed text lines. Accepts plain strings
# or objects shaped like OcrWindows.ps1 lines (.Text and optional .Words).
# When word boxes exist they win, so rebuilt spacing replaces the raw Text.
function ConvertTo-SendTextLines {
    param($OcrLines, [double]$GapRatio = 0.5)
    $out = @()
    foreach ($ln in @($OcrLines)) {
        if ($null -eq $ln) { continue }
        $text = ''
        if ($ln -is [string]) {
            $text = $ln
        } else {
            $words = $null
            if ($ln.PSObject.Properties.Name -contains 'Words') { $words = @($ln.Words) }
            if ($null -ne $words -and $words.Count -gt 0) {
                $text = Get-SendLineTextFromWords $words $GapRatio
                # word .Text reads can come back empty on some WinRT
                # projections while the line .Text still works -- fall back
                if ([string]::IsNullOrWhiteSpace($text) -and
                    $ln.PSObject.Properties.Name -contains 'Text') {
                    $text = [string]$ln.Text
                }
            } elseif ($ln.PSObject.Properties.Name -contains 'Text') {
                $text = [string]$ln.Text
            }
        }
        $text = ([string]$text).TrimEnd()
        if (-not [string]::IsNullOrWhiteSpace($text)) { $out += $text }
    }
    # no comma protection: callers wrap in @(...) and @( ,@($arr) ) nests
    # in PS 5.1 (every image then counts as ONE "line")
    return $out
}

# Rebuilds TERMINAL rows from the word boxes of ALL OCR lines of one
# image. The engine fragments one terminal row into several OCR "lines"
# (field: ~187 OCR lines for a ~40-row screen), so a row label and its
# record can land in different fragments and no per-line matcher can see
# them together. All words are re-clustered by vertical center
# (tolerance = 0.6 x the median word height) and each cluster becomes one
# line, left-to-right, via the spacing rebuild. Falls back to
# ConvertTo-SendTextLines when no word boxes are available.
function ConvertTo-SendRowLines {
    param($OcrLines, [double]$GapRatio = 0.5)
    $allWords = @()
    foreach ($ln in @($OcrLines)) {
        if ($null -eq $ln) { continue }
        if ($ln -isnot [string] -and $ln.PSObject.Properties.Name -contains 'Words') {
            foreach ($w in @($ln.Words)) { if ($null -ne $w) { $allWords += $w } }
        }
    }
    if ($allWords.Count -eq 0) { return ConvertTo-SendTextLines $OcrLines $GapRatio }

    $hs = @($allWords | ForEach-Object { [double]$_.Height } | Sort-Object)
    $medH = [double]$hs[[int][Math]::Floor($hs.Count / 2)]
    if ($medH -le 0) { $medH = 10.0 }
    $tol = $medH * 0.6

    $sorted = @($allWords | Sort-Object { [double]$_.Y + ([double]$_.Height / 2.0) })
    $rows = @()
    $cur = @()
    $curCenter = 0.0
    foreach ($w in $sorted) {
        $c = [double]$w.Y + ([double]$w.Height / 2.0)
        if ($cur.Count -eq 0) {
            $cur = @($w); $curCenter = $c
        } elseif (($c - $curCenter) -le $tol) {
            $cur += $w
            $curCenter += (($c - $curCenter) / $cur.Count)   # running mean
        } else {
            $rows += ,@($cur)
            $cur = @($w); $curCenter = $c
        }
    }
    if ($cur.Count -gt 0) { $rows += ,@($cur) }

    $out = @()
    foreach ($r in $rows) {
        $t = ([string](Get-SendLineTextFromWords $r $GapRatio)).TrimEnd()
        if (-not [string]::IsNullOrWhiteSpace($t)) { $out += $t }
    }
    return $out
}

# Detects the standard 0-byte screenshot pattern.
# TODO(Stage 2): tighten the default once representative 0-byte SEND
# screenshots are available; override via SendVsGift.ZeroBytePattern.
function Test-SendZeroByteText {
    param([string[]]$TextLines, [string]$Pattern = '')
    if ([string]::IsNullOrWhiteSpace($Pattern)) {
        $Pattern = '(^|[^0-9])0\s*(byte|bytes)([^a-z]|$)'
    }
    foreach ($t in @($TextLines)) {
        if ([string]$t -match $Pattern) { return $true }
    }
    return $false
}

# Host list screens print a leading row number on each record line; the
# largest leading integer is the best row-count guess. 0 = unknown.
function Get-SendRowNumberGuess {
    param([string[]]$TextLines)
    $max = 0
    foreach ($t in @($TextLines)) {
        $m = [regex]::Match([string]$t, '^\s*(\d{1,9})\b')
        if (-not $m.Success) { continue }
        $v = 0
        if ([int]::TryParse($m.Groups[1].Value, [ref]$v) -and $v -gt $max) { $max = $v }
    }
    return $max
}

# Assembles one send_metadata.csv row, field-parallel to gift_metadata.csv.
# Confidence is heuristic: the Windows OCR API exposes no per-word score,
# so we score by how many comparison fields were actually parsed.
function Build-SendMetadataRecord {
    param(
        [string]$CorrelIdS,
        [string]$ExcelName,
        [int]$ImageCount,
        [string[]]$TextLines,
        [string]$ZeroBytePattern = ''
    )
    $lines = @(@($TextLines) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $first = if ($lines.Count -gt 0) { [string]$lines[0] } else { '' }
    $last  = if ($lines.Count -gt 0) { [string]$lines[$lines.Count - 1] } else { '' }
    $zero  = Test-SendZeroByteText $lines $ZeroBytePattern
    $rows  = Get-SendRowNumberGuess $lines

    $conf = 0.0
    if ($lines.Count -gt 0) { $conf += 0.4 }
    if ($rows -gt 0 -or $zero) { $conf += 0.3 }
    if ((Get-SendFirstToken $first) -ne '') { $conf += 0.3 }

    [pscustomobject]@{
        CorrelIdS        = [string]$CorrelIdS
        ExcelName        = [string]$ExcelName
        ImageCount       = [int]$ImageCount
        OcrLineCount     = [int]$lines.Count
        ZeroByte         = [bool]$zero
        RowNumberGuess   = [int]$rows
        FirstRecord      = $first
        LastRecord       = $last
        FirstRecordToken = Get-SendFirstToken $first
        LastRecordToken  = Get-SendFirstToken $last
        Confidence       = [Math]::Round($conf, 2)
        MetadataVersion  = '2'
    }
}

# ---- Stage 2 evidence-verdict helpers (pure; unit-tested) ----
# These implement the operator's review rules for the host screenshots:
#   0-byte file  : dataset-info screen shows 'used CYLINDERS : 0', OR the
#                  begin-of-data and end-of-data markers sit on the SAME
#                  image with no 000001 record line.
#   non-0-byte   : the zero-padded max row number (from gift_metadata) must
#                  appear in the OCR text, and the first/last records are
#                  compared by first space-free token (exact) with a
#                  prefix-similarity fallback for OCR noise (~80% rule).

# Japanese fragments built from code points (ASCII-only source; see
# ProjectLabels.ps1 for the policy).
function Get-SendZeroByteLabels {
    @{
        # 'shiyou' (used) U+4F7F U+7528 - the dataset-info 'used CYLINDERS' row
        Shiyou = [string]([char]0x4F7F) + [char]0x7528
        # 'no hajime' U+306E U+59CB U+3081 - tail of 'de-ta no hajime' (begin-of-data)
        Begin  = [string]([char]0x306E) + [char]0x59CB + [char]0x3081
        # 'no owa..' U+306E U+7D42 - tail of 'de-ta no owari' (end-of-data, both spellings)
        End    = [string]([char]0x306E) + [char]0x7D42
    }
}

# Removes every space (ASCII + full-width U+3000). The ja recognizer
# returns one word per CHARACTER and the word-box spacing rebuild can
# over-insert ('002640' -> '0 0 2 6 4 0'), so every matcher below also
# runs against this compact form of each line.
function ConvertTo-SendCompactLine {
    param([string]$Text)
    # .NET \s covers the ideographic space U+3000 too
    return ([string]$Text -replace '\s+', '')
}

# Zero-padded row label as printed on host list screens (000001, 004644, ...).
function Get-SendRowLabel {
    param([int]$Number, [int]$MinWidth = 6)
    $w = [Math]::Max($MinWidth, $Number.ToString().Length)
    return $Number.ToString('D' + $w)
}

# True when the row label appears anywhere in the lines as a standalone
# number (not part of a longer digit run). Falls back to the compact form:
# host list screens print the row number at line start, so a compact line
# beginning with the label counts even when the label glues to the record.
function Test-SendRowNumberPresent {
    param([string[]]$TextLines, [string]$RowLabel)
    if ([string]::IsNullOrWhiteSpace($RowLabel)) { return $false }
    $re = '(^|[^0-9])' + [regex]::Escape($RowLabel) + '([^0-9]|$)'
    foreach ($t in @($TextLines)) {
        if ([string]$t -match $re) { return $true }
    }
    foreach ($t in @($TextLines)) {
        if ((ConvertTo-SendCompactLine $t).StartsWith($RowLabel)) { return $true }
    }
    return $false
}

# Returns the record text after a leading row label, or $null when no line
# starts with that label. Tolerates the OCR gluing the label to the data.
# Falls back to the compact line form (returns a compact record then;
# Compare-SendRecordCheck compacts both sides so that stays comparable).
function Find-SendRecordByRowNumber {
    param([string[]]$TextLines, [string]$RowLabel)
    if ([string]::IsNullOrWhiteSpace($RowLabel)) { return $null }
    $re = '^\s*' + [regex]::Escape($RowLabel) + '(?![0-9])\s*(.*)$'
    foreach ($t in @($TextLines)) {
        $m = [regex]::Match([string]$t, $re)
        if ($m.Success) {
            $rest = ([string]$m.Groups[1].Value).Trim()
            if ($rest -ne '') { return $rest }
        }
    }
    foreach ($t in @($TextLines)) {
        $c = ConvertTo-SendCompactLine $t
        if ($c.Length -gt $RowLabel.Length -and $c.StartsWith($RowLabel)) {
            return $c.Substring($RowLabel.Length)
        }
    }
    return $null
}

# Per-IMAGE 0-byte evidence (operator rule). A custom regex overrides both
# built-in rules. Rule B needs begin+end on the same image because a
# non-empty file shows the begin marker on its head image and the end
# marker on its tail image, never both together.
function Test-SendZeroByteImage {
    param([string[]]$TextLines, [string]$Pattern = '')
    $lines = @(@($TextLines) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($lines.Count -eq 0) { return $false }
    # every rule also runs against the compact (space-stripped) form: the
    # ja recognizer's per-character words can spread 'CYLINDERS' or the
    # begin/end markers as 'C Y L I N D E R S'
    $compact = @($lines | ForEach-Object { ConvertTo-SendCompactLine $_ })
    if (-not [string]::IsNullOrWhiteSpace($Pattern)) {
        foreach ($t in @($lines) + @($compact)) { if ([string]$t -match $Pattern) { return $true } }
        return $false
    }
    $L = Get-SendZeroByteLabels
    # Rule A: dataset-info screen, 'used CYLINDERS . . : 0'
    $cyl = [regex]::Escape($L.Shiyou) + '\s*CYLINDERS[^0-9]*0([^0-9]|$)'
    foreach ($t in @($lines) + @($compact)) { if ([string]$t -match $cyl) { return $true } }
    # Rule B: begin + end markers on the same image, no 000001 record line.
    $hasBegin = $false
    $hasEnd = $false
    foreach ($t in @($lines) + @($compact)) {
        $s = [string]$t
        if ($s.Contains($L.Begin)) { $hasBegin = $true }
        if ($s.Contains($L.End)) { $hasEnd = $true }
    }
    if ($hasBegin -and $hasEnd) {
        return (-not (Test-SendRowNumberPresent $lines '000001'))
    }
    return $false
}

# Levenshtein similarity of the first PrefixLength chars (case-sensitive).
# 1.0 = identical prefixes; the default 20-char window keeps the DP cheap
# while still covering 'a part of the record' per the review rule.
function Get-SendPrefixSimilarity {
    param([string]$A, [string]$B, [int]$PrefixLength = 20)
    $a = [string]$A
    $b = [string]$B
    if ($PrefixLength -gt 0) {
        if ($a.Length -gt $PrefixLength) { $a = $a.Substring(0, $PrefixLength) }
        if ($b.Length -gt $PrefixLength) { $b = $b.Substring(0, $PrefixLength) }
    }
    if ($a.Length -eq 0 -and $b.Length -eq 0) { return 1.0 }
    if ($a.Length -eq 0 -or $b.Length -eq 0) { return 0.0 }
    $n = $a.Length
    $m = $b.Length
    $prev = New-Object int[] ($m + 1)
    for ($j = 0; $j -le $m; $j++) { $prev[$j] = $j }
    for ($i = 1; $i -le $n; $i++) {
        $cur = New-Object int[] ($m + 1)
        $cur[0] = $i
        for ($j = 1; $j -le $m; $j++) {
            $cost = if ($a[$i - 1] -ceq $b[$j - 1]) { 0 } else { 1 }
            $cur[$j] = [Math]::Min([Math]::Min($cur[$j - 1] + 1, $prev[$j] + 1), $prev[$j - 1] + $cost)
        }
        $prev = $cur
    }
    $den = [Math]::Max($n, $m)
    return [Math]::Round(1.0 - ($prev[$m] / [double]$den), 4)
}

# One record check: exact first-token match wins; prefix similarity >=
# threshold is a 'fuzzy' pass (OCR noise tolerated); both records present
# but neither passing is a 'mismatch'; missing evidence is 'unknown'.
function Compare-SendRecordCheck {
    param([string]$Name, [string]$SendRecord, [string]$GiftRecord,
          [double]$Threshold = 0.8, [int]$PrefixLength = 20)
    $sDisp = [string]$SendRecord
    $gDisp = [string]$GiftRecord
    if ($sDisp.Length -gt $PrefixLength) { $sDisp = $sDisp.Substring(0, $PrefixLength) + '..' }
    if ($gDisp.Length -gt $PrefixLength) { $gDisp = $gDisp.Substring(0, $PrefixLength) + '..' }
    if ([string]::IsNullOrWhiteSpace($SendRecord) -or [string]::IsNullOrWhiteSpace($GiftRecord)) {
        return [pscustomobject]@{ Name = $Name; Send = $sDisp; Gift = $gDisp; Status = 'unknown' }
    }
    $sTok = Get-SendFirstToken $SendRecord
    $gTok = Get-SendFirstToken $GiftRecord
    $status = 'mismatch'
    if ($sTok -ne '' -and $sTok -ceq $gTok) {
        $status = 'match'
    } else {
        $sim = Get-SendPrefixSimilarity $SendRecord $GiftRecord $PrefixLength
        if ($sim -lt $Threshold) {
            # compact compare: OCR spacing is unreliable in both directions
            # (over-inserted per-character spaces or dropped real ones)
            $sim = Get-SendPrefixSimilarity (ConvertTo-SendCompactLine $SendRecord) (ConvertTo-SendCompactLine $GiftRecord) $PrefixLength
        }
        if ($sim -ge $Threshold) { $status = 'fuzzy' }
    }
    return [pscustomobject]@{ Name = $Name; Send = $sDisp; Gift = $gDisp; Status = $status }
}

# Full per-correl verdict. ImageTextSets = one string[] of OCR lines per
# exported picture (sheet order). Returns Verdict 'ok' / 'ng' / 'unknown'
# plus the per-field checks for console display.
#   ok      -> evidence agrees with gift_metadata (auto-markable as 1)
#   ng      -> positive disagreement (auto-markable as 2, operator follows up)
#   unknown -> not enough OCR evidence; manual review decides
function Compare-SendGiftEvidence {
    param(
        $GiftRow,
        [object[]]$ImageTextSets,
        [double]$SimilarityThreshold = 0.8,
        [int]$PrefixLength = 20,
        [string]$ZeroBytePattern = ''
    )
    $sets = @()
    foreach ($s in @($ImageTextSets)) {
        if ($null -eq $s) { continue }
        $sets += ,@(@($s) | ForEach-Object { [string]$_ })
    }
    $allLines = @()
    foreach ($s in $sets) { $allLines += $s }

    $checks = @()
    $giftSize = 0L
    try {
        if ($GiftRow.PSObject.Properties.Name -contains 'SizeBytes') { $giftSize = [long]$GiftRow.SizeBytes }
    } catch {}

    if ($giftSize -eq 0) {
        $zeroSeen = $false
        foreach ($s in $sets) {
            if (Test-SendZeroByteImage -TextLines $s -Pattern $ZeroBytePattern) { $zeroSeen = $true; break }
        }
        $hasRow1 = Test-SendRowNumberPresent $allLines '000001'
        $st = 'unknown'
        if ($zeroSeen -and -not $hasRow1) { $st = 'match' }
        elseif ($hasRow1) { $st = 'mismatch' }
        $checks += [pscustomobject]@{
            Name   = 'ZeroByte'
            Send   = ('zeroEvidence={0} row000001={1}' -f $zeroSeen, $hasRow1)
            Gift   = '0 bytes'
            Status = $st
        }
        $verdict = switch ($st) { 'match' { 'ok' } 'mismatch' { 'ng' } default { 'unknown' } }
        return [pscustomobject]@{
            Verdict = $verdict; Checks = @($checks)
            RowLabel = ''; FirstRecordSend = ''; LastRecordSend = ''
        }
    }

    $maxRows = 0
    try {
        if ($GiftRow.PSObject.Properties.Name -contains 'MaxRowNumber') { $maxRows = [int]$GiftRow.MaxRowNumber }
    } catch {}
    $maxLabel = ''
    if ($maxRows -gt 0) { $maxLabel = Get-SendRowLabel $maxRows }

    $maxFound = ($maxLabel -ne '') -and (Test-SendRowNumberPresent $allLines $maxLabel)
    $checks += [pscustomobject]@{
        Name   = 'MaxRowNumber'
        Send   = $(if ($maxFound) { $maxLabel } else { '(not found)' })
        Gift   = [string]$maxRows
        Status = $(if ($maxFound) { 'match' } else { 'unknown' })
    }

    $firstSend = Find-SendRecordByRowNumber $allLines (Get-SendRowLabel 1)
    $lastSend = $null
    if ($maxLabel -ne '') { $lastSend = Find-SendRecordByRowNumber $allLines $maxLabel }

    $giftFirst = ''
    $giftLast = ''
    try {
        if ($GiftRow.PSObject.Properties.Name -contains 'FirstRecord') { $giftFirst = [string]$GiftRow.FirstRecord }
        if ($GiftRow.PSObject.Properties.Name -contains 'LastRecord') { $giftLast = [string]$GiftRow.LastRecord }
    } catch {}

    $firstCheck = Compare-SendRecordCheck 'FirstRecord' ([string]$firstSend) $giftFirst $SimilarityThreshold $PrefixLength
    $lastCheck  = Compare-SendRecordCheck 'LastRecord' ([string]$lastSend) $giftLast $SimilarityThreshold $PrefixLength
    $checks += $firstCheck
    $checks += $lastCheck

    $anyMismatch = @($checks | Where-Object { $_.Status -eq 'mismatch' }).Count -gt 0
    $recordsOk = (@('match', 'fuzzy') -contains $firstCheck.Status) -and (@('match', 'fuzzy') -contains $lastCheck.Status)
    $verdict = 'unknown'
    if ($anyMismatch) { $verdict = 'ng' }
    elseif ($maxFound -and $recordsOk) { $verdict = 'ok' }

    return [pscustomobject]@{
        Verdict = $verdict; Checks = @($checks)
        RowLabel = $maxLabel
        FirstRecordSend = [string]$firstSend
        LastRecordSend = [string]$lastSend
    }
}

# Compares a send record (Build-SendMetadataRecord) against one
# gift_metadata.csv row. Each check is match / mismatch / unknown;
# absence of OCR evidence is "unknown", never "mismatch".
# Verdict: any mismatch -> mismatch; zero-byte agreement or at least
# MinMatches matches -> match; otherwise unknown (manual review stays).
function Compare-SendGiftMetadata {
    param($Send, $Gift, [int]$MinMatches = 2)
    $checks = @()

    $giftZero = $false
    try {
        if ($Gift.PSObject.Properties.Name -contains 'SizeBytes') { $giftZero = ([long]$Gift.SizeBytes -eq 0) }
    } catch {}
    $sendZero = $false
    if ($Send.PSObject.Properties.Name -contains 'ZeroByte') {
        $zv = [string]$Send.ZeroByte
        $sendZero = ($zv -eq 'True' -or $zv -eq 'true' -or $zv -eq '1')
    }
    $zeroMatched = $false
    if ($sendZero -or $giftZero) {
        $st = 'unknown'
        if ($sendZero -and $giftZero) { $st = 'match'; $zeroMatched = $true }
        elseif ($sendZero -and -not $giftZero) { $st = 'mismatch' }
        $checks += [pscustomobject]@{ Name = 'ZeroByte'; Send = [string]$sendZero; Gift = [string]$giftZero; Status = $st }
    }

    $sendRows = 0
    try {
        if ($Send.PSObject.Properties.Name -contains 'RowNumberGuess') { $sendRows = [int]$Send.RowNumberGuess }
    } catch {}
    $giftRows = 0
    try {
        if ($Gift.PSObject.Properties.Name -contains 'MaxRowNumber') { $giftRows = [int]$Gift.MaxRowNumber }
    } catch {}
    $st = 'unknown'
    if ($sendRows -gt 0 -and $giftRows -gt 0) {
        $st = if ($sendRows -eq $giftRows) { 'match' } else { 'mismatch' }
    }
    $checks += [pscustomobject]@{ Name = 'RowNumber'; Send = [string]$sendRows; Gift = [string]$giftRows; Status = $st }

    foreach ($tok in @('FirstRecordToken', 'LastRecordToken')) {
        $sv = ''
        $gv = ''
        if ($Send.PSObject.Properties.Name -contains $tok) { $sv = [string]$Send.$tok }
        if ($Gift.PSObject.Properties.Name -contains $tok) { $gv = [string]$Gift.$tok }
        $st = 'unknown'
        if ($sv -ne '' -and $gv -ne '') {
            $st = if ($sv -eq $gv) { 'match' } else { 'mismatch' }
        }
        $checks += [pscustomobject]@{ Name = $tok; Send = $sv; Gift = $gv; Status = $st }
    }

    $matchCount    = @($checks | Where-Object { $_.Status -eq 'match' }).Count
    $mismatchCount = @($checks | Where-Object { $_.Status -eq 'mismatch' }).Count
    $unknownCount  = @($checks | Where-Object { $_.Status -eq 'unknown' }).Count

    $verdict = 'unknown'
    if ($mismatchCount -gt 0) { $verdict = 'mismatch' }
    elseif ($zeroMatched -or $matchCount -ge $MinMatches) { $verdict = 'match' }

    [pscustomobject]@{
        Checks        = @($checks)
        Verdict       = $verdict
        MatchCount    = [int]$matchCount
        MismatchCount = [int]$mismatchCount
        UnknownCount  = [int]$unknownCount
    }
}
