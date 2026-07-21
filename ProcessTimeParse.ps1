# ============================================================
#  ProcessTimeParse.ps1
#
#  PURE library for the ProcessTime phase -- NO Excel COM, NO OCR calls,
#  NO mapping I/O. Dot-source only (no param() block).
#
#  ProcessTime.ps1 extracts each correl's HM batch processing start/end
#  time (and derives the duration) from the GIFT/GFIX receive-result
#  evidence screenshot already inserted into the evidence workbook
#  (sheet SheetGiftRecv / SheetGfixRecv, see ProjectLabels.ps1). Two
#  sources feed the same shape of result, cheapest/most-accurate first:
#    1. the archived Ctrl+A page text saved by HmSnap.ps1 at snap time
#       (WorkDir\snap\GIFT_HM\<correl>.txt / GFIX_HM), re-parsed with
#       SnapVerify.ps1's ConvertFrom-HmPageText (exact, TAB-anchored).
#    2. OCR of the evidence picture itself (ProcessTime.ps1's candidate
#       export + Windows OCR), when no archived text is available (e.g.
#       only the delivered J4 workbook is on hand). OCR output never
#       carries real TAB characters AND injects spaces INSIDE time tokens
#       ('10 :58 :20', '00 :00 : 0 1' -- observed on a real office-PC run,
#       v2.12.2), so this file's ConvertFrom-ProcessTimeOcrLines first
#       normalizes every loose time token back to HH:mm:ss and then
#       anchors on datetime tokens per line rather than column position.
#
#  Both tiers produce objects carrying StartTime/EndTime/Status, so
#  Get-NewestProcessTimeRow (newest-by-StartTime, matching this project's
#  established newest-wins convention -- see SnapVerify.Test-HmAbend)
#  works the same regardless of which tier produced the rows.
#
#  Convention: functions return plain arrays -- never return ,@(...)
#  because callers wrap calls in @() and that nests in PS 5.1.
# ============================================================

# Japanese status literals (source stays ASCII; see SnapVerify.ps1's
# identical constants -- duplicated here on purpose so this file has no
# cross-file dependency and stays independently unit-testable).
$script:PT_Normal = [char]0x6B63 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # seijo-shuuryo (normal end)
$script:PT_Abend  = [char]0x7570 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # ijo-shuuryo (abend)
$script:PT_Shuryo = [char]0x7D42 + [char]0x4E86                                 # shuuryo (the trailing '...end')
$script:PT_SeiChar = [string][char]0x6B63                                       # sei (normal marker)
$script:PT_IChar   = [string][char]0x7570                                       # i (abend marker)

# ---------------------------------------------------------------------------
# Get-ProcessDurationText
#   Formats the elapsed time between two [datetime] values as HH:mm:ss
#   (HH is not clamped to 24 -- a run can span more than a day). Returns ''
#   when either value is missing or EndTime is before StartTime (garbled
#   OCR read; never invent a negative duration).
# ---------------------------------------------------------------------------
function Get-ProcessDurationText {
    param($StartTime, $EndTime)
    if ($null -eq $StartTime -or $null -eq $EndTime) { return '' }
    $span = $EndTime - $StartTime
    if ($span.TotalSeconds -lt 0) { return '' }
    $totalHours = [int][Math]::Floor($span.TotalHours)
    return ('{0:00}:{1:00}:{2:00}' -f $totalHours, $span.Minutes, $span.Seconds)
}

# ---------------------------------------------------------------------------
# ConvertTo-ProcessTimeNormalizedLine
#   Repairs OCR-injected whitespace inside time tokens: the ja recognizer
#   reads the HM page's '10:58:20' as '10 :58 :20' and '00:00:01' as
#   '00 :00 : 0 1' (real office-PC dumps). Every loose H:M:S cluster --
#   1-2 digit fields, optional spaces around the colons and even between
#   the two digits of one field -- is rewritten as a canonical zero-padded
#   HH:mm:ss. Everything else on the line (14-digit datetimes, '1 ,036'
#   record counts) is left untouched: only colon-joined clusters qualify.
# ---------------------------------------------------------------------------
function ConvertTo-ProcessTimeNormalizedLine {
    param([string]$Line)
    if ([string]::IsNullOrEmpty($Line)) { return '' }
    $loose = '(?<!\d)(\d{1,2})\s*:\s*(\d(?:\s?\d)?)\s*:\s*(\d(?:\s?\d)?)(?!\d)'
    $eval = {
        param($m)
        $h  = ($m.Groups[1].Value -replace '\s', '')
        $mi = ($m.Groups[2].Value -replace '\s', '')
        $s  = ($m.Groups[3].Value -replace '\s', '')
        ('{0}:{1}:{2}' -f $h.PadLeft(2, '0'), $mi.PadLeft(2, '0'), $s.PadLeft(2, '0'))
    }
    return [regex]::Replace($Line, $loose, $eval)
}

# ---------------------------------------------------------------------------
# ConvertTo-ProcessTimeCorrelKey
#   Folds the glyphs Windows OCR most often confuses on the HM page's
#   fixed-pitch correl id into a single canonical digit, so a correl id can
#   still be recognized when the recognizer misread it. Observed on a real
#   office-PC run: 'JIGPKB1S' comes back as 'JIGPKBIS' (the digit 1 read as
#   letter I), which an exact substring test rejected -- discarding the
#   correct HM screenshot. Only the letter<->digit pairs that actually occur
#   in these ids are folded (I/L/|/!->1, O/Q->0, S->5, B->8, Z->2); the
#   distinguishing letters (J G D M K P C R F Y U) and the digits 3/4/6/7/9
#   are left alone so two genuinely different correls do not collapse
#   together. Both the correl id and the OCR line are folded the same way
#   before comparing, so real matches are preserved.
# ---------------------------------------------------------------------------
function ConvertTo-ProcessTimeCorrelKey {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $map = @{
        'I' = '1'; 'L' = '1'; '|' = '1'; '!' = '1'
        'O' = '0'; 'Q' = '0'
        'S' = '5'
        'B' = '8'
        'Z' = '2'
    }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToUpperInvariant().ToCharArray()) {
        $c = [string]$ch
        if ($map.ContainsKey($c)) { [void]$sb.Append($map[$c]) }
        else { [void]$sb.Append($c) }
    }
    return $sb.ToString()
}

# ---------------------------------------------------------------------------
# Test-ProcessTimeCorrelSeen
#   $true when $CorrelId appears in $Line, either verbatim (case-insensitive)
#   or after OCR glyph folding (ConvertTo-ProcessTimeCorrelKey) -- so a
#   screenshot whose correl id OCR misread (1<->I etc.) still counts as
#   belonging to this correl. The id is 8 chars, so a folded substring hit
#   is very unlikely to be coincidental.
# ---------------------------------------------------------------------------
function Test-ProcessTimeCorrelSeen {
    param([string]$Line, [string]$CorrelId)
    if ([string]::IsNullOrWhiteSpace($CorrelId) -or [string]::IsNullOrEmpty($Line)) { return $false }
    if ($Line.IndexOf($CorrelId, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    $keyId = ConvertTo-ProcessTimeCorrelKey $CorrelId
    if ([string]::IsNullOrEmpty($keyId)) { return $false }
    $keyLine = ConvertTo-ProcessTimeCorrelKey $Line
    return ($keyLine.IndexOf($keyId) -ge 0)
}

# ---------------------------------------------------------------------------
# Get-ProcessTimeRecordCount
#   Extracts the HM row's record count (shori-kensu) from an OCR'd line: the
#   comma-grouped integer that sits AFTER the data-creation datestamp and
#   before the result / correl-id columns (e.g. '11,262', '2,370', '0'). OCR
#   splits the digits with spaces ('1 1 ,262') so any internal whitespace is
#   removed. $SearchFrom lets the caller start the scan after the row's end
#   datetime so the proc-time (HH:mm:ss) column is never taken for the count.
#   Anchoring on the 8-14 digit datestamp is required (returns '' without it)
#   so a partial/garbled row never guesses a count out of some other column.
# ---------------------------------------------------------------------------
function Get-ProcessTimeRecordCount {
    param([string]$Line, [int]$SearchFrom = 0)
    if ([string]::IsNullOrEmpty($Line)) { return '' }
    $tail = if ($SearchFrom -gt 0 -and $SearchFrom -lt $Line.Length) { $Line.Substring($SearchFrom) } else { $Line }
    $mStamp = [regex]::Match($tail, '(?<!\d)\d{8,14}(?!\d)')
    if (-not $mStamp.Success) { return '' }
    $tail = $tail.Substring($mStamp.Index + $mStamp.Length)
    $mCount = [regex]::Match($tail, '\d[\d ,]*\d|\d')
    if (-not $mCount.Success) { return '' }
    return ($mCount.Value -replace '\s', '')
}

# ---------------------------------------------------------------------------
# ConvertFrom-ProcessTimeOcrLines
#   Anchor-based reader for OCR'd HM page text: a real table column split
#   is unreliable once OCR has collapsed variable-width whitespace, so
#   instead of splitting fields this normalizes each line's time tokens
#   (ConvertTo-ProcessTimeNormalizedLine) and scans for
#   'yyyy/MM/dd HH:mm:ss' datetime tokens, matching this project's general
#   OCR philosophy of anchoring on distinctive substrings rather than
#   trusting column position (see GfixLog.ps1 / Compare-SendRecordCheck).
#   The date-to-time gap is capped at 5 spaces so a date whose own time
#   was dropped by OCR cannot pair up with a time from a later column.
#
#   $Lines should already be RECONSTRUCTED rows (one table row per string),
#   e.g. via SendMetadata.ps1's ConvertTo-SendRowLines against the OCR
#   result's word boxes -- a raw OcrResult.Text line split can fragment one
#   wide HM row across several lines.
#
#   Row kinds:
#     full    -- two parsed datetime tokens (start, end in page order).
#     partial -- exactly ONE parsed datetime AND the line carries a
#                status-ish literal (...shuuryo): the row was seen but the
#                other time was unreadable. Partial=$true, EndTime=$null.
#                Reported so the operator knows WHICH time is missing
#                instead of a blanket 'not detected'.
#
#   Extra per-row fields (all optional for downstream consumers):
#     Status       canonical normal-end/abend literal. OCR garbles the
#                  middle characters ('sei-TEI-shuuryo' observed for
#                  normal-end), so after the exact literals miss, a fuzzy
#                  match on '...shuuryo' classifies by the preceding
#                  characters (sei->normal, i->abend).
#     PageDuration the page's own proc-time column (first standalone
#                  HH:mm:ss after the last datetime token), for
#                  cross-checking the derived duration. '' when unread.
#     CorrelSeen   $true when -CorrelId was given and appears in the line,
#                  verbatim OR after OCR glyph folding (Test-ProcessTime
#                  CorrelSeen: 1<->I etc.) -- lets the caller verify the
#                  OCR'd picture belongs to this correl even when the
#                  recognizer misread a digit in the id.
#     RecordCount  the HM row's record count (shori-kensu, comma-grouped),
#                  read after the data-creation datestamp. '' when unread.
#
#   Returns plain array of PSCustomObject
#     { StartTime; EndTime; Status; PageDuration; RecordCount; Partial;
#       CorrelSeen; RawLine }.
# ---------------------------------------------------------------------------
function ConvertFrom-ProcessTimeOcrLines {
    param([string[]]$Lines, [string]$CorrelId = '')

    $dtToken = '\d{4}/\d{2}/\d{2}[ \t]{1,5}\d{2}:\d{2}:\d{2}'
    $timeToken = '(?<!\d)\d{2}:\d{2}:\d{2}(?!\d)'
    $dtFmt   = 'yyyy/MM/dd HH:mm:ss'
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $out     = [System.Collections.Generic.List[object]]::new()

    foreach ($rawLine in @($Lines)) {
        if ([string]::IsNullOrWhiteSpace($rawLine)) { continue }
        $line = ConvertTo-ProcessTimeNormalizedLine $rawLine

        # Parsed datetimes in page order (skip tokens whose date is garbage).
        $parsed = @()
        foreach ($m in @([regex]::Matches($line, $dtToken))) {
            $txt = ($m.Value -replace '\s+', ' ').Trim()
            $dt = $null
            try { $dt = [datetime]::ParseExact($txt, $dtFmt, $culture) } catch {}
            if ($null -ne $dt) { $parsed += @{ Time = $dt; End = ($m.Index + $m.Length) } }
        }
        if ($parsed.Count -eq 0) { continue }

        # Status: exact literals first, then fuzzy '...shuuryo' (OCR garbles
        # the middle characters; classify by the char(s) just before).
        $status = ''
        if ($line.IndexOf($script:PT_Normal) -ge 0) { $status = $script:PT_Normal }
        elseif ($line.IndexOf($script:PT_Abend) -ge 0) { $status = $script:PT_Abend }
        else {
            $ix = $line.IndexOf($script:PT_Shuryo)
            if ($ix -gt 0) {
                $from = [Math]::Max(0, $ix - 2)
                $before = $line.Substring($from, $ix - $from)
                if ($before.IndexOf($script:PT_IChar) -ge 0) { $status = $script:PT_Abend }
                elseif ($before.IndexOf($script:PT_SeiChar) -ge 0) { $status = $script:PT_Normal }
            }
        }

        $isPartial = $false
        if ($parsed.Count -lt 2) {
            # One readable datetime only: keep it as a PARTIAL row when the
            # line also looks like a status row, so 'end time unreadable'
            # can be reported instead of dropping the whole row silently.
            if ($line.IndexOf($script:PT_Shuryo) -lt 0) { continue }
            $isPartial = $true
        }

        $startTime = $parsed[0].Time
        $endTime   = $null
        $searchFrom = [int]$parsed[0].End
        if (-not $isPartial) {
            $endTime = $parsed[1].Time
            $searchFrom = [int]$parsed[1].End
        }

        # The page's own proc-time column: first standalone time after the
        # last consumed datetime token (cross-check for the derived duration).
        $pageDuration = ''
        if ($searchFrom -lt $line.Length) {
            $pm = [regex]::Match($line.Substring($searchFrom), $timeToken)
            if ($pm.Success) { $pageDuration = $pm.Value }
        }

        $correlSeen = $false
        if (-not [string]::IsNullOrWhiteSpace($CorrelId)) {
            $correlSeen = (Test-ProcessTimeCorrelSeen -Line $line -CorrelId $CorrelId)
        }

        # Record count (shori-kensu), scanned after the row's end datetime so
        # the proc-time column is never mistaken for it.
        $recordCount = Get-ProcessTimeRecordCount -Line $line -SearchFrom $searchFrom

        $out.Add([PSCustomObject]@{
            StartTime    = $startTime
            EndTime      = $endTime
            Status       = $status
            PageDuration = $pageDuration
            RecordCount  = $recordCount
            Partial      = $isPartial
            CorrelSeen   = $correlSeen
            RawLine      = $rawLine
        })
    }

    return $out.ToArray()
}

# ---------------------------------------------------------------------------
# Get-ProcessTimeRowRank
#   Confidence rank of one OCR row for candidate selection:
#     3 = full row (start+end) AND the correl id was seen on the line
#     2 = full row, correl id not seen (or none was asked for)
#     1 = partial row (one time only), correl id seen
#     0 = partial row, correl id not seen
#   Rows from the archived-text tier lack the Partial/CorrelSeen fields
#   and rank as full/unseen (2).
# ---------------------------------------------------------------------------
function Get-ProcessTimeRowRank {
    param($Row)
    if ($null -eq $Row) { return -1 }
    $partial = $false; $seen = $false
    if ($Row.PSObject.Properties['Partial'])    { $partial = [bool]$Row.Partial }
    if ($Row.PSObject.Properties['CorrelSeen']) { $seen = [bool]$Row.CorrelSeen }
    $rank = 0
    if (-not $partial) { $rank += 2 }
    if ($seen) { $rank += 1 }
    return $rank
}

# ---------------------------------------------------------------------------
# Select-ProcessTimeRow
#   Picks the best row out of one OCR'd picture's parsed rows: highest
#   Get-ProcessTimeRowRank first (a full row always beats a partial one,
#   correl-seen beats unseen), newest StartTime among equals (this
#   project's newest-wins convention). -RequireCorrelSeen drops every row
#   whose line did not carry the correl id -- used for relaxed picture
#   candidates where position alone cannot prove the picture belongs to
#   this correl. Returns $null when nothing qualifies.
# ---------------------------------------------------------------------------
function Select-ProcessTimeRow {
    param([object[]]$Rows, [switch]$RequireCorrelSeen)
    $rows = @($Rows | Where-Object { $null -ne $_ -and $null -ne $_.StartTime })
    if ($RequireCorrelSeen) {
        $rows = @($rows | Where-Object { $_.PSObject.Properties['CorrelSeen'] -and [bool]$_.CorrelSeen })
    }
    if ($rows.Count -eq 0) { return $null }
    $best = $null; $bestRank = -1
    foreach ($r in $rows) {
        $rk = Get-ProcessTimeRowRank $r
        if ($rk -gt $bestRank -or ($rk -eq $bestRank -and $r.StartTime -gt $best.StartTime)) {
            $best = $r; $bestRank = $rk
        }
    }
    return $best
}

# ---------------------------------------------------------------------------
# Get-ProcessTimeOcrMissNote
#   One short diagnostic string explaining WHY a pooled OCR read produced
#   no usable time rows -- e.g. the en-US recognizer reads the date columns
#   but drops the time-of-day entirely (real office-PC dumps), which is
#   indistinguishable from 'blank picture' without this. Printed next to
#   'not detected' so the operator sees what is actually missing.
# ---------------------------------------------------------------------------
function Get-ProcessTimeOcrMissNote {
    param([string[]]$Lines)
    $dtToken   = '\d{4}/\d{2}/\d{2}[ \t]{1,5}\d{2}:\d{2}:\d{2}'
    $dateToken = '\d{4}/\d{2}/\d{2}'
    $total = 0; $dateNoTime = 0; $oneDt = 0
    foreach ($rawLine in @($Lines)) {
        if ([string]::IsNullOrWhiteSpace($rawLine)) { continue }
        $total++
        $line = ConvertTo-ProcessTimeNormalizedLine $rawLine
        $dtCount = @([regex]::Matches($line, $dtToken)).Count
        if ($dtCount -eq 1) { $oneDt++ }
        elseif ($dtCount -eq 0 -and $line -match $dateToken) { $dateNoTime++ }
    }
    if ($total -eq 0) { return 'no OCR lines' }
    $parts = @(("{0} OCR line(s)" -f $total))
    if ($dateNoTime -gt 0) { $parts += ("{0} with a date but no readable time-of-day" -f $dateNoTime) }
    if ($oneDt -gt 0) { $parts += ("{0} with only one datetime" -f $oneDt) }
    if ($dateNoTime -eq 0 -and $oneDt -eq 0) { $parts += 'no date/time tokens recognized' }
    return ($parts -join ', ')
}

# ---------------------------------------------------------------------------
# Get-NewestProcessTimeRow
#   Returns the row with the latest StartTime (newest-wins, matching
#   Test-HmAbend / GIFT_MQ's established convention), or $null when -Rows
#   is empty. Works uniformly on rows from ConvertFrom-HmPageText (archived
#   text tier) or ConvertFrom-ProcessTimeOcrLines (OCR tier) -- both carry
#   a StartTime property.
# ---------------------------------------------------------------------------
function Get-NewestProcessTimeRow {
    param([object[]]$Rows)
    $rows = @($Rows | Where-Object { $null -ne $_ -and $null -ne $_.StartTime })
    if ($rows.Count -eq 0) { return $null }
    return @($rows | Sort-Object StartTime -Descending)[0]
}
