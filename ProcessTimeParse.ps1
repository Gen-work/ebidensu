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
        # OCR frequently inserts a space between the leading J and the rest
        # of the fixed-width id ("J IDSCS4S", "J IGPM05S"). Formatting
        # whitespace is not part of an HM correlation id.
        if ([char]::IsWhiteSpace($ch)) { continue }
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
#   Normally the 8-14 digit datestamp anchors the scan. JDL rows where OCR
#   drops that blank column use the result diamond as a strict fallback
#   anchor, so a partial/garbled row never guesses from the batch id.
# ---------------------------------------------------------------------------
function Get-ProcessTimeRecordCount {
    param([string]$Line, [int]$SearchFrom = 0)
    if ([string]::IsNullOrEmpty($Line)) { return '' }
    $tail = if ($SearchFrom -gt 0 -and $SearchFrom -lt $Line.Length) { $Line.Substring($SearchFrom) } else { $Line }
    $mStamp = [regex]::Match($tail, '(?<!\d)\d{8,14}(?!\d)')
    if ($mStamp.Success) {
        $tail = $tail.Substring($mStamp.Index + $mStamp.Length)
        $mCount = [regex]::Match($tail, '(?<![A-Z0-9])(\d[\d ,]*\d|\d)(?![A-Z0-9])', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($mCount.Success) { return ($mCount.Value -replace '\s', '') }
    }
    # Some JDL OCR rows lose the blank data-creation column completely, but
    # retain the count immediately before the black result diamond. This
    # anchor is unambiguous and avoids mistaking digits in the batch id.
    $diamond = [string][char]0x25C6
    $beforeResult = [regex]::Match($tail, ('(\d[\d ,]*\d|\d)\s*{0}' -f [regex]::Escape($diamond)))
    if ($beforeResult.Success) { return ($beforeResult.Groups[1].Value -replace '\s', '') }
    return ''
}

# ---------------------------------------------------------------------------
# Get-ProcessTimeDateHints
#   Parses every 14-digit data-creation datestamp (yyyyMMddHHmmss) out of a
#   set of OCR lines into [datetime] values. The datestamp equals the batch
#   start datetime and is a pure Latin-digit run, so the en-US recognizer
#   reads it cleanly even when the ja recognizer misreads a DATE digit in the
#   same row's start-time column (observed: ja reads '2026/05/29' as
#   '2026/05/23' while the times of day are correct). Feed the en-US-only
#   lines here and pass the result as -StartDateHints to
#   ConvertFrom-ProcessTimeOcrLines to correct a row's date without losing
#   the ja recognizer's more complete time-of-day read. Returns plain array.
# ---------------------------------------------------------------------------
function Get-ProcessTimeDateHints {
    param([string[]]$Lines)
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $styles  = [System.Globalization.DateTimeStyles]::None
    $hints   = [System.Collections.Generic.List[datetime]]::new()
    foreach ($ln in @($Lines)) {
        if ([string]::IsNullOrWhiteSpace($ln)) { continue }
        foreach ($m in @([regex]::Matches($ln, '(?<!\d)\d{14}(?!\d)'))) {
            $dt = [datetime]::MinValue
            if ([datetime]::TryParseExact($m.Value, 'yyyyMMddHHmmss', $culture, $styles, [ref]$dt)) {
                $hints.Add($dt)
            }
        }
    }
    return $hints.ToArray()
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
#     DateCorrected $true when the row's date was replaced from a -StartDateHints
#                  entry (a ja date-digit misread fixed from the en-US
#                  datestamp; the time of day is kept).
#
#   -StartDateHints (optional): trusted start [datetime]s, e.g. from
#     Get-ProcessTimeDateHints over the en-US-only lines. A row whose start
#     time-of-day matches a hint's but whose date differs adopts the hint's
#     date (end time shifted by the same delta).
#
#   Returns plain array of PSCustomObject
#     { StartTime; EndTime; Status; PageDuration; RecordCount; Partial;
#       CorrelSeen; DateCorrected; RawLine }.
# ---------------------------------------------------------------------------
function ConvertFrom-ProcessTimeOcrLines {
    param([string[]]$Lines, [string]$CorrelId = '', [datetime[]]$StartDateHints = @())

    $dtToken = '\d{4}/\d{2}/\d{2}[ \t]{1,5}\d{2}:\d{2}:\d{2}'
    $timeToken = '(?<!\d)\d{2}:\d{2}:\d{2}(?!\d)'
    $dtFmt   = 'yyyy/MM/dd HH:mm:ss'
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $out     = [System.Collections.Generic.List[object]]::new()
    $countHints = @{}

    # en-US often reads the 14-digit creation stamp and count cleanly but
    # drops both time-of-day columns. Preserve those counts and join them to
    # the ja timestamp row by the creation stamp (= start datetime).
    foreach ($hintLine in @($Lines)) {
        foreach ($stamp in @([regex]::Matches($hintLine, '(?<!\d)\d{14}(?!\d)'))) {
            $after = $hintLine.Substring($stamp.Index + $stamp.Length)
            $countMatch = [regex]::Match($after, '(?<![A-Z0-9])(\d[\d ,]*\d|\d)(?![A-Z0-9])', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($countMatch.Success) {
                $countHints[$stamp.Value] = ($countMatch.Value -replace '\s', '')
            }
        }
    }

    # Windows OCR reconstructs a wide HM table as several visual rows. In
    # particular, the correlation id is frequently returned on a different
    # line from the start/end timestamps. Candidate ownership is therefore
    # evidence about the complete exported picture, not only the timestamp
    # line. Keep the per-line check below, but promote it when the id appears
    # anywhere in this picture's OCR output.
    $pictureCorrelSeen = $false
    if (-not [string]::IsNullOrWhiteSpace($CorrelId)) {
        foreach ($candidateLine in @($Lines)) {
            if (Test-ProcessTimeCorrelSeen -Line $candidateLine -CorrelId $CorrelId) {
                $pictureCorrelSeen = $true
                break
            }
        }
    }

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

        # Date correction: the ja recognizer sometimes misreads a DATE digit
        # in the start-time column ('2026/05/29' -> '2026/05/23') while the
        # time of day is correct. When a trusted en-US start-datetime hint
        # (from the clean 14-digit datestamp) shares this row's start
        # time-of-day but carries a different date, adopt the hint's date and
        # shift the end time by the same delta -- keeping the ja time-of-day.
        $dateCorrected = $false
        if ($StartDateHints.Count -gt 0 -and $null -ne $startTime) {
            foreach ($hint in $StartDateHints) {
                if ($hint.TimeOfDay -eq $startTime.TimeOfDay) {
                    if ($hint.Date -ne $startTime.Date) {
                        $delta = $hint.Date - $startTime.Date
                        $startTime = $startTime.Add($delta)
                        if ($null -ne $endTime) { $endTime = $endTime.Add($delta) }
                        $dateCorrected = $true
                    }
                    break
                }
            }
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
            $correlSeen = $pictureCorrelSeen -or (Test-ProcessTimeCorrelSeen -Line $line -CorrelId $CorrelId)
        }

        # Record count (shori-kensu), scanned after the row's end datetime so
        # the proc-time column is never mistaken for it.
        $recordCount = Get-ProcessTimeRecordCount -Line $line -SearchFrom $searchFrom
        if ([string]::IsNullOrWhiteSpace($recordCount) -and $null -ne $startTime) {
            $stampKey = $startTime.ToString('yyyyMMddHHmmss')
            if ($countHints.ContainsKey($stampKey)) { $recordCount = [string]$countHints[$stampKey] }
        }

        $out.Add([PSCustomObject]@{
            StartTime     = $startTime
            EndTime       = $endTime
            Status        = $status
            PageDuration  = $pageDuration
            RecordCount   = $recordCount
            Partial       = $isPartial
            CorrelSeen    = $correlSeen
            DateCorrected = $dateCorrected
            RawLine       = $rawLine
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
#   this correl. -MinimumTimeOfDay (default 09:00) drops HM history rows
#   from before the operator's actual working window -- pass [timespan]::Zero
#   to disable. Returns $null when nothing qualifies.
# ---------------------------------------------------------------------------
function Select-ProcessTimeRow {
    param([object[]]$Rows, [switch]$RequireCorrelSeen, [timespan]$MinimumTimeOfDay = ([timespan]'09:00:00'))
    $rows = @($Rows | Where-Object { $null -ne $_ -and $null -ne $_.StartTime })
    if ($MinimumTimeOfDay -gt [timespan]::Zero) {
        $rows = @($rows | Where-Object { $_.StartTime.TimeOfDay -ge $MinimumTimeOfDay })
    }
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
#   a StartTime property. -MinimumTimeOfDay (default 09:00) rejects HM
#   history rows the same way the OCR-tier Select-ProcessTimeRow does.
# ---------------------------------------------------------------------------
function Get-NewestProcessTimeRow {
    param([object[]]$Rows, [timespan]$MinimumTimeOfDay = ([timespan]'09:00:00'))
    $rows = @($Rows | Where-Object { $null -ne $_ -and $null -ne $_.StartTime })
    if ($MinimumTimeOfDay -gt [timespan]::Zero) {
        $rows = @($rows | Where-Object { $_.StartTime.TimeOfDay -ge $MinimumTimeOfDay })
    }
    if ($rows.Count -eq 0) { return $null }
    return @($rows | Sort-Object StartTime -Descending)[0]
}

# ---------------------------------------------------------------------------
# Resolve-ProcessTimeRowPlan
#   Decides which stage(s) of the ProcessTime phase one mapping row still
#   needs, given -Stage (which stage(s) THIS RUN is allowed to touch: 'Ocr' |
#   'Write' | 'Both') and three completion signals:
#     -SidecarExists : a snap\ProcessTime\<correl>\result.json cache from a
#                      previous OCR pass exists for this correl (the
#                      per-correl, filesystem-based signal -- NOT the shared
#                      output workbook -- so a row's OCR-done state can be
#                      known without opening/trusting the single output
#                      .xlsx many rows write into).
#     -OcrDone       : bit 1 of the mapping's ProcessTime_Inserted column is
#                      already set (this correl's OCR result was cached AND
#                      recorded on the row, by this run or a previous one).
#     -WriteDone     : bit 2 of ProcessTime_Inserted is already set (the row
#                      was already written into an output workbook by a
#                      previous run).
#   ProcessTime_Inserted is a BITMASK (v2.15.0), matching this project's
#   isReplaced/isMarked/isReviewed convention: bit 1 = OCR'd, bit 2 =
#   written, 3 = both done. Callers compute -OcrDone/-WriteDone via
#   MappingStore.ps1's Test-BitDone against that column; a legacy plain '1'
#   value (pre-v2.15.0, meaning "written") is migrated to '3' by
#   Get-ProcessTimeMigratedInsertedValue before this function ever sees it,
#   so no special-casing of the old value is needed here.
#   OCR is considered done when EITHER the sidecar exists OR the OcrDone bit
#   is set -- a sidecar can exist without the bit yet being persisted (e.g.
#   a mapping save that failed after the sidecar write succeeded).
#   -Force ignores all completion signals for whichever stage(s) -Stage
#   selects.
#   Returns [pscustomobject]{ NeedsOcr; NeedsWrite; Touch } (Touch = either).
# ---------------------------------------------------------------------------
function Resolve-ProcessTimeRowPlan {
    param(
        [bool]$SidecarExists,
        [bool]$OcrDone,
        [bool]$WriteDone,
        [string]$Stage = 'Both',
        [bool]$Force = $false
    )
    $wantOcr   = ($Stage -eq 'Ocr')   -or ($Stage -eq 'Both')
    $wantWrite = ($Stage -eq 'Write') -or ($Stage -eq 'Both')
    $ocrIsDone = $SidecarExists -or $OcrDone
    $needsOcr   = $wantOcr   -and ($Force -or -not $ocrIsDone)
    $needsWrite = $wantWrite -and ($Force -or -not $WriteDone)
    return [pscustomobject]@{
        NeedsOcr   = $needsOcr
        NeedsWrite = $needsWrite
        Touch      = ($needsOcr -or $needsWrite)
    }
}

# ---------------------------------------------------------------------------
# Get-ProcessTimeMigratedInsertedValue
#   One-way migration of the ProcessTime_Inserted column from its pre-
#   v2.15.0 plain 0/1 shape (1 = written into the output workbook; OCR
#   completion was tracked only by sidecar-file existence, not on the row
#   itself) to the v2.15.0 bitmask shape (bit 1 = OCR'd, bit 2 = written).
#   A legacy '1' always meant "written", and writing can only happen after
#   OCR succeeded, so it is migrated to '3' (both bits) -- never just bit 2
#   -- so an old fully-done row does not appear to need a fresh OCR pass
#   under the new scheme. Any other value ('0', '', '2', '3', ...) passes
#   through unchanged (nothing else was ever written by the old code).
# ---------------------------------------------------------------------------
function Get-ProcessTimeMigratedInsertedValue {
    param([string]$Value)
    if ($Value -eq '1') { return '3' }
    return $Value
}

# ---------------------------------------------------------------------------
# Get-ProcessTimeOutputTag
#   Classifies one result row's output bucket from its mapping Excel_NAME,
#   by first-match order against -Tags (each tested as a literal substring,
#   not a regex -- Excel_NAME letters are plain ASCII tags like 'JDL'/'JRV'/
#   'JDS', not pattern metacharacters). Falls back to -UnclassifiedTag when
#   no configured tag matches, instead of throwing -- an operator's mapping
#   can legitimately contain an Excel_NAME the configured tag list does not
#   yet cover (e.g. a new project prefix), and one such row must never abort
#   writing every OTHER already-resolved tag's workbook.
# ---------------------------------------------------------------------------
function Get-ProcessTimeOutputTag {
    param([string]$ExcelName, [string[]]$Tags, [string]$UnclassifiedTag = 'Other')
    foreach ($t in @($Tags)) {
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        if ($ExcelName.IndexOf($t, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $t }
    }
    return $UnclassifiedTag
}

# ---------------------------------------------------------------------------
# Get-ProcessTimeOutputFileName
#   Output workbook file name for one tag: '<Label>(<Tag>).xlsx' in Split
#   mode, or a single untagged '<Label>.xlsx' when -Single (OutputMode
#   'Single' writes every result row into one workbook regardless of tag).
# ---------------------------------------------------------------------------
function Get-ProcessTimeOutputFileName {
    param([string]$Label, [string]$Tag, [bool]$Single = $false)
    if ($Single -or [string]::IsNullOrWhiteSpace($Tag)) { return ("{0}.xlsx" -f $Label) }
    return ("{0}({1}).xlsx" -f $Label, $Tag)
}

# ---------------------------------------------------------------------------
# Resolve-ProcessTimeOutputDir
#   Destination directory for one tag's output workbook: -DirByTag's entry
#   for -Tag when present and non-blank (ProcessTime.OutputDirectoryByTag --
#   lets each tag route to its own real J4-style destination folder instead
#   of every tag landing in the same directory), else -DefaultDir.
# ---------------------------------------------------------------------------
function Resolve-ProcessTimeOutputDir {
    param([string]$Tag, [hashtable]$DirByTag, [string]$DefaultDir)
    if ($null -ne $DirByTag -and $DirByTag.ContainsKey($Tag)) {
        $v = [string]$DirByTag[$Tag]
        if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
    }
    return $DefaultDir
}

# ---------------------------------------------------------------------------
# ConvertTo-ProcessTimeBucketArray
#   Materializes an output bucket without using @($Buckets[$tag]). Windows
#   PowerShell 5.1 can throw "Argument types do not match" when its dynamic
#   binder enumerates a generic List[object] obtained through a hashtable
#   index. Keeping this workaround in a pure helper makes the exact runtime
#   regression directly testable instead of relying on visual inspection of
#   ProcessTime.ps1.
# ---------------------------------------------------------------------------
function ConvertTo-ProcessTimeBucketArray {
    param([Parameter(Mandatory = $true)][System.Collections.Generic.List[object]]$Bucket)
    # Unary comma prevents PowerShell's function pipeline from unrolling the
    # array (and turning a one-row bucket back into a scalar).
    return ,$Bucket.ToArray()
}

# ---------------------------------------------------------------------------
# Get-ProcessTimeCheckSummaryLine
#   One-line, human-readable "needs manual check" note for a correl whose
#   GIFT and/or GFIX side was not matched -- printed in the end-of-run
#   summary so the operator knows exactly which ids to open and verify by
#   hand, instead of having to scroll back through the whole OCR log.
#   Returns '' when both sides matched (nothing to report).
# ---------------------------------------------------------------------------
function Get-ProcessTimeCheckSummaryLine {
    param([string]$CorrelId, [bool]$GiftMatched, [bool]$GfixMatched, [string]$GiftNote = '', [string]$GfixNote = '')
    $parts = New-Object System.Collections.Generic.List[string]
    if (-not $GiftMatched) {
        $parts.Add(("GIFT: {0}" -f $(if ([string]::IsNullOrWhiteSpace($GiftNote)) { 'not detected' } else { $GiftNote })))
    }
    if (-not $GfixMatched) {
        $parts.Add(("GFIX: {0}" -f $(if ([string]::IsNullOrWhiteSpace($GfixNote)) { 'not detected' } else { $GfixNote })))
    }
    if ($parts.Count -eq 0) { return '' }
    return ("{0}  -- {1}" -f $CorrelId, ($parts -join '; '))
}
