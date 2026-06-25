# ============================================================
#  SnapVerify.ps1
#
#  PURE snap-phase NG detection library -- NO Excel COM, NO
#  SendKeys, NO mapping I/O.  Dot-source only (no param() block).
#
#  Covers M1 (shared base for all SnapVerify milestones):
#    - HM page text parsing + abend verdict (F1)
#    - MQ page text parsing + record verdict (F2)
#    - Jenkins file list parsing + file verdict (F3/F4)
#    - Page-kind sentinel (A3)
#    - Batch run-time pure logic (2.2)
#
#  Convention: functions return plain arrays -- never return ,@(...)
#  because callers wrap calls in @() and that nests in PS 5.1.
# ============================================================

# ---------------------------------------------------------------------------
# Japanese string constants built from [char] code points (source stays ASCII)
# ---------------------------------------------------------------------------
$script:SV_Normal    = [char]0x6B63 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # seijo-shuuryo (normal end)
$script:SV_Abend     = [char]0x7570 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # ijo-shuuryo (abend)
# HM page title: batch shori jokyo ichiran
$script:SV_HmTitle   = [char]0x30D0 + [char]0x30C3 + [char]0x30C1 + [char]0x51E6 `
                     + [char]0x7406 + [char]0x72B6 + [char]0x6CC1 + [char]0x4E00 + [char]0x89A7
# HM table header col1: kaishi-nichiji (start datetime)
$script:SV_HmHdrCol1 = [char]0x958B + [char]0x59CB + [char]0x65E5 + [char]0x6642
# HM table header col2: shuuryo-nichiji (end datetime)
$script:SV_HmHdrCol2 = [char]0x7D42 + [char]0x4E86 + [char]0x65E5 + [char]0x6642
# Jenkins file-list row suffix: sansho (reference)
$script:SV_Ref       = [char]0x53C2 + [char]0x7167
# GIFT System outer-frame signature
$script:SV_GiftSys   = 'GIFT System'
# MQ result page signatures
$script:SV_MqResult  = 'Transfer status inquiry results'
$script:SV_MqNumRec  = 'Number of records'
$script:SV_NoData    = 'No Data!'

# ---------------------------------------------------------------------------
# ConvertFrom-HmPageText
#   Parses the Ctrl+A clipboard text of an HM batch-status page.
#   Returns an array of PSCustomObjects, one per data row.
#
#   Field layout (TAB-separated per spec appendix A):
#     [0] start-dt  [1] end-dt  [2] proc-time  [3] batch-id  [4] SS
#     [5..] variable -- status = FIRST field equal to normal-end/abend
#     LAST field = correl-id
#
#   Returns objects with:
#     StartTime  [datetime|$null]
#     EndTime    [datetime|$null]
#     BatchId    [string]
#     Status     [string]  normal-end | abend
#     CorrelId   [string]
#     RawLine    [string]
# ---------------------------------------------------------------------------
function ConvertFrom-HmPageText {
    param([string]$Text)

    $dtFmt   = 'yyyy/MM/dd HH:mm:ss'
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($rawLine in ($Text -split "`r?`n")) {
        # Data rows start with a datetime followed immediately by a TAB
        if ($rawLine -notmatch '^\d{4}/\d{2}/\d{2} \d{2}:\d{2}:\d{2}\t') { continue }

        $fields = $rawLine -split "`t"
        if ($fields.Count -lt 7) { continue }

        # Parse start / end times
        $startTime = $null; $endTime = $null
        try { $startTime = [datetime]::ParseExact($fields[0].Trim(), $dtFmt, $culture) } catch {}
        try { $endTime   = [datetime]::ParseExact($fields[1].Trim(), $dtFmt, $culture) } catch {}

        # status: first field matching normal-end or abend
        $status = ''
        foreach ($f in $fields) {
            if ($f -eq $script:SV_Normal -or $f -eq $script:SV_Abend) {
                $status = $f; break
            }
        }
        if (-not $status) { continue }  # not a recognised status row

        # correl-id: last non-empty field (trim trailing empties)
        $correlId = ''
        for ($i = $fields.Count - 1; $i -ge 0; $i--) {
            if ($fields[$i].Trim() -ne '') { $correlId = $fields[$i].Trim(); break }
        }

        $results.Add([PSCustomObject]@{
            StartTime = $startTime
            EndTime   = $endTime
            BatchId   = $fields[3].Trim()
            Status    = $status
            CorrelId  = $correlId
            RawLine   = $rawLine
        })
    }

    # Return plain array (never comma-wrapped -- see file header note)
    return $results.ToArray()
}

# ---------------------------------------------------------------------------
# Test-HmAbend
#   Given parsed HM rows, a target CorrelId, an optional Expected time, and a
#   tolerance, returns a verdict hashtable.
#
#   Parameters:
#     Rows           PSCustomObject[]  output of ConvertFrom-HmPageText
#     CorrelId       string            Correl_ID_S to filter on
#     Expected       datetime|$null    $null = no time check
#     ToleranceMin   int               minutes either side of Expected
#
#   Returns @{ Verdict='ok'|'ng'|'warn'|'ask'; Reason=[string]; Warnings=[string[]] }
#
#   Verdict rules (per spec 2.3 / 4.F1):
#     - 0 rows for CorrelId          -> ask
#     - Expected is $null (no-time)  -> any abend row -> ask; else ok
#     - Expected set:
#         windowRows = rows with |StartTime - Expected| <= ToleranceMin
#         0 window rows              -> ask
#         newest (by StartTime) is normal-end -> ok  (warn about earlier window abends)
#         newest is abend             -> ng
#       Out-of-window abend rows always generate a warn entry (historic records).
# ---------------------------------------------------------------------------
function Test-HmAbend {
    param(
        [object[]]$Rows,
        [string]$CorrelId,
        [object]$Expected,        # [datetime] or $null
        [int]$ToleranceMin = 30
    )

    $matchRows = @($Rows | Where-Object { $_.CorrelId -eq $CorrelId })

    if ($matchRows.Count -eq 0) {
        return @{ Verdict = 'ask'; Reason = "no rows found for correl $CorrelId"; Warnings = @() }
    }

    # No time check mode
    if ($null -eq $Expected) {
        $abendRows = @($matchRows | Where-Object { $_.Status -eq $script:SV_Abend })
        if ($abendRows.Count -gt 0) {
            return @{ Verdict = 'ask'; Reason = 'abend row found; no time window to determine if current'; Warnings = @() }
        }
        return @{ Verdict = 'ok'; Reason = 'all rows normal termination'; Warnings = @() }
    }

    # Time-window mode
    $windowRows  = @($matchRows | Where-Object {
        $null -ne $_.StartTime -and [Math]::Abs(($_.StartTime - $Expected).TotalMinutes) -le $ToleranceMin
    })
    $outsideRows = @($matchRows | Where-Object {
        $null -eq $_.StartTime -or [Math]::Abs(($_.StartTime - $Expected).TotalMinutes) -gt $ToleranceMin
    })

    $warnings = [System.Collections.Generic.List[string]]::new()
    foreach ($r in ($outsideRows | Where-Object { $_.Status -eq $script:SV_Abend })) {
        $warnings.Add(("historic abend outside window: {0}" -f $r.StartTime))
    }

    if ($windowRows.Count -eq 0) {
        return @{ Verdict = 'ask'; Reason = 'no rows inside time window'; Warnings = $warnings.ToArray() }
    }

    # Sort window rows by StartTime descending; newest = first
    $sorted  = @($windowRows | Sort-Object { $_.StartTime } -Descending)
    $newest  = $sorted[0]

    if ($newest.Status -eq $script:SV_Normal) {
        # Any earlier abend in window -> warn but verdict is ok
        $earlierAbend = @($sorted | Select-Object -Skip 1 | Where-Object { $_.Status -eq $script:SV_Abend })
        foreach ($r in $earlierAbend) {
            $warnings.Add(("retried, last run ok; earlier window abend: {0}" -f $r.StartTime))
        }
        return @{ Verdict = 'ok'; Reason = 'newest run in window is normal termination'; Warnings = $warnings.ToArray() }
    }

    # Newest is abend
    return @{ Verdict = 'ng'; Reason = ("abend is the most recent run in window: {0}" -f $newest.StartTime); Warnings = $warnings.ToArray() }
}

# ---------------------------------------------------------------------------
# ConvertFrom-MqPageText
#   Parses the Ctrl+A clipboard text of an MQ transfer-status page.
#   Absorbs Parse-GiftMq.ps1 regex logic as a unit-testable function.
#
#   Format (Appendix B-1): records occupy two lines each.
#   Line 1: No<TAB>SendNode<TAB>RecvNode<TAB>CorrelId<TAB>SendDate<TAB>
#           Tmode<TAB>RecvDate<TAB>Rtncd<TAB>Rsncd
#   Line 2: Msgid<TAB>Reccnt<TAB>FileSize  (not parsed; reserved for future)
#
#   Returns @{ NumRecords=[int]; Rows=[object[]] }
#   Each row: No, SendNode, RecvNode, CorrelId, SendDate, Tmode, RecvDate, Rtncd, Rsncd
# ---------------------------------------------------------------------------
function ConvertFrom-MqPageText {
    param([string]$Text)

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $dtFmt   = 'yyyy/MM/dd HH:mm:ss'

    $m = [regex]::Match($Text, "$($script:SV_MqNumRec)\s+(\d+)")
    $numRec = if ($m.Success) { [int]$m.Groups[1].Value } else { -1 }

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($line in ($Text -split "`r?`n")) {
        $line = $line.Trim()
        # Matches line-1 of each record (TAB or space separated, trailing whitespace ok)
        if ($line -match '^(\d+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2})\s+(\S+)\s+(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2})\s+(\d+)\s+(\d+)\s*$') {
            $sendDt = $null; $recvDt = $null
            try { $sendDt = [datetime]::ParseExact($Matches[5], $dtFmt, $culture) } catch {}
            try { $recvDt = [datetime]::ParseExact($Matches[7], $dtFmt, $culture) } catch {}
            $rows.Add([PSCustomObject]@{
                No       = [int]$Matches[1]
                SendNode = $Matches[2]
                RecvNode = $Matches[3]
                CorrelId = $Matches[4]
                SendDate = $sendDt
                Tmode    = $Matches[6]
                RecvDate = $recvDt
                Rtncd    = [int]$Matches[8]
                Rsncd    = [int]$Matches[9]
            })
        }
    }

    return @{ NumRecords = $numRec; Rows = $rows.ToArray() }
}

# ---------------------------------------------------------------------------
# Test-MqRecord
#   Given parsed MQ data, a target CorrelId, optional Expected time, and
#   tolerance, returns a verdict hashtable.
#
#   NG conditions (per spec 2.4 -- all three checked):
#     1. No matching row, OR page was MqNoData            -> ng
#     2. Matching row(s) found but RecvDate outside window -> ng
#     3. Rtncd or Rsncd non-zero                           -> ng
#
#   When multiple rows share the same CorrelId, newest-wins by RecvDate.
#
#   Parameters:
#     Parsed       output of ConvertFrom-MqPageText (@{ NumRecords; Rows })
#     CorrelId     string
#     Expected     datetime|$null  ($null = no time check, skip condition 2)
#     ToleranceMin int
#     IsNoData     bool  pass $true when page kind is MqNoData
#
#   Returns @{ Verdict='ok'|'ng'; Reason=[string]; MatchedRow=[object]|$null }
# ---------------------------------------------------------------------------
function Test-MqRecord {
    param(
        [hashtable]$Parsed,
        [string]$CorrelId,
        [object]$Expected,
        [int]$ToleranceMin = 30,
        [bool]$IsNoData = $false
    )

    # Condition 1a: page showed "No Data!"
    if ($IsNoData) {
        return @{ Verdict = 'ng'; Reason = 'page shows No Data!'; MatchedRow = $null }
    }

    $matchRows = @($Parsed.Rows | Where-Object { $_.CorrelId -eq $CorrelId })

    # Condition 1b: no matching row
    if ($matchRows.Count -eq 0) {
        return @{ Verdict = 'ng'; Reason = "no row found for correl $CorrelId"; MatchedRow = $null }
    }

    # Newest-wins: pick row with latest RecvDate
    $target = ($matchRows | Sort-Object { $_.RecvDate } -Descending)[0]

    # Condition 2: time window (skip when Expected is $null)
    if ($null -ne $Expected) {
        $diffMin = [Math]::Abs(($target.RecvDate - $Expected).TotalMinutes)
        if ($diffMin -gt $ToleranceMin) {
            return @{
                Verdict    = 'ng'
                Reason     = ("RecvDate {0} is {1:F1} min outside window of Expected {2} +-{3} min" -f
                              $target.RecvDate, $diffMin, $Expected, $ToleranceMin)
                MatchedRow = $target
            }
        }
    }

    # Condition 3: non-zero return codes
    if ($target.Rtncd -ne 0 -or $target.Rsncd -ne 0) {
        return @{
            Verdict    = 'ng'
            Reason     = ("non-zero return code: Rtncd={0} Rsncd={1}" -f $target.Rtncd, $target.Rsncd)
            MatchedRow = $target
        }
    }

    return @{ Verdict = 'ok'; Reason = 'record found, in window, return codes zero'; MatchedRow = $target }
}

# ---------------------------------------------------------------------------
# ConvertFrom-JenkinsListText
#   Parses the Ctrl+A text of a Jenkins file-list page.
#   Absorbs Parse-JenkinsList.ps1 regex logic as a unit-testable function.
#   Returns a plain array of PSCustomObjects: Name, DateTime, Size.
# ---------------------------------------------------------------------------
function ConvertFrom-JenkinsListText {
    param([string]$Text)

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $refWord = $script:SV_Ref  # sansho (reference)
    $results = [System.Collections.Generic.List[object]]::new()

    foreach ($rawLine in ($Text -split "`r?`n")) {
        $line = $rawLine.Trim()
        if ($line -match ('^(\S+)\s+(\d{4}/\d{2}/\d{2})\s+(\d{2}:\d{2}:\d{2})\s+(.+?)\s+' + $refWord + '$')) {
            $dt = $null
            try {
                $dt = [datetime]::ParseExact(
                    ("{0} {1}" -f $Matches[2], $Matches[3]),
                    'yyyy/MM/dd HH:mm:ss', $culture)
            } catch {}
            $results.Add([PSCustomObject]@{
                Name     = $Matches[1]
                DateTime = $dt
                Size     = $Matches[4]
            })
        }
    }

    return $results.ToArray()
}

# ---------------------------------------------------------------------------
# Test-JenkinsFile
#   Given parsed Jenkins file entries, checks whether the expected file
#   exists (or does not exist for NoGfix mode) within the time window.
#
#   Parameters:
#     Files        object[]  output of ConvertFrom-JenkinsListText
#     CorrelId     string    filename to look for (exact match)
#     Expected     datetime|$null
#     ToleranceMin int
#     ExpectExists bool      $true = file should exist (F3); $false = NoGfix (F4)
#
#   Returns @{ Verdict='ok'|'ng'; Reason=[string]; File=[object]|$null }
# ---------------------------------------------------------------------------
function Test-JenkinsFile {
    param(
        [object[]]$Files,
        [string]$CorrelId,
        [object]$Expected,
        [int]$ToleranceMin = 30,
        [bool]$ExpectExists = $true
    )

    $target = $Files | Where-Object { $_.Name -eq $CorrelId } | Select-Object -First 1

    if (-not $ExpectExists) {
        # NoGfix mode: expecting NO file for this correl
        if ($null -eq $target) {
            return @{ Verdict = 'ok'; Reason = 'no file found as expected (NoGfix)'; File = $null }
        }
        return @{
            Verdict = 'ng'
            Reason  = ("unexpected file found (may be past data): {0} {1}" -f $target.Name, $target.DateTime)
            File    = $target
        }
    }

    # Normal mode: file should exist
    if ($null -eq $target) {
        return @{ Verdict = 'ng'; Reason = 'file not in list'; File = $null }
    }

    # No time check
    if ($null -eq $Expected) {
        return @{ Verdict = 'ok'; Reason = 'file found (no time check)'; File = $target }
    }

    $diffMin = [Math]::Abs(($target.DateTime - $Expected).TotalMinutes)
    if ($diffMin -gt $ToleranceMin) {
        return @{
            Verdict = 'ng'
            Reason  = ("file found but DateTime {0} is {1:F1} min outside window" -f $target.DateTime, $diffMin)
            File    = $target
        }
    }

    return @{ Verdict = 'ok'; Reason = 'file found within time window'; File = $target }
}

# ---------------------------------------------------------------------------
# Get-SnapPageKind
#   Classifies the Ctrl+A text of a snap-phase page into one of:
#     HmResult      HM batch-status page (expected for HM phase)
#     MqResult      MQ transfer-status result page (expected for MQ phase)
#     MqNoData      MQ page with no records (legitimate MQ terminal state)
#     JenkinsResult Jenkins file-list page (expected for Jenkins phase)
#     OuterFrame    GIFT outer-menu frame (wrong focus)
#     Empty         Text is blank/whitespace (clipboard read failed)
#     Unknown       Does not match any known pattern
#
#   Parameters:
#     Phase   string  'Hm' | 'Mq' | 'Jenkins'  (informational; not used to filter)
#     Text    string
#
#   Returns a string (one of the kind names above).
# ---------------------------------------------------------------------------
function Get-SnapPageKind {
    param(
        [string]$Phase,
        [string]$Text
    )

    # Empty / whitespace
    if ([string]::IsNullOrWhiteSpace($Text)) { return 'Empty' }

    $trimmed = $Text.Trim()

    # MqNoData: full text (trimmed) is exactly "No Data!"
    if ($trimmed -eq $script:SV_NoData) { return 'MqNoData' }

    # OuterFrame: contains the GIFT system outer-menu signature
    if ($Text.Contains($script:SV_GiftSys)) { return 'OuterFrame' }

    # HmResult: contains the HM page title or the table header (start-dt<TAB>end-dt)
    $hmHeader = $script:SV_HmHdrCol1 + "`t" + $script:SV_HmHdrCol2
    if ($Text.Contains($script:SV_HmTitle) -or $Text.Contains($hmHeader)) { return 'HmResult' }

    # MqResult: contains MQ result page markers
    if ($Text.Contains($script:SV_MqResult) -or $Text.Contains($script:SV_MqNumRec)) { return 'MqResult' }

    # JenkinsResult: contains at least one sansho/reference (file-list row suffix)
    if ($Text.Contains($script:SV_Ref)) { return 'JenkinsResult' }

    return 'Unknown'
}

# ---------------------------------------------------------------------------
# Resolve-SnapRunTime
#   Pure logic for the batch run-time inquiry (spec 2.2).
#   Parses user's typed input (already captured by the caller via Read-Host)
#   and returns a validated time configuration.
#
#   Parameters:
#     TimeInput         string   '' or [Enter] = use Now; 'n'/'N' = no time check;
#                                'yyyy/MM/dd HH:mm:ss' / 'yyyy/MM/dd HH:mm' /
#                                'yyyy/MM/dd' = fixed date-time; 'HH:mm:ss' /
#                                'HH:mm' (1- or 2-digit hour) = time-only, anchored
#                                to Now's date. Anything else -> Ok=$false.
#     ToleranceInput    string   '' = use DefaultToleranceMin; digit string = override
#     DefaultTolerance  int      default tolerance in minutes (typically 30)
#     Now               datetime reference clock (default: Get-Date)
#
#   Returns @{ Ok=[bool]; TimeMode='fixed'|'none'; Time=[datetime]|$null;
#              ToleranceMinutes=[int]; Error=[string] }
# ---------------------------------------------------------------------------
function Resolve-SnapRunTime {
    param(
        [string]$TimeInput,
        [string]$ToleranceInput = '',
        [int]$DefaultTolerance  = 30,
        [datetime]$Now          = (Get-Date)
    )

    # Parse tolerance. A digit string overrides the default; blank or anything
    # non-numeric keeps the default (so a stray Enter never zeroes tolerance).
    $tol     = $DefaultTolerance
    $tolTrim = if ($null -eq $ToleranceInput) { '' } else { $ToleranceInput.Trim() }
    if ($tolTrim -match '^\d+$') { $tol = [int]$tolTrim }

    # No time check
    if ($TimeInput -match '^[Nn]$') {
        return @{ Ok = $true; TimeMode = 'none'; Time = $null; ToleranceMinutes = $tol; Error = '' }
    }

    # Empty / whitespace -> use current time
    if ([string]::IsNullOrWhiteSpace($TimeInput)) {
        return @{ Ok = $true; TimeMode = 'fixed'; Time = $Now; ToleranceMinutes = $tol; Error = '' }
    }

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $styles  = [System.Globalization.DateTimeStyles]::None
    $t       = $TimeInput.Trim()
    $parsed  = [datetime]::MinValue

    # 1) Full date+time formats are used as-is.
    $dateTimeFmts = @('yyyy/MM/dd HH:mm:ss', 'yyyy/MM/dd HH:mm', 'yyyy/MM/dd')
    foreach ($fmt in $dateTimeFmts) {
        if ([datetime]::TryParseExact($t, $fmt, $culture, $styles, [ref]$parsed)) {
            return @{ Ok = $true; TimeMode = 'fixed'; Time = $parsed; ToleranceMinutes = $tol; Error = '' }
        }
    }

    # 2) Time-only formats (HH:mm:ss / HH:mm, 1- or 2-digit hour) are anchored
    #    to today's date ($Now.Date) -- the common case where the operator only
    #    types the run clock time.
    $timeOnlyFmts = @('HH:mm:ss', 'H:mm:ss', 'HH:mm', 'H:mm')
    $tod = [datetime]::MinValue
    foreach ($fmt in $timeOnlyFmts) {
        if ([datetime]::TryParseExact($t, $fmt, $culture, $styles, [ref]$tod)) {
            $combined = $Now.Date.Add($tod.TimeOfDay)
            return @{ Ok = $true; TimeMode = 'fixed'; Time = $combined; ToleranceMinutes = $tol; Error = '' }
        }
    }

    return @{
        Ok               = $false
        TimeMode         = 'none'
        Time             = $null
        ToleranceMinutes = $tol
        Error            = "cannot parse time input '$TimeInput'; expected yyyy/MM/dd HH:mm:ss, HH:mm:ss, HH:mm, or n"
    }
}

# ---------------------------------------------------------------------------
# ConvertTo-ExpectedDateTime
#   Parses a per-row Expected_Time cell into a [datetime], or $null when the
#   cell is empty/blank/unparseable. Used by the snap phases to turn the
#   mapping's Expected_Time column value into the -Expected argument of the
#   Test-* verdict functions. A $null result means "no time window" for that
#   row (the verdict functions skip the window check when Expected is $null).
#
#   Parameters:
#     Value   string    the raw cell text (may be '', whitespace, or a date)
#     Format  string    preferred parse format (others are tried as fallback)
#
#   Returns [datetime] or $null.
# ---------------------------------------------------------------------------
function ConvertTo-ExpectedDateTime {
    param(
        [string]$Value,
        [string]$Format = 'yyyy/MM/dd HH:mm:ss'
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $parsed  = [datetime]::MinValue
    $fmts    = @($Format, 'yyyy/MM/dd HH:mm:ss', 'yyyy/MM/dd HH:mm', 'yyyy/MM/dd')
    foreach ($f in $fmts) {
        if ([datetime]::TryParseExact($Value.Trim(), $f, $culture,
                [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            return $parsed
        }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Set-EmptyRunTimeCells
#   Batch-applies a run time to every row whose -Field cell is empty/blank,
#   leaving rows that already carry a value untouched (spec 2.2 step 3:
#   "Expected_Time empty rows get the chosen time; existing values kept").
#   Pure in-memory mutation -- the caller persists via Export-MappingAtomic.
#
#   Parameters:
#     Rows    object[]  rows to fill (typically the pending subset)
#     Field   string    column name (e.g. 'Expected_Time')
#     Value   string    the time string to write into empty cells
#
#   Returns [int] the number of cells filled.
# ---------------------------------------------------------------------------
function Set-EmptyRunTimeCells {
    param(
        [object[]]$Rows,
        [string]$Field,
        [string]$Value
    )
    $filled = 0
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $has = ($r.PSObject.Properties.Name -contains $Field)
        $cur = if ($has) { [string]$r.$Field } else { '' }
        if ([string]::IsNullOrWhiteSpace($cur)) {
            if ($has) {
                $r.$Field = $Value
            } else {
                $r | Add-Member -NotePropertyName $Field -NotePropertyValue $Value -Force
            }
            $filled++
        }
    }
    return $filled
}

# ===========================================================================
#  M5 / F5 -- object-data pixel localisation (pure geometry, no image I/O)
#
#  Produces a rectangle (screenshot pixel coords) for the data row a snap
#  verdict is about, so the Mark phase can draw a red box exactly on it.
#  Two geometry legs (per docs/SnapVerify-Plan.md sec 4.F5):
#    - HM / MQ form pages : fixed geometry Row1Top + (n-1)*RowHeight, where the
#      screen row n comes from the clipboard parse order (Get-MatchedRowIndex).
#      Same model as Find-Abend.ps1 (Calibrate-HmGeometry.ps1 measures it).
#    - Jenkins list page  : the Ctrl+F active-match band measured from the PNG
#      by Find-ActiveHighlightRow.ps1; Get-JenkinsHighlightRect turns the
#      Top/Bottom band into a rect (pure; the image scan stays in the wiring).
#  The result is written next to the PNG as <correl>.loc.json by the snap
#  scripts; ReplaceEvidence -> AltText -> Mark consumes it (M6 pipeline).
# ===========================================================================

# ---------------------------------------------------------------------------
# Get-MatchedRowIndex
#   Returns the 1-based SCREEN row index (position in parse order, which equals
#   the on-screen top-to-bottom order) of the data row a verdict selects, so the
#   pixel box lands on the physically-displayed row. The SELECTION mirrors
#   Test-HmAbend / Test-MqRecord (newest-wins inside the time window, else
#   newest overall); the RETURNED index is the row's original parse position,
#   NOT its rank after the newest-wins re-sort.
#
#   Parameters:
#     Rows          object[]  parser output (ConvertFrom-HmPageText / -MqPageText)
#     CorrelId      string    correl to filter on
#     DateProperty  string    'StartTime' (HM) | 'RecvDate' (MQ)
#     Expected      datetime|$null  $null = no window (newest overall)
#     ToleranceMin  int       window half-width in minutes
#
#   Returns [int]: the 1-based screen index, or 0 when no row matches CorrelId.
# ---------------------------------------------------------------------------
function Get-MatchedRowIndex {
    param(
        [object[]]$Rows,
        [string]$CorrelId,
        [string]$DateProperty = 'StartTime',
        [object]$Expected     = $null,
        [int]$ToleranceMin    = 30
    )

    if ($null -eq $Rows) { return 0 }

    # Collect matches keeping their original (screen) position.
    $matched = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Rows.Count; $i++) {
        $r = $Rows[$i]
        if ($null -eq $r) { continue }
        if ([string]$r.CorrelId -eq $CorrelId) {
            $matched.Add([PSCustomObject]@{ Index = $i + 1; Date = $r.$DateProperty })
        }
    }
    if ($matched.Count -eq 0) { return 0 }

    # Candidates that carry a usable timestamp.
    $dated = @($matched | Where-Object { $null -ne $_.Date })
    if ($dated.Count -eq 0) {
        # No timestamps to rank by -- box the first matched row.
        return [int]$matched[0].Index
    }

    # Time window: newest inside it wins; if none inside, fall back to newest overall.
    if ($null -ne $Expected) {
        $window = @($dated | Where-Object {
            [Math]::Abs(($_.Date - $Expected).TotalMinutes) -le $ToleranceMin
        })
        if ($window.Count -gt 0) {
            return [int](@($window | Sort-Object { $_.Date } -Descending)[0]).Index
        }
    }

    return [int](@($dated | Sort-Object { $_.Date } -Descending)[0]).Index
}

# ---------------------------------------------------------------------------
# Get-RowPixelRect
#   Fixed-geometry rectangle for an HM / MQ form-page data row (same model as
#   Find-Abend.ps1: roiTop = Row1Top + (RowIndex-1)*RowHeight). Coordinates are
#   returned in the FINAL (already-cropped) PNG pixel space: pass the per-side
#   crop pixels removed by Invoke-CropPng so the rect tracks the cropped image.
#
#   Parameters:
#     RowIndex   int  1-based screen row (from Get-MatchedRowIndex)
#     Row1Top    int  top pixel of row 1 in the pre-crop capture
#     RowHeight  int  per-row height in pixels (MQ: a 2-line record height)
#     ColLeft    int  left pixel of the column to box (HM status / MQ recv-date)
#     ColWidth   int  width of that column in pixels
#     CropLeft   int  pixels removed from the left by Invoke-CropPng (default 0)
#     CropTop    int  pixels removed from the top  by Invoke-CropPng (default 0)
#
#   Returns @{ x; y; w; h } (ints; x/y clamped to >= 0). Throws on bad geometry.
# ---------------------------------------------------------------------------
function Get-RowPixelRect {
    param(
        [int]$RowIndex,
        [int]$Row1Top,
        [int]$RowHeight,
        [int]$ColLeft,
        [int]$ColWidth,
        [int]$CropLeft = 0,
        [int]$CropTop  = 0
    )
    if ($RowIndex -lt 1) { throw "Get-RowPixelRect: RowIndex must be >= 1 (got $RowIndex)" }
    if ($RowHeight -le 0) { throw "Get-RowPixelRect: RowHeight must be > 0 (got $RowHeight)" }
    if ($ColWidth -le 0)  { throw "Get-RowPixelRect: ColWidth must be > 0 (got $ColWidth)" }

    $rowTop = $Row1Top + ($RowIndex - 1) * $RowHeight
    $x = $ColLeft - $CropLeft
    $y = $rowTop  - $CropTop
    if ($x -lt 0) { $x = 0 }
    if ($y -lt 0) { $y = 0 }

    return @{ x = [int]$x; y = [int]$y; w = [int]$ColWidth; h = [int]$RowHeight }
}

# ---------------------------------------------------------------------------
# Get-JenkinsHighlightRect
#   Turns a Find-ActiveHighlightRow.ps1 band (@{ Top; Bottom }) into a rect.
#   The orange Ctrl+F active-match spans the whole matched file row; box the
#   file-time column horizontally (ColLeft/ColWidth) or the full row width.
#
#   Parameters:
#     Top         int  top pixel of the highlight band
#     Bottom      int  bottom pixel of the highlight band (>= Top)
#     ColLeft     int  left pixel of the box (default 0)
#     ColWidth    int  box width; 0 = derive from ImageWidth - ColLeft
#     ImageWidth  int  PNG width (used when ColWidth = 0)
#     Pad         int  pixels added above/below the band (default 0)
#
#   Returns @{ x; y; w; h } (ints; x/y clamped to >= 0). Throws on bad input.
# ---------------------------------------------------------------------------
function Get-JenkinsHighlightRect {
    param(
        [int]$Top,
        [int]$Bottom,
        [int]$ColLeft    = 0,
        [int]$ColWidth   = 0,
        [int]$ImageWidth = 0,
        [int]$Pad        = 0
    )
    if ($Bottom -lt $Top) { throw "Get-JenkinsHighlightRect: Bottom ($Bottom) < Top ($Top)" }

    $w = 0
    if ($ColWidth -gt 0) {
        $w = $ColWidth
    } elseif ($ImageWidth -gt 0) {
        $w = $ImageWidth - $ColLeft
    }
    if ($w -le 0) { throw "Get-JenkinsHighlightRect: provide ColWidth or ImageWidth (> ColLeft)" }

    $x = $ColLeft;        if ($x -lt 0) { $x = 0 }
    $y = $Top - $Pad;     if ($y -lt 0) { $y = 0 }
    $h = ($Bottom - $Top + 1) + (2 * $Pad)

    return @{ x = [int]$x; y = [int]$y; w = [int]$w; h = [int]$h }
}

# ---------------------------------------------------------------------------
# New-SnapLocRect
#   Builds the .loc.json sidecar payload for one localised row. The minimum
#   contract is { x, y, w, h } in screenshot pixel space (plan 4.F5); the extra
#   fields let the Mark phase convert pixels -> points (it needs imageWidth to
#   compute Shape.Width / imageWidth) and let reviewers trace the source.
#
#   Parameters:
#     CorrelId / X / Y / W / H
#     Source        string  'hm-geometry' | 'mq-geometry' | 'jenkins-highlight'
#     RowIndex      int     screen row the box came from (0 = n/a)
#     ImageWidth    int     final PNG width  (for pixel->point scaling)
#     ImageHeight   int     final PNG height
#     CreatedUtc    string  ISO-8601 stamp; '' = now (injectable for tests)
#
#   Returns an [ordered] hashtable ready for ConvertTo-Json.
# ---------------------------------------------------------------------------
function New-SnapLocRect {
    param(
        [string]$CorrelId,
        [int]$X,
        [int]$Y,
        [int]$W,
        [int]$H,
        [string]$Source      = 'geometry',
        [int]$RowIndex       = 0,
        [int]$ImageWidth     = 0,
        [int]$ImageHeight    = 0,
        [string]$CreatedUtc  = ''
    )
    $stamp = if ([string]::IsNullOrWhiteSpace($CreatedUtc)) {
        (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    } else { $CreatedUtc }

    return [ordered]@{
        correl      = $CorrelId
        source      = $Source
        rowIndex    = [int]$RowIndex
        x           = [int]$X
        y           = [int]$Y
        w           = [int]$W
        h           = [int]$H
        imageWidth  = [int]$ImageWidth
        imageHeight = [int]$ImageHeight
        created     = $stamp
    }
}

# ---------------------------------------------------------------------------
# Save-SnapLocSidecar
#   Writes a New-SnapLocRect payload to <Path> as UTF-8 (no BOM) JSON, the same
#   encoding the snap scripts use for the A1 .txt archive. Thin I/O wrapper so
#   the wiring stays one line; the geometry itself is the pure part above.
#   Returns the path written.
# ---------------------------------------------------------------------------
function Save-SnapLocSidecar {
    param(
        [Parameter(Mandatory)] $Loc,
        [Parameter(Mandatory)][string]$Path
    )
    $json = $Loc | ConvertTo-Json -Depth 5
    $enc  = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $json, $enc)
    return $Path
}
