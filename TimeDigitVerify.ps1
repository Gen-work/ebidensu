# ============================================================
#  TimeDigitVerify.ps1
#
#  PURE 3<->9 digit-risk analysis for OCR'd HM processing times.
#  NO Excel COM, NO OCR, NO file I/O. Dot-source only (no param()).
#  ASCII source (this module needs no Japanese literals at all).
#
#  WHY THIS EXISTS
#  ---------------
#  The Windows ja recognizer misreads MS Gothic '9' as '3' (and back) in
#  the HM batch-status page's time columns. The misread yields a
#  FORMAT-VALID timestamp, so it cannot be spotted by a shape check.
#
#  Earlier attempts leaned the other way and were too eager: a heuristic
#  that "corrects" a plausibly-read digit can turn a CORRECT reading into
#  a wrong one, and silently overriding the page's own printed processing
#  time with one derived from a suspect digit shipped wrong values
#  (observed: start ss=02, end ss=03, page duration 00:00:01 -- the end
#  was misread as 09, the derived 00:00:07 replaced the printed 00:00:01).
#
#  So this module only ever does what is DETERMINISTIC, and flags the
#  rest for a human:
#
#    1. IMPOSSIBLE-POSITION REPAIR (Repair-ImpossibleTimeDigit)
#       A digit that makes its field fall outside the field's legal range
#       is repaired ONLY when exactly one single 3<->9 substitution inside
#       that field brings it back in range. '10:93:20' -> '10:33:20' is
#       forced (a minute tens digit cannot be 9); '2026/19/03' is left
#       alone (no 3<->9 swap makes month 19 legal). Never a guess.
#
#    2. ARITHMETIC DISAMBIGUATION (Resolve-ProcessTimeDurationConflict)
#       start, end and the page's own processing-time column are three
#       independent readings of the same fact. When they disagree and
#       EXACTLY ONE single 3<->9 substitution across start/end reconciles
#       them, that substitution is forced by arithmetic, not guessed --
#       apply it. Zero or several reconciling candidates means we do not
#       know: change nothing, flag the row.
#
#    3. EVERYTHING ELSE IS FLAGGED, NOT FIXED (Get-TimeDigitRisk)
#       A 3 or 9 sitting where the opposite digit would ALSO be legal is
#       an undetectable misread. Those cells are marked by the output
#       workbook's conditional formatting (Get-ProcessTimeDigitFormatRule)
#       so the operator can check them against the snap image. Their
#       VALUES are never touched.
#
#  Unit-tested by Tests\Test-TimeDigitVerify.ps1.
# ============================================================

# The only OCR confusion this module models. Returns the opposite digit,
# or $null when the character is not part of the confusable pair.
function Get-TimeDigitSwapChar {
    param($Char)
    switch ([string]$Char) {
        '3'     { return '9' }
        '9'     { return '3' }
        default { return $null }
    }
}

# Legal range of one named time field. DurHour is deliberately unbounded
# in practice (Get-ProcessDurationText does not clamp a duration to 24h).
function Get-TimeDigitFieldLimit {
    param([string]$Name)
    switch ([string]$Name) {
        'Year'    { return @{ Min = 1900; Max = 2999 } }
        'Month'   { return @{ Min = 1;    Max = 12 } }
        'Day'     { return @{ Min = 1;    Max = 31 } }
        'Hour'    { return @{ Min = 0;    Max = 23 } }
        'Minute'  { return @{ Min = 0;    Max = 59 } }
        'Second'  { return @{ Min = 0;    Max = 59 } }
        'DurHour' { return @{ Min = 0;    Max = 9999 } }
        default   { return @{ Min = 0;    Max = 9999 } }
    }
}

# Classifies a time-ish string so the field spec knows which layout to use.
# 'HH:mm:ss' is reported as TimeOfDay; a caller holding a DURATION must say
# so with -Kind Duration, otherwise a legitimate 30-hour duration looks like
# an out-of-range hour.
function Get-TimeDigitTextKind {
    param([string]$Text)
    $t = ([string]$Text).Trim()
    if ([string]::IsNullOrEmpty($t)) { return 'Unknown' }
    if ($t -match '^\d{4}/\d{2}/\d{2}[ \t]+\d{2}:\d{2}:\d{2}$') { return 'DateTime' }
    if ($t -match '^\d{14}$')                                   { return 'Stamp' }
    if ($t -match '^\d{2}:\d{2}:\d{2}$')                        { return 'TimeOfDay' }
    if ($t -match '^\d{1,4}:\d{2}:\d{2}$')                      { return 'Duration' }
    return 'Unknown'
}

# Splits a time-ish string into its named numeric fields, each carrying the
# character offset it occupies in the TRIMMED text plus its legal range.
# Returns a plain array (empty when the text does not match the layout).
function Get-TimeDigitFieldSpec {
    param([string]$Text, [string]$Kind = 'Auto')
    $t = ([string]$Text).Trim()
    if ([string]::IsNullOrEmpty($t)) { return @() }

    $k = [string]$Kind
    if ([string]::IsNullOrWhiteSpace($k) -or $k -eq 'Auto') { $k = Get-TimeDigitTextKind $t }

    $rx = ''
    $names = @()
    switch ($k) {
        'DateTime'  { $rx = '^(\d{4})/(\d{2})/(\d{2})[ \t]+(\d{2}):(\d{2}):(\d{2})$'
                      $names = @('Year', 'Month', 'Day', 'Hour', 'Minute', 'Second') }
        'Stamp'     { $rx = '^(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})$'
                      $names = @('Year', 'Month', 'Day', 'Hour', 'Minute', 'Second') }
        'TimeOfDay' { $rx = '^(\d{2}):(\d{2}):(\d{2})$'
                      $names = @('Hour', 'Minute', 'Second') }
        'Duration'  { $rx = '^(\d{1,4}):(\d{2}):(\d{2})$'
                      $names = @('DurHour', 'Minute', 'Second') }
        default     { return @() }
    }

    $m = [regex]::Match($t, $rx)
    if (-not $m.Success) { return @() }

    $out = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $names.Count; $i++) {
        $g = $m.Groups[$i + 1]
        $lim = Get-TimeDigitFieldLimit $names[$i]
        $out.Add([pscustomobject]@{
            Name   = $names[$i]
            Index  = [int]$g.Index
            Length = [int]$g.Length
            Value  = [string]$g.Value
            Min    = [int]$lim.Min
            Max    = [int]$lim.Max
        })
    }
    return $out.ToArray()
}

# True when the field's digits parse to a number inside its legal range.
function Test-TimeDigitFieldValue {
    param([string]$Value, [int]$Min, [int]$Max)
    $n = 0
    if (-not [int]::TryParse(([string]$Value), [ref]$n)) { return $false }
    return ($n -ge $Min -and $n -le $Max)
}

# ---------------------------------------------------------------------------
# Repair-ImpossibleTimeDigit
#   Fixes digits that CANNOT be what OCR read, using nothing but the field's
#   own legal range: for each out-of-range field, every single 3<->9
#   substitution inside that field is tried, and the repair is applied only
#   when EXACTLY ONE of them lands back in range.
#
#     '10:93:20'            -> '10:33:20'  (minute tens cannot be 9)
#     '20260723 09:99:07'   -> minute 99: swapping the tens gives 39 (legal),
#                              swapping the units gives 93 (still illegal)
#                              -> unique -> '09:39:07'
#     '2026/19/03 ...'      -> month 19: 13 is still illegal -> untouched,
#                              Invalid = $true (a human decides)
#
#   Returns a hashtable:
#     Text      the (possibly repaired) string -- unchanged unless Repaired
#     Repaired  $true when at least one digit was forced back in range
#     Invalid   $true when a field stayed out of range (unrepairable)
#     Changes   human-readable 'Field old->new' notes, plain array
# ---------------------------------------------------------------------------
function Repair-ImpossibleTimeDigit {
    param([string]$Text, [string]$Kind = 'Auto')

    $t = ([string]$Text).Trim()
    $res = @{ Text = $t; Repaired = $false; Invalid = $false; Changes = @() }
    $fields = @(Get-TimeDigitFieldSpec -Text $t -Kind $Kind)
    if ($fields.Count -eq 0) { return $res }

    $chars = $t.ToCharArray()
    $changes = [System.Collections.Generic.List[string]]::new()
    foreach ($f in $fields) {
        if (Test-TimeDigitFieldValue -Value $f.Value -Min $f.Min -Max $f.Max) { continue }

        $fixes = [System.Collections.Generic.List[object]]::new()
        for ($p = 0; $p -lt $f.Length; $p++) {
            $alt = Get-TimeDigitSwapChar $f.Value[$p]
            if ($null -eq $alt) { continue }
            $cand = $f.Value.Remove($p, 1).Insert($p, [string]$alt)
            if (Test-TimeDigitFieldValue -Value $cand -Min $f.Min -Max $f.Max) {
                $fixes.Add([pscustomobject]@{ Pos = $p; Char = [string]$alt; Value = $cand })
            }
        }

        if ($fixes.Count -eq 1) {
            $fix = $fixes[0]
            $chars[$f.Index + [int]$fix.Pos] = ([string]$fix.Char)[0]
            $changes.Add(('{0} {1}->{2}' -f $f.Name, $f.Value, $fix.Value))
        } else {
            # 0 fixes = not a 3<->9 problem; >1 = we cannot tell which digit.
            # Either way the value stays exactly as OCR read it.
            $res.Invalid = $true
        }
    }

    if ($changes.Count -gt 0) {
        $res.Text = (-join $chars)
        $res.Repaired = $true
        $res.Changes = $changes.ToArray()
    }
    return $res
}

# ---------------------------------------------------------------------------
# Get-TimeDigitSwapVariants
#   Every string that differs from -Text by exactly ONE 3<->9 substitution.
#   The search space for the arithmetic disambiguation below; also the basis
#   for "could this reading have been something else" questions.
#   Returns a plain array of { Index; From; To; Text }.
# ---------------------------------------------------------------------------
function Get-TimeDigitSwapVariants {
    param([string]$Text)
    $t = [string]$Text
    $out = [System.Collections.Generic.List[object]]::new()
    if ([string]::IsNullOrEmpty($t)) { return @() }
    for ($i = 0; $i -lt $t.Length; $i++) {
        $alt = Get-TimeDigitSwapChar $t[$i]
        if ($null -eq $alt) { continue }
        $out.Add([pscustomobject]@{
            Index = $i
            From  = [string]$t[$i]
            To    = [string]$alt
            Text  = $t.Remove($i, 1).Insert($i, [string]$alt)
        })
    }
    return $out.ToArray()
}

# ---------------------------------------------------------------------------
# Get-TimeDigitRisk
#   Classifies a reading WITHOUT changing it:
#     'none'    no confusable digit sits anywhere it could hide
#     'suspect' at least one 3/9 sits where the opposite digit would ALSO be
#               legal -- an undetectable misread, so a human must look
#     'invalid' a field is out of range and no single 3<->9 swap repairs it
#     'unknown' the text does not match a known time layout (nothing to say)
#   Positions holds the character offsets of the suspect digits, Fields the
#   names of the fields they sit in. Pure: nothing is rewritten here.
# ---------------------------------------------------------------------------
function Get-TimeDigitRisk {
    param([string]$Text, [string]$Kind = 'Auto')

    $t = ([string]$Text).Trim()
    $res = @{ Level = 'unknown'; Positions = @(); Fields = @() }
    $fields = @(Get-TimeDigitFieldSpec -Text $t -Kind $Kind)
    if ($fields.Count -eq 0) { return $res }

    $positions = [System.Collections.Generic.List[int]]::new()
    $names     = [System.Collections.Generic.List[string]]::new()

    # An out-of-range field that the positional repair CAN force back in
    # range is not a risk (it is fixed deterministically upstream); one it
    # cannot is a real unknown.
    $repair  = Repair-ImpossibleTimeDigit -Text $t -Kind $Kind
    $invalid = [bool]$repair.Invalid

    foreach ($f in $fields) {
        if (-not (Test-TimeDigitFieldValue -Value $f.Value -Min $f.Min -Max $f.Max)) { continue }
        for ($p = 0; $p -lt $f.Length; $p++) {
            $alt = Get-TimeDigitSwapChar $f.Value[$p]
            if ($null -eq $alt) { continue }
            $cand = $f.Value.Remove($p, 1).Insert($p, [string]$alt)
            # Both spellings legal -> the text alone cannot tell them apart.
            if (Test-TimeDigitFieldValue -Value $cand -Min $f.Min -Max $f.Max) {
                $positions.Add($f.Index + $p)
                if (-not $names.Contains([string]$f.Name)) { $names.Add([string]$f.Name) }
            }
        }
    }

    $res.Positions = $positions.ToArray()
    $res.Fields    = $names.ToArray()
    if     ($invalid)                { $res.Level = 'invalid' }
    elseif ($positions.Count -gt 0)  { $res.Level = 'suspect' }
    else                             { $res.Level = 'none' }
    return $res
}

# Parses this project's formatted stamp ('yyyy/MM/dd HH:mm:ss') into a
# [datetime], or $null. Local to keep the module self-contained.
function ConvertTo-TimeDigitDateTime {
    param([string]$Text)
    $t = ([string]$Text).Trim()
    if ([string]::IsNullOrEmpty($t)) { return $null }
    # Explicitly typed so the binder picks the string[] ParseExact overload.
    [string[]]$fmts = @('yyyy/MM/dd HH:mm:ss', 'yyyy/MM/dd H:mm:ss')
    $dt = [datetime]::MinValue
    if ([datetime]::TryParseExact($t, $fmts, [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return $dt
    }
    return $null
}

# 'HH:mm:ss' (hours NOT clamped to 24) -> total seconds, or $null.
function ConvertTo-TimeDigitDurationSeconds {
    param([string]$Text)
    $t = ([string]$Text).Trim()
    if ([string]::IsNullOrEmpty($t)) { return $null }
    $m = [regex]::Match($t, '^(\d{1,4}):(\d{2}):(\d{2})$')
    if (-not $m.Success) { return $null }
    $h = [int]$m.Groups[1].Value
    $mi = [int]$m.Groups[2].Value
    $s = [int]$m.Groups[3].Value
    if ($mi -gt 59 -or $s -gt 59) { return $null }
    return ($h * 3600 + $mi * 60 + $s)
}

# Total seconds -> 'HH:mm:ss', hours not clamped to 24 (mirrors
# ProcessTimeParse.ps1's Get-ProcessDurationText output shape).
function Format-TimeDigitDuration {
    param([int]$TotalSeconds)
    if ($TotalSeconds -lt 0) { return '' }
    $h = [int][Math]::Floor($TotalSeconds / 3600)
    $m = [int][Math]::Floor(($TotalSeconds % 3600) / 60)
    $s = [int]($TotalSeconds % 60)
    return ('{0:00}:{1:00}:{2:00}' -f $h, $m, $s)
}

# ---------------------------------------------------------------------------
# Resolve-ProcessTimeDurationConflict
#   The arithmetic disambiguator. start, end and the page's own printed
#   processing-time column are three readings of one fact, so they constrain
#   each other:
#
#     start 11:19:02, end 11:19:09, page 00:00:01
#       -> swapping end's '9' for '3' gives 11:19:03, and 03 - 02 == 1 == page
#       -> that substitution is FORCED by arithmetic, so apply it.
#
#     start 11:19:03, end 11:19:09, page 00:00:00
#       -> start's 3->9 works AND end's 9->3 works: two different records
#          both satisfy the page. We do not know which -> change nothing.
#
#   The page column is the anchor rather than a fourth unknown: it is real
#   on-page evidence, and this project's operator confirmed a printed short
#   duration ('00:00:01') is the reading least likely to be misread.
#
#   Returns a hashtable:
#     Status   'agree'     derived duration already equals the page column
#              'repaired'  exactly one substitution reconciled them (applied)
#              'ambiguous' several substitutions reconcile -- nothing changed
#              'conflict'  none reconciles -- some other discrepancy
#              'unknown'   an input was unparseable (no opinion)
#     Start/End  the (possibly repaired) stamps
#     Duration   the agreed duration text, '' when not agreed
#     Derived    the duration implied by Start/End as READ
#     Note       one-line explanation for the row's Note column
# ---------------------------------------------------------------------------
function Resolve-ProcessTimeDurationConflict {
    param([string]$Start, [string]$End, [string]$PageDuration)

    $res = @{ Status = 'unknown'; Start = [string]$Start; End = [string]$End
              Duration = ''; Derived = ''; Note = '' }

    $st = ConvertTo-TimeDigitDateTime $Start
    $en = ConvertTo-TimeDigitDateTime $End
    $pd = ConvertTo-TimeDigitDurationSeconds $PageDuration
    if ($null -eq $st -or $null -eq $en) { return $res }

    $derivedSec = [int][Math]::Round(($en - $st).TotalSeconds)
    if ($derivedSec -ge 0) { $res.Derived = Format-TimeDigitDuration $derivedSec }
    if ($null -eq $pd) { return $res }

    if ($derivedSec -eq $pd) {
        $res.Status = 'agree'
        $res.Duration = Format-TimeDigitDuration $derivedSec
        return $res
    }

    # Enumerate every single 3<->9 substitution of start and of end, keeping
    # the ones that make the record agree with the printed duration.
    $cands = [System.Collections.Generic.List[object]]::new()
    $seen  = @{}
    foreach ($v in @(Get-TimeDigitSwapVariants $Start)) {
        $alt = ConvertTo-TimeDigitDateTime $v.Text
        if ($null -eq $alt) { continue }
        $d = [int][Math]::Round(($en - $alt).TotalSeconds)
        if ($d -lt 0 -or $d -ne $pd) { continue }
        $key = ('{0}|{1}' -f $v.Text, $End)
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $cands.Add([pscustomobject]@{ Side = 'start'; Start = $v.Text; End = [string]$End
                                      Change = ('start {0}->{1} at {2}' -f $v.From, $v.To, $v.Index) })
    }
    foreach ($v in @(Get-TimeDigitSwapVariants $End)) {
        $alt = ConvertTo-TimeDigitDateTime $v.Text
        if ($null -eq $alt) { continue }
        $d = [int][Math]::Round(($alt - $st).TotalSeconds)
        if ($d -lt 0 -or $d -ne $pd) { continue }
        $key = ('{0}|{1}' -f $Start, $v.Text)
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $cands.Add([pscustomobject]@{ Side = 'end'; Start = [string]$Start; End = $v.Text
                                      Change = ('end {0}->{1} at {2}' -f $v.From, $v.To, $v.Index) })
    }

    $pageText = Format-TimeDigitDuration $pd
    if ($cands.Count -eq 1) {
        $c = $cands[0]
        $res.Status   = 'repaired'
        $res.Start    = [string]$c.Start
        $res.End      = [string]$c.End
        $res.Duration = $pageText
        $res.Note     = ('3/9 misread fixed by arithmetic: {0} (derived {1} -> page {2})' -f `
                         $c.Change, $res.Derived, $pageText)
        return $res
    }
    if ($cands.Count -gt 1) {
        $res.Status = 'ambiguous'
        $res.Note   = ('page duration {0} != derived {1}; {2} different 3/9 fixes would explain it -- left as read, needs a human' -f `
                       $pageText, $res.Derived, $cands.Count)
        return $res
    }
    $res.Status = 'conflict'
    $res.Note   = ('page duration {0} != derived {1} and no single 3/9 fix explains it -- left as read, needs a human' -f `
                   $pageText, $res.Derived)
    return $res
}

# ---------------------------------------------------------------------------
# Get-ProcessTimeDigitFormatRule
#   The conditional-formatting spec for the ProcessTime output workbook: mark
#   every start/end/duration cell whose SECONDS digit is a 3 or a 9, because
#   that is exactly the digit this OCR confusion can flip without leaving any
#   other trace. The operator scans the red cells against the linked snap.
#
#   The formula deliberately routes through TEXT(cell,"ss") so ONE rule
#   covers both spellings the writer can produce: a real date/time serial and
#   the plain-text fallback (Excel coerces a 'yyyy/MM/dd HH:mm:ss' or
#   'HH:mm:ss' string to a serial, so TEXT yields the seconds digits either
#   way; an uncoercible string comes back unchanged and its last character is
#   still the seconds units digit). A blank cell reads as 0 -> '00' and never
#   fires, and the IFERROR wrapper keeps any remaining surprise from showing
#   up as an error instead of simply not marking.
#
#   Colors are Excel BGR Longs: font 255 = pure red, interior 13551615 =
#   Excel's own "light red fill" -- the pairing operators already know from
#   the built-in Highlight Cells rules.
#
#   Returns a plain array of { Column; FormulaTemplate; FontColor;
#   InteriorColor; StopIfTrue }. New-ProcessTimeDigitFormatFormula fills the
#   template in; the COM side (ProcessTime.ps1's Set-ProcessTimeDigitFormat)
#   applies it.
# ---------------------------------------------------------------------------
function Get-ProcessTimeDigitFormatRule {
    param([string[]]$Columns = @('D', 'E', 'F'))
    $tpl = '=IFERROR(OR(RIGHT(TEXT(${0}{1},"ss"),1)="3",RIGHT(TEXT(${0}{1},"ss"),1)="9"),FALSE)'
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($col in @($Columns)) {
        if ([string]::IsNullOrWhiteSpace($col)) { continue }
        $out.Add(@{
            Column          = [string]$col
            FormulaTemplate = $tpl
            FontColor       = 255
            InteriorColor   = 13551615
            StopIfTrue      = $false
        })
    }
    return $out.ToArray()
}

# Fills a conditional-format template: {0} = column letter, {1} = the range's
# top data row (Excel rewrites the relative reference down the range).
function New-ProcessTimeDigitFormatFormula {
    param([string]$Template, [string]$Column, [int]$Row)
    return ([string]$Template -f ([string]$Column), $Row)
}
