#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'TimeDigitVerify.ps1')

Reset-Tests 'TimeDigitVerify'

# ---------------------------------------------------------------------------
# Get-TimeDigitSwapChar : the only confusion this module models.
# ---------------------------------------------------------------------------
Assert-Equal '9' (Get-TimeDigitSwapChar '3') '3 swaps to 9'
Assert-Equal '3' (Get-TimeDigitSwapChar '9') '9 swaps to 3'
Assert-True  ($null -eq (Get-TimeDigitSwapChar '5')) '5 is not confusable'
Assert-True  ($null -eq (Get-TimeDigitSwapChar ':')) 'a separator is not confusable'

# ---------------------------------------------------------------------------
# Get-TimeDigitTextKind
# ---------------------------------------------------------------------------
Assert-Equal 'DateTime'  (Get-TimeDigitTextKind '2026/07/16 10:59:29') 'full stamp detected'
Assert-Equal 'Stamp'     (Get-TimeDigitTextKind '20260716105929')      '14-digit datestamp detected'
Assert-Equal 'TimeOfDay' (Get-TimeDigitTextKind '10:59:29')            'time of day detected'
Assert-Equal 'Duration'  (Get-TimeDigitTextKind '123:04:05')           'long duration detected'
Assert-Equal 'Unknown'   (Get-TimeDigitTextKind 'not a time')          'garbage is Unknown'
Assert-Equal 'Unknown'   (Get-TimeDigitTextKind '')                    'empty is Unknown'

# ---------------------------------------------------------------------------
# Get-TimeDigitFieldSpec : offsets + legal ranges.
# ---------------------------------------------------------------------------
$fields = @(Get-TimeDigitFieldSpec -Text '2026/07/16 10:59:29')
Assert-Equal 6 $fields.Count 'full stamp splits into 6 fields'
Assert-Equal 'Second' $fields[5].Name  'last field is Second'
Assert-Equal 17       $fields[5].Index 'Second sits at offset 17'
Assert-Equal 59       $fields[5].Max   'Second max is 59'
Assert-Equal 23       $fields[3].Max   'Hour max is 23'

$durFields = @(Get-TimeDigitFieldSpec -Text '30:00:01' -Kind 'Duration')
Assert-Equal 3    $durFields.Count 'duration splits into 3 fields'
Assert-Equal 9999 $durFields[0].Max 'duration hours are effectively unbounded'

Assert-Equal 0 (@(Get-TimeDigitFieldSpec -Text 'nonsense')).Count 'unparseable text yields no fields'

# ---------------------------------------------------------------------------
# Repair-ImpossibleTimeDigit : ONLY digits that cannot be what OCR read.
# This is the "9 in a position a 9 can never appear" case the operator asked
# to have fixed outright.
# ---------------------------------------------------------------------------
$r1 = Repair-ImpossibleTimeDigit -Text '2026/07/16 10:93:20'
Assert-True  $r1.Repaired                      'minute tens 9 is impossible -> repaired'
Assert-Equal '2026/07/16 10:33:20' $r1.Text    'minute 93 forced to 33'
Assert-True  (-not $r1.Invalid)                'a forced repair is not flagged invalid'
Assert-Equal 'Minute 93->33' $r1.Changes[0]    'the change is reported'

$r2 = Repair-ImpossibleTimeDigit -Text '2026/07/16 10:59:93'
Assert-Equal '2026/07/16 10:59:33' $r2.Text    'second 93 forced to 33'

# Both digits illegal as read: only ONE single swap lands back in range.
$r3 = Repair-ImpossibleTimeDigit -Text '2026/07/16 10:99:20'
Assert-True  $r3.Repaired                      'minute 99 has a unique single-swap fix'
Assert-Equal '2026/07/16 10:39:20' $r3.Text    'minute 99 -> 39 (93 is still illegal)'

# Nothing wrong -> nothing touched. The whole point of the rework.
$r4 = Repair-ImpossibleTimeDigit -Text '2026/07/16 10:59:29'
Assert-True  (-not $r4.Repaired)               'a legal reading is never rewritten'
Assert-Equal '2026/07/16 10:59:29' $r4.Text    'a legal reading is returned unchanged'
Assert-True  (-not $r4.Invalid)                'a legal reading is not invalid'

# Out of range but NOT a 3/9 problem -> left alone, flagged for a human.
$r5 = Repair-ImpossibleTimeDigit -Text '2026/19/03 10:59:29'
Assert-True  (-not $r5.Repaired)               'month 19 is not repaired (13 is still illegal)'
Assert-True  $r5.Invalid                       'month 19 is reported invalid'
Assert-Equal '2026/19/03 10:59:29' $r5.Text    'month 19 keeps the value OCR read'

# A duration must be told what it is, or a 30-hour run looks like a bad hour.
$r6 = Repair-ImpossibleTimeDigit -Text '30:00:01' -Kind 'Duration'
Assert-True  (-not $r6.Repaired)               'a 30-hour duration is legal'
Assert-True  (-not $r6.Invalid)                'a 30-hour duration is not invalid'
$r7 = Repair-ImpossibleTimeDigit -Text '00:93:07' -Kind 'Duration'
Assert-Equal '00:33:07' $r7.Text               'duration minute 93 forced to 33'

# 14-digit datestamps use the same field rules.
$r8 = Repair-ImpossibleTimeDigit -Text '20260716109320'
Assert-Equal '20260716103320' $r8.Text         'datestamp minute 93 forced to 33'

Assert-Equal 'zzz' (Repair-ImpossibleTimeDigit -Text 'zzz').Text 'unparseable text is returned as-is'

# ---------------------------------------------------------------------------
# Get-TimeDigitSwapVariants
# ---------------------------------------------------------------------------
$v = @(Get-TimeDigitSwapVariants '10:59:23')
Assert-Equal 2 $v.Count 'two confusable digits -> two variants'
Assert-Equal '10:53:23' $v[0].Text 'first variant flips the 9'
Assert-Equal '10:59:29' $v[1].Text 'second variant flips the 3'
Assert-Equal 0 (@(Get-TimeDigitSwapVariants '10:57:24')).Count 'no 3 or 9 -> no variants'
Assert-Equal 0 (@(Get-TimeDigitSwapVariants '')).Count 'empty text -> no variants'

# ---------------------------------------------------------------------------
# Get-TimeDigitRisk : classify, never rewrite.
# ---------------------------------------------------------------------------
# The seconds units digit could legally be either 3 or 9 -> undetectable.
$k1 = Get-TimeDigitRisk -Text '2026/07/16 10:57:23'
Assert-Equal 'suspect' $k1.Level 'a 3 in the seconds units digit is suspect'
Assert-True  ($k1.Positions -contains 18) 'the suspect position is reported'

# A 3 in the seconds TENS digit is safe: a 9 could never be there, so a 3
# read in that slot cannot be hiding one -- and if OCR had flipped it to 9,
# Repair-ImpossibleTimeDigit would have forced it straight back.
$k2 = Get-TimeDigitRisk -Text '2026/07/16 10:57:30'
Assert-Equal 'none' $k2.Level 'a seconds TENS 3 is not suspect (9 is impossible there)'
Assert-True  (-not ($k2.Positions -contains 17)) 'the seconds TENS position is not marked'

$k3 = Get-TimeDigitRisk -Text '2026/07/16 10:57:20'
Assert-Equal 'none' $k3.Level 'no confusable digit anywhere -> none'
Assert-Equal 0 (@($k3.Positions)).Count 'no positions reported'

$k4 = Get-TimeDigitRisk -Text '2026/19/03 10:57:20'
Assert-Equal 'invalid' $k4.Level 'an unrepairable out-of-range field is invalid'

Assert-Equal 'unknown' (Get-TimeDigitRisk -Text 'nope').Level 'unparseable text has no opinion'

# ---------------------------------------------------------------------------
# ConvertTo-TimeDigitDurationSeconds / Format-TimeDigitDuration
# ---------------------------------------------------------------------------
Assert-Equal 1     (ConvertTo-TimeDigitDurationSeconds '00:00:01') 'one second parses'
Assert-Equal 3661  (ConvertTo-TimeDigitDurationSeconds '01:01:01') 'h/m/s parse'
Assert-Equal 90000 (ConvertTo-TimeDigitDurationSeconds '25:00:00') 'hours past 24 parse'
Assert-True  ($null -eq (ConvertTo-TimeDigitDurationSeconds '00:99:00')) 'illegal minute does not parse'
Assert-True  ($null -eq (ConvertTo-TimeDigitDurationSeconds ''))         'empty does not parse'
Assert-Equal '00:00:01' (Format-TimeDigitDuration 1)     'one second formats'
Assert-Equal '25:00:00' (Format-TimeDigitDuration 90000) 'hours are not clamped to 24'

# ---------------------------------------------------------------------------
# Resolve-ProcessTimeDurationConflict : the arithmetic disambiguator.
# This is the operator-reported regression: start ss=02, end ss=03 and a
# printed duration of 00:00:01, where OCR flipped the end's 3 to a 9 and the
# old code silently shipped the derived 00:00:07.
# ---------------------------------------------------------------------------
$c1 = Resolve-ProcessTimeDurationConflict -Start '2026/07/16 11:19:02' `
                                          -End   '2026/07/16 11:19:09' `
                                          -PageDuration '00:00:01'
Assert-Equal 'repaired' $c1.Status 'a unique 3/9 fix reconciles start/end with the page'
Assert-Equal '2026/07/16 11:19:03' $c1.End 'the end second is corrected 09 -> 03'
Assert-Equal '2026/07/16 11:19:02' $c1.Start 'the start is left alone'
Assert-Equal '00:00:01' $c1.Duration 'the duration becomes the printed one'
Assert-Equal '00:00:07' $c1.Derived  'the as-read derived value is reported too'

# Already consistent -> hands off.
$c2 = Resolve-ProcessTimeDurationConflict -Start '2026/07/16 11:19:02' `
                                          -End   '2026/07/16 11:19:03' `
                                          -PageDuration '00:00:01'
Assert-Equal 'agree' $c2.Status 'consistent readings agree'
Assert-Equal '2026/07/16 11:19:03' $c2.End 'an agreeing record is never rewritten'

# Two different single swaps both satisfy the page -> we do not know.
$c3 = Resolve-ProcessTimeDurationConflict -Start '2026/07/16 11:19:03' `
                                          -End   '2026/07/16 11:19:09' `
                                          -PageDuration '00:00:00'
Assert-Equal 'ambiguous' $c3.Status 'two competing fixes -> ambiguous'
Assert-Equal '2026/07/16 11:19:03' $c3.Start 'an ambiguous record keeps the start as read'
Assert-Equal '2026/07/16 11:19:09' $c3.End   'an ambiguous record keeps the end as read'
Assert-Equal '' $c3.Duration 'an ambiguous record agrees on no duration'

# No 3/9 fix explains the gap -> some other problem, still hands off.
$c4 = Resolve-ProcessTimeDurationConflict -Start '2026/07/16 11:19:02' `
                                          -End   '2026/07/16 11:19:08' `
                                          -PageDuration '00:00:01'
Assert-Equal 'conflict' $c4.Status 'an unexplainable gap is a conflict'
Assert-Equal '2026/07/16 11:19:08' $c4.End 'a conflicting record keeps the end as read'

# The misread can be on the START side too. Here the true start second was
# 03, OCR read 09, and only start 09 -> 03 reproduces the printed 6 seconds
# (the competing candidates would all imply a negative run, which is
# rejected before it can be counted).
$c5 = Resolve-ProcessTimeDurationConflict -Start '2026/07/16 11:19:09' `
                                          -End   '2026/07/16 11:19:09' `
                                          -PageDuration '00:00:06'
Assert-Equal 'repaired' $c5.Status 'start 09 -> 03 gives the printed 6-second run'
Assert-Equal '2026/07/16 11:19:03' $c5.Start 'the start second is corrected'
Assert-Equal '2026/07/16 11:19:09' $c5.End   'the end is left alone'

# Missing / unparseable inputs -> no opinion at all.
Assert-Equal 'unknown' (Resolve-ProcessTimeDurationConflict -Start '' -End '' -PageDuration '').Status `
    'empty inputs yield no opinion'
Assert-Equal 'unknown' (Resolve-ProcessTimeDurationConflict -Start '2026/07/16 11:19:02' `
    -End '2026/07/16 11:19:09' -PageDuration '').Status 'no page column -> no opinion'

# ---------------------------------------------------------------------------
# Get-ProcessTimeDigitFormatRule / New-ProcessTimeDigitFormatFormula
# ---------------------------------------------------------------------------
$rules = @(Get-ProcessTimeDigitFormatRule)
Assert-Equal 3 $rules.Count 'start / end / duration get a rule each'
Assert-Equal 'D' $rules[0].Column 'first rule targets the start column'
Assert-Equal 'F' $rules[2].Column 'last rule targets the duration column'
Assert-Equal 255 $rules[0].FontColor 'the marked digits go red'

$f = New-ProcessTimeDigitFormatFormula -Template $rules[0].FormulaTemplate -Column 'D' -Row 2
Assert-Equal '=IFERROR(OR(RIGHT(TEXT($D2,"ss"),1)="3",RIGHT(TEXT($D2,"ss"),1)="9"),FALSE)' $f `
    'the D-column rule fires on a 3 or 9 in the seconds digit'
$f2 = New-ProcessTimeDigitFormatFormula -Template $rules[2].FormulaTemplate -Column 'F' -Row 5
Assert-Equal '=IFERROR(OR(RIGHT(TEXT($F5,"ss"),1)="3",RIGHT(TEXT($F5,"ss"),1)="9"),FALSE)' $f2 `
    'the template follows the column and first data row'

Assert-Equal 1 (@(Get-ProcessTimeDigitFormatRule -Columns @('D'))).Count 'the column set is configurable'

exit (Complete-Tests)
