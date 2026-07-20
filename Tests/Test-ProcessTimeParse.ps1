#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'ProcessTimeParse.ps1')

Reset-Tests 'ProcessTimeParse'

$normal = [char]0x6B63 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # seijo-shuuryo (normal end)
$abend  = [char]0x7570 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # ijo-shuuryo (abend)

# ---------------------------------------------------------------------------
# Get-ProcessDurationText
# ---------------------------------------------------------------------------
$s1 = [datetime]'2026/07/16 09:00:00'
$e1 = [datetime]'2026/07/16 09:01:23'
Assert-Equal '00:01:23' (Get-ProcessDurationText $s1 $e1) 'duration under an hour'

$s2 = [datetime]'2026/07/15 09:00:00'
$e2 = [datetime]'2026/07/16 12:15:30'
Assert-Equal '27:15:30' (Get-ProcessDurationText $s2 $e2) 'duration over 24h is not clamped/wrapped'

Assert-Equal '' (Get-ProcessDurationText $null $e1) 'null start yields empty duration'
Assert-Equal '' (Get-ProcessDurationText $s1 $null) 'null end yields empty duration'
Assert-Equal '' (Get-ProcessDurationText $e1 $s1) 'end before start yields empty duration (never negative)'

# ---------------------------------------------------------------------------
# ConvertFrom-ProcessTimeOcrLines
# ---------------------------------------------------------------------------
$rows = @(ConvertFrom-ProcessTimeOcrLines @(
    "2026/07/16 09:00:00   2026/07/16 09:01:23   JIDSM01S   $normal",
    'no timestamps on this line at all',
    "2026/07/16   09:10:00    2026/07/16   09:12:45   JIDSM01S   $abend"
))
Assert-Equal 2 $rows.Count 'two rows carry two datetime tokens each'
Assert-Equal '2026/07/16 09:00:00' $rows[0].StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'first row start time parsed'
Assert-Equal '2026/07/16 09:01:23' $rows[0].EndTime.ToString('yyyy/MM/dd HH:mm:ss') 'first row end time parsed'
Assert-Equal $normal $rows[0].Status 'first row status literal found'
Assert-Equal $abend  $rows[1].Status 'second row (extra internal whitespace) status literal found'

$noMatch = @(ConvertFrom-ProcessTimeOcrLines @('nothing here', ''))
Assert-Equal 0 $noMatch.Count 'lines without two datetime tokens are skipped'

$badDate = @(ConvertFrom-ProcessTimeOcrLines @('2026/13/40 09:00:00   2026/07/16 09:01:23   x'))
Assert-Equal 0 $badDate.Count 'an unparseable datetime token drops the row instead of throwing'

# ---------------------------------------------------------------------------
# ConvertTo-ProcessTimeNormalizedLine (OCR-injected spaces inside time tokens)
# ---------------------------------------------------------------------------
Assert-Equal '10:58:20' (ConvertTo-ProcessTimeNormalizedLine '10 :58 :20') 'spaces around colons are stripped'
Assert-Equal '00:00:01' (ConvertTo-ProcessTimeNormalizedLine '00 :00 : 0 1') 'space INSIDE a two-digit field is stripped'
Assert-Equal '09:05:07' (ConvertTo-ProcessTimeNormalizedLine '9: 5 : 7') 'one-digit fields are zero-padded'
Assert-Equal 'x 1 ,036 y' (ConvertTo-ProcessTimeNormalizedLine 'x 1 ,036 y') 'record counts without colons are untouched'
Assert-Equal '20260523105820' (ConvertTo-ProcessTimeNormalizedLine '20260523105820') '14-digit datetime column is untouched'
Assert-Equal '10:58:20 2026/05/23' (ConvertTo-ProcessTimeNormalizedLine '10 :58 :20 2026/05/23') 'normalization stops at the next column'

# ---------------------------------------------------------------------------
# Real office-PC OCR samples (v2.12.2): the ja recognizer injects spaces
# into every time token and garbles normal-end to sei-TEI-shuuryo; the
# en-US recognizer reads the dates but drops the time-of-day entirely.
# ---------------------------------------------------------------------------
$garbledNormal = [char]0x6B63 + [char]0x5E1D + [char]0x7D42 + [char]0x4E86   # sei-TEI-shuuryo (OCR garble)
$garbledAbend  = [char]0x7570 + [char]0x5E1D + [char]0x7D42 + [char]0x4E86   # i-TEI-shuuryo (OCR garble)
$diamond = [string][char]0x25C6

$jaLine1 = '2026/05/23  10 :58 :20       2026/05/23  10 :58 :21       00 :00 : 0 1       IDSLB053      C   ' + $garbledNormal + '       20260523105820                         1 ,036   ' + $diamond + '            JIDSC48S'
$jaLine2 = '2026/05/23  01 :21 :43       2026/05/23  01 :21 :50       00 :00 : 0 1       IDSLB053      C   ' + $garbledNormal + '       20260523012143                   0   ' + $diamond + '            JIDSC48S'
$enLine  = '2026/05/29                2026/05/29                              IDSLB053                            20260529105820            1,036                 JIDSC48S'

$real = @(ConvertFrom-ProcessTimeOcrLines -Lines @($jaLine1, $jaLine2, $enLine) -CorrelId 'JIDSC48S')
Assert-Equal 2 $real.Count 'two ja rows parse; the date-only en row is skipped (no invented midnight times)'
Assert-Equal '2026/05/23 10:58:20' $real[0].StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'spaced start time normalized and parsed'
Assert-Equal '2026/05/23 10:58:21' $real[0].EndTime.ToString('yyyy/MM/dd HH:mm:ss') 'spaced end time normalized and parsed'
Assert-Equal $normal $real[0].Status 'garbled sei-TEI-shuuryo classifies as the CANONICAL normal-end literal'
Assert-Equal '00:00:01' $real[0].PageDuration 'page proc-time column captured for cross-checking'
Assert-True $real[0].CorrelSeen 'correl id seen on the line'
Assert-True (-not $real[0].Partial) 'two datetimes = full row'

$abRow = @(ConvertFrom-ProcessTimeOcrLines -Lines @('2026/05/23  10:00:00   2026/05/23  10:00:05   ' + $garbledAbend + '   JIDSC48S'))
Assert-Equal $abend $abRow[0].Status 'garbled i-TEI-shuuryo classifies as the canonical abend literal'

$partial = @(ConvertFrom-ProcessTimeOcrLines -Lines @('2026/05/23  10 :58 :20    IDSLB053  ' + $garbledNormal + '   00 :00 : 0 1   JIDSC48S') -CorrelId 'JIDSC48S')
Assert-Equal 1 $partial.Count 'one datetime + a status literal is kept as a PARTIAL row'
Assert-True $partial[0].Partial 'partial row flagged'
Assert-True ($null -eq $partial[0].EndTime) 'partial row has no end time'
Assert-Equal '2026/05/23 10:58:20' $partial[0].StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'partial row keeps the readable start time'
Assert-Equal '00:00:01' $partial[0].PageDuration 'partial row still captures the page duration column'

$noStatusOneDt = @(ConvertFrom-ProcessTimeOcrLines -Lines @('2026/05/23  10:58:20    IDSLB053   plain text'))
Assert-Equal 0 $noStatusOneDt.Count 'one datetime WITHOUT a status literal is not a row (header/furniture)'

# ---------------------------------------------------------------------------
# Get-ProcessTimeRowRank / Select-ProcessTimeRow
# ---------------------------------------------------------------------------
Assert-Equal 3 (Get-ProcessTimeRowRank $real[0]) 'full row + correl seen ranks 3'
Assert-Equal 1 (Get-ProcessTimeRowRank $partial[0]) 'partial row + correl seen ranks 1'
Assert-Equal 2 (Get-ProcessTimeRowRank ([PSCustomObject]@{ StartTime = $s1; EndTime = $e1 })) 'archived-tier row (no Partial/CorrelSeen fields) ranks 2'

$selReal = Select-ProcessTimeRow -Rows $real
Assert-Equal '2026/05/23 10:58:20' $selReal.StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'newest full row wins among equal ranks'

$mix = @($partial[0], $real[1])   # partial (10:58, rank 1) vs full (01:21, rank 3)
$selMix = Select-ProcessTimeRow -Rows $mix
Assert-True (-not $selMix.Partial) 'a full row beats a NEWER partial row'

$unseen = @(ConvertFrom-ProcessTimeOcrLines -Lines @($jaLine1) -CorrelId 'JIDSXXXX')
Assert-True (-not $unseen[0].CorrelSeen) 'different correl id is not seen'
Assert-True ($null -eq (Select-ProcessTimeRow -Rows $unseen -RequireCorrelSeen)) 'RequireCorrelSeen drops correl-unseen rows'
Assert-True ($null -ne (Select-ProcessTimeRow -Rows $unseen)) 'without the switch the row is still selectable'

# ---------------------------------------------------------------------------
# Get-ProcessTimeOcrMissNote
# ---------------------------------------------------------------------------
Assert-Equal 'no OCR lines' (Get-ProcessTimeOcrMissNote -Lines @()) 'empty input'
$missNote = Get-ProcessTimeOcrMissNote -Lines @($enLine, 'HONDA', '')
Assert-True ($missNote -match 'no readable time-of-day') 'date-only en lines are called out'
$missNone = Get-ProcessTimeOcrMissNote -Lines @('HONDA', 'IDSXA041')
Assert-True ($missNone -match 'no date/time tokens recognized') 'furniture-only reads are called out'

# ---------------------------------------------------------------------------
# Get-NewestProcessTimeRow
# ---------------------------------------------------------------------------
$older = [PSCustomObject]@{ StartTime = [datetime]'2026/07/16 08:00:00'; EndTime = [datetime]'2026/07/16 08:05:00' }
$newer = [PSCustomObject]@{ StartTime = [datetime]'2026/07/16 09:00:00'; EndTime = [datetime]'2026/07/16 09:05:00' }
$picked = Get-NewestProcessTimeRow -Rows @($older, $newer)
Assert-Equal $newer.StartTime.ToString('yyyy/MM/dd HH:mm:ss') $picked.StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'newest-by-StartTime wins among multiple rows'

Assert-True ($null -eq (Get-NewestProcessTimeRow -Rows @())) 'empty input returns null'
Assert-True ($null -eq (Get-NewestProcessTimeRow -Rows @($null))) 'array of only nulls returns null'

exit (Complete-Tests)
