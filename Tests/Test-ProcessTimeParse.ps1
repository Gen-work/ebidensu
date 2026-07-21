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

# ---------------------------------------------------------------------------
# ConvertTo-ProcessTimeCorrelKey / Test-ProcessTimeCorrelSeen (OCR glyph fold)
# Real office-PC run (v2.12.3): the HM OCR reads 'JIGPKB1S' as 'JIGPKBIS'
# (digit 1 -> letter I), which the old exact substring test rejected --
# discarding the correct HM screenshot.
# ---------------------------------------------------------------------------
Assert-Equal (ConvertTo-ProcessTimeCorrelKey 'JIGPKBIS') (ConvertTo-ProcessTimeCorrelKey 'JIGPKB1S') 'JIGPKBIS and JIGPKB1S fold to the same key'
Assert-True (Test-ProcessTimeCorrelSeen -Line '...  1 1 ,262   JIGPKBIS' -CorrelId 'JIGPKB1S') 'correl id seen through the 1<->I OCR confusion'
Assert-True (Test-ProcessTimeCorrelSeen -Line 'a row for JIDSM48S here' -CorrelId 'JIDSM48S') 'exact correl id still matches'
Assert-True (-not (Test-ProcessTimeCorrelSeen -Line 'a row for JIDSK48S here' -CorrelId 'JIDSM48S')) 'a genuinely different correl (K vs M) is not a false match'
Assert-True (-not (Test-ProcessTimeCorrelSeen -Line 'no id anywhere' -CorrelId 'JIGPKB1S')) 'absent correl id is not seen'
Assert-True (-not (Test-ProcessTimeCorrelSeen -Line 'anything' -CorrelId '')) 'blank correl id is never seen'

# ---------------------------------------------------------------------------
# Get-ProcessTimeRecordCount (shori-kensu) -- ja HM rows split the count digits
# ('1 1 ,262'); the count is read after the 14-digit data-creation stamp.
# ---------------------------------------------------------------------------
$countLine = '2026/05/23  10 :58 :50       2026/05/23  10 :58 :57       00 :00 :07      IGPLB073        K   ' + $normal + '       20260523105850                       1 1 ,262   ' + $diamond + '          JIGPKBIS'
$rc = @(ConvertFrom-ProcessTimeOcrLines -Lines @($countLine) -CorrelId 'JIGPKB1S')
Assert-Equal 1 $rc.Count 'record-count line parses as one full row'
Assert-Equal '2026/05/23 10:58:50' $rc[0].StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'spaced start time normalized'
Assert-Equal '11,262' $rc[0].RecordCount 'record count read after the data-creation stamp, internal spaces stripped'
Assert-True $rc[0].CorrelSeen 'correl id seen via glyph fold on the count row'

$zeroLine = '2026/05/23  01 :21 :43       2026/05/23  01 :21 :50       00 :00 :07      IDSLB053      C   ' + $normal + '       20260523012143                   0   ' + $diamond + '            JIDSC48S'
$rz = @(ConvertFrom-ProcessTimeOcrLines -Lines @($zeroLine) -CorrelId 'JIDSC48S')
Assert-Equal '0' $rz[0].RecordCount 'a zero record count is read'

Assert-Equal '' (Get-ProcessTimeRecordCount -Line 'no digits at all here') 'no datestamp -> empty record count (never guesses)'

# ---------------------------------------------------------------------------
# Get-ProcessTimeDateHints + -StartDateHints date correction (v2.12.3):
# the ja recognizer misreads a DATE digit ('2026/05/29' -> '2026/05/23') while
# the time of day is correct; the en-US datestamp is the trusted date source.
# ---------------------------------------------------------------------------
$hints = @(Get-ProcessTimeDateHints -Lines @('start 20260529105850 count 11,262 JIGPKBIS', 'no stamp here', 'x 20260529075413 y'))
Assert-Equal 2 $hints.Count '14-digit datestamps parsed to datetimes (others ignored)'
Assert-Equal '2026/05/29 10:58:50' $hints[0].ToString('yyyy/MM/dd HH:mm:ss') 'datestamp parsed as yyyyMMddHHmmss'

$jaWrongDate = '2026/05/23  10 :58 :50       2026/05/23  10 :58 :57       00 :00 :07      IGPLB073        K   ' + $normal + '       20260523105850                       1 1 ,262   ' + $diamond + '          JIGPKBIS'
$enHint = @(Get-ProcessTimeDateHints -Lines @('2026/05/29   2026/05/29   IGPLB073   20260529105850   11,262   JIGPKBIS'))
$corr = @(ConvertFrom-ProcessTimeOcrLines -Lines @($jaWrongDate) -CorrelId 'JIGPKB1S' -StartDateHints $enHint)
Assert-Equal '2026/05/29 10:58:50' $corr[0].StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'ja date-digit misread corrected from the en-US datestamp (time kept)'
Assert-Equal '2026/05/29 10:58:57' $corr[0].EndTime.ToString('yyyy/MM/dd HH:mm:ss') 'end time shifted to the corrected date'
Assert-True $corr[0].DateCorrected 'date-corrected flag is set'

$agree = @(ConvertFrom-ProcessTimeOcrLines -Lines @($jaWrongDate) -CorrelId 'JIGPKB1S' -StartDateHints @([datetime]'2026/05/23 10:58:50'))
Assert-True (-not $agree[0].DateCorrected) 'no correction when the hint date already matches the read date'
Assert-Equal '2026/05/23 10:58:50' $agree[0].StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'a matching-date hint leaves the row unchanged'

$noHint = @(ConvertFrom-ProcessTimeOcrLines -Lines @($jaWrongDate) -CorrelId 'JIGPKB1S')
Assert-True (-not $noHint[0].DateCorrected) 'no hints -> not corrected'
Assert-Equal '2026/05/23 10:58:50' $noHint[0].StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'no hints leaves the ja date as read'

# ---------------------------------------------------------------------------
# Resolve-ProcessTimeRowPlan (v2.13.0: staged Ocr/Write + sidecar-based re-run)
# ---------------------------------------------------------------------------
$fresh = Resolve-ProcessTimeRowPlan -SidecarExists $false -Inserted $false -Stage 'Both' -Force $false
Assert-True $fresh.NeedsOcr 'fresh row (Both, no Force): needs OCR'
Assert-True $fresh.NeedsWrite 'fresh row (Both, no Force): needs write'
Assert-True $fresh.Touch 'fresh row: touched'

$ocrDoneWritePending = Resolve-ProcessTimeRowPlan -SidecarExists $true -Inserted $false -Stage 'Both' -Force $false
Assert-True (-not $ocrDoneWritePending.NeedsOcr) 'sidecar cached, not yet inserted: OCR skipped (re-run reuses the cache)'
Assert-True $ocrDoneWritePending.NeedsWrite 'sidecar cached, not yet inserted: write still pending'

$legacyDone = Resolve-ProcessTimeRowPlan -SidecarExists $false -Inserted $true -Stage 'Both' -Force $false
Assert-True (-not $legacyDone.NeedsOcr) 'legacy-inserted row with no sidecar: OCR not re-triggered (backward compat)'
Assert-True (-not $legacyDone.NeedsWrite) 'legacy-inserted row: write not pending'
Assert-True (-not $legacyDone.Touch) 'fully-done row (legacy or new): not touched at all'

$bothDone = Resolve-ProcessTimeRowPlan -SidecarExists $true -Inserted $true -Stage 'Both' -Force $false
Assert-True (-not $bothDone.Touch) 'sidecar + inserted: fully done, skipped'

$forced = Resolve-ProcessTimeRowPlan -SidecarExists $true -Inserted $true -Stage 'Both' -Force $true
Assert-True $forced.NeedsOcr '-Force redoes OCR even when sidecar + inserted'
Assert-True $forced.NeedsWrite '-Force redoes write even when sidecar + inserted'

$ocrOnlyStage = Resolve-ProcessTimeRowPlan -SidecarExists $false -Inserted $false -Stage 'Ocr' -Force $false
Assert-True $ocrOnlyStage.NeedsOcr 'Stage=Ocr: OCR needed for a fresh row'
Assert-True (-not $ocrOnlyStage.NeedsWrite) 'Stage=Ocr: write never requested regardless of Inserted'

$writeOnlyStageReady = Resolve-ProcessTimeRowPlan -SidecarExists $true -Inserted $false -Stage 'Write' -Force $false
Assert-True (-not $writeOnlyStageReady.NeedsOcr) 'Stage=Write: OCR never requested regardless of cache state'
Assert-True $writeOnlyStageReady.NeedsWrite 'Stage=Write: write needed when not yet inserted'

$writeOnlyStageNoCache = Resolve-ProcessTimeRowPlan -SidecarExists $false -Inserted $false -Stage 'Write' -Force $false
Assert-True $writeOnlyStageNoCache.NeedsWrite 'Stage=Write: still marked pending even with no cache (caller reports a MISS and skips it at run time)'

$ocrStageLegacyRow = Resolve-ProcessTimeRowPlan -SidecarExists $false -Inserted $true -Stage 'Ocr' -Force $false
Assert-True (-not $ocrStageLegacyRow.Touch) 'Stage=Ocr on an already-inserted legacy row: nothing to do without -Force'

$ocrStageLegacyRowForced = Resolve-ProcessTimeRowPlan -SidecarExists $false -Inserted $true -Stage 'Ocr' -Force $true
Assert-True $ocrStageLegacyRowForced.NeedsOcr 'Stage=Ocr -Force: backfills a sidecar for a legacy row on demand'
Assert-True (-not $ocrStageLegacyRowForced.NeedsWrite) 'Stage=Ocr -Force: still never touches write'

exit (Complete-Tests)
