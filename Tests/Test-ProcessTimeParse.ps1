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
# -MinimumTimeOfDay disabled here: this checks rank-vs-recency (full beats a
# newer partial regardless of clock time), not the 09:00 history filter --
# $real[1]'s 01:21 start would otherwise be dropped before rank is compared.
$selMix = Select-ProcessTimeRow -Rows $mix -MinimumTimeOfDay ([timespan]::Zero)
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
$midnightArchived = [PSCustomObject]@{ StartTime = [datetime]'2026/07/16 00:17:21'; EndTime = [datetime]'2026/07/16 00:17:21' }
Assert-True ($null -eq (Get-NewestProcessTimeRow -Rows @($midnightArchived))) 'archived tier also rejects pre-09:00 history'

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
Assert-True (Test-ProcessTimeCorrelSeen -Line 'result cell: J IDSCS4S' -CorrelId 'JIDSCS4S') 'OCR whitespace inside correl id is ignored'
Assert-True (Test-ProcessTimeCorrelSeen -Line 'result cell: J IGPF05S' -CorrelId 'JIGPFO5S') 'whitespace and O/0 glyph confusion compose safely'

# The OCR API commonly places the far-right correl cell on a separate visual
# line from the timestamps. Ownership must apply to the complete picture.
$splitCorrel = @(ConvertFrom-ProcessTimeOcrLines -Lines @(
    "2026/05/12 10:58:00  2026/05/12 10:58:07  00:00:07  $normal",
    'right edge cell: JIDSCS4S'
) -CorrelId 'JIDSCS4S')
Assert-Equal 1 $splitCorrel.Count 'timestamp row parses when correl is on another OCR line'
Assert-True $splitCorrel[0].CorrelSeen 'correl seen anywhere in the exported picture owns its timestamp rows'

$splitWrong = @(ConvertFrom-ProcessTimeOcrLines -Lines @(
    "2026/05/12 10:58:00  2026/05/12 10:58:07  00:00:07  $normal",
    'right edge cell: JIDSMS4S'
) -CorrelId 'JIDSCS4S')
Assert-True (-not $splitWrong[0].CorrelSeen) 'a different correl on another OCR line does not claim the picture'

# Real multi-row JRV sample: newest daytime row must beat a valid midnight
# history row from the same picture.
$dayAndMidnight = @(ConvertFrom-ProcessTimeOcrLines -Lines @(
    "2026/05/15  13 :44 :25  2026/05/15  13 :44 :25  00 :00 :00  IGPLB133 F $normal 20260515134425 0 J IGPF05S",
    "2026/05/15 00 : 17 :21  2026/05/15 00 : 17 :21  00 :00 :00  IGPLB133 F $normal 20260515001721 0 JIGPF05S"
) -CorrelId 'JIGPFO5S')
Assert-Equal 2 $dayAndMidnight.Count 'both real JRV history rows parse'
$pickedDay = Select-ProcessTimeRow -Rows $dayAndMidnight -RequireCorrelSeen
Assert-Equal '2026/05/15 13:44:25' $pickedDay.StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'daytime row is selected instead of midnight history'
Assert-True ($null -eq (Select-ProcessTimeRow -Rows @($dayAndMidnight[1]) -RequireCorrelSeen)) 'a picture containing only pre-09:00 history is rejected'

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
Assert-Equal '' (Get-ProcessTimeRecordCount -Line '20260515001721 JIGPFO5S') 'digit embedded in correl id is not mistaken for count'
$diamondOnly = '2026/05/12 10:58:02 2026/05/12 10:58:02 00:00:00 IDSLDO13 C ' + $normal + ' 0 ' + $diamond + ' J IDSCS4S'
Assert-Equal '0' (Get-ProcessTimeRecordCount -Line $diamondOnly -SearchFrom 40) 'JDL count before result diamond is read without a creation stamp'

$countFromEn = @(ConvertFrom-ProcessTimeOcrLines -Lines @(
    '2026/05/15 2026/05/15 IGPLB073 20260515134743 82 JIGPQBIS',
    '2026/05/15 13 :47 :43 2026/05/15 13 :47 :43 00 :00 :00 IGPLB073 ' + $normal + ' ' + $diamond + ' JIGPQBIS'
) -CorrelId 'JIGPQB1S')
Assert-Equal '82' $countFromEn[0].RecordCount 'count from en-US date-only row joins to ja timestamp row by creation stamp'

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
# Resolve-ProcessTimeRowPlan (v2.15.0: ProcessTime_Inserted is now a bitmask --
# -OcrDone/-WriteDone replace the old plain -Inserted bool)
# ---------------------------------------------------------------------------
$fresh = Resolve-ProcessTimeRowPlan -SidecarExists $false -OcrDone $false -WriteDone $false -Stage 'Both' -Force $false
Assert-True $fresh.NeedsOcr 'fresh row (Both, no Force): needs OCR'
Assert-True $fresh.NeedsWrite 'fresh row (Both, no Force): needs write'
Assert-True $fresh.Touch 'fresh row: touched'

$ocrDoneWritePending = Resolve-ProcessTimeRowPlan -SidecarExists $true -OcrDone $false -WriteDone $false -Stage 'Both' -Force $false
Assert-True (-not $ocrDoneWritePending.NeedsOcr) 'sidecar cached, bit not yet set: OCR skipped (re-run reuses the cache)'
Assert-True $ocrDoneWritePending.NeedsWrite 'sidecar cached, write bit not set: write still pending'

$ocrBitOnly = Resolve-ProcessTimeRowPlan -SidecarExists $false -OcrDone $true -WriteDone $false -Stage 'Both' -Force $false
Assert-True (-not $ocrBitOnly.NeedsOcr) 'OCR bit set (no sidecar file): OCR not re-triggered'
Assert-True $ocrBitOnly.NeedsWrite 'OCR bit set, write bit not set: write still pending'

$migratedLegacyDone = Resolve-ProcessTimeRowPlan -SidecarExists $false -OcrDone $true -WriteDone $true -Stage 'Both' -Force $false
Assert-True (-not $migratedLegacyDone.NeedsOcr) 'both bits set (e.g. a migrated legacy row): OCR not re-triggered'
Assert-True (-not $migratedLegacyDone.NeedsWrite) 'both bits set: write not pending'
Assert-True (-not $migratedLegacyDone.Touch) 'fully-done row (both bits): not touched at all'

$bothDone = Resolve-ProcessTimeRowPlan -SidecarExists $true -OcrDone $true -WriteDone $true -Stage 'Both' -Force $false
Assert-True (-not $bothDone.Touch) 'sidecar + both bits: fully done, skipped'

$forced = Resolve-ProcessTimeRowPlan -SidecarExists $true -OcrDone $true -WriteDone $true -Stage 'Both' -Force $true
Assert-True $forced.NeedsOcr '-Force redoes OCR even when sidecar + both bits set'
Assert-True $forced.NeedsWrite '-Force redoes write even when sidecar + both bits set'

$ocrOnlyStage = Resolve-ProcessTimeRowPlan -SidecarExists $false -OcrDone $false -WriteDone $false -Stage 'Ocr' -Force $false
Assert-True $ocrOnlyStage.NeedsOcr 'Stage=Ocr: OCR needed for a fresh row'
Assert-True (-not $ocrOnlyStage.NeedsWrite) 'Stage=Ocr: write never requested regardless of bits'

$writeOnlyStageReady = Resolve-ProcessTimeRowPlan -SidecarExists $true -OcrDone $true -WriteDone $false -Stage 'Write' -Force $false
Assert-True (-not $writeOnlyStageReady.NeedsOcr) 'Stage=Write: OCR never requested regardless of cache state'
Assert-True $writeOnlyStageReady.NeedsWrite 'Stage=Write: write needed when write bit not set'

$writeOnlyStageNoCache = Resolve-ProcessTimeRowPlan -SidecarExists $false -OcrDone $false -WriteDone $false -Stage 'Write' -Force $false
Assert-True $writeOnlyStageNoCache.NeedsWrite 'Stage=Write: still marked pending even with no cache (caller reports a MISS and skips it at run time)'

$ocrStageLegacyRow = Resolve-ProcessTimeRowPlan -SidecarExists $false -OcrDone $true -WriteDone $true -Stage 'Ocr' -Force $false
Assert-True (-not $ocrStageLegacyRow.Touch) 'Stage=Ocr on an already fully-done row: nothing to do without -Force'

$ocrStageLegacyRowForced = Resolve-ProcessTimeRowPlan -SidecarExists $false -OcrDone $true -WriteDone $true -Stage 'Ocr' -Force $true
Assert-True $ocrStageLegacyRowForced.NeedsOcr 'Stage=Ocr -Force: backfills a sidecar for an already-done row on demand'
Assert-True (-not $ocrStageLegacyRowForced.NeedsWrite) 'Stage=Ocr -Force: still never touches write'

# ---------------------------------------------------------------------------
# Get-ProcessTimeMigratedInsertedValue (v2.15.0: legacy plain 0/1 -> bitmask)
# ---------------------------------------------------------------------------
Assert-Equal '3' (Get-ProcessTimeMigratedInsertedValue '1') 'legacy plain 1 (written) migrates to bit 3 (both done)'
Assert-Equal '0' (Get-ProcessTimeMigratedInsertedValue '0') 'legacy 0 stays 0'
Assert-Equal '' (Get-ProcessTimeMigratedInsertedValue '') 'blank stays blank'
Assert-Equal '2' (Get-ProcessTimeMigratedInsertedValue '2') 'an already-bitmask value 2 passes through unchanged'
Assert-Equal '3' (Get-ProcessTimeMigratedInsertedValue '3') 'an already-bitmask value 3 passes through unchanged'

# ---------------------------------------------------------------------------
# Get-ProcessTimeOutputTag / Get-ProcessTimeOutputFileName / Resolve-ProcessTimeOutputDir
# (v2.15.0: generalized, config-driven output classification -- no longer
# hardcoded to just JDL/JRV, and never throws on an unrecognized tag)
# ---------------------------------------------------------------------------
Assert-Equal 'JDL' (Get-ProcessTimeOutputTag -ExcelName 'CJDLWDFL' -Tags @('JDL', 'JRV')) 'classifies by configured tag substring'
Assert-Equal 'JRV' (Get-ProcessTimeOutputTag -ExcelName 'KJRVWD64' -Tags @('JDL', 'JRV')) 'classifies a second configured tag'
Assert-Equal 'JDS' (Get-ProcessTimeOutputTag -ExcelName 'KJDSWD01' -Tags @('JDL', 'JRV', 'JDS')) 'a tag list can be extended beyond JDL/JRV (e.g. JDS)'
Assert-Equal 'Other' (Get-ProcessTimeOutputTag -ExcelName 'ZZZZZZZZ' -Tags @('JDL', 'JRV')) 'an unrecognized Excel_NAME falls back to the unclassified bucket instead of throwing'
Assert-Equal 'Misc' (Get-ProcessTimeOutputTag -ExcelName 'ZZZZZZZZ' -Tags @('JDL', 'JRV') -UnclassifiedTag 'Misc') '-UnclassifiedTag is configurable'

Assert-Equal (([string][char]0x51E6 + [char]0x7406 + [char]0x6642 + [char]0x9593) + '(JDL).xlsx') `
    (Get-ProcessTimeOutputFileName -Label ([string][char]0x51E6 + [char]0x7406 + [char]0x6642 + [char]0x9593) -Tag 'JDL') 'Split-mode file name carries the (Tag) suffix'
Assert-Equal 'Label.xlsx' (Get-ProcessTimeOutputFileName -Label 'Label' -Tag 'JDL' -Single $true) 'Single mode drops the tag suffix entirely'

Assert-Equal 'C:\J4\JDS' (Resolve-ProcessTimeOutputDir -Tag 'JDS' -DirByTag @{ JDS = 'C:\J4\JDS' } -DefaultDir 'C:\Default') 'a configured per-tag directory wins'
Assert-Equal 'C:\Default' (Resolve-ProcessTimeOutputDir -Tag 'JRV' -DirByTag @{ JDS = 'C:\J4\JDS' } -DefaultDir 'C:\Default') 'a tag with no override falls back to the default directory'
Assert-Equal 'C:\Default' (Resolve-ProcessTimeOutputDir -Tag 'JRV' -DirByTag @{ JRV = '' } -DefaultDir 'C:\Default') 'a blank per-tag override is treated as unset'

# ProcessTime workbook formatting is COM-bound and cannot be executed in the
# pure test suite.  Guard the office-PC compatibility fix statically: passing
# two Cells COM proxies to Worksheet.Range caused DISP_E_TYPEMISMATCH on a
# real Windows PowerShell/Excel run; A1-address Range calls avoid that binder.
$processTimeSource = Get-Content -LiteralPath (Join-Path (Split-Path $here -Parent) 'ProcessTime.ps1') -Raw
Assert-True ($processTimeSource -notmatch '\$ws\.Range\(\$ws\.Cells\.Item') 'ProcessTime formatting never passes Cells COM proxies to Worksheet.Range'
Assert-True ($processTimeSource -match '\$ws\.Range\(\(''A1:H\{0\}'' -f \$finalRow\)\)') 'ProcessTime table formatting uses a COM-safe A1 address'

# A COM formatting step failing must never be silently swallowed -- each
# formatting try/catch block around the table write has to log a Write-Warning
# (with the failing step + output path) instead of a bare `catch {}`, so a
# real Excel/PowerShell incompatibility is diagnosable instead of showing up
# only as "the saved workbook just isn't formatted".
$formattingRegion = if ($processTimeSource -match '(?s)# Renumber retained \+ appended records.*?if \(\$isNew\)') { $Matches[0] } else { '' }
Assert-True (-not [string]::IsNullOrEmpty($formattingRegion)) 'ProcessTime workbook formatting region is present for the empty-catch regression check'
Assert-True ($formattingRegion -notmatch 'catch\s*\{\s*\}') 'ProcessTime workbook formatting no longer uses a bare catch {} that silently swallows COM errors'
Assert-Equal 7 (([regex]::Matches($formattingRegion, 'Write-Warning \("ProcessTime workbook formatting:')).Count) 'every formatting sub-step (range resolve, autofilter/borders, header fill/font, font/row-height, alignment, row fill, column widths) warns on failure'

# Windows PowerShell 5.1 can fail when @() directly wraps a List[object]
# fetched through a hashtable index. Exercise the production workaround with
# the same collection shape used by ProcessTime.ps1's Split output buckets.
$bucketTable = @{}
$bucketTable['JDL'] = New-Object System.Collections.Generic.List[object]
$bucketTable['JDL'].Add([pscustomobject]@{ CorrelId = 'JIDSA01S' })
$bucketTable['JDL'].Add([pscustomobject]@{ CorrelId = 'JIDSA02S' })
$bucketRows = ConvertTo-ProcessTimeBucketArray -Bucket $bucketTable['JDL']
Assert-Equal 2 $bucketRows.Count 'generic output bucket materializes without the PowerShell 5.1 hashtable/list binder crash'
Assert-Equal 'JIDSA01S' $bucketRows[0].CorrelId 'bucket materialization preserves row order'
$oneRowBucket = New-Object System.Collections.Generic.List[object]
$oneRowBucket.Add([pscustomobject]@{ CorrelId = 'JIDSA03S' })
$oneBucketRows = ConvertTo-ProcessTimeBucketArray -Bucket $oneRowBucket
Assert-True ($oneBucketRows -is [array]) 'one-row bucket remains an array instead of being pipeline-unrolled to a scalar'
Assert-Equal 1 $oneBucketRows.Count 'one-row bucket keeps an accurate Count'

# ---------------------------------------------------------------------------
# Get-ProcessTimeCheckSummaryLine (v2.15.0: end-of-run manual-check summary)
# ---------------------------------------------------------------------------
Assert-Equal '' (Get-ProcessTimeCheckSummaryLine -CorrelId 'JIDSM01S' -GiftMatched $true -GfixMatched $true) 'both sides matched -> nothing to report'
Assert-Equal 'JIDSM01S  -- GIFT: not detected' `
    (Get-ProcessTimeCheckSummaryLine -CorrelId 'JIDSM01S' -GiftMatched $false -GfixMatched $true) 'GIFT-only miss reports just the GIFT side'
Assert-Equal 'JIDSM01S  -- GFIX: not detected' `
    (Get-ProcessTimeCheckSummaryLine -CorrelId 'JIDSM01S' -GiftMatched $true -GfixMatched $false) 'GFIX-only miss reports just the GFIX side'
Assert-Equal 'JIDSM01S  -- GIFT: no exportable picture; GFIX: end time not read from OCR' `
    (Get-ProcessTimeCheckSummaryLine -CorrelId 'JIDSM01S' -GiftMatched $false -GfixMatched $false -GiftNote 'no exportable picture' -GfixNote 'end time not read from OCR') 'both sides miss with their own notes'
Assert-Equal 'JIDSM01S  -- GIFT: not detected' `
    (Get-ProcessTimeCheckSummaryLine -CorrelId 'JIDSM01S' -GiftMatched $false -GfixMatched $true -GiftNote '') 'a blank note falls back to the generic not-detected text'

exit (Complete-Tests)
