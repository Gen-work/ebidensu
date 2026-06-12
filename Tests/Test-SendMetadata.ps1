#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'SendMetadata.ps1')

Reset-Tests 'SendMetadata'

# -- Get-SendFirstToken --
Assert-Equal 'JIDSF48S' (Get-SendFirstToken '  JIDSF48S  20260529 OK ') 'first token from padded line'
Assert-Equal ''         (Get-SendFirstToken '')                          'empty line -> empty token'
Assert-Equal ''         (Get-SendFirstToken '   ')                       'whitespace line -> empty token'

# -- Get-SendLineTextFromWords: rebuild spacing from word boxes --
function New-Word([string]$Text, [double]$X, [double]$Width) {
    [pscustomobject]@{ Text = $Text; X = $X; Y = 0.0; Width = $Width; Height = 10.0 }
}
$words = @(
    (New-Word 'AAA' 0 30),     # charW = 10
    (New-Word 'BBB' 50 30)     # gap 20 -> 2 spaces
)
Assert-Equal 'AAA  BBB' (Get-SendLineTextFromWords $words) 'gap of 2 char widths -> 2 spaces'

$words2 = @(
    (New-Word 'AAA' 0 30),
    (New-Word 'CC' 31 20)      # gap 1 < 0.5*charW -> glued
)
Assert-Equal 'AAACC' (Get-SendLineTextFromWords $words2) 'sub-threshold gap -> no space'

$words3 = @(
    (New-Word 'BBB' 50 30),
    (New-Word 'AAA' 0 30)      # out of order -> sorted by X
)
Assert-Equal 'AAA  BBB' (Get-SendLineTextFromWords $words3) 'words sorted by X before join'
Assert-Equal '' (Get-SendLineTextFromWords @()) 'no words -> empty line'

# -- ConvertTo-SendTextLines: strings, Text objects, Words objects --
$mixed = @(
    'PLAIN LINE',
    [pscustomobject]@{ Text = 'TEXTONLY'; Words = @() },
    [pscustomobject]@{ Text = 'IGNORED'; Words = $words },
    '',
    $null,
    '   '
)
$lines = ConvertTo-SendTextLines $mixed
Assert-Equal 3 $lines.Count 'blank/null lines dropped'
Assert-Equal 'PLAIN LINE' $lines[0] 'plain string kept'
Assert-Equal 'TEXTONLY' $lines[1] 'empty Words falls back to .Text'
Assert-Equal 'AAA  BBB' $lines[2] 'word boxes win over raw .Text'

# -- Test-SendZeroByteText --
Assert-True (Test-SendZeroByteText @('FILE SIZE 0 BYTE'))        'default pattern: 0 BYTE'
Assert-True (Test-SendZeroByteText @('size=0bytes'))             'default pattern: 0bytes glued'
Assert-True (-not (Test-SendZeroByteText @('SIZE 10 BYTES')))    '10 bytes is not zero'
Assert-True (-not (Test-SendZeroByteText @()))                   'no lines -> not zero'
Assert-True (Test-SendZeroByteText @('EMPTY-MARK') 'EMPTY-MARK') 'custom pattern override'

# -- Get-SendRowNumberGuess --
$recordLines = @(
    'HEADER NO DIGITS',
    '000001 JIDSF48S DATA',
    '000123 JIDSF48S DATA',
    '99 TRAILING IGNORED BECAUSE SMALLER'
)
Assert-Equal 123 (Get-SendRowNumberGuess $recordLines) 'max leading integer wins'
Assert-Equal 0 (Get-SendRowNumberGuess @('NO NUMBERS HERE')) 'no leading integer -> 0'

# -- Build-SendMetadataRecord --
$meta = Build-SendMetadataRecord -CorrelIdS 'JIDSF48S' -ExcelName 'LJRVWD64' -ImageCount 2 -TextLines $recordLines
Assert-Equal 'JIDSF48S' $meta.CorrelIdS          'record carries correl id'
Assert-Equal 4 $meta.OcrLineCount                'line count'
Assert-Equal 123 $meta.RowNumberGuess            'row guess wired in'
Assert-Equal 'HEADER' $meta.FirstRecordToken     'first token of first line'
Assert-Equal '99' $meta.LastRecordToken          'first token of last line (mirrors gift side)'
Assert-Equal 'False' ([string]$meta.ZeroByte)    'no zero-byte pattern'
Assert-Equal '1' ([string]$meta.Confidence)      'all three confidence parts present -> 1'
Assert-Equal '2' $meta.MetadataVersion           'metadata version 2'

$metaEmpty = Build-SendMetadataRecord -CorrelIdS 'X' -ExcelName 'Y' -ImageCount 0 -TextLines @()
Assert-Equal '0' ([string]$metaEmpty.Confidence) 'no OCR evidence -> confidence 0'
Assert-Equal '' $metaEmpty.FirstRecordToken      'empty input -> empty tokens'

# -- Compare-SendGiftMetadata --
function New-GiftRow([long]$Size, [int]$Rows, [string]$FirstTok, [string]$LastTok) {
    [pscustomobject]@{
        FileName = 'JIDSF48S'; SizeBytes = $Size; MaxRowNumber = $Rows
        FirstRecordToken = $FirstTok; LastRecordToken = $LastTok
    }
}

# full match
$send = Build-SendMetadataRecord -CorrelIdS 'JIDSF48S' -ExcelName 'E' -ImageCount 1 -TextLines @('000001 A', '000123 B')
$gift = New-GiftRow 1024 123 '000001' '000123'
$cmp = Compare-SendGiftMetadata $send $gift
Assert-Equal 'match' $cmp.Verdict 'rows + tokens agree -> match'
Assert-Equal 3 $cmp.MatchCount    'three checks matched'
Assert-Equal 0 $cmp.MismatchCount 'no mismatches'

# row count disagrees -> mismatch wins
$giftBad = New-GiftRow 1024 999 '000001' '000123'
$cmpBad = Compare-SendGiftMetadata $send $giftBad
Assert-Equal 'mismatch' $cmpBad.Verdict 'row mismatch -> verdict mismatch'

# no OCR evidence -> unknown, never mismatch
$cmpEmpty = Compare-SendGiftMetadata $metaEmpty $gift
Assert-Equal 'unknown' $cmpEmpty.Verdict       'empty send side -> unknown'
Assert-Equal 0 $cmpEmpty.MismatchCount         'absence of evidence is not mismatch'

# zero-byte agreement alone is a match
$sendZero = Build-SendMetadataRecord -CorrelIdS 'J' -ExcelName 'E' -ImageCount 1 -TextLines @('FILE SIZE 0 BYTE')
$giftZero = New-GiftRow 0 0 '' ''
$cmpZero = Compare-SendGiftMetadata $sendZero $giftZero
Assert-Equal 'match' $cmpZero.Verdict 'both sides zero-byte -> match'

# send says zero but gift has data -> mismatch
$cmpZeroBad = Compare-SendGiftMetadata $sendZero $gift
Assert-Equal 'mismatch' $cmpZeroBad.Verdict 'send zero vs gift data -> mismatch'

# CSV round-trip: Import-Csv turns everything into strings; compare must cope
$csvSend = $send | Select-Object *
$csvSend.ZeroByte = 'False'
$csvSend.RowNumberGuess = '123'
$cmpCsv = Compare-SendGiftMetadata $csvSend $gift
Assert-Equal 'match' $cmpCsv.Verdict 'string-typed send row (CSV round-trip) still matches'

# ============================================================
# Stage 2 evidence-verdict helpers (operator review rules)
# ============================================================
# Japanese fragments built from code points (test source stays ASCII).
$shiyou  = [string]([char]0x4F7F) + [char]0x7528                                  # 'used'
$dataNo  = [string]([char]0x30C7) + [char]0x30FC + [char]0x30BF + [char]0x306E   # 'de-ta no'
$hajime  = $dataNo + [char]0x59CB + [char]0x3081                                  # 'de-ta no hajime'
$owari   = $dataNo + [char]0x7D42 + [char]0x308F + [char]0x308A                   # 'de-ta no owari'
$owari2  = $dataNo + [char]0x7D42 + [char]0x308A                                  # alt spelling 'owari'

# -- Get-SendRowLabel --
Assert-Equal '000001' (Get-SendRowLabel 1)       'row 1 -> 000001'
Assert-Equal '004644' (Get-SendRowLabel 4644)    'row 4644 -> 004644'
Assert-Equal '1234567' (Get-SendRowLabel 1234567) 'wider than 6 digits kept'

# -- Test-SendRowNumberPresent --
Assert-True (Test-SendRowNumberPresent @('000003 X2 DATA') '000003')        'leading row label found'
Assert-True (Test-SendRowNumberPresent @('xx 004644 tail') '004644')        'mid-line row label found'
Assert-True (-not (Test-SendRowNumberPresent @('0000031 DATA') '000003'))   'longer digit run does not count'
Assert-True (-not (Test-SendRowNumberPresent @('no labels') '000003'))      'absent label'

# -- Find-SendRecordByRowNumber --
$ocrHead = @(
    'VIEW LJOD.C.VER.JOD155',
    '000001 0000001001B40500015A FOO',
    '000002 0000002001B40500015A'
)
Assert-Equal '0000001001B40500015A FOO' (Find-SendRecordByRowNumber $ocrHead '000001') 'record text after row label'
Assert-Equal 'X2REST' (Find-SendRecordByRowNumber @('000003X2REST') '000003') 'glued label still split'
Assert-True ($null -eq (Find-SendRecordByRowNumber $ocrHead '000009')) 'missing label -> null'
Assert-True ($null -eq (Find-SendRecordByRowNumber @('0000031 DATA') '000003')) 'lookahead rejects longer digit run'

# -- Test-SendZeroByteImage --
$cylZero = @('LJOD.C.VER.JOD382', ($shiyou + ' CYLINDERS . . : 0'))
$cylUsed = @(($shiyou + ' CYLINDERS . . : 1'))
Assert-True (Test-SendZeroByteImage $cylZero)            'used CYLINDERS 0 -> zero'
Assert-True (-not (Test-SendZeroByteImage $cylUsed))     'used CYLINDERS 1 -> not zero'
$beginEndSame = @($hajime, $owari2)
Assert-True (Test-SendZeroByteImage $beginEndSame)       'begin+end markers, no 000001 -> zero'
$beginEndData = @($hajime, '000001 SOMEDATA', $owari)
Assert-True (-not (Test-SendZeroByteImage $beginEndData)) 'begin+end with 000001 -> not zero'
Assert-True (-not (Test-SendZeroByteImage @($hajime)))    'begin only (head image) -> not zero'
Assert-True (Test-SendZeroByteImage @('EMPTY-MARK') 'EMPTY-MARK') 'custom pattern override'

# -- Get-SendPrefixSimilarity --
Assert-Equal 1 ([int](Get-SendPrefixSimilarity 'ABCDEF' 'ABCDEF')) 'identical -> 1.0'
Assert-True ((Get-SendPrefixSimilarity '0000001001B40500015AXX' '0000801001B40500015AXX' 20) -ge 0.9) 'one OCR slip in 20 chars stays high'
Assert-True ((Get-SendPrefixSimilarity 'TOTALLYDIFFERENT' '0000001001B4050001' 20) -lt 0.5) 'different strings score low'
Assert-Equal 0 ([int](Get-SendPrefixSimilarity '' 'ABC')) 'one empty side -> 0'

# -- Compare-SendGiftEvidence: 0-byte gift --
function New-GiftMetaRow([long]$Size, [int]$Rows, [string]$First, [string]$Last) {
    [pscustomobject]@{
        FileName = 'JIDSC03S'; SizeBytes = $Size; MaxRowNumber = $Rows
        FirstRecord = $First; LastRecord = $Last
        FirstRecordToken = (Get-SendFirstToken $First); LastRecordToken = (Get-SendFirstToken $Last)
    }
}
$giftZero2 = New-GiftMetaRow 0 0 '' ''
$cmpZeroOk = Compare-SendGiftEvidence -GiftRow $giftZero2 -ImageTextSets @(,@($cylZero))
Assert-Equal 'ok' $cmpZeroOk.Verdict 'gift 0 bytes + CYLINDERS-0 image -> ok'
$cmpZeroNg = Compare-SendGiftEvidence -GiftRow $giftZero2 -ImageTextSets @(,@(@('000001 DATA HERE')))
Assert-Equal 'ng' $cmpZeroNg.Verdict 'gift 0 bytes but send shows 000001 -> ng'
$cmpZeroUnk = Compare-SendGiftEvidence -GiftRow $giftZero2 -ImageTextSets @(,@(@('NOISE ONLY')))
Assert-Equal 'unknown' $cmpZeroUnk.Verdict 'gift 0 bytes, no evidence either way -> unknown'

# -- Compare-SendGiftEvidence: data gift --
$gift3 = New-GiftMetaRow 300 3 '0000001001B40500015A TAIL' '0001548003X2 END'
$headImg = @('000001 0000001001B40500015A TAIL', '000002 MIDDLE', '000003 0001548003X2 END')
$cmpOk = Compare-SendGiftEvidence -GiftRow $gift3 -ImageTextSets @(,@($headImg))
Assert-Equal 'ok' $cmpOk.Verdict 'max row found + first/last tokens match -> ok'

$badImg = @('000001 ZZZZZZZZZZZZZZZZZZZZZ', '000003 0001548003X2 END')
$cmpNg = Compare-SendGiftEvidence -GiftRow $gift3 -ImageTextSets @(,@($badImg))
Assert-Equal 'ng' $cmpNg.Verdict 'first record disagrees -> ng'

$noMaxImg = @('000001 0000001001B40500015A TAIL', '000002 MIDDLE')
$cmpUnk = Compare-SendGiftEvidence -GiftRow $gift3 -ImageTextSets @(,@($noMaxImg))
Assert-Equal 'unknown' $cmpUnk.Verdict 'max row number not visible -> unknown'

# fuzzy: token differs by one OCR slip but 20-char prefix stays >= 80%
$fuzzyImg = @('000001 0000001001B40500015B TAIL', '000003 0001548003X2 END')
$cmpFuzzy = Compare-SendGiftEvidence -GiftRow $gift3 -ImageTextSets @(,@($fuzzyImg))
Assert-Equal 'ok' $cmpFuzzy.Verdict 'single OCR slip passes via prefix similarity'
$fuzzyCheck = @($cmpFuzzy.Checks | Where-Object { $_.Name -eq 'FirstRecord' })[0]
Assert-Equal 'fuzzy' $fuzzyCheck.Status 'slip is reported as fuzzy, not exact match'

# two images (head strip + tail strip), tail carries the max row
$gift4644 = New-GiftMetaRow 1672702 4644 '0000001001B40500015A X' '0001548003X2'
$head4644 = @('000001 0000001001B40500015A X', '000017 SOMETHING')
$tail4644 = @('004628 NEARTAIL', '004644 0001548003X2')
$cmpTwo = Compare-SendGiftEvidence -GiftRow $gift4644 -ImageTextSets @(@($head4644), @($tail4644))
Assert-Equal 'ok' $cmpTwo.Verdict 'head+tail images combine to ok'

# CSV round-trip gift row (strings everywhere)
$giftCsv = $gift3 | Select-Object *
$giftCsv.SizeBytes = '300'
$giftCsv.MaxRowNumber = '3'
$cmpCsv2 = Compare-SendGiftEvidence -GiftRow $giftCsv -ImageTextSets @(,@($headImg))
Assert-Equal 'ok' $cmpCsv2.Verdict 'string-typed gift row (CSV round-trip) still ok'

# -- compact-form fallbacks: ja OCR returns one word per CHARACTER, so the
# spacing rebuild can yield '0 0 2 6 4 0 5 1 1 2 ...' lines --
Assert-Equal '002640X' (ConvertTo-SendCompactLine ' 0 0 2 6 4 0 X ') 'compact strips all spaces'

$spaced = @('0 0 2 6 4 0 5 1 1 2 7 2 0 0 1 9')
Assert-True (Test-SendRowNumberPresent $spaced '002640') 'per-char spaced row label found via compact form'
Assert-Equal '5112720019' (Find-SendRecordByRowNumber $spaced '002640') 'compact record extracted after glued label'

$sendRec = '5112720019999999990604'
$giftRec = '5 1 1 2 7 2 0 0 1 9 9 9 9 9 9 9 9 9 9 0 6 0 4'
$chkCompact = Compare-SendRecordCheck 'FirstRecord' $sendRec $giftRec 0.8 20
Assert-Equal 'fuzzy' $chkCompact.Status 'compact prefix similarity rescues spaced record'

# end-to-end: per-char spaced head+tail images against a 3-row gift file
$giftSp = New-GiftMetaRow 300 3 '5112ABC' '5112XYZ'
$headSp = @('0 0 0 0 0 1 5 1 1 2 A B C', 'H E A D E R')
$tailSp = @('0 0 0 0 0 3 5 1 1 2 X Y Z')
$cmpSp = Compare-SendGiftEvidence -GiftRow $giftSp -ImageTextSets @(@($headSp), @($tailSp))
Assert-Equal 'ok' $cmpSp.Verdict 'per-char spaced OCR lines still verify to ok'

# zero-byte rule A with per-char spacing ('used CYLINDERS : 0')
$Lz = Get-SendZeroByteLabels
$cylSpaced = ([string]$Lz.Shiyou[0] + ' ' + [string]$Lz.Shiyou[1] + '   C Y L I N D E R S . . : 0')
Assert-True (Test-SendZeroByteImage @($cylSpaced)) 'spaced CYLINDERS 0 detected via compact form'

exit (Complete-Tests)
