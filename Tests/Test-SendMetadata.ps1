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

exit (Complete-Tests)
