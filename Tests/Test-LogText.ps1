#Requires -Version 5.1
# Test-LogText.ps1 -- kernel/LogText.ps1 against the real gfix-recv samples
# (Tests/fixtures/gfix-recv: a trimmed GFIXReceive.log with the SJIS colon
# byte left in, one GoAnywhere job log, a Jenkins report page excerpt).

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'LogText.ps1')

Reset-Tests 'LogText'
$fx = Join-Path (Join-Path $here 'fixtures') 'gfix-recv'

# ---------------------------------------------------------------- mixed bytes
$colon = [char]0xFF1A
$b = [byte[]](0x61, 0x81, 0x46, 0x62)                                     # a <SJIS colon> b
$r = ConvertFrom-EbiMixedBytes -Bytes $b
Assert-Equal ('a' + $colon + 'b') $r['text'] 'mixed: a stray CP932 pair decodes as CP932'
Assert-Equal 1 $r['foreignPairs'] 'mixed: the pair is counted'
$utf = (New-Object System.Text.UTF8Encoding($false)).GetBytes([string][char]0x5E74 + 'x')   # year kanji in UTF-8
$r = ConvertFrom-EbiMixedBytes -Bytes $utf
Assert-Equal ([string][char]0x5E74 + 'x') $r['text'] 'mixed: valid UTF-8 stays UTF-8'
Assert-Equal 0 $r['foreignPairs'] 'mixed: no foreign pair in clean UTF-8'
$r = ConvertFrom-EbiMixedBytes -Bytes ([byte[]](0x41, 0xFF, 0x42))
Assert-Equal ('A' + [char]0xFFFD + 'B') $r['text'] 'mixed: a lone bad byte becomes U+FFFD, not dropped'
Assert-Equal 1 $r['badBytes'] 'mixed: and is counted'
$r = ConvertFrom-EbiMixedBytes -Bytes ([byte[]](0xEF, 0xBB, 0xBF, 0x41))
Assert-Equal 'A' $r['text'] 'mixed: a UTF-8 BOM is skipped'
Assert-Equal '' (ConvertFrom-EbiMixedBytes -Bytes ([byte[]]@()))['text'] 'mixed: empty in, empty out'

$logBytes = [System.IO.File]::ReadAllBytes((Join-Path $fx 'GFIXReceive.sample.log'))
$dec = ConvertFrom-EbiMixedBytes -Bytes $logBytes
Assert-True ($dec['foreignPairs'] -ge 2) 'mixed: the real log has its SJIS colons decoded'
Assert-Equal 0 $dec['badBytes'] 'mixed: the real log has no undecodable byte left'
Assert-True ($dec['text'].Contains('update count' + $colon + '1')) 'mixed: "update count<colon>1" reads as the operator sees it'
$sj = Get-EbiCodePage -CodePage 932
Assert-Equal 932 $sj.CodePage 'codepage: 932 is available'

# ---------------------------------------------------------------- lines
Assert-Equal 3 @(Get-EbiTextLines -Text "a`r`nb`nc`n").Count 'lines: CRLF and LF split; the final newline is not a line'
Assert-Equal 0 @(Get-EbiTextLines -Text '').Count 'lines: empty -> none'
Assert-Equal 2 @(Get-EbiTextLines -Text "a`n`n").Count 'lines: an inner empty line is kept'

# ---------------------------------------------------------------- blocks
$lines = @(Get-EbiTextLines -Text $dec['text'])
$sp = '^(?<tag>\S+) START (?<key>\S+)\s*$'
$ep = '^(?<tag>\S+) END (?<key>\S+)\s*$'
$all = Get-EbiLogBlocks -Lines $lines -StartPattern $sp -EndPattern $ep -AfterEnd 1
Assert-Equal 3 @($all['blocks']).Count 'blocks: three STARTs, three blocks'
$k = 'JJPCRS1220260706110052987442'
$one = Get-EbiLogBlocks -Lines $lines -StartPattern $sp -EndPattern $ep -Keys @($k) -AfterEnd 1
Assert-Equal 1 @($one['blocks']).Count 'blocks: keyed -> just that one'
$blk = $one['blocks'][0]
Assert-True ([bool]$blk['complete']) 'blocks: START..END is complete'
Assert-Equal ('GFIXReceiver START ' + $k) $blk['lines'][0] 'blocks: first line is the START line'
Assert-True ($blk['lines'][$blk['lines'].Count - 1] -match '^2026') 'blocks: AfterEnd 1 keeps the date stamp after END'
Assert-Equal ('GFIXReceiver END ' + $k) $blk['lines'][$blk['lines'].Count - 2] 'blocks: END is the line before it'
$dang = @($all['blocks'] | Where-Object { $_['key'] -like '*DANGLING*' })
Assert-Equal 1 $dang.Count 'blocks: a START without END is still returned'
Assert-True (-not [bool]$dang[0]['complete']) 'blocks: ... flagged incomplete'
$miss = Get-EbiLogBlocks -Lines $lines -StartPattern $sp -EndPattern $ep -Keys @('NOPE') -AfterEnd 1
Assert-Equal 'NOPE' (@($miss['missing']) -join ',') 'blocks: a key with no block is reported missing'

# ---------------------------------------------------------------- line hits
$hits = @(Find-EbiLineHits -Lines $blk['lines'] -Patterns @('FileGIFTGetBLBean\.execute\(\) fileName', 'file stored'))
Assert-Equal 2 $hits.Count 'hits: the two marker lines of the block'
Assert-True ($hits[0]['text'] -match 'execute\(\) fileName : \[/appl/JPC/gfix/recv/JJPCRS1220260706110052987442\]') 'hits: fileName line names this transfer'
Assert-True ($hits[1]['text'] -match 'file stored : \[/files/JPCVer/data/report/F\d{12}\.csv\]') 'hits: file stored line names the report file'
Assert-True ($hits[0]['index'] -lt $hits[1]['index']) 'hits: in line order'
$job = @(Get-EbiTextLines -Text ([System.IO.File]::ReadAllText((Join-Path $fx 'job-1000004619654.log'))))
$jh = @(Find-EbiLineHits -Lines $job -Patterns @("successfully uploaded to", "INFO\s+Command: "))
Assert-Equal 2 $jh.Count 'hits: job log -> upload line + Command line'
Assert-Equal 23 $jh[0]['index'] 'hits: upload line is line 24 (0-based 23)'
Assert-Equal 33 $jh[1]['index'] 'hits: Command line is line 34 (0-based 33)'

# ---------------------------------------------------------------- highlight width (calibrated on the sample workbook)
Assert-Equal 1 (Get-EbiTextCellUnits -Text 'a') 'units: ASCII is 1'
Assert-Equal 2 (Get-EbiTextCellUnits -Text ([string][char]0x5E74)) 'units: kanji is 2'
Assert-Equal 1 (Get-EbiTextCellUnits -Text ([string][char]0xFF71)) 'units: half-width katakana is 1'
$t237 = 'x' * 237; $t139 = 'x' * 139; $t254 = 'x' * 254
Assert-Equal 80 (Get-EbiHighlightEndColumn -Text $t237 -StartColumn 2) 'width: 237 units from B ends at CB (sample row 32)'
Assert-Equal 48 (Get-EbiHighlightEndColumn -Text $t139 -StartColumn 2) 'width: 139 units from B ends at AV (sample row 42)'
Assert-Equal 86 (Get-EbiHighlightEndColumn -Text $t254 -StartColumn 2) 'width: 254 units from B ends at CH (sample row 88)'
Assert-Equal 2 (Get-EbiHighlightEndColumn -Text '' -StartColumn 2) 'width: empty text still one column'
Assert-Equal 50 (Get-EbiHighlightEndColumn -Text $t237 -StartColumn 2 -MaxColumn 50) 'width: capped at MaxColumn'
Assert-Equal 81 (Get-EbiHighlightEndColumn -Text $t237 -StartColumn 2 -PadColumns 1) 'width: padding widens'
Assert-Equal 80 (Get-EbiHighlightEndColumn -Text ($t237 + '   ') -StartColumn 2) 'width: trailing blanks do not count'

# ---------------------------------------------------------------- clock / window
Assert-Equal '11:45:00' (ConvertTo-EbiClockText -Value '11:44:59.9999999999984025') 'clock: float artefact rounds up'
Assert-Equal '13:15:00' (ConvertTo-EbiClockText -Value '13:15:00.0000000000031950') 'clock: float artefact rounds down'
Assert-Equal '10:30:00' (ConvertTo-EbiClockText -Value '10:30:00.000') 'clock: plain'
Assert-Equal '13:20:27' (ConvertTo-EbiClockText -Value '13:20:26.99999999999675925') 'clock: seconds kept'
Assert-Equal '11:45:00' (ConvertTo-EbiClockText -Value 0.48958333333) 'clock: a day fraction'
Assert-Equal '11:45:00' (ConvertTo-EbiClockText -Value 46304.4895833333) 'clock: a full serial'
Assert-Equal '' (ConvertTo-EbiClockText -Value '-') 'clock: "-" is unreadable'
Assert-Equal '09:05:00' (ConvertTo-EbiClockText -Value '9:05') 'clock: a one-digit hour reads'
$w = Get-EbiTimeAround -Date '2026-10-08' -Clock '13:44:59.9999999999968050' -BeforeMinutes 2 -AfterMinutes 13
Assert-True ([bool]$w['ok']) 'around: ok'
Assert-Equal '2026-10-08T13:45:00' $w['at'] 'around: at'
Assert-Equal '2026-10-08T13:43:00' $w['from'] 'around: from = at - 2 min'
Assert-Equal '2026-10-08T13:58:00' $w['to'] 'around: to = at + 13 min'
Assert-True (-not (Get-EbiTimeAround -Date 'x' -Clock '1:00')['ok']) 'around: a bad date is not ok'
Assert-True (-not (Get-EbiTimeAround -Date '2026-10-08' -Clock '')['ok']) 'around: a bad time is not ok'

# ---------------------------------------------------------------- pairs
$L = @(@{ name = 'F202608270023.csv'; lines = 5; order = 'F202608270023' }, @{ name = 'F202608270024.csv'; lines = 30; order = 'F202608270024' })
$R = @(@{ name = 'F202610080002.csv'; lines = 30; order = 'F202610080002' }, @{ name = 'F202610080001.csv'; lines = 5; order = 'F202610080001' })
$p = Select-EbiFilePairs -Left $L -Right $R
Assert-Equal 'ok' $p['code'] 'pairs: in order with equal counts -> ok'
Assert-Equal 'F202610080001.csv' $p['pairs'][0]['right'] 'pairs: first left with first-arrived right'
$R2 = @(@{ name = 'B1'; lines = 30; order = '1' }, @{ name = 'B2'; lines = 5; order = '2' })
$p = Select-EbiFilePairs -Left $L -Right $R2
Assert-Equal 'ok' $p['code'] 'pairs: order disagrees but counts are unique -> paired by count'
Assert-Equal 'B2' (@($p['pairs'] | Where-Object { $_['left'] -eq 'F202608270023.csv' })[0]['right']) 'pairs: the 5-line file goes with the 5-line file'
$R3 = @(@{ name = 'C1'; lines = 7; order = '1' }, @{ name = 'C2'; lines = 5; order = '2' })
$p = Select-EbiFilePairs -Left $L -Right $R3
Assert-Equal 'unknown' $p['code'] 'pairs: a count with no partner -> unknown'
Assert-Equal 'C1' (@($p['rightOver']) -join ',') 'pairs: the leftover is named'
$Ldup = @(@{ name = 'A'; lines = 5; order = '1' }, @{ name = 'B'; lines = 5; order = '2' })
$Rdup = @(@{ name = 'X'; lines = 5; order = '1' })
$p = Select-EbiFilePairs -Left $Ldup -Right $Rdup
Assert-Equal 'unknown' $p['code'] 'pairs: two same-count files for one partner -> never a guess'
$p = Select-EbiFilePairs -Left @() -Right $R
Assert-Equal 'unknown' $p['code'] 'pairs: nothing on one side -> unknown'

# ---------------------------------------------------------------- compare
$c = Compare-EbiTextLines -Left "a`r`nb`r`n" -Right "a`nb"
Assert-True ([bool]$c['identical']) 'compare: CRLF vs LF and the final newline are ignored'
$c = Compare-EbiTextLines -Left "a`nb`nc" -Right "a`nX`nc"
Assert-True (-not [bool]$c['identical']) 'compare: a changed line differs'
Assert-Equal 2 $c['firstDiff'] 'compare: first difference at line 2'
$c = Compare-EbiTextLines -Left "a`nb" -Right "a`nb`nc"
Assert-Equal 3 $c['firstDiff'] 'compare: an extra line differs where it starts'

# ---------------------------------------------------------------- columns
Assert-Equal 80 (ConvertTo-EbiColumnNumber 'CB') 'col: CB is 80'
Assert-Equal 2 (ConvertTo-EbiColumnNumber 'b') 'col: lower case'
Assert-Equal 7 (ConvertTo-EbiColumnNumber 7) 'col: a number passes'
Assert-Equal 0 (ConvertTo-EbiColumnNumber 'A1') 'col: a cell address is not a column'
Assert-Equal 'CB' (ConvertTo-EbiColumnLetter 80) 'col: 80 is CB'
Assert-Equal 'AV' (ConvertTo-EbiColumnLetter 48) 'col: 48 is AV'
Assert-Equal 'ZZ' (ConvertTo-EbiColumnLetter 702) 'col: 702 is ZZ'

exit (Complete-Tests)
