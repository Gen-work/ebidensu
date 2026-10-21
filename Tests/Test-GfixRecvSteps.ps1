#Requires -Version 5.1
# Test-GfixRecvSteps.ps1 -- the steps the gfix-recv workflows chain, run for
# real (not dry) on the real samples in Tests/fixtures/gfix-recv, in the
# order the workflows call them:
#   track:    GoAnywhere text -> parse -> this job's rows -> Receive rows +
#             time span -> Jenkins text -> the received file -> pairing with
#             the GIFT side -> line compare -> Teams picture region
#   logs:     job log -> transfer file name
#   evidence: receive log -> the transfer's START..END block + marker lines
#             -> picture stack plan; Jenkins row boxes from ink bands
# plus kernel/Layout.ps1 and kernel/RichClip.ps1. UI / COM halves are not
# here (they need Windows); their DryRun contract is Test-StepDryRun.ps1.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path $repoRoot 'kernel/Json.ps1')
. (Join-Path $repoRoot 'kernel/Registry.ps1')
. (Join-Path $repoRoot 'kernel/Profile.ps1')
. (Join-Path $repoRoot 'kernel/Layout.ps1')
. (Join-Path $repoRoot 'kernel/RichClip.ps1')
. (Join-Path $repoRoot 'kernel/LogText.ps1')
. (Join-Path $repoRoot 'kernel/Key.ps1')
. (Join-Path $repoRoot 'kernel/Worklist.ps1')

Reset-Tests 'GfixRecvSteps'

$fx = Join-Path (Join-Path $here 'fixtures') 'gfix-recv'
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-gfix-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$prof = (Read-EbiProfile -Dir (Join-Path $repoRoot 'profiles/gfix-recv')).value
$reg = New-EbiRegistry -ModulesRoot (Join-Path $repoRoot 'modules')
$log = New-Object PSObject
$log | Add-Member -MemberType NoteProperty -Name Lines -Value (New-Object System.Collections.ArrayList)
$log | Add-Member -MemberType ScriptMethod -Name Info -Value { param($m) [void]$this.Lines.Add($m) }
$log | Add-Member -MemberType ScriptMethod -Name Warn -Value { param($m) [void]$this.Lines.Add($m) }
$log | Add-Member -MemberType ScriptMethod -Name Debug -Value { param($m) }
$ctx = @{ WorkDir = $tmp; RunId = 'test'; Profile = $prof; Log = $log; DryRun = $false; Session = @{}; Item = $null; KeyColumns = @('Excel_NAME') }

# Every step is loaded ONCE, dot-sourced at script scope (the runner does
# the same): a step's helpers and the kernel files it dot-sources must stay
# visible for every later call.
$loaded = @{}
foreach ($u in @('table.upsert', 'verify.time_window', 'verify.parse_text', 'verify.filter_records', 'screen.row_region', 'file.list', 'verify.pair_files',
                 'file.compare', 'file.read_text', 'file.convert_encoding', 'file.extract_blocks', 'excel.stack_plan', 'screen.list_rects',
                 'verify.derive_fields', 'human.paste', 'screen.launch_capture', 'browser.wait_for', 'browser.download_each', 'file.download', 'excel.tidy',
                 'verify.record_counts', 'verify.nearest_time', 'verify.owner_pattern', 'excel.open', 'verify.assert_each', 'flow.require',
                 'verify.find_token', 'screen.ocr', 'human.gate', 'excel.append_row', 'human.draft_mail', 'human.share',
                 'screen.read_line_numbers', 'verify.df_ranges', 'excel.export_pictures', 'screen.compare_strip')) {
    $r = . Import-EbiStep -Registry $reg -Use $u
    if (-not $r['ok']) { throw ('cannot load ' + $u + ': ' + $r['message']) }
    $loaded[$u] = $r['Entry']
}

function Invoke-TestStep {
    # Check inputs like the runner does, run, return the step's hashtable.
    param([string]$Use, [hashtable]$With)
    $e = $loaded[$Use]
    $res = Resolve-EbiStepInputs -Manifest $e['Manifest'] -With $With -Session $ctx['Session']
    if (-not $res['ok']) { throw ($Use + ' inputs: ' + $res['message']) }
    return (& $e['Invoke'] $res['In'] $ctx)
}

try {
    # ============================================================ track
    $gaText = [System.IO.File]::ReadAllText((Join-Path $fx 'goanywhere-list.sample.txt'))
    $win = Invoke-TestStep 'verify.time_window' @{ date = '2026-10-08'; clock = '13:44:59.9999999999968050'; beforeMinutes = $prof['pages']['goAnywhere']['windowBefore']; afterMinutes = $prof['pages']['goAnywhere']['windowAfter'] }
    Assert-Equal '2026-10-08 13:45' $win['atText'] 'track: the GoAnywhere wait text is the scheduled minute as the page prints it'
    Assert-True ($gaText.Contains($win['atText'])) 'track: ... and the sample page does contain it'
    Assert-Equal 16 @($win['minuteTexts']).Count 'track: every minute of the -2..+13 window is a ready text'
    Assert-Equal '2026-10-08 13:43|2026-10-08 13:58' (@($win['minuteTexts'])[0] + '|' + @($win['minuteTexts'])[15]) 'track: first and last ready minute'
    $rec = Invoke-TestStep 'verify.parse_text' @{ text = $gaText; grammar = $prof['grammar']['goAnywhere'] }
    Assert-Equal 12 $rec['recordCount'] 'track: 12 GoAnywhere rows parsed'
    Assert-Equal 0 $rec['unrecognized'] 'track: no GoAnywhere line left unrecognised'
    # A page that is not the expected one (one non-blank line, nothing the
    # grammar recognises) is a no_records failure, not a crash: counting a
    # single filtered line once threw under StrictMode on PS 5.1 and 7.
    $bad = Invoke-TestStep 'verify.parse_text' @{ text = "not a job list`r`n`r`n"; grammar = $prof['grammar']['goAnywhere'] }
    Assert-True ((-not $bad['ok']) -and $bad['failure'] -eq 'no_records' -and $bad['message'].Contains(' 1 non-blank line')) ('parse: an unrelated page is no_records over its 1 non-blank line: ' + $bad['message'])
    $mine = Invoke-TestStep 'verify.filter_records' @{ records = $rec['records']; where = @(@{ field = 'startTime'; op = 'within'; value = $win['window'] }) }
    Assert-Equal 2 $mine['matched'] 'track: the 13:45 job is two rows (Send + Receive)'
    Assert-Equal 1 $mine['first'] 'track: ... at the top of the list (newest first)'
    $recv = Invoke-TestStep 'verify.filter_records' @{ records = $mine['records']; where = @(@{ field = 'folder'; op = 'equals'; value = $prof['pages']['goAnywhere']['receiveFolder'] }); pluck = 'key'; spanFrom = 'startTime'; spanTo = 'endTime'; padBeforeSec = 2; padAfterSec = 10; expect = '1' }
    Assert-Equal '1000004619654' (@($recv['plucked']) -join ',') 'track: the Receive job number'
    Assert-Equal '2026-10-08T13:45:22' $recv['span']['from'] 'track: span starts 2 s before the Receive start'
    Assert-Equal '2026-10-08T13:45:37' $recv['span']['to'] 'track: span ends 10 s after the Receive end'
    Assert-True (-not $recv.Contains('warnings') -or @($recv['warnings']).Count -eq 0) 'track: one file expected, one found: no warning'
    $w11 = Invoke-TestStep 'verify.time_window' @{ date = '2026-10-08'; clock = '11:15:00.000'; beforeMinutes = 2; afterMinutes = 13 }
    $m11 = Invoke-TestStep 'verify.filter_records' @{ records = $rec['records']; where = @(@{ field = 'startTime'; op = 'within'; value = $w11['window'] }) }
    Assert-Equal 11 $m11['first'] 'track: the 11:15 job is rows 11-12 further down'
    $none = Invoke-TestStep 'verify.filter_records' @{ records = $rec['records']; where = @(@{ field = 'startTime'; op = 'within'; value = @{ from = '2026-10-08T15:00:00'; to = '2026-10-08T15:10:00' } }) }
    Assert-True (-not $none['ok'] -and $none['failure'] -eq 'not_found') 'track: a slot with no row yet is not_found (transient: refresh and retry)'

    $reg2 = Invoke-TestStep 'screen.row_region' @{ left = 590; right = 1560; top = 146; firstRowTop = 268; rowHeight = 31; first = $mine['first']; rows = $mine['matched'] }
    Assert-Equal 590 $reg2['x'] 'track: Teams picture x'
    Assert-Equal 146 $reg2['y'] 'track: Teams picture y (panel title)'
    Assert-Equal 970 $reg2['width'] 'track: Teams picture width (to the start-time column)'
    Assert-Equal 186 $reg2['height'] 'track: Teams picture height = title..2 rows + pad'
    $reg11 = Invoke-TestStep 'screen.row_region' @{ left = 590; right = 1560; top = 146; firstRowTop = 268; rowHeight = 31; first = 11; rows = 2 }
    Assert-Equal 10 $reg11['skippedRows'] 'track: rows lower in the list are reported as skipped above'

    $jkText = [System.IO.File]::ReadAllText((Join-Path $fx 'jenkins-report.sample.txt'))
    $jk = Invoke-TestStep 'verify.parse_text' @{ text = $jkText; grammar = $prof['grammar']['jenkinsReport'] }
    Assert-Equal 24 $jk['recordCount'] 'track: 24 Jenkins rows parsed'
    $jkMine = Invoke-TestStep 'verify.filter_records' @{ records = $jk['records']; where = @(@{ field = 'time'; op = 'within'; value = $recv['span'] }, @{ field = 'key'; op = 'matches'; value = $prof['pages']['jenkinsReport']['fileNamePattern'] }); pluck = 'key'; expect = 1 }
    Assert-Equal 'F202610080006.csv' (@($jkMine['plucked']) -join ',') 'track: the received file is the one stored inside the Receive span'

    # GIFT side vs GFIX side: two files each, order and line counts
    $giftDir = Join-Path $tmp 'DATA/GIFT/RJDSJM40'; $gfixDir = Join-Path $tmp 'DATA/GFIX/RJDSWM40'
    New-Item -ItemType Directory -Path $giftDir, $gfixDir -Force | Out-Null
    $short = "S,Q`r`nHDR,2026/05/15`r`nCOL,A,B`r`nNO DATA`r`n"
    $long = (1..40 | ForEach-Object { 'ROW' + $_ + ',x' }) -join "`r`n"
    [System.IO.File]::WriteAllText((Join-Path $giftDir 'F202608270023.csv'), $short)
    [System.IO.File]::WriteAllText((Join-Path $giftDir 'F202608270024.csv'), $long)
    [System.IO.File]::WriteAllText((Join-Path $gfixDir 'F202610080001.csv'), $short.Replace("`r`n", "`n"))
    [System.IO.File]::WriteAllText((Join-Path $gfixDir 'F202610080002.csv'), $long)
    $gift = Invoke-TestStep 'file.list' @{ dir = $giftDir; glob = '*.csv'; countLines = $true }
    $gfix = Invoke-TestStep 'file.list' @{ dir = $gfixDir; glob = '*.csv'; countLines = $true }
    Assert-Equal '4,40' (@($gift['files'] | ForEach-Object { $_['lines'] }) -join ',') 'track: line counts of the GIFT files'
    $pair = Invoke-TestStep 'verify.pair_files' @{ left = $gift['files']; right = $gfix['files']; longOver = 16 }
    Assert-Equal 'ok' $pair['code'] 'track: two files each, same counts in order -> paired'
    Assert-Equal 'F202610080001.csv' $pair['pairs'][0]['rightName'] 'track: first GIFT file with first GFIX file'
    Assert-True ((-not [bool]$pair['pairs'][0]['long']) -and [bool]$pair['pairs'][1]['long']) 'track: the 4-line file fits one DF screen, the 40-line one does not'
    $cmp = Invoke-TestStep 'file.compare' @{ pairs = $pair['pairs'] }
    Assert-Equal 'ok' $cmp['code'] 'track: identical content (CRLF vs LF ignored) -> ok'
    [System.IO.File]::WriteAllText((Join-Path $gfixDir 'F202610080002.csv'), $long.Replace('ROW7,', 'ROW7X,'))
    $cmp2 = Invoke-TestStep 'file.compare' @{ pairs = $pair['pairs'] }
    Assert-Equal 'ng' $cmp2['code'] 'track: one changed line -> ng'
    Assert-True ($cmp2['reason'] -match 'line 7') 'track: ... and the reason names the line'
    Assert-Equal 0 $cmp2['repaired'] 'track: without repairByContent the pairing is never touched'
    Assert-Equal 'F202610080002.csv' $cmp2['pairs'][1]['rightName'] 'track: ... and pairs echoes the input'

    # ---- several files of one job (10-13: JJDSJM10 x2, JJDSJM51 x2), the
    # first FileCount = 2 day. The GIFT side is named by the operator, the
    # GFIX side by arrival (F + timestamp); when both files have the same
    # line count the pairing follows the names, which may not be the order
    # the HOST sent them: every pair differs although every file arrived.
    $mFirst = @(foreach ($i in 1..12) { ('M' + $i + ',' + ($i * 7)) }) -join "`r`n"
    $mSecond = @(foreach ($i in 1..12) { ('N' + $i + ',' + ($i * 3)) }) -join "`r`n"
    $giftDir2 = Join-Path $tmp 'DATA/GIFT/JJDSJM51'; $gfixDir2 = Join-Path $tmp 'DATA/GFIX/JJDSWM51'
    New-Item -ItemType Directory -Path $giftDir2, $gfixDir2 -Force | Out-Null
    [System.IO.File]::WriteAllText((Join-Path $giftDir2 'ALOC9002.csv'), $mFirst)
    [System.IO.File]::WriteAllText((Join-Path $giftDir2 'ALOC9004.csv'), $mSecond)
    [System.IO.File]::WriteAllText((Join-Path $gfixDir2 'F202610130001.csv'), $mSecond)   # the second GIFT file arrived first
    [System.IO.File]::WriteAllText((Join-Path $gfixDir2 'F202610130002.csv'), $mFirst)
    $gift2 = Invoke-TestStep 'file.list' @{ dir = $giftDir2; glob = '*.csv'; countLines = $true; expect = 2 }
    Assert-True (-not $gift2.Contains('warnings') -or @($gift2['warnings']).Count -eq 0) 'multi: two GIFT files expected, two found: no warning'
    $gift2w = Invoke-TestStep 'file.list' @{ dir = $giftDir2; glob = '*.csv'; countLines = $true; expect = 3 }
    Assert-Equal 'count_mismatch' $gift2w['warnings'][0]['code'] 'multi: FileCount 3 but two files -> count_mismatch warning, not a failure'
    Assert-True ([bool]$gift2w['ok'] -and $gift2w['warnings'][0]['message'] -match 'expected 3 .* found 2') 'multi: ... naming expected and found'
    $gift2f = Invoke-TestStep 'file.list' @{ dir = $giftDir2; glob = '*.csv'; countLines = $true; expect = -1; expectFallback = '2' }
    Assert-True (-not $gift2f.Contains('warnings') -or @($gift2f['warnings']).Count -eq 0) 'multi: FileCount unknown, the Teams count (2) stands in: no warning'
    $gift2g = Invoke-TestStep 'file.list' @{ dir = $giftDir2; glob = '*.csv'; countLines = $true; expect = -1; expectFallback = '' }
    Assert-True (-not $gift2g.Contains('warnings') -or @($gift2g['warnings']).Count -eq 0) 'multi: no count known at all: nothing to check'
    $gfix2 = Invoke-TestStep 'file.list' @{ dir = $gfixDir2; glob = '*.csv'; countLines = $true }
    $pair2 = Invoke-TestStep 'verify.pair_files' @{ left = $gift2['files']; right = $gfix2['files']; longOver = 16 }
    Assert-Equal 'ok' $pair2['code'] 'multi: same line counts -> paired by name order (ALOC9002 with the first arrival)'
    $cmpA = Invoke-TestStep 'file.compare' @{ pairs = $pair2['pairs'] }
    Assert-Equal 'ng' $cmpA['code'] 'multi: compared as named, both pairs differ'
    $cmpB = Invoke-TestStep 'file.compare' @{ pairs = $pair2['pairs']; repairByContent = $true }
    Assert-Equal 'ok' $cmpB['code'] 'multi: repairByContent: each GIFT file has exactly one identical GFIX file -> re-paired, ok'
    Assert-Equal 2 $cmpB['repaired'] 'multi: ... two pairs re-paired'
    Assert-Equal 'F202610130002.csv' $cmpB['pairs'][0]['rightName'] 'multi: ALOC9002 now goes with the file that holds its content'
    Assert-Equal 'F202610130001.csv' $cmpB['pairs'][0]['repairedFrom'] 'multi: ... and remembers which file it was listed with'
    Assert-Equal 'ALOC9002.csv' $cmpB['pairs'][0]['leftName'] 'multi: the left side and every other field stay'
    Assert-Equal 12 $cmpB['pairs'][0]['lines'] 'multi: the line count stays (same on both sides by construction)'
    Assert-Equal 'repaired_by_content' $cmpB['warnings'][0]['code'] 'multi: the re-pairing is a warning in the run summary'
    Assert-True ($cmpB['reason'] -match 'ALOC9002.csv <- F202610130002.csv \(was F202610130001.csv\)') ('multi: the reason names the move: ' + $cmpB['reason'])
    Assert-Equal 2 $cmpB['identical'] 'multi: the second pass compared the new pairs'
    # not provable: the second GFIX file is damaged -> no re-pairing, ng as compared
    [System.IO.File]::WriteAllText((Join-Path $gfixDir2 'F202610130002.csv'), $mFirst.Replace('M3,21', 'M3,22'))
    $cmpC = Invoke-TestStep 'file.compare' @{ pairs = $pair2['pairs']; repairByContent = $true }
    Assert-Equal 'ng' $cmpC['code'] 'multi: one GIFT file matches no GFIX file -> no re-pairing, ng'
    Assert-Equal 0 $cmpC['repaired'] 'multi: ... nothing moved'
    Assert-Equal 'F202610130001.csv' $cmpC['pairs'][0]['rightName'] 'multi: ... pairs are the input pairs'
    # not provable either: two identical GIFT files (the same data set sent twice) -> a guess, refused
    [System.IO.File]::WriteAllText((Join-Path $gfixDir2 'F202610130002.csv'), $mSecond)
    $cmpD = Invoke-TestStep 'file.compare' @{ pairs = $pair2['pairs']; repairByContent = $true }
    Assert-Equal 'ng' $cmpD['code'] 'multi: ALOC9002 matches nothing, ALOC9004 matches both -> refused, ng'
    Assert-Equal 0 $cmpD['repaired'] 'multi: ... nothing moved'
    [System.IO.File]::WriteAllText((Join-Path $gfixDir2 'F202610130002.csv'), $mFirst)
    # a single pair never re-pairs (nothing to swap with)
    $cmpE = Invoke-TestStep 'file.compare' @{ pairs = @($pair2['pairs'][0]); repairByContent = $true }
    Assert-Equal 'ng' $cmpE['code'] 'multi: one differing pair alone stays ng'

    # the GoAnywhere rows of a two-file job: Send, Receive, Receive (and
    # whatever else ran in the window between them) -- the screenshot must
    # run from the first matched row to the last, not count matched rows
    $gaRows = @(@{ key = '1'; folder = '/jpcusr/Receive'; t = 'x' }, @{ key = '2'; folder = '/other'; t = 'x' }, @{ key = '3'; folder = '/jpcusr/Receive'; t = 'x' }, @{ key = '4'; folder = '/jpcusr/Send'; t = 'x' })
    $rv = Invoke-TestStep 'verify.filter_records' @{ records = $gaRows; where = @(@{ field = 'folder'; op = 'equals'; value = '/jpcusr/Receive' }); pluck = 'key' }
    Assert-Equal '2|1|3|3' ('' + $rv['matched'] + '|' + $rv['first'] + '|' + $rv['last'] + '|' + $rv['extent']) 'multi: two Receive rows with another job between: matched 2, rows 1..3, extent 3'
    Assert-Equal '2|1|2|2' ('' + $mine['matched'] + '|' + $mine['first'] + '|' + $mine['last'] + '|' + $mine['extent']) 'track: one job, two adjacent rows: extent = matched'
    Assert-Equal '0|0' ('' + $none['last'] + '|' + $none['extent']) 'track: nothing matched: last / extent 0'
    $rvF = Invoke-TestStep 'verify.filter_records' @{ records = $gaRows; where = @(@{ field = 'folder'; op = 'equals'; value = '/jpcusr/Receive' }); pluck = 'key'; expect = -1; expectFallback = '2' }
    Assert-True (-not $rvF.Contains('warnings') -or @($rvF['warnings']).Count -eq 0) 'multi: FileCount unknown, Teams says 2, 2 Receive rows: no warning'
    $rvG = Invoke-TestStep 'verify.filter_records' @{ records = $gaRows; where = @(@{ field = 'folder'; op = 'equals'; value = '/jpcusr/Receive' }); pluck = 'key'; expect = -1; expectFallback = '3' }
    Assert-Equal 'count_mismatch' $rvG['warnings'][0]['code'] 'multi: Teams says 3, 2 Receive rows: count_mismatch (from the fallback count)'
    Assert-True ($rvG['warnings'][0]['message'] -match 'fallback') 'multi: ... and the message says the count came from the fallback'
    $rvH = Invoke-TestStep 'verify.filter_records' @{ records = $gaRows; where = @(@{ field = 'folder'; op = 'equals'; value = '/jpcusr/Receive' }); pluck = 'key'; expect = 2; expectFallback = '3' }
    Assert-Equal 'expect_disagree' $rvH['warnings'][0]['code'] 'multi: mapping says 2, Teams says 3 -> expect_disagree, checked against 2'
    Assert-Equal 1 @($rvH['warnings']).Count 'multi: ... and 2 rows match the 2, so no count_mismatch on top'
    $rvI = Invoke-TestStep 'verify.filter_records' @{ records = $gaRows; where = @(@{ field = 'folder'; op = 'equals'; value = '/jpcusr/Receive' }); pluck = 'key'; expect = 2; expectFallback = 'abc' }
    Assert-True (-not $rvI.Contains('warnings') -or @($rvI['warnings']).Count -eq 0) 'multi: a non-numeric fallback (no Teams message) is ignored'

    # ============================================================ logs
    $job = Invoke-TestStep 'file.read_text' @{ paths = @((Join-Path $fx 'job-1000004619654.log')) }
    $fn = Invoke-TestStep 'verify.parse_text' @{ text = $job['text']; grammar = $prof['grammar']['jobLog'] }
    Assert-Equal 'JJPCRS1220260706110052987442' (@($fn['names']) -join ',') 'logs: the transfer file name from the job log'
    $raw = Join-Path $tmp 'log/GFIXReceive/1008.utf8.log'
    New-Item -ItemType Directory -Path (Split-Path $raw) -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $fx 'GFIXReceive.sample.log') -Destination $raw
    $conv = Invoke-TestStep 'file.convert_encoding' @{ path = $raw; saveAs = 'log/GFIXReceive/1008.log'; from = 'mixed'; to = 'cp932' }
    Assert-True ([bool]$conv['ok'] -and [int]$conv['foreignPairs'] -ge 2 -and [int]$conv['lost'] -eq 0) 'logs: mixed UTF-8 -> SJIS, the SJIS colons recovered, nothing lost'
    $back = (Get-EbiCodePage -CodePage 932).GetString([System.IO.File]::ReadAllBytes((Join-Path $tmp 'log/GFIXReceive/1008.log')))
    Assert-True ($back.Contains('update count' + [char]0xFF1A + '1')) 'logs: the SJIS file reads back with the full-width colon'
    $conv2 = Invoke-TestStep 'file.convert_encoding' @{ path = $raw; saveAs = 'log/GFIXReceive/1008.log'; from = 'mixed'; to = 'cp932' }
    Assert-True ([System.IO.File]::ReadAllText((Join-Path $tmp 'log/GFIXReceive/1008.log'), (Get-EbiCodePage -CodePage 932)) -eq $back) 'logs: converting again from the kept original gives the same file'

    # ============================================================ evidence
    $blk = Invoke-TestStep 'file.extract_blocks' @{ paths = @('log/GFIXReceive/1008.log', 'log/GFIXReceive/1008-unzip.log'); blockKeys = $fn['names']; encoding = 'cp932'; markers = $prof['layout']['evidence']['highlight']['receiveLog']; saveTo = 'capture/gfix/RJDSWM40/receive.txt' }
    Assert-True ([bool]$blk['ok']) 'evidence: the block was cut (the missing unzip log is skipped)'
    Assert-Equal 148 $blk['lineCount'] 'evidence: START..END + date line = 148 lines'
    Assert-Equal 2 @($blk['markers']).Count 'evidence: two marker lines (fileName, file stored)'
    Assert-True ($blk['markers'][1]['text'] -match 'F202610080006\.csv') 'evidence: the stored file is the Jenkins file found in track'
    $lines = @(Get-EbiTextLines -Text ([System.IO.File]::ReadAllText($blk['path'])))
    Assert-Equal ('GFIXReceiver START JJPCRS1220260706110052987442') $lines[0] 'evidence: the block starts with START'
    $endCol = Get-EbiHighlightEndColumn -Text $lines[$blk['markers'][1]['line'] - 1] -StartColumn 2
    Assert-True ($endCol -gt 60 -and $endCol -lt 100) ('evidence: the file stored highlight ends around column BZ (got ' + (ConvertTo-EbiColumnLetter $endCol) + ')')
    $jobHits = @(Find-EbiLineHits -Lines @(Get-EbiTextLines -Text $job['text']) -Patterns $prof['layout']['evidence']['highlight']['jobLog'])
    Assert-Equal 2 $jobHits.Count 'evidence: the job log has exactly the upload line and the Command line to highlight'
    $miss = Invoke-TestStep 'file.extract_blocks' @{ paths = @('log/GFIXReceive/1008.log'); blockKeys = @('JJPCRS12NOTHERE'); saveTo = 'capture/x.txt' }
    Assert-True (-not $miss['ok'] -and $miss['failure'] -eq 'not_found') 'evidence: a transfer with no block is not_found'

    $plan = Invoke-TestStep 'excel.stack_plan' @{ sets = @(@{ first = 'a.png'; last = '' }, @{ first = 'b1.png'; last = 'b2.png' }); separator = 'wave.png'; separatorColumn = 'Z'; markRects = $prof['layout']['df']['marks'] }
    Assert-Equal 'a.png|b1.png|wave.png|b2.png' (@($plan['pictures'] | ForEach-Object { $_['path'] }) -join '|') 'evidence: short file one shot; long file first, wave, last'
    Assert-Equal '2|0|0|2' (@($plan['pictures'] | ForEach-Object { @($_['rects']).Count }) -join '|') 'evidence: the boxes go on the picture showing the end'
    Assert-Equal 'Z' $plan['pictures'][2]['column'] 'evidence: the wave sits in column Z'
    Assert-Equal 1 $plan['pictures'][1]['gapRows'] 'evidence: one blank row between files'

    $names = @($jk['names'])
    $bands = @(for ($i = 0; $i -lt 30; $i++) { @{ top = 100 + $i * 20; bottom = 112 + $i * 20; height = 13; center = 106 + $i * 20 } })
    $lr = Invoke-TestStep 'screen.list_rects' @{ names = $names; targets = @('F202610080005.csv', 'F202610080006.csv'); bands = $bands; x = 445; width = 892; height = 25 }
    Assert-Equal 1 @($lr['rects']).Count 'evidence: two consecutive files -> one box'
    $k = [array]::IndexOf($names, 'F202610080005.csv')
    $expectTop = [int][Math]::Round((106 + (30 - ($names.Count - $k)) * 20) - 12.5, [System.MidpointRounding]::AwayFromZero)
    Assert-Equal $expectTop $lr['rects'][0]['y'] 'evidence: the box sits on the band counted from the bottom'
    Assert-Equal 45 $lr['rects'][0]['h'] 'evidence: two rows tall'
    $lr2 = Invoke-TestStep 'screen.list_rects' @{ names = $names; targets = @('F202610070002.csv', 'F202610080006.csv'); x = 445; width = 892; height = 25; lastRowCenterY = 718; rowPitch = 20 }
    Assert-Equal 'pitch' $lr2['source'] 'evidence: without bands the fixed pitch is used'
    Assert-Equal 2 @($lr2['rects']).Count 'evidence: two files apart -> two boxes'

    # ============================================================ plan: upsert
    $wl = @{ path = ''; columns = @('Excel_NAME', 'JOB', 'GFIX_DATE', 'GFIX_TIME', 'FileCount', 'track', 'logs', 'evidence'); rows = @() }
    $ctx['Session']['wl'] = @{ kind = 'worklist'; value = $wl; registeredBy = 'test' }
    $flds = @{ Excel_NAME = 'Excel_NAME'; JOB = 'JOB'; GFIX_DATE = 'GFIX_DATE'; GFIX_TIME = 'GFIX_TIME' }
    $reset = @{ watch = 'GFIX_DATE'; clear = @('track', 'logs', 'evidence') }
    $rows1 = @(@{ Excel_NAME = 'JJDSWM51'; JOB = 'JJDSJM51'; GFIX_DATE = '2026-10-08'; GFIX_TIME = '10:00:00' }, @{ Excel_NAME = 'JJDSWM51'; JOB = 'JJDSJM51'; GFIX_DATE = '2026-10-08'; GFIX_TIME = '10:00:00' })
    $u1 = Invoke-TestStep 'table.upsert' @{ worklist = 'wl'; rows = $rows1; fields = $flds; countAs = 'FileCount'; overwrite = @('GFIX_DATE', 'GFIX_TIME'); resetOnChange = $reset }
    Assert-Equal '1|2' ('' + $u1['added'] + '|' + $wl['rows'][0]['FileCount']) 'plan: two mapping rows of one job -> one worklist row, FileCount 2'
    $wl['rows'][0]['track'] = 'ok'
    $u2 = Invoke-TestStep 'table.upsert' @{ worklist = 'wl'; rows = $rows1; fields = $flds; countAs = 'FileCount'; overwrite = @('GFIX_DATE', 'GFIX_TIME'); resetOnChange = $reset }
    Assert-Equal 'ok' $wl['rows'][0]['track'] 'plan: a rerun the same day keeps the progress'
    $rows2 = @(@{ Excel_NAME = 'JJDSWM51'; JOB = 'JJDSJM51'; GFIX_DATE = '2026-10-09'; GFIX_TIME = '11:00:00' })
    $u3 = Invoke-TestStep 'table.upsert' @{ worklist = 'wl'; rows = $rows2; fields = $flds; countAs = 'FileCount'; overwrite = @('GFIX_DATE', 'GFIX_TIME'); resetOnChange = $reset }
    Assert-Equal '|2026-10-09|11:00:00|1' ('' + $wl['rows'][0]['track'] + '|' + $wl['rows'][0]['GFIX_DATE'] + '|' + $wl['rows'][0]['GFIX_TIME'] + '|' + $wl['rows'][0]['FileCount']) 'plan: planned again on another day -> progress cleared, new date / time'

    # addNew false (10/14): with operator.owner set the plan only refreshes the jobs the WBS gave
    # this operator -- mapping.xlsx lists everyone's, and the other owner's rows used to be added
    $rows3 = @(@{ Excel_NAME = 'JJDSWM51'; JOB = 'JJDSJM51'; GFIX_DATE = '2026-10-09'; GFIX_TIME = '11:30:00' }, @{ Excel_NAME = 'JJDSWMAJ'; JOB = 'JJDSJMAJ'; GFIX_DATE = '2026-10-09'; GFIX_TIME = '10:30:00' })
    $u4 = Invoke-TestStep 'table.upsert' @{ worklist = 'wl'; rows = $rows3; fields = $flds; countAs = 'FileCount'; overwrite = @('GFIX_DATE', 'GFIX_TIME'); addNew = $false }
    Assert-Equal '0|1|JJDSWMAJ|1|11:30:00' ('' + $u4['added'] + '|' + $u4['updated'] + '|' + (@($u4['skippedNew']) -join ',') + '|' + @($wl['rows']).Count + '|' + $wl['rows'][0]['GFIX_TIME']) 'plan: addNew false -> the existing row refreshed, the unknown job not added'
    $u5 = Invoke-TestStep 'table.upsert' @{ worklist = 'wl'; rows = $rows3; fields = $flds; countAs = 'FileCount'; overwrite = @('GFIX_DATE', 'GFIX_TIME') }
    Assert-Equal '1|2' ('' + $u5['added'] + '|' + @($wl['rows']).Count) 'plan: addNew default (true) adds it as before'

    # ---- owner (kernel/Worklist.ps1): the same arrow rule the plan's WBS filter uses
    $la = [string][char]0x2190; $ra = [string][char]0x2192
    $op = Get-EbiOwnerPattern -Owner 'yan'
    Assert-Equal '' (Get-EbiOwnerPattern -Owner '  ') 'owner: blank owner -> no pattern'
    Assert-True (('yan' -match $op) -and (('yan' + $la + 'nin') -match $op) -and (('nin' + $ra + 'yan') -match $op)) 'owner: bare name, name<-other, other->name are mine'
    Assert-True (-not (('yan' + $ra + 'nin') -match $op) -and -not ('nin' -match $op)) 'owner: name->other and another name are not'
    $orows = @(@{ Excel_NAME = 'A'; Owner = 'yan' }, @{ Excel_NAME = 'B'; Owner = 'nin' }, @{ Excel_NAME = 'C'; Owner = '' }, @{ Excel_NAME = 'D'; Owner = ('nin' + $ra + 'yan') })
    $own = Select-EbiOwnedRows -Rows $orows -Column 'Owner' -Pattern $op -KeyColumns @('Excel_NAME')
    Assert-Equal 'A,C,D|B (nin)' ((@($own['rows'] | ForEach-Object { $_['Excel_NAME'] }) -join ',') + '|' + (@($own['leftOut']) -join ',')) 'owner: someone else''s row left out, a blank owner kept'
    Assert-Equal 'Owner' (Get-EbiWorklistOwnerColumn -Profile $prof) 'owner: the gfix-recv worklist declares its owner column'

    # ============================================================ kernel/Layout.ps1
    $bs = @(Get-EbiInkBands -Counts @(0, 0, 3, 4, 0, 5, 0, 0, 0, 2) -MinInk 1 -MergeGap 1)
    Assert-Equal '2-5|9-9' ((@($bs | ForEach-Object { '' + $_['top'] + '-' + $_['bottom'] })) -join '|') 'layout: a one-row gap merges, a three-row gap splits'
    $a = Resolve-EbiAnchoredRect -Rect @{ x = 117; y = 19; w = 112; h = 19; anchor = 'br' } -ImageWidth 1133 -ImageHeight 429
    Assert-Equal '1016,410' ('' + $a['x'] + ',' + $a['y']) 'layout: a bottom-right anchored box lands on the status bar cell'
    $sr = ConvertTo-EbiSheetRect -PicLeft 10 -PicTop 100 -PicWidth 849.75 -PicHeight 321.75 -ImageWidth 1133 -ImageHeight 429 -Rect @{ x = 100; y = 40; w = 48; h = 18 }
    Assert-Equal '85,130,36,13.5' ('' + $sr['left'] + ',' + $sr['top'] + ',' + $sr['width'] + ',' + $sr['height']) 'layout: image px -> sheet points through the picture scale'

    # ============================================================ human.share front (10/19)
    $wins = @(@{ handle = 1; title = 'Inbox - Outlook'; processName = 'OUTLOOK' },
              @{ handle = 2; title = 'J4 chat | Microsoft Teams'; processName = 'ms-teams' },
              @{ handle = 3; title = 'Teams (classic)'; processName = 'Teams' })
    Assert-Equal 2 (HumanShare-PickWindow -Windows $wins -Pattern 'ms-teams|^Teams$|Microsoft Teams')['handle'] 'share front: the frontmost Teams window is picked'
    Assert-Equal 3 (HumanShare-PickWindow -Windows $wins -Pattern '^Teams$')['handle'] 'share front: process name matches too'
    Assert-True ($null -eq (HumanShare-PickWindow -Windows $wins -Pattern 'Slack')) 'share front: no match -> null'
    Assert-True ($null -eq (HumanShare-PickWindow -Windows $wins -Pattern '')) 'share front: empty pattern -> null (console stays in front)'

    # ============================================================ kernel/RichClip.ps1
    $png = [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAIAAAADCAYAAAC56t6BAAAAEUlEQVR42mP8z8Dwn4EIwDiqEAD8xwMBmkAeOAAAAABJRU5ErkJggg==')
    $pi = Get-EbiPngInfo -Bytes $png
    Assert-Equal '2x3' ('' + $pi['width'] + 'x' + $pi['height']) 'clip: PNG size from IHDR'
    Assert-True (-not (Get-EbiPngInfo -Bytes ([byte[]](1, 2, 3)))['ok']) 'clip: not a PNG -> not ok'
    $msg = [string][char]0x524D + [char]0x5F8C + [char]0x4E00 + [char]0x81F4   # zen-go-itchi
    $frag = New-EbiShareHtml -Lines @($msg, 'a<b') -Pictures @(@{ bytes = $png; width = 2; height = 3 })
    Assert-True ($frag.Contains('a&lt;b') -and $frag.Contains('data:image/png;base64,')) 'clip: text escaped, picture inline'
    $cf = New-EbiCfHtml -Fragment $frag
    $u8 = New-Object System.Text.UTF8Encoding($false)
    $b = $u8.GetBytes($cf)
    $sf = [int]([regex]::Match($cf, 'StartFragment:(\d+)').Groups[1].Value); $ef = [int]([regex]::Match($cf, 'EndFragment:(\d+)').Groups[1].Value)
    Assert-Equal $frag ($u8.GetString($b, $sf, $ef - $sf)) 'clip: CF_HTML byte offsets cut out exactly the fragment (Japanese included)'
    $eh = [int]([regex]::Match($cf, 'EndHTML:(\d+)').Groups[1].Value)
    Assert-Equal $b.Length $eh 'clip: EndHTML is the byte length'
    $rtf = New-EbiShareRtf -Lines @($msg) -Pictures @(@{ bytes = $png; width = 2; height = 3 })
    Assert-True ($rtf.Contains(('\u' + '21069?')) -and $rtf.Contains('\pngblip') -and $rtf.Contains('\picwgoal30')) 'clip: RTF text as \uN?, picture as pngblip in twips'
    Assert-Equal '\\\{x\}' (ConvertTo-EbiRtfText -Text '\{x}') 'clip: RTF escapes'

    Write-Host '  -- plan from WBS alone + the start time from the pasted Teams message'
    $wbsRecs = @(@{ job = 'SJDSJM40'; start = '2026-10-09' }, @{ job = 'JJMRJE6C'; start = '2026-10-09' }, @{ job = 'ODD'; start = '2026-10-09' })
    $dv = Invoke-TestStep 'verify.derive_fields' @{ records = $wbsRecs; set = @(@{ to = 'Excel_NAME'; from = 'job'; pattern = '^(.{4})J'; replace = '${1}W' }, @{ to = 'JOB'; from = 'job' }, @{ to = 'GFIX_DATE'; value = '2026-10-09' }) }
    Assert-Equal 'SJDSWM40|JJMRWE6C|ODD' ((@($dv['records']) | ForEach-Object { $_['Excel_NAME'] }) -join '|') 'plan: W name = J name with the 5th character J -> W (the leader table names SJDSWM40, JJMRWE6C)'
    Assert-Equal 'SJDSJM40' $dv['records'][0]['JOB'] 'plan: JOB keeps the J name (DATA/GIFT/<J>)'
    Assert-Equal '2026-10-09' $dv['records'][1]['GFIX_DATE'] 'plan: a constant on every row'
    Assert-True (@($dv['unchanged']).Count -eq 1 -and @($dv['warnings']).Count -eq 1) 'plan: a name that does not follow the rule is copied as is AND warned about'
    $bad = Invoke-TestStep 'verify.derive_fields' @{ records = $wbsRecs; set = @(@{ to = 'X' }) }
    Assert-True ((-not $bad['ok']) -and $bad['failure'] -eq 'input_invalid') 'plan: a rule with neither value nor from is refused'

    $pat = $prof['vocabulary']['startPattern']
    $jobW = 'JJMRWE6K'
    $kana = -join @([char]0x30B8, [char]0x30E7, [char]0x30D6)          # job (katakana)
    $wo = [string][char]0x3092; $jisshi = -join @([char]0x5B9F, [char]0x65BD, [char]0x3057, [char]0x307E, [char]0x3059, [char]0x3002)
    $yotei = -join @([char]0x9001, [char]0x4FE1, [char]0x4E88, [char]0x5B9A); $ken = [string][char]0x4EF6
    $msgOk = $kana + ':' + $jobW + $wo + $jisshi + '(' + $yotei + ':6' + $ken + ')'
    $c = HumanPaste-Check -Text $msgOk -Pattern $pat -Expect $jobW -ExpectGroup 'job'
    Assert-True ($c['ok'] -and $c['fields']['job'] -eq $jobW -and $c['fields']['counts'] -eq ('6' + $ken)) ('paste: the Teams start message gives job + send count: ' + $c['reason'])
    $msgFw = $kana + [char]0xFF1A + (-join ($jobW.ToCharArray() | ForEach-Object { [char]([int]$_ + 0xFEE0) })) + $wo + $jisshi
    $c = HumanPaste-Check -Text $msgFw -Pattern $pat -Expect $jobW -ExpectGroup 'job'
    Assert-True ($c['ok'] -and $c['fields']['counts'] -eq '') 'paste: full-width colon / letters fold; no send count -> counts is "" (still templatable)'
    # 10-13 JJDSWM10: two files -> two counts in one message
    $ten = [string][char]0x3001
    $msg2 = $kana + ':JJDSWM10' + $wo + $jisshi + '(' + $yotei + ':815' + $ken + $ten + '46' + $ken + ')' + "`r`n2 more lines"
    $c = HumanPaste-Check -Text $msg2 -Pattern $pat -Expect 'JJDSWM10' -ExpectGroup 'job'
    Assert-True ($c['ok'] -and $c['fields']['counts'] -eq ('815' + $ken + $ten + '46' + $ken)) ('paste: a two-file message keeps the whole count list: ' + $c['fields']['counts'])

    # ---- verify.record_counts: the list -> file count; files -> the reply text
    $rc = Invoke-TestStep 'verify.record_counts' @{ text = $c['fields']['counts']; unit = $ken; joiner = $ten; label = 'recv:' }
    Assert-Equal '815,46|2' ((@($rc['expected']) -join ',') + '|' + $rc['fileCount']) 'counts: 815, 46 -> two files'
    Assert-Equal ('815' + $ken + $ten + '46' + $ken) $rc['text'] 'counts: without files the text is the message''s counts'
    Assert-Equal (' (recv:815' + $ken + $ten + '46' + $ken + ')') $rc['suffix'] 'counts: the reply suffix'
    $rc0 = Invoke-TestStep 'verify.record_counts' @{ text = ''; unit = $ken }
    Assert-Equal '-1|' ('' + $rc0['fileCount'] + '|' + $rc0['suffix']) 'counts: no message -> fileCount -1, no suffix'
    $filesOk = @(@{ name = 'F1.csv'; lines = 816 }, @{ name = 'F2.csv'; lines = 46 })
    $rc1 = Invoke-TestStep 'verify.record_counts' @{ text = $c['fields']['counts']; files = $filesOk; unit = $ken; joiner = $ten; label = 'recv:' }
    Assert-True ([bool]$rc1['matched'] -and $rc1['text'] -eq ('815' + $ken + $ten + '46' + $ken)) 'counts: a header line more (816 vs 815) still agrees; text shows the message counts'
    Assert-True (-not $rc1.Contains('warnings') -or @($rc1['warnings']).Count -eq 0) 'counts: ... no warning'
    $filesNg = @(@{ name = 'F1.csv'; lines = 0 }, @{ name = 'F2.csv'; lines = 1 })
    $rc2 = Invoke-TestStep 'verify.record_counts' @{ text = ('2' + $ken + $ten + '4' + $ken); files = $filesNg; unit = $ken; joiner = $ten; label = 'recv:' }
    Assert-Equal ('0' + $ken + '(2' + $ken + ')' + $ten + '1' + $ken + '(4' + $ken + ')') $rc2['text'] 'counts: a file that disagrees shows lines(message count) -- the operator''s example'
    Assert-True ((-not [bool]$rc2['matched']) -and $rc2['warnings'][0]['code'] -eq 'count_disagree') 'counts: ... and warns count_disagree'
    $rc2s = Invoke-TestStep 'verify.record_counts' @{ text = ('1' + $ken + $ten + '127' + $ken + $ten + '50' + $ken); files = @(@{ name = 'R.csv'; lines = 0 }, @{ name = 'S.csv'; lines = 126 }, @{ name = 'T.csv'; lines = 47 }); unit = $ken; joiner = $ten; shortOk = $true }
    Assert-Equal ('0' + $ken + '(1' + $ken + ')' + $ten + '126' + $ken + '(127' + $ken + ')' + $ten + '47' + $ken + '(50' + $ken + ')') $rc2s['text'] 'counts shortOk (10/15 R/S): short files still show lines(count) in the reply'
    Assert-True ([bool]$rc2s['matched'] -and (-not $rc2s.Contains('warnings') -or @($rc2s['warnings']).Count -eq 0)) 'counts shortOk: 1, 1 and 3 lines short are expected -> matched, no count_disagree'
    Assert-Equal '1,1,3' (@($rc2s['short']) -join ',') 'counts shortOk: short says by how many lines'
    $rc2o = Invoke-TestStep 'verify.record_counts' @{ text = ('10' + $ken); files = @(@{ name = 'U.csv'; lines = 13 }); unit = $ken; shortOk = $true }
    Assert-True (-not [bool]$rc2o['matched'] -and $rc2o['warnings'][0]['code'] -eq 'count_disagree') 'counts shortOk: MORE lines than count + header is still a disagreement'
    $rcN = Invoke-TestStep 'verify.record_counts' @{ text = ''; files = @(@{ name = 'F.csv'; lines = 61 }); unit = $ken }
    Assert-True ($rcN['ok'] -and $rcN['matched'] -and -not $rcN.Contains('warnings')) 'record_counts: no start message (row time used) -> nothing to compare, no count_disagree warning'
    $rc3 = Invoke-TestStep 'verify.record_counts' @{ text = ('0' + $ken); files = @(@{ name = 'F.csv'; lines = 1 }); unit = $ken }
    Assert-True ([bool]$rc3['matched']) 'counts: an empty transfer may be one empty line'
    $rc4 = Invoke-TestStep 'verify.record_counts' @{ text = ('5' + $ken); files = $filesOk; unit = $ken; joiner = $ten }
    Assert-True ((-not [bool]$rc4['matched']) -and $rc4['text'] -like ('*' + $ken + '(?)')) 'counts: more files than counts -> the extra file shows (?)'

    # ---- verify.nearest_time: the typed time was off, rows are on the page
    $keep = Invoke-TestStep 'verify.nearest_time' @{ found = $true; date = '2026-10-08'; clock = '13:20:00'; records = $rec['records'] }
    Assert-True ($keep['clock'] -eq '13:20:00' -and -not [bool]$keep['changed'] -and $keep['code'] -eq 'ok') 'near: the window matched -> the clock is kept'
    $nr = Invoke-TestStep 'verify.nearest_time' @{ found = $false; date = '2026-10-08'; clock = '13:40:00'; records = $rec['records']; withinMinutes = 7 }
    Assert-True ([bool]$nr['changed'] -and $nr['code'] -eq 'unknown') 'near: not found -> the nearest row is proposed as unknown (a gate asks)'
    Assert-True ($nr['clock'] -like '13:45:*') ('near: ... the 13:45 job is the nearest to 13:40: ' + $nr['clock'])
    Assert-True ([int]$nr['offsetMinutes'] -ge 5 -and [int]$nr['offsetMinutes'] -le 6 -and [bool]$nr['near']) 'near: ... +5 min, within the near range'
    Assert-True ($nr['reason'] -like '*13:45:*' -and $nr['reason'] -like '*Enter = go on with*') 'near: the reason tells the operator what Enter does'
    $far = Invoke-TestStep 'verify.nearest_time' @{ found = $false; date = '2026-10-08'; clock = '10:45:20'; records = $rec['records']; withinMinutes = 5 }
    Assert-True ([bool]$far['changed'] -and -not [bool]$far['near'] -and $far['reason'] -like '*farther than 5 min*') 'near: a row far away is still proposed, flagged as far'
    $nothing = Invoke-TestStep 'verify.nearest_time' @{ found = $false; date = '2026-10-08'; clock = '10:45:20'; records = @() }
    Assert-True ((-not [bool]$nothing['changed']) -and $nothing['clock'] -eq '10:45:20') 'near: no rows at all -> the clock is kept'

    # ---- a FAILED job on the OPEN side (10-13 JJDSWM41): the status icon is not in
    # the page text; the operator adds the view's resolved column, which reads No
    # on the failed row only
    $rsText = [System.IO.File]::ReadAllText((Join-Path $fx 'goanywhere-list-resolved.sample.txt'))
    $rs = Invoke-TestStep 'verify.parse_text' @{ text = $rsText; grammar = $prof['grammar']['goAnywhere'] }
    Assert-Equal 6 $rs['recordCount'] 'resolved: the page with the extra column parses (6 rows)'
    Assert-Equal 'No' (@($rs['records'] | Where-Object { $_['key'] -eq '1000004712517' })[0]['resolved']) 'resolved: the failed row reads No'
    Assert-Equal '' (@($rs['records'] | Where-Object { $_['key'] -eq '1000004712355' })[0]['resolved']) 'resolved: a good row is blank'
    Assert-Equal 12 $rec['recordCount'] 'resolved: the OLD page without the column still parses every row (optionalTail)'
    Assert-Equal '' $rec['records'][0]['resolved'] 'resolved: ... with resolved blank'
    $w45 = Invoke-TestStep 'verify.time_window' @{ date = '2026-10-13'; clock = '10:45:20'; beforeMinutes = 2; afterMinutes = 13 }
    $m45 = Invoke-TestStep 'verify.filter_records' @{ records = $rs['records']; where = @(@{ field = 'startTime'; op = 'within'; value = $w45['window'] }) }
    $hs = Invoke-TestStep 'verify.assert_each' @{ records = $m45['records']; rules = $prof['rules']['goAnywhereJob'] }
    Assert-Equal 'ng' $hs['code'] 'failed job: the 10:45 job has a Send row with resolved = No -> ng'
    Assert-True ($hs['reason'] -like '1000004712517:*' -and $hs['failedCount'] -eq 1 -and $hs['judged'] -eq 1) ('failed job: the reason names the row: ' + $hs['reason'])
    $w11 = Invoke-TestStep 'verify.time_window' @{ date = '2026-10-13'; clock = '11:00:27'; beforeMinutes = 2; afterMinutes = 13 }
    $m11 = Invoke-TestStep 'verify.filter_records' @{ records = $rs['records']; where = @(@{ field = 'startTime'; op = 'within'; value = $w11['window'] }) }
    $hs2 = Invoke-TestStep 'verify.assert_each' @{ records = $m11['records']; rules = $prof['rules']['goAnywhereJob'] }
    Assert-True ($hs2['code'] -eq 'ok' -and $hs2['judged'] -eq 4 -and $hs2['failedCount'] -eq 0 -and $hs2['reason'] -eq '') 'good job: 2 Send + 2 Receive rows, none unresolved -> ok'
    # ---- 10-14: the view had NO user column (jikkou-user) -- by position every
    # column after status moved left, endTime read the seconds, the Receive
    # span came out empty and the Jenkins filter let 2183 files through. The
    # grammar now reads the columns off the page's own header line.
    $nuText = [System.IO.File]::ReadAllText((Join-Path $fx 'goanywhere-list-nouser.sample.txt'))
    $nu = Invoke-TestStep 'verify.parse_text' @{ text = $nuText; grammar = $prof['grammar']['goAnywhere'] }
    Assert-Equal 8 $nu['recordCount'] 'no user column: 8 rows'
    $nu0 = @($nu['records'] | Where-Object { $_['key'] -eq '1000004734879' })[0]
    Assert-Equal '2026-10-14 11:00:29|2026-10-14 11:00:32|/jpcusr/Receive|' ($nu0['startTime'] + '|' + $nu0['endTime'] + '|' + $nu0['folder'] + '|' + $nu0['user']) 'no user column: start / end / folder read by header, user blank'
    $nuW = Invoke-TestStep 'verify.time_window' @{ date = '2026-10-14'; clock = '11:00:00'; beforeMinutes = 2; afterMinutes = 13 }
    $nuMine = Invoke-TestStep 'verify.filter_records' @{ records = $nu['records']; where = @(@{ field = 'startTime'; op = 'within'; value = $nuW['window'] }) }
    $nuRecv = Invoke-TestStep 'verify.filter_records' @{ records = $nuMine['records']; where = @(@{ field = 'folder'; op = 'equals'; value = '/jpcusr/Receive' }); pluck = 'key'; spanFrom = 'startTime'; spanTo = 'endTime'; padBeforeSec = 2; padAfterSec = 10 }
    Assert-Equal '2026-10-14T11:00:27|2026-10-14T11:00:42' ($nuRecv['span']['from'] + '|' + $nuRecv['span']['to']) 'no user column: the Receive span is the real 11:00:29..11:00:32 (+pads)'
    Assert-Equal 6 $rs['recordCount'] 'header with a leading empty cell (10-13 page) still maps every column'
    Assert-Equal '2026-10-08 13:45:24' $rec['records'][0]['startTime'] 'the 10-08 page (user column, no resolved) still reads by header'
    $hdrBad = Invoke-TestStep 'verify.parse_text' @{ text = ($nuText.Replace("`t" + (-join [char[]](0x958B, 0x59CB, 0x6642, 0x9593)), "`tX")); grammar = $prof['grammar']['goAnywhere'] }
    Assert-True ((-not $hdrBad['ok']) -and $hdrBad['failure'] -eq 'grammar_invalid' -and $hdrBad['message'] -like '*startTime*') ('a header without the start column stops the parse: ' + $hdrBad['message'])
    # the guards behind it: an empty window and an unreadable span are stops, not pass-throughs
    $wm = Invoke-TestStep 'verify.filter_records' @{ records = $nu['records']; where = @(@{ field = 'startTime'; op = 'within'; value = @{} }) }
    Assert-True ((-not $wm['ok']) -and $wm['failure'] -eq 'window_missing') 'filter: within with an empty window -> window_missing'
    $su = Invoke-TestStep 'verify.filter_records' @{ records = @(@{ key = 'a'; startTime = '2026-10-14 11:00:32'; endTime = '3.52' }); spanFrom = 'startTime'; spanTo = 'endTime' }
    Assert-True ((-not $su['ok']) -and $su['failure'] -eq 'span_unreadable' -and $su['message'] -like '*3.52*') ('filter: an end time that is not a time -> span_unreadable: ' + $su['message'])
    $tm = Invoke-TestStep 'file.download' @{ baseUrl = 'http://127.0.0.1:9/'; names = @('a', 'b', 'c'); dir = $tmp; maxFiles = 2 }
    Assert-True ((-not $tm['ok']) -and $tm['failure'] -eq 'too_many' -and $tm['total'] -eq 0) 'download: more names than maxFiles -> too_many, nothing fetched'
    $hs3 = Invoke-TestStep 'verify.assert_each' @{ records = @(); rules = $prof['rules']['goAnywhereJob'] }
    Assert-Equal 'ok' $hs3['code'] 'assert_each: no records -> ok'
    $hsMix = Invoke-TestStep 'verify.assert_each' @{ records = @(@{ key = 'a'; resolved = '' }, @{ key = 'b'; resolved = 'No' }, @{ key = 'c'; resolved = 'No' }); rules = $prof['rules']['goAnywhereJob'] }
    Assert-True ($hsMix['code'] -eq 'ng' -and $hsMix['failedCount'] -eq 2 -and $hsMix['reason'] -like 'b:*(+1 more)') 'assert_each: several failing rows -> the first named, the rest counted'
    $hsBad = Invoke-TestStep 'verify.assert_each' @{ records = @(); rules = @{ rules = 'x' } }
    Assert-True ((-not $hsBad['ok']) -and $hsBad['failure'] -eq 'rules_invalid') 'assert_each: a broken rule table is rules_invalid'
    $rq = Invoke-TestStep 'flow.require' @{ value = 'ok'; equals = 'ok'; message = 'm' }
    Assert-True ([bool]$rq['ok'] -and [bool]$rq['met']) 'require: ok passes'
    $rq2 = Invoke-TestStep 'flow.require' @{ value = 'ng'; equals = 'ok'; message = 'the track judged this job ng' }
    Assert-True ((-not $rq2['ok']) -and $rq2['failure'] -eq 'not_met' -and $rq2['message'] -like 'the track judged this job ng (value: "ng"*') 'require: ng is not_met with the reason and the value'
    $rq3 = Invoke-TestStep 'flow.require' @{ value = ''; equals = 'ok' }
    Assert-True ((-not $rq3['ok']) -and $rq3['failure'] -eq 'not_met') 'require: an empty value (no track.json field) is not_met too'

    # ---- verify.owner_pattern: the WBS owner cell (name, name<-other, other->name)
    $gen = [string][char]0x53B3                     # the operator's name as the WBS writes it
    $other = -join @([char]0x30CB, [char]0x30F3, [char]0x30CC)
    $la = [string][char]0x2190; $ra = [string][char]0x2192
    $op = Invoke-TestStep 'verify.owner_pattern' @{ owner = $gen }
    Assert-True (-not [bool]$op['everyone'] -and $op['pattern'] -ne '') 'owner: a name gives a pattern'
    foreach ($cell in @($gen, (' ' + $gen + ' '), ($gen + $la + $other), ($other + $ra + $gen))) { Assert-True ([regex]::IsMatch($cell, $op['pattern'])) ('owner: mine: ' + $cell) }
    foreach ($cell in @($other, ($other + $la + $gen), ($gen + $ra + $other), '')) { Assert-True (-not [regex]::IsMatch($cell, $op['pattern'])) ('owner: not mine: ' + $cell) }
    $opAll = Invoke-TestStep 'verify.owner_pattern' @{ owner = '  ' }
    Assert-True ([bool]$opAll['everyone'] -and $opAll['pattern'] -eq '' -and [regex]::IsMatch($other, $opAll['pattern'])) 'owner: blank -> everyone (the empty pattern matches every cell)'
    $opRx = Invoke-TestStep 'verify.owner_pattern' @{ owner = 'a.b' }
    Assert-True ((-not [regex]::IsMatch('axb', $opRx['pattern'])) -and [regex]::IsMatch('a.b', $opRx['pattern'])) 'owner: regex characters in a name are escaped'

    # ---- excel.open: who holds the workbook (Excel's ~$ owner file)
    $who = 'JP222093'
    $lockBytes = New-Object byte[] (55 + 2 * $who.Length + 8)
    $lockBytes[0] = [byte]$who.Length
    [System.Text.Encoding]::ASCII.GetBytes($who).CopyTo($lockBytes, 1)
    $lockBytes[54] = [byte]$who.Length
    [System.Text.Encoding]::Unicode.GetBytes($who).CopyTo($lockBytes, 55)
    Assert-Equal $who (ExcelOpen-LockOwner -Bytes $lockBytes) 'lock: the UTF-16 name from the ~$ file'
    $short = New-Object byte[] 10; $short[0] = 3; [System.Text.Encoding]::ASCII.GetBytes('abc').CopyTo($short, 1)
    Assert-Equal 'abc' (ExcelOpen-LockOwner -Bytes $short) 'lock: a short file falls back to the ANSI name'
    Assert-Equal '' (ExcelOpen-LockOwner -Bytes (New-Object byte[] 0)) 'lock: empty -> no name'
    $lockDir = Join-Path $tmp 'lock'; New-Item -ItemType Directory -Path $lockDir -Force | Out-Null
    [System.IO.File]::WriteAllBytes((Join-Path $lockDir ('~$' + 'J4_X.xlsx')), $lockBytes)
    Assert-Equal $who (ExcelOpen-HeldBy -Path (Join-Path $lockDir 'J4_X.xlsx')) 'lock: found next to the workbook'
    Assert-Equal '' (ExcelOpen-HeldBy -Path (Join-Path $lockDir 'J4_Y.xlsx')) 'lock: no owner file -> empty, no throw'

    # ---- file.extract_blocks: two transfers logged at the same time (10-13 JJDSWM10)
    $il = @('GFIXReceiver START KEYA', 'a1', 'GFIXReceiver START KEYB', 'b1', 'a2', 'GFIXReceiver END KEYA', '2026-10-13 10:30:30', 'b2', 'GFIXReceiver END KEYB', '2026-10-13 10:30:31', 'GFIXReceiver START KEYC', 'c1', 'GFIXReceiver END KEYC', '2026-10-13 10:40:00')
    $ilPath = Join-Path $tmp 'log/il.log'
    [System.IO.File]::WriteAllText($ilPath, (($il -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
    $ib = Invoke-TestStep 'file.extract_blocks' @{ paths = @($ilPath); blockKeys = @('KEYA', 'KEYB'); encoding = 'utf8'; saveTo = 'capture/il.txt' }
    Assert-True ([bool]$ib['ok']) 'interleaved: ok'
    Assert-Equal 'True,True' (@($ib['blocks'] | ForEach-Object { [string]$_['complete'] }) -join ',') 'interleaved: both blocks reach their own END (no block_incomplete)'
    Assert-Equal 10 $ib['lineCount'] 'interleaved: START A .. END B + date written ONCE (10 lines), not A cut at START B'
    $ilOut = @(Get-EbiTextLines -Text ([System.IO.File]::ReadAllText($ib['path'])))
    Assert-Equal 'GFIXReceiver START KEYA' $ilOut[0] 'interleaved: begins at START A'
    Assert-Equal '2026-10-13 10:30:31' $ilOut[9] 'interleaved: ends with the date line after END B'
    Assert-Equal 1 @($ib['warnings'] | Where-Object { $_['code'] -eq 'block_interleaved' }).Count 'interleaved: one block_interleaved warning names the pair'
    Assert-Equal 0 @($ib['warnings'] | Where-Object { $_['code'] -eq 'block_incomplete' }).Count 'interleaved: ... and no block_incomplete'
    $ic = Invoke-TestStep 'file.extract_blocks' @{ paths = @($ilPath); blockKeys = @('KEYC'); encoding = 'utf8'; saveTo = 'capture/il2.txt' }
    Assert-Equal 4 $ic['lineCount'] 'interleaved: a block after the pair is cut on its own as before'
    $c = HumanPaste-Check -Text $msgOk -Pattern $pat -Expect 'SJDSWM40' -ExpectGroup 'job'
    Assert-True ((-not $c['ok']) -and $c['reason'].Contains('JJMRWE6K') -and $c['reason'].Contains('SJDSWM40')) 'paste: a message for another job is refused and says both names'
    $c = HumanPaste-Check -Text 'hello' -Pattern $pat -Expect '' -ExpectGroup 'job'
    Assert-True ((-not $c['ok']) -and $c['fields'].Contains('counts') -and $c['fields'].Contains('job')) 'paste: not the message -> refused, every group still present as ""'
    Assert-Equal 'time|10:30:00' ((HumanPaste-Typed -Answer '10:30')['kind'] + '|' + (HumanPaste-Typed -Answer '10:30')['clock']) 'paste: a typed time replaces now'
    Assert-Equal '09:05:00' (HumanPaste-Typed -Answer '905')['clock'] 'paste: 905 -> 09:05:00'
    $kinds = @(foreach ($a in @('', 'n', 'k', 's', 'zz')) { (HumanPaste-Typed -Answer $a)['kind'] })
    Assert-Equal 'enter|n|k|s|text' ($kinds -join '|') 'paste: Enter / n / k / s / anything else is text, checked like the clipboard'
    $d = HumanPaste-Decide -Typed (HumanPaste-Typed -Answer '') -Default '11:45:29' -NowClock '13:00:00'
    Assert-True ($d['clock'] -eq '11:45:29' -and $d['source'] -eq 'worklist' -and -not $d['needMessage']) 'paste: Enter with a time on the row takes the row (a rerun / an old day needs no message)'
    $d = HumanPaste-Decide -Typed (HumanPaste-Typed -Answer '') -Default '' -NowClock '13:00:00'
    Assert-True ($d['clock'] -eq '13:00:00' -and $d['source'] -eq 'now' -and $d['needMessage']) 'paste: Enter without one is NOW and needs the start message'
    $d = HumanPaste-Decide -Typed (HumanPaste-Typed -Answer 'n') -Default '11:45:29' -NowClock '13:00:00'
    Assert-True ($d['clock'] -eq '13:00:00' -and $d['needMessage']) 'paste: n overrides the row with NOW'
    $d = HumanPaste-Decide -Typed (HumanPaste-Typed -Answer '10:30') -Default '11:45:29' -NowClock '13:00:00'
    Assert-True ($d['clock'] -eq '10:30:00' -and $d['source'] -eq 'typed' -and -not $d['needMessage']) 'paste: a typed time wins over the row'

    Write-Host '  -- first office run of the track: DF window, GIFT folder, waiting'
    $wins = @(@{ handle = '100'; title = 'DF - [C:\x\OLD1.csv  - C:\y\OLD2.csv]' }, @{ handle = '200'; title = 'DF - [C:\a\F202608310006.csv  - C:\b\F202610090033.csv]' }, @{ handle = '300'; title = 'Teams' })
    Assert-Equal '200' ([string](ScreenLaunchCapture-PickWindow -Windows $wins -Before @('100') -Title 'DF - ' -Names @('F202608310006', 'F202610090033'))) 'df: the new window showing the pair, not the old one left open'
    $cut = @(@{ handle = '400'; title = 'DF - [\\srv\...\GIFT\JJMRWE6L\F202608310006.csv  -  \\srv\...\GFIX\JJMRWE6L\F202610090033.cs]' })
    Assert-Equal '400' ([string](ScreenLaunchCapture-PickWindow -Windows $cut -Before @() -Title 'DF - ' -Names @('F202608310006', 'F202610090033'))) 'df: a title cut short by the program (".cs]") still matches by the file stems (office run 3)'
    Assert-Equal '0' ([string][int](ScreenLaunchCapture-PickWindow -Windows @($wins[0], $wins[2]) -Before @('100') -Title 'DF - ' -Names @('F1.csv', 'F2.csv'))) 'df: only the old window there -> none (never capture it)'
    $reused = @(@{ handle = '100'; title = 'DF - [C:\a\F1.csv  - C:\b\F2.csv]' })
    Assert-Equal '100' ([string](ScreenLaunchCapture-PickWindow -Windows $reused -Before @('100') -Title 'DF - ' -Names @('F1.csv', 'F2.csv'))) 'df: a single-instance program that reused its window for THIS pair is taken'
    Assert-True (BrowserWaitFor-IsPast -Iso '2000-01-01T00:00:00') 'wait: a window that ended long ago is settled (one read decides)'
    Assert-True (-not (BrowserWaitFor-IsPast -Iso '2999-01-01T00:00:00') -and -not (BrowserWaitFor-IsPast -Iso '')) 'wait: a future window / none -> keep polling'
    $giftRoot = Join-Path $tmp 'GIFT'
    New-Item -ItemType Directory -Path (Join-Path $giftRoot 'JJMRWE6F') -Force | Out-Null
    [System.IO.File]::WriteAllText((Join-Path (Join-Path $giftRoot 'JJMRWE6F') 'F1.csv'), "a`r`nb`r`n")
    $gl = Invoke-TestStep 'file.list' @{ dir = (Join-Path $giftRoot 'JJMRJE6F'); alsoDirs = @((Join-Path $giftRoot 'JJMRWE6F')); glob = '*.csv' }
    Assert-True ($gl['ok'] -and $gl['total'] -eq 1 -and $gl['dir'].EndsWith('JJMRWE6F')) 'gift: no folder under the J name -> the W-name one is listed'
    $gl = Invoke-TestStep 'file.list' @{ dir = (Join-Path $giftRoot 'NONE1'); alsoDirs = @((Join-Path $giftRoot 'NONE2'), ($giftRoot + '/')); glob = '*.csv' }
    Assert-True ((-not $gl['ok']) -and $gl['failure'] -eq 'file_not_found' -and $gl['message'].Contains('NONE1') -and $gl['message'].Contains('NONE2')) 'gift: neither folder -> stop and say every name tried (an unfilled name ending in / is never the GIFT root)'
    $gl = Invoke-TestStep 'file.list' @{ dir = (Join-Path $giftRoot 'JJMRJE6K'); alsoDirs = @((Join-Path $giftRoot 'JJMRWE6K')); glob = '*.csv'; createIfMissing = $true; requireFiles = $true }
    Assert-True ((-not $gl['ok']) -and $gl['failure'] -eq 'no_files' -and (Test-Path -LiteralPath (Join-Path $giftRoot 'JJMRJE6K') -PathType Container)) 'gift: neither folder + createIfMissing -> the J-name folder is made and the step stops (put the files in, r)'
    $gl = Invoke-TestStep 'file.list' @{ dir = (Join-Path $giftRoot 'JJMRJE6K'); glob = '*.csv'; requireFiles = $true }
    Assert-True ((-not $gl['ok']) -and $gl['failure'] -eq 'no_files') 'gift: the folder is there but empty -> still stops (retryable)'
    [System.IO.File]::WriteAllText((Join-Path (Join-Path $giftRoot 'JJMRJE6K') 'F2.csv'), "x`r`n")
    $gl = Invoke-TestStep 'file.list' @{ dir = (Join-Path $giftRoot 'JJMRJE6K'); glob = '*.csv'; requireFiles = $true }
    Assert-True ($gl['ok'] -and $gl['total'] -eq 1) 'gift: after the files are put in, r goes on'
    $d = HumanPaste-Decide -Typed (HumanPaste-Typed -Answer 'k') -Default '10:30:25' -NowClock '14:56:48'
    Assert-True ($d['clock'] -eq '10:30:25' -and $d['source'] -eq 'without') 'paste: k on a row with a time keeps that time (the second office run took now and overwrote it)'

    Write-Host '  -- 10/20: the start message / its counts typed or pasted AT the prompt'
    $no = [string][char]0x306E; $iine = -join @([char]0x3044, [char]0x3044, [char]0x306D)
    $msgSJ = $kana + ':SJODWSPQ' + $wo + $jisshi + '(' + $yotei + ':5' + $ken + ')'
    $ty = HumanPaste-Typed -Answer ($msgSJ + ' 13:30')
    Assert-True ($ty['kind'] -eq 'text' -and $ty['clock'] -eq '13:30:00' -and $ty['text'] -eq $msgSJ) ('paste: "<message> 13:30" -> the time comes out, the message stays: ' + $ty['kind'] + ' ' + $ty['clock'])
    $c = HumanPaste-Check -Text $ty['text'] -Pattern $pat -Expect 'SJODWSPQ' -ExpectGroup 'job'
    Assert-True ($c['ok'] -and $c['fields']['counts'] -eq ('5' + $ken)) 'paste: ... and it is checked like the clipboard (job + counts)'
    $ty = HumanPaste-Typed -Answer $msgSJ
    Assert-True ($ty['kind'] -eq 'text' -and $ty['clock'] -eq '') 'paste: the message with no time -> text, no clock (the caller takes now)'
    $withReaction = $msgSJ + "`r`n2 " + $ken + $no + ' ' + $iine + "!`r`n2"
    $c = HumanPaste-Check -Text $withReaction -Pattern $pat -Expect 'SJODWSPQ' -ExpectGroup 'job'
    Assert-True ($c['ok'] -and $c['fields']['counts'] -eq ('5' + $ken)) ('paste: Teams reaction lines ("2 ' + 'ken of likes") are not counts: ' + $c['fields']['counts'])
    $ty = HumanPaste-Typed -Answer ('5' + $ken + ' 13:30')
    Assert-True ($ty['kind'] -eq 'counts' -and $ty['counts'] -eq ('5' + $ken) -and $ty['clock'] -eq '13:30:00') 'paste: "5 ken 13:30" -> counts and time typed by hand'
    $ty = HumanPaste-Typed -Answer ('13:30 815' + $ken + $ten + '46 ' + $ken)
    Assert-True ($ty['kind'] -eq 'counts' -and $ty['counts'] -eq ('815' + $ken + $ten + '46' + $ken)) ('paste: time first, two counts -> both counts: ' + $ty['counts'])
    $ty = HumanPaste-Typed -Answer ('5' + $ken)
    Assert-True ($ty['kind'] -eq 'counts' -and $ty['clock'] -eq '') 'paste: counts alone -> no clock (the caller takes now)'
    Assert-Equal 'time' (HumanPaste-Typed -Answer '13:30')['kind'] 'paste: a time alone is still the time answer'

    Write-Host '  -- GIFT Shift_JIS vs GFIX UTF-8 (office run 3: DF said same content, the compare said ng at line 1)'
    $line1 = -join @([char]0xFF33, [char]0xFF33, ',', [char]0x8A08, [char]0x753B, [char]0x5206, [char]0x985E, ',', [char]0xFF33, [char]0xFF30, [char]0x90E8, [char]0x756A, [char]0x51E6, [char]0x7406)
    $body = $line1 + "`r`nJ,35,322003T2 JE01`r`n"
    $sj = Get-EbiCodePage -CodePage 932
    $pg = Join-Path $tmp 'gift.csv'; $pf = Join-Path $tmp 'gfix.csv'
    [System.IO.File]::WriteAllBytes($pg, $sj.GetBytes($body))
    [System.IO.File]::WriteAllBytes($pf, (New-Object System.Text.UTF8Encoding($false)).GetBytes($body.Replace("`r`n", "`n")))
    $cmp = Invoke-TestStep 'file.compare' @{ pairs = @(@{ left = $pg; right = $pf }) }
    Assert-True ($cmp['code'] -eq 'ok' -and $cmp['reason'].Contains('cp932 vs utf8')) ('compare: the same text in Shift_JIS and UTF-8 is the same content, and the encodings are named: ' + $cmp['reason'])
    $minusSj = $line1 + [char]0xFF0D + "1`r`n"; $minusU = $line1 + [char]0x2212 + "1`r`n"
    [System.IO.File]::WriteAllBytes($pg, $sj.GetBytes($minusSj)); [System.IO.File]::WriteAllBytes($pf, (New-Object System.Text.UTF8Encoding($false)).GetBytes($minusU))
    $cmp = Invoke-TestStep 'file.compare' @{ pairs = @(@{ left = $pg; right = $pf }) }
    Assert-Equal 'ok' $cmp['code'] 'compare: U+FF0D (CP932 minus) vs U+2212 (JIS-mapped minus) is the same character to a Shift_JIS diff'
    [System.IO.File]::WriteAllBytes($pf, (New-Object System.Text.UTF8Encoding($false)).GetBytes($line1 + 'X1' + "`r`n"))
    $cmp = Invoke-TestStep 'file.compare' @{ pairs = @(@{ left = $pg; right = $pf }) }
    Assert-True ($cmp['code'] -eq 'ng' -and $cmp['reason'].Contains('col 15: U+FF0D vs U+0058')) ('compare: a real difference names line, column and both characters: ' + $cmp['reason'])
    $d1 = ConvertFrom-EbiTextBytes -Bytes $sj.GetBytes($body)
    Assert-True ($d1['encoding'] -eq 'cp932' -and $d1['text'] -eq $body) 'compare: a Shift_JIS file decodes as one Shift_JIS text'

    Write-Host '  -- job log already in Downloads (office run: downloaded by hand, r still waited for a NEW file)'
    $dlDir = Join-Path $tmp 'Downloads'; New-Item -ItemType Directory -Path $dlDir -Force | Out-Null
    [System.IO.File]::WriteAllText((Join-Path $dlDir '1000004638518.log'), 'x')
    [System.IO.File]::WriteAllText((Join-Path $dlDir '1000004638518.log.crdownload'), 'x')
    [System.IO.File]::WriteAllText((Join-Path $dlDir '1000004639999.log'), 'x')
    $a = BrowserDownloadEach-Already -Dir $dlDir -Glob '*.log' -Term '1000004638518'
    Assert-True ($null -ne $a -and $a.Name -eq '1000004638518.log') 'joblog: the hand-downloaded file with the job number is taken'
    Assert-True ($null -eq (BrowserDownloadEach-Already -Dir $dlDir -Glob '*.log' -Term '1000004638521')) 'joblog: another job number is not'
    # ---------------------------------------------------------------- DF 19-line rule (10-12 review)
    $sp = { param($n) ScreenLaunchCapture-SecondPlan -Set @{ lines = $n; long = $true } -ScreenLines 19 -LinesField 'lines' -SecondField 'long' -KeysMid 'MID' -KeysSecond 'END' }
    Assert-True (-not (& $sp 0)['two']) 'df rule: 0 lines -> one picture'
    Assert-True (-not (& $sp 19)['two']) 'df rule: 19 lines -> one picture'
    Assert-Equal 'MID' (& $sp 20)['keys'] 'df rule: 20 lines -> second picture from line 20 (keysMid)'
    Assert-Equal 'MID' (& $sp 38)['keys'] 'df rule: 38 lines -> still keysMid'
    Assert-Equal 'END' (& $sp 39)['keys'] 'df rule: 39 lines -> the END view'
    $old = ScreenLaunchCapture-SecondPlan -Set @{ lines = 30; long = $true } -ScreenLines 0 -LinesField 'lines' -SecondField 'long' -KeysMid 'MID' -KeysSecond 'END'
    Assert-True ($old['two'] -and $old['keys'] -eq 'END') 'df rule: off (screenLines 0) -> the long flag and keysSecond as before'
    Assert-Equal 'end' (& $sp 39)['view'] 'df rule: the END view is named, so the window is resized to heightEnd first (10/19)'
    Assert-Equal 'mid' (& $sp 20)['view'] 'df rule: the PgDn page keeps the first height'
    # 10/19 measured: 18 px a line, 476 = 17 lines; the HOST shows 17 records, HOST n = DF n-1
    $lay = (Read-EbiJson -Path (Join-Path $repoRoot 'profiles/gfix-recv/layout.json'))['value']['df']
    # 10/21 measured sweep: first view shows 16 clean lines at ~430-434; the END view shows one line fewer than the window holds, 17 at 464-480
    Assert-True ([int]$lay['height'] -ge 430 -and [int]$lay['height'] -le 434) 'df size: first height in the measured 16-clean-lines range 430..434 (458 showed 17 + part of the 18th)'
    Assert-True ([int]$lay['heightEnd'] -ge 464 -and [int]$lay['heightEnd'] -le 480) 'df size: END height in the measured 17-lines range 464..480'
    # 10/21 fit (SJODWSPN: the 458 first picture showed the 17th line): bands of the line-number strip
    $bands = { param([int]$n, [int]$lastH) @(for ($i = 0; $i -lt $n; $i++) { @{ top = 18 * $i; bottom = 18 * $i + 11; height = $(if ($i -eq $n - 1 -and $lastH -gt 0) { $lastH } else { 12 }) } }) }
    $fp = ScreenLaunchCapture-FitPlan -Bands (& $bands 16 0) -Want 16
    Assert-True ($fp['done'] -and $fp['full'] -eq 16) 'df fit: 16 whole lines, wanted 16 -> done'
    $fp = ScreenLaunchCapture-FitPlan -Bands (& $bands 17 0) -Want 16
    Assert-True ((-not $fp['done']) -and $fp['delta'] -eq -6) 'df fit: a 17th whole line (the 10/20 upper picture) -> shrink by a step, not a whole line (the 17th may be only partly in the window)'
    $fp = ScreenLaunchCapture-FitPlan -Bands (& $bands 17 5) -Want 16
    Assert-True ((-not $fp['done']) -and $fp['partial'] -and $fp['full'] -eq 16 -and $fp['delta'] -eq -6) 'df fit: 16 lines + a sliver of the 17th -> shrink until it is gone'
    $fp = ScreenLaunchCapture-FitPlan -Bands (& $bands 15 0) -Want 16
    Assert-True ((-not $fp['done']) -and $fp['delta'] -eq 6) 'df fit: 15 lines, wanted 16 -> grow by a step'
    $fp = ScreenLaunchCapture-FitPlan -Bands (& $bands 13 0) -Want 17
    Assert-Equal 60 $fp['delta'] 'df fit: 4 short -> 3 whole lines and a step (18*3+6)'
    $fp = ScreenLaunchCapture-FitPlan -Bands (@(@{ top = 0; bottom = 4; height = 5 }) + (& $bands 16 0)) -Want 17
    Assert-True $fp['done'] 'df fit: a short FIRST band is the strip edge, still a line'
    Assert-True (-not (ScreenLaunchCapture-FitPlan -Bands @() -Want 16)['done']) 'df fit: no band read -> not done, no resize (delta 0)'
    $lc = @{ screenLines = 16; linesField = 'lines'; linesEnd = 17 }
    Assert-Equal 16 (ScreenLaunchCapture-Want -Set @{ lines = 61 } -In $lc -View 'first') 'df fit: the first picture of a long file = 16 lines'
    Assert-Equal 9 (ScreenLaunchCapture-Want -Set @{ lines = 9 } -In $lc -View 'first') 'df fit: a 9-line file shows its 9 lines, blank below'
    Assert-Equal 17 (ScreenLaunchCapture-Want -Set @{ lines = 61 } -In $lc -View 'end') 'df fit: the END picture = linesEnd (the HOST last 17)'
    Assert-Equal 0 (ScreenLaunchCapture-Want -Set @{ lines = 61 } -In @{ screenLines = 0; linesField = 'lines'; linesEnd = 0 } -View 'first') 'df fit: no line rule -> no fit'
    Assert-Equal 17 ([int]$lay['linesEnd']) 'df size: linesEnd 17 (HOST n = DF n-1, the HOST last picture shows 17)'

    # ---------------------------------------------------------------- HOST ranges -> DF pictures (10/21)
    $ocr = @('VIEW LJOD.S.VE  Columns 00001 00080', 'Command ===>   Scroll ===> CSR', '****** ***** Top of Data *****', '000001 20260518,D,1,ATH2520', '000002 20260518,D,1,ATH2521', '0OO003 20260518,D,1,ATH2522', '000004 20260518,D,1,ATH2523', '000005')
    $ln = @(ScreenReadLineNumbers-Parse -Lines $ocr)
    Assert-True ($ln.Count -eq 1 -and $ln[0]['ok'] -and $ln[0]['first'] -eq 1 -and $ln[0]['last'] -eq 5) 'host lines: numbers at the start of a line only (header Columns 00001 ignored), O/0 folded'
    $ln = @(ScreenReadLineNumbers-Parse -Lines @('005666 a', '005667 b', '005688 c', '005669 d', '005670 e', '005671 f'))
    Assert-True ($ln.Count -eq 1 -and $ln[0]['ok'] -and $ln[0]['first'] -eq 5666 -and $ln[0]['last'] -eq 5671) 'host lines: a misread number inside the run is corrected'
    $ln = @(ScreenReadLineNumbers-Parse -Lines @('005666 a', '005667 b', '005668 c', '005669 d', '005679 e'))
    Assert-True ($ln[0]['ok'] -and $ln[0]['last'] -eq 5670) 'host lines: a misread LAST number is the next line (a shot lists lines one after another)'
    $ln = @(ScreenReadLineNumbers-Parse -Lines @('007559 a', '007551 b', '007552 c', '007553 d', '007554 e'))
    Assert-True ($ln[0]['ok'] -and $ln[0]['first'] -eq 7550) 'host lines: a misread FIRST number is the line before the run under it'
    Assert-True (-not @(ScreenReadLineNumbers-Parse -Lines @('SOLID data', 'x'))[0]['ok']) 'host lines: no numbers -> not ok'
    # KJODWWBV 10/21: both HOST shots in ONE grouped image -> two runs
    $two = @(1..19 | ForEach-Object { '{0:000000} 0090,K*SP**' -f $_ }) + @('Command ===>') + @(7550..7570 | ForEach-Object { '{0:000000} 0090,K3KD42' -f $_ })
    $ln = @(ScreenReadLineNumbers-Parse -Lines $two)
    Assert-True ($ln.Count -eq 2 -and $ln[0]['first'] -eq 1 -and $ln[0]['last'] -eq 19 -and $ln[1]['first'] -eq 7550 -and $ln[1]['last'] -eq 7570) 'host lines: two shots in one image -> two runs (1-19, 7550-7570)'
    $ln = @(ScreenReadLineNumbers-Parse -Lines @('000001 a', '000002 b', '000004 d', '000005 e', '000006 f'))
    Assert-True ($ln.Count -eq 1 -and $ln[0]['first'] -eq 1 -and $ln[0]['last'] -eq 6) 'host lines: a line OCR did not see does not split the listing'
    $rg = { param($a, $b) @{ ok = $true; first = $a; last = $b; count = ($b - $a + 1) } }
    # MJODWDBC 10/21: HOST 1-19 and 5666-5685, the file 5684 lines -> DF 1-18 / 5665-5684 (20)
    $g = VerifyDfRanges-Sets -Ranges @((& $rg 1 19), (& $rg 5666 5685))
    $pl = VerifyDfRanges-Plan -Lines 5684 -Set $g['sets'][0]
    Assert-True ($pl['mode'] -eq 'align' -and $pl['first'] -eq 18 -and $pl['end'] -eq 20) 'df ranges: 19-line HOST shots -> 18 / 20 DF lines'
    # the usual shape: HOST 1-17 / 46-62, 61 lines -> 16 / 17
    $pl = VerifyDfRanges-Plan -Lines 61 -Set (VerifyDfRanges-Sets -Ranges @((& $rg 1 17), (& $rg 46 62)))['sets'][0]
    Assert-True ($pl['mode'] -eq 'align' -and $pl['first'] -eq 16 -and $pl['end'] -eq 17) 'df ranges: the usual 17-line HOST shots -> 16 / 17 (SJODWSPN)'
    # KJODWWBV 10/21: HOST 007550-007568 (007568 under a red box, OCR read up to 007567), the file 7568 lines -> DF 07549-07568 either way
    $pl = VerifyDfRanges-Plan -Lines 7568 -Set (VerifyDfRanges-Sets -Ranges @((& $rg 1 19), (& $rg 7550 7568)))['sets'][0]
    Assert-True ($pl['mode'] -eq 'align' -and $pl['first'] -eq 18 -and $pl['end'] -eq 20) 'df ranges: the file runs 1 line past the HOST -> END picture 07549-07568 (20)'
    $pl = VerifyDfRanges-Plan -Lines 7568 -Set (VerifyDfRanges-Sets -Ranges @((& $rg 1 19), (& $rg 7550 7567)))['sets'][0]
    Assert-True ($pl['mode'] -eq 'align' -and $pl['end'] -eq 20) 'df ranges: OCR missed the last HOST line under a red box -> the same 20 lines'
    Assert-Equal 'free' (VerifyDfRanges-Plan -Lines 7568 -Set (VerifyDfRanges-Sets -Ranges @((& $rg 1 19), (& $rg 7550 7563)))['sets'][0])['mode'] 'df ranges: more than maxExtra lines past the HOST -> not lined up'
    Assert-Equal 'free' (VerifyDfRanges-Plan -Lines 7560 -Set (VerifyDfRanges-Sets -Ranges @((& $rg 1 19), (& $rg 7550 7568)))['sets'][0])['mode'] 'df ranges: the HOST longer than the file -> not lined up'
    # empty lines dropped on the receive side: HOST ends far beyond the file -> not lined up
    $pl = VerifyDfRanges-Plan -Lines 900 -Set (VerifyDfRanges-Sets -Ranges @((& $rg 1 17), (& $rg 1180 1196)))['sets'][0]
    Assert-Equal 'free' $pl['mode'] 'df ranges: HOST last - 1 <> file lines (empty lines lost) -> free, the default pictures'
    $pl = VerifyDfRanges-Plan -Lines 2 -Set (VerifyDfRanges-Sets -Ranges @((& $rg 1 2)))['sets'][0]
    Assert-True ($pl['mode'] -eq 'align' -and $pl['first'] -eq 2 -and $pl['end'] -eq 0) 'df ranges: RJODWDBC -- one HOST shot 000001-000002, a 2-line file -> one DF picture of 2 lines'
    $ln = @(ScreenReadLineNumbers-Parse -Lines @('000001', '000002 a'))
    Assert-True ($ln[0]['ok'] -and $ln[0]['last'] -eq 2) 'host lines: a 2-line shot is trusted (minLines 2)'
    $pl = VerifyDfRanges-Plan -Lines 12 -Set (VerifyDfRanges-Sets -Ranges @((& $rg 1 13)))['sets'][0]
    Assert-True ($pl['mode'] -eq 'align' -and $pl['first'] -eq 12 -and $pl['end'] -eq 0) 'df ranges: one HOST picture = the whole file -> one DF picture'
    $g = VerifyDfRanges-Sets -Ranges @((& $rg 1 17), (& $rg 46 62), (& $rg 1 19), (& $rg 100 118))
    Assert-Equal 2 @($g['sets']).Count 'df ranges: a picture from 000001 opens the next file pair'
    Assert-Equal 'default' (VerifyDfRanges-Plan -Lines 61 -Set $null)['mode'] 'df ranges: no HOST set -> default'
    $g = VerifyDfRanges-Sets -Ranges @(@{ ok = $false }, (& $rg 1 17))
    Assert-Equal 1 $g['bad'] 'df ranges: an unreadable picture is counted'
    # launch_capture follows an align plan: start heights moved by 18 px a line
    $in = @{ screenLines = 16; linesEnd = 17; height = 432; heightEnd = 476; fitPitchPx = 18; linesField = 'lines'; secondField = 'long'; keysMid = 'MID'; keysSecond = 'END' }
    $sp = ScreenLaunchCapture-ShotPlan -Set @{ lines = 5684 } -In $in -Plan @{ mode = 'align'; first = 18; end = 20 }
    Assert-True ($sp['aligned'] -and $sp['wantFirst'] -eq 18 -and $sp['wantLast'] -eq 20 -and $sp['firstHeight'] -eq 468 -and $sp['lastHeight'] -eq 530 -and $sp['view'] -eq 'end') 'df shots: align 18 / 20 -> 468 / 530, END view'
    $sp = ScreenLaunchCapture-ShotPlan -Set @{ lines = 61 } -In $in -Plan @{ mode = 'free' }
    Assert-True ((-not $sp['aligned']) -and $sp['wantFirst'] -eq 16 -and $sp['wantLast'] -eq 17 -and $sp['lastHeight'] -eq 476) 'df shots: free -> the line rule (16 / 17)'
    $sp = ScreenLaunchCapture-ShotPlan -Set @{ lines = 20 } -In $in -Plan $null
    Assert-True ($sp['view'] -eq 'mid' -and $sp['wantLast'] -eq 0) 'df shots: no plan, 20 lines -> the PgDn page, not fitted'
    $sp = ScreenLaunchCapture-ShotPlan -Set @{ lines = 12 } -In $in -Plan @{ mode = 'align'; first = 12; end = 0 }
    Assert-True ((-not $sp['two']) -and $sp['wantFirst'] -eq 12) 'df shots: align with end 0 -> one picture'
    $dr = Invoke-TestStep 'verify.df_ranges' @{ pairs = @(@{ lines = 5684; rightName = 'F1.csv' }, @{ lines = 900; rightName = 'F2.csv' }); ranges = @((& $rg 1 19), (& $rg 5666 5685), (& $rg 1 17), (& $rg 1180 1196)) }
    Assert-True ($dr['ok'] -and $dr['aligned'] -eq 1 -and $dr['plans'][1]['mode'] -eq 'free' -and $dr['alert'] -eq '') 'df ranges step: pair 1 lined up, pair 2 free (empty lines) with no alert'
    $dr = Invoke-TestStep 'verify.df_ranges' @{ pairs = @(@{ lines = 61; rightName = 'F1.csv' }); ranges = @(@{ ok = $false; reason = 'x' }) }
    Assert-True ($dr['plans'][0]['mode'] -eq 'default' -and $dr['alert'] -match 'NOT lined up') 'df ranges step: HOST pictures there but unreadable -> default pictures + a red line on the gate'
    # ---------------------------------------------------------------- picture 1: job column by pixels (10/21)
    $gg = Get-EbiStripGeometry -WidthA 1200 -HeightA 300 -WidthB 600 -HeightB 150 -FromFrac 0.8 -ToFrac 1.0
    Assert-True ($gg['ok'] -and $gg['x0'] -eq 960 -and $gg['x1'] -eq 1200 -and $gg['scaledWidthB'] -eq 1200) 'pic1 strip: B shown at half size is scaled to A''s height, the job strip at 960..1200'
    $gg = Get-EbiStripGeometry -WidthA 1200 -HeightA 300 -WidthB 500 -HeightB 150 -FromFrac 0.8 -ToFrac 1.0
    Assert-True (-not $gg['ok']) 'pic1 strip: the other picture narrower than the strip -> not compared'
    Assert-True (-not (Get-EbiStripGeometry -WidthA 100 -HeightA 50 -WidthB 100 -HeightB 50 -FromFrac 0 -ToFrac 0)['ok']) 'pic1 strip: no job column found (0..0) -> not compared'
    Assert-Equal 'ok' (ScreenCompareStrip-Verdict -Score 3 -Ink 0.05 -OkMax 8 -NgMin 25 -MinInk 0.005) 'pic1 cmp: small difference -> ok'
    Assert-Equal 'ng' (ScreenCompareStrip-Verdict -Score 40 -Ink 0.05 -OkMax 8 -NgMin 25 -MinInk 0.005) 'pic1 cmp: big difference -> ng'
    Assert-Equal 'unknown' (ScreenCompareStrip-Verdict -Score 15 -Ink 0.05 -OkMax 8 -NgMin 25 -MinInk 0.005) 'pic1 cmp: in between -> unknown'
    Assert-Equal 'unknown' (ScreenCompareStrip-Verdict -Score 0 -Ink 0 -OkMax 8 -NgMin 25 -MinInk 0.005) 'pic1 cmp: a blank strip says nothing'
    $cs = Invoke-TestStep 'screen.compare_strip' @{ a = @(); b = 'x.png' }
    Assert-True ($cs['ok'] -and $cs['code'] -eq 'unknown' -and $cs['alert'] -ne '') 'pic1 cmp step: the new picture missing -> unknown + alert, never a stop'
    # 10/21 FJODWWBV: OCR frame and the pixel search
    $fr = Get-EbiOcrFrame -Width 6000 -Height 1300
    Assert-True ($fr['cropW'] -eq 1400 -and $fr['outW'] -le 2500 -and $fr['outH'] -le 2500) 'ocr frame: a wide shot -> its left 1400 px, within the recognizer size'
    $fr = Get-EbiOcrFrame -Width 1200 -Height 4000
    Assert-True ($fr['outH'] -eq 2500 -and $fr['outW'] -eq 750) 'ocr frame: a tall picture is scaled down to fit'
    [void](Get-EbiImageMath)
    $pat = [byte[]](0, 255, 255, 0)
    $hay = New-Object byte[] 60; for ($i = 0; $i -lt 60; $i++) { $hay[$i] = 255 }; $hay[33] = 0; $hay[54] = 0   # pattern at x=13, y=1 in 20x3
    $bd = [EbiImageMath]::BestDiff($pat, 2, 2, $hay, 20, 3, 9, 1, 10, 1)
    Assert-True ([double]$bd[0] -eq 0 -and [int]$bd[1] -eq 13 -and [int]$bd[2] -eq 1) 'pixel search: the same cells found wherever they sit across the other picture'
    $cr = [EbiImageMath]::Crop([byte[]](1, 2, 3, 4, 5, 6), 3, 1, 1, 2, 1)
    Assert-True ($cr.Length -eq 2 -and $cr[0] -eq 5 -and $cr[1] -eq 6) 'pixel search: the best place cut out of the other picture (saved for a person)'
    $bin = [EbiImageMath]::Threshold([byte[]](10, 10, 10, 200, 200, 200), 6, 2, 1, [byte[]](10, 200))
    Assert-True ($bin[0] -eq 255 -and $bin[3] -eq 0) 'ocr bin: a dark ground -> the bright text turns black on white'
    $sls = @(@{ match = '-unzip'; label = 'GFIXReceiveUnzip.log' }, @{ match = '.'; label = 'GFIXReceive.log' })
    Assert-Equal 'GFIXReceiveUnzip.log' (FileExtractBlocks-Label -Sources @('log/GFIXReceive/1021-unzip.log') -SourceLabels $sls -Default 'X') 'recv label: a block from the unzip log is GFIXReceiveUnzip.log (10/21)'
    Assert-Equal 'GFIXReceive.log / GFIXReceiveUnzip.log' (FileExtractBlocks-Label -Sources @('a/1021.log', 'a/1021-unzip.log', 'a/1021.log') -SourceLabels $sls -Default 'X') 'recv label: blocks from both logs name both, in order'
    Assert-Equal 'X' (FileExtractBlocks-Label -Sources @() -SourceLabels $sls -Default 'X') 'recv label: no block -> the default'
    # 10/21: the 同一内容 cell found in the picture (Find-EbiTextPanel)
    $sw = 70; $sh = 14; $sg = New-Object byte[] ($sw * $sh); for ($i = 0; $i -lt $sg.Length; $i++) { $sg[$i] = 240 }
    foreach ($y in 0..13) { $sg[$y * $sw + 24] = 150; $sg[$y * $sw + 62] = 150 }                      # the cell's borders
    foreach ($y in 3..10) { foreach ($x in 5..14) { $sg[$y * $sw + $x] = 20 }; foreach ($x in 30..50) { if ($x % 3 -ne 0) { $sg[$y * $sw + $x] = 20 } } }   # 移動 | 同一内容
    $tp = Find-EbiTextPanel -Gray $sg -W $sw -H $sh
    Assert-True ($tp['ok'] -and $tp['left'] -eq 24 -and $tp['right'] -eq 62 -and $tp['top'] -eq 0 -and $tp['bottom'] -eq 13) 'df status: the box is the cell around the LAST text (not 移動), border to border'
    $sg2 = New-Object byte[] ($sw * $sh); for ($i = 0; $i -lt $sg2.Length; $i++) { $sg2[$i] = 240 }
    Assert-True (-not (Find-EbiTextPanel -Gray $sg2 -W $sw -H $sh)['ok']) 'df status: no text -> not found (the fixed box is the fallback)'
    $old = Join-Path $tmp 'old__first.png'; [System.IO.File]::WriteAllText($old, 'x')
    $sp = ScreenLaunchCapture-ShotPlan -Set @{ lines = 12 } -In $in -Plan $null
    Assert-True (ScreenLaunchCapture-CanReuse -Old @{ first = $old; last = ''; firstLines = 12 } -Plan $sp) 'df shots: an earlier picture with the lines wanted is kept'
    Assert-True (-not (ScreenLaunchCapture-CanReuse -Old @{ first = $old; last = ''; firstLines = 11 } -Plan $sp)) 'df shots: other line count -> captured again'
    Assert-True (-not (ScreenLaunchCapture-CanReuse -Old @{ first = $old; last = '' } -Plan $sp)) 'df shots: a picture from before the fit (no firstLines) -> captured again'

    # ---------------------------------------------------------------- home cell (10-12)
    $send = [string][char]0x9001 + [char]0x4FE1 + [char]0x21D2   # send-arrow sheet
    $recv = [string][char]0x53D7 + [char]0x4FE1 + [char]0x21D2   # receive-arrow sheet
    $homes = @{}; $homes[$send] = 'A1'; $homes[$recv] = 'A1'
    Assert-Equal 'A3' (ExcelTidy-CellFor -Sheet 'GFIX' -Default 'A3' -Map $homes) 'tidy: an ordinary sheet gets the home cell'
    Assert-Equal 'A1' (ExcelTidy-CellFor -Sheet ($send + 'X') -Default 'A3' -Map $homes) 'tidy: a sheet starting with a cellFor key gets its cell'
    Assert-Equal 'A1' (ExcelTidy-CellFor -Sheet $recv -Default 'A3' -Map $homes) 'tidy: ... the receive sheet too'
    Assert-Equal 'A1' (ExcelTidy-CellFor -Sheet 'x' -Default '' -Map $null) 'tidy: nothing configured -> A1'

    # ---------------------------------------------------------------- gzip body (10-12)
    # Jenkins sent a 770-byte CSV gzip-encoded; it was saved as 478 bytes
    # starting 1F 8B, and DF / file.compare saw garbage
    $csvText = "SS,Y`r`nbase,2026/05/15`r`n"
    $plainBytes = [System.Text.Encoding]::ASCII.GetBytes($csvText)
    $gzPath = Join-Path $tmp 'F202610090024.csv'
    $ms = New-Object System.IO.MemoryStream
    $gzw = New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Compress, $true)
    $gzw.Write($plainBytes, 0, $plainBytes.Length); $gzw.Dispose()
    [System.IO.File]::WriteAllBytes($gzPath, $ms.ToArray())
    Assert-True (FileDownload-Gunzip -Path $gzPath) 'download: a gzip body under a .csv name is decompressed'
    Assert-Equal $csvText ([System.IO.File]::ReadAllText($gzPath)) 'download: ... back to the plain text'
    Assert-True (-not (FileDownload-Gunzip -Path $gzPath)) 'download: plain text is left alone'
    [System.IO.File]::WriteAllBytes((Join-Path $tmp 'a.gz'), $ms.ToArray())
    Assert-True (-not (FileDownload-Gunzip -Path (Join-Path $tmp 'a.gz'))) 'download: a real .gz file is kept as it is'
    [System.IO.File]::WriteAllBytes((Join-Path $tmp 'empty.csv'), [byte[]]@())
    Assert-True (-not (FileDownload-Gunzip -Path (Join-Path $tmp 'empty.csv'))) 'download: a 0-byte file is left alone'
    # ============================================================ evidence: picture 1 is this job's row (10/15)
    $ft = Invoke-TestStep 'verify.find_token' @{ text = "Host JOD - Open JPC - JJPCRS13 IF6000 001 JOD_PRD1_IF6000_FTP TEXT No SJODJDB6"; expect = 'SJODJDB6'; label = 'picture 1' }
    Assert-True ($ft['ok'] -and $ft['code'] -eq 'ok' -and $ft['alert'] -eq '') ('find_token: the J name on the picture -> ok: ' + $ft['reason'])
    $ft = Invoke-TestStep 'verify.find_token' @{ text = "JJPCRS13 IF6000 TEXT No SJ0D JDB6"; expect = 'SJODJDB6' }
    Assert-Equal 'ok' $ft['code'] 'find_token: OCR look-alike (0 for O) and a space inside the word still match'
    $ft = Invoke-TestStep 'verify.find_token' @{ text = "JJPCRS13 IF6000 TEXT No QJODJDC2"; expect = 'SJODJDB6'; label = 'picture 1' }
    Assert-True ($ft['code'] -eq 'ng' -and $ft['alert'] -like '*does NOT show SJODJDB6*QJODJDC2*') ('find_token: another job -> ng, alert names what was read: ' + $ft['alert'])
    $ft = Invoke-TestStep 'verify.find_token' @{ text = "JJPCRS13 IF6000 TEXT No RJODJDB6"; expect = 'SJODJDB6' }
    Assert-True ($ft['code'] -eq 'unknown' -and $ft['alert'] -like '*one character off*') ('find_token: one letter off (R vs S: misread or the neighbour job) -> unknown, still red: ' + $ft['alert'])
    $ft = Invoke-TestStep 'verify.find_token' @{ text = ''; expect = 'SJODJDB6' }
    Assert-True ($ft['code'] -eq 'unknown' -and $ft['alert'] -eq '') 'find_token: no OCR text (no OCR on this PC) -> unknown, no red line'
    $ft = Invoke-TestStep 'verify.find_token' @{ text = 'SJODJDB6'; expect = '' }
    Assert-True ($ft['code'] -eq 'unknown' -and $ft['reason'] -like '*no expected name*') 'find_token: a row with no JOB -> unknown'
    Assert-Equal 1 (VerifyFindToken-Distance 'SJODJDB6' 'RJODJDB6') 'find_token: distance of one substitution'
    $oc = Invoke-TestStep 'screen.ocr' @{ path = '' }
    Assert-True ($oc['ok'] -and -not $oc['available'] -and $oc['text'] -eq '') 'screen.ocr: no picture exported -> ok, nothing read, no failure'
    if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
        $png = Join-Path $tmp 'pic1.png'; [System.IO.File]::WriteAllBytes($png, [byte[]]@(137, 80, 78, 71))
        $oc = Invoke-TestStep 'screen.ocr' @{ path = $png }
        Assert-True ($oc['ok'] -and -not $oc['available'] -and @($oc['warnings']).Count -eq 1) 'screen.ocr: no Windows OCR here -> ok, available=false, one warning'
    }
    $oc = Invoke-TestStep 'screen.ocr' @{ path = (Join-Path $tmp 'nope.png') }
    Assert-Equal 'file_not_found' $oc['failure'] 'screen.ocr: a named picture that is not there is a failure'
    $gw = HumanGate-WhatLines -Code 'unknown' -Reason 'picture 1 shows SJODJDB6' -Checklist @('B3 is this job', '', 'red boxes') -Alert @('', 'picture 1 does NOT show SJODJDB6 (read: QJODJDC2): another job''s picture? This line is long enough to wrap onto a second panel row')
    Assert-True ($gw['what'][0] -like '!! picture 1 does NOT*') 'gate: the alert comes first, marked !!'
    Assert-Equal 2 @($gw['what'] | Where-Object { $_ -like 'check: *' }).Count 'gate: one check line per non-empty checklist point'
    Assert-True (@($gw['alert']).Count -ge 2) ('gate: a long alert is red on every wrapped row: ' + @($gw['alert']).Count + ' fragment(s)')
    $panel = @(Format-EbiGatePanel -Title 'GATE X' -What $gw['what'])
    $red = @($panel | Where-Object { (Get-EbiGateLineColor -Line $_ -Alert $gw['alert']) -eq 'Red' })
    Assert-Equal @($gw['alert']).Count $red.Count 'gate: exactly the alert rows of the rendered panel are red'
    $gw = HumanGate-WhatLines -Code 'unknown' -Reason '' -Checklist @() -Alert ''
    Assert-True (@($gw['alert']).Count -eq 0 -and $gw['what'][0] -eq 'verdict so far: unknown') 'gate: no alert, no red'
    # ============================================================ deliver (10/15): check sheet row + mail draft
    $p0 = ExcelAppendRow-Plan -LastRow 358 -FirstRow 19 -KeyTexts @{ 355 = 'J4_book_RJODWDB6.xlsx'; 356 = ' j4_book_sjodwdb6.xlsx ' } -Key 'J4_book_SJODWDB6.xlsx'
    Assert-True ($p0['existing'] -and $p0['row'] -eq 356 -and -not $p0['added']) 'append_row: a book already listed (10/14, by hand) -> nothing added, its row reported (trim + case-insensitive)'
    $p0 = ExcelAppendRow-Plan -LastRow 358 -FirstRow 19 -KeyTexts @{ 355 = 'J4_book_RJODWDB6.xlsx' } -Key 'J4_book_KJODWDBC.xlsx'
    Assert-True ($p0['added'] -and $p0['row'] -eq 359) 'append_row: a new book goes under the last used row'
    Assert-Equal 19 (ExcelAppendRow-Plan -LastRow 2 -FirstRow 19 -KeyTexts @{} -Key '')['row'] 'append_row: never above firstRow'
    Assert-Equal '2026-10-16' ((ExcelAppendRow-ParseDate '2026-10-16').ToString('yyyy-MM-dd')) 'append_row: run.date -> a real date'
    Assert-True ($null -eq (ExcelAppendRow-ParseDate 'JAVA')) 'append_row: text is not a date'
    Assert-True ((ExcelAppendRow-Said -Sheet 'S' -Row 356 -Existing $true -When '2026/10/14') -like 'already listed on S row 356 (2026/10/14)*') 'append_row: the gate line names the row and its date'

    $at = [string][char]64
    $addrTo = 'reviewer' + $at + 'example.invalid'; $addrCc = 'lead' + $at + 'example.invalid'
    $tplObj = @{ token = 'MJDSWM39'; subject = ' [GIFT] BIX_J4 review (MJDSWM39)'; body = "Hello`r`nJ4_book_ MJDSWM39.xlsx`r`n"
                 html = '<p>J4_book_<span style=''x''> MJDSWM39</span>.xlsx</p><p>MJDSWM39 again</p>'; recipients = @(@{ name = 'R'; address = $addrTo; type = 1 }) }
    $tplJson = Join-Path $tmp 'review-mail.json'
    [void](Write-EbiJson -Path $tplJson -Value $tplObj)
    $rt = Read-EbiMailTemplate -Path $tplJson
    Assert-True ($rt['ok'] -and $rt['template']['token'] -eq 'MJDSWM39' -and @($rt['template']['recipients']).Count -eq 1) ('mail: a .json template reads: ' + $rt['message'])
    $bat = Join-Path $tmp 'sendj4mail.bat'
    $b64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes((ConvertTo-EbiJson -Value $tplObj)))
    [System.IO.File]::WriteAllText($bat, ("@echo off`r`nrem ...`r`n#==POWERSHELL==`r`n# code`r`n#==TEMPLATE_JSON==`r`n" + $b64 + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
    $rb = Read-EbiMailTemplate -Path $bat
    Assert-True ($rb['ok'] -and $rb['template']['subject'] -eq $tplObj['subject']) ('mail: the template embedded in sendj4mail.bat reads the same: ' + $rb['message'])
    # 10/16: the .bat's base64 body pasted into a .json file reads as the template too
    $tplB64 = Join-Path $tmp 'review-mail-b64.json'
    [System.IO.File]::WriteAllText($tplB64, ($b64.Substring(0, 40) + "`r`n" + $b64.Substring(40) + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
    $r64 = Read-EbiMailTemplate -Path $tplB64
    Assert-True ($r64['ok'] -and $r64['template']['token'] -eq 'MJDSWM39' -and $r64['template']['subject'] -eq $tplObj['subject']) ('mail: a .json holding the .bat''s base64 body (line breaks too) reads as the template: ' + $r64['message'])
    $tplBad = Join-Path $tmp 'review-mail-bad.json'
    [System.IO.File]::WriteAllText($tplBad, 'QUJDREVGRw==', (New-Object System.Text.UTF8Encoding($false)))
    $rbad = Read-EbiMailTemplate -Path $tplBad
    Assert-True (-not $rbad['ok'] -and $rbad['message'] -like '*holds base64*sendj4mail.bat*') ('mail: base64 that is not a template says so: ' + $rbad['message'])
    $x = Expand-EbiMailTemplate -Template $rb['template'] -Value 'RJODWDB6' -TrimBefore
    Assert-True ($x['ok'] -and $x['subject'] -like '*review (RJODWDB6)') 'mail: the token becomes the job in the subject'
    Assert-True ($x['body'] -like '*J4_book_RJODWDB6.xlsx*') ('mail: "_ MJDSWM39.xlsx" -> "_RJODWDB6.xlsx" (the stray space dropped): ' + $x['body'])
    Assert-True ($x['html'] -like '*J4_book_<span style=''x''>RJODWDB6</span>.xlsx*' -and $x['html'] -like '*<p>RJODWDB6 again*') ('mail: html too, the space inside the span dropped, other places replaced: ' + $x['html'])
    Assert-True (-not (Expand-EbiMailTemplate -Template $rb['template'] -Token 'NOPE' -Value 'X')['ok']) 'mail: a token found nowhere is refused (the mail would go out unchanged)'
    $rc = @(HumanDraftMail-Recipients -Recipients $rb['template']['recipients'] -To @() -Cc @($addrCc))
    Assert-True ($rc.Count -eq 2 -and $rc[0]['address'] -eq $addrTo -and $rc[1]['type'] -eq 2) 'mail: the template To kept, a CC added'
    $rc = @(HumanDraftMail-Recipients -Recipients $rb['template']['recipients'] -To @($addrCc) -Cc @())
    Assert-True ($rc.Count -eq 1 -and $rc[0]['address'] -eq $addrCc) 'mail: a given To replaces the template To'
    $ctx['DryRun'] = $true
    $dm = Invoke-TestStep 'human.draft_mail' @{ template = $bat; value = 'RJODWDB6' }
    Assert-True ($dm['ok'] -and $dm['subject'] -like '*(RJODWDB6)' -and @($dm['to']).Count -eq 1 -and -not $dm['saved']) 'draft_mail: a dry run shows the subject it would use, saves nothing'
    $ctx['DryRun'] = $false
    $dm = Invoke-TestStep 'human.draft_mail' @{ template = (Join-Path $tmp 'none.json'); value = 'X' }
    Assert-Equal 'file_not_found' $dm['failure'] 'draft_mail: no template file -> file_not_found (set paths.mailTemplate)'
    [System.IO.File]::WriteAllText((Join-Path $tmp 'bad.bat'), "@echo off`r`n", (New-Object System.Text.UTF8Encoding($false)))
    $dm = Invoke-TestStep 'human.draft_mail' @{ template = (Join-Path $tmp 'bad.bat'); value = 'X' }
    Assert-True ($dm['failure'] -eq 'template_invalid' -and $dm['message'] -like '*TEMPLATE_JSON*') 'draft_mail: a .bat without the template section -> template_invalid'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

exit (Complete-Tests)
