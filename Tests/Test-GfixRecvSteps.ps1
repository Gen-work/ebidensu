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
                 'verify.derive_fields', 'human.paste', 'screen.launch_capture', 'browser.wait_for', 'browser.download_each')) {
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

    # ============================================================ kernel/Layout.ps1
    $bs = @(Get-EbiInkBands -Counts @(0, 0, 3, 4, 0, 5, 0, 0, 0, 2) -MinInk 1 -MergeGap 1)
    Assert-Equal '2-5|9-9' ((@($bs | ForEach-Object { '' + $_['top'] + '-' + $_['bottom'] })) -join '|') 'layout: a one-row gap merges, a three-row gap splits'
    $a = Resolve-EbiAnchoredRect -Rect @{ x = 117; y = 19; w = 112; h = 19; anchor = 'br' } -ImageWidth 1133 -ImageHeight 429
    Assert-Equal '1016,410' ('' + $a['x'] + ',' + $a['y']) 'layout: a bottom-right anchored box lands on the status bar cell'
    $sr = ConvertTo-EbiSheetRect -PicLeft 10 -PicTop 100 -PicWidth 849.75 -PicHeight 321.75 -ImageWidth 1133 -ImageHeight 429 -Rect @{ x = 100; y = 40; w = 48; h = 18 }
    Assert-Equal '85,130,36,13.5' ('' + $sr['left'] + ',' + $sr['top'] + ',' + $sr['width'] + ',' + $sr['height']) 'layout: image px -> sheet points through the picture scale'

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
    Assert-True ($c['ok'] -and $c['fields']['job'] -eq $jobW -and $c['fields']['count'] -eq '6') ('paste: the Teams start message gives job + send count: ' + $c['reason'])
    $msgFw = $kana + [char]0xFF1A + (-join ($jobW.ToCharArray() | ForEach-Object { [char]([int]$_ + 0xFEE0) })) + $wo + $jisshi
    $c = HumanPaste-Check -Text $msgFw -Pattern $pat -Expect $jobW -ExpectGroup 'job'
    Assert-True ($c['ok'] -and $c['fields']['count'] -eq '') 'paste: full-width colon / letters fold; no send count -> count is "" (still templatable)'
    $c = HumanPaste-Check -Text $msgOk -Pattern $pat -Expect 'SJDSWM40' -ExpectGroup 'job'
    Assert-True ((-not $c['ok']) -and $c['reason'].Contains('JJMRWE6K') -and $c['reason'].Contains('SJDSWM40')) 'paste: a message for another job is refused and says both names'
    $c = HumanPaste-Check -Text 'hello' -Pattern $pat -Expect '' -ExpectGroup 'job'
    Assert-True ((-not $c['ok']) -and $c['fields'].Contains('count') -and $c['fields'].Contains('job')) 'paste: not the message -> refused, every group still present as ""'
    Assert-Equal 'time|10:30:00' ((HumanPaste-Typed -Answer '10:30')['kind'] + '|' + (HumanPaste-Typed -Answer '10:30')['clock']) 'paste: a typed time replaces now'
    Assert-Equal '09:05:00' (HumanPaste-Typed -Answer '905')['clock'] 'paste: 905 -> 09:05:00'
    $kinds = @(foreach ($a in @('', 'n', 'k', 's', 'zz')) { (HumanPaste-Typed -Answer $a)['kind'] })
    Assert-Equal 'enter|n|k|s|other' ($kinds -join '|') 'paste: Enter / n / k / s / anything else'
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
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

exit (Complete-Tests)
