# Test-VerifySteps.ps1 -- kernel/Parse.ps1 (P1-30 / P1-31), verify.parse_text,
# verify.match_record (P1-32), verify.assert (P1-33) and the human steps
# on kernel/Gate.ps1 (P1-34), driven with a scripted reader so nothing
# blocks. ASCII source; Japanese via [char]. No param() block.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_TestCommon.ps1')
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'kernel/Registry.ps1')
. (Join-Path $repoRoot 'kernel/Parse.ps1')
. (Join-Path $repoRoot 'kernel/Gate.ps1')
. (Join-Path $repoRoot 'kernel/Key.ps1')

Reset-Tests 'VerifySteps'
$tab = "`t"
$jaStatus = '' + [char]0x72B6 + [char]0x614B          # status label
$jaEnd    = '' + [char]0x7D42 + [char]0x4E86 + [char]0x6642 + [char]0x523B   # end-time label
$jaNormal = '' + [char]0x6B63 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86   # normal end
$jaRef    = '' + [char]0x53C2 + [char]0x7167          # reference (Jenkins list)
$jaFile   = '' + [char]0x30D5 + [char]0x30A1 + [char]0x30A4 + [char]0x30EB + [char]0x540D   # file name header
$jaTime   = '' + [char]0x66F4 + [char]0x65B0 + [char]0x65E5 + [char]0x6642   # update time header

function New-TestLog {
    $log = New-Object PSObject
    $log | Add-Member -MemberType NoteProperty -Name Lines -Value (New-Object System.Collections.ArrayList)
    $log | Add-Member -MemberType ScriptMethod -Name Info  -Value { param($m) [void]$this.Lines.Add('info:' + $m) }
    $log | Add-Member -MemberType ScriptMethod -Name Warn  -Value { param($m) [void]$this.Lines.Add('warn:' + $m) }
    $log | Add-Member -MemberType ScriptMethod -Name Debug -Value { param($m) }
    return $log
}
function New-TestCtx { param([bool]$DryRun = $false, $Item = $null) return @{ WorkDir = [System.IO.Path]::GetTempPath(); RunId = 'r1'; Profile = @{}; Log = (New-TestLog); DryRun = $DryRun; Session = @{}; Item = $Item; KeyColumns = @('Correl_ID_S') } }
$reg = New-EbiRegistry -ModulesRoot (Join-Path $repoRoot 'modules')
foreach ($u in @('verify.parse_text', 'verify.match_record', 'verify.assert', 'human.prepare', 'human.gate', 'human.choose')) {
    $r = . Import-EbiStep -Registry $reg -Use $u
    Assert-True $r['ok'] ('loads: ' + $u + ' ' + $r['message'])
}
function Step { param([string]$Use, [hashtable]$With, $Ctx) return (& (Get-EbiStep -Registry $reg -Use $Use)['Invoke'] $With $Ctx) }

# --- 1. time parsing: single-digit hours ---------------------------------------------
Write-Host '  -- ConvertTo-EbiDateTime'
$t = ConvertTo-EbiDateTime -Text '2026/08/24 9:50:03'
Assert-True ($t['ok'] -and $t['value'].Hour -eq 9 -and $t['value'].Second -eq 3) 'REGRESSION: a single-digit hour parses (the old \d{2} lost every row before 10:00)'
$t = ConvertTo-EbiDateTime -Text '2026/08/24 10:02:11'
Assert-True ($t['ok'] -and $t['value'].Hour -eq 10) 'two-digit hour parses too'
$t = ConvertTo-EbiDateTime -Text '2026/8/4 9:5:3'
Assert-True (-not $t['ok'] -or $t['value'].Minute -eq 5) 'single-digit minutes: parse or refuse, never mis-read'
$t = ConvertTo-EbiDateTime -Text '9:50:03' -Date ([datetime]'2026-08-24')
Assert-True ($t['ok'] -and $t['value'] -eq [datetime]'2026-08-24T09:50:03') 'time only is placed on the given date'
$t = ConvertTo-EbiDateTime -Text '2026/08/24  9:50:03'
Assert-True $t['ok'] 'doubled inner whitespace is collapsed'
$t = ConvertTo-EbiDateTime -Text '20260824095003'
Assert-True ($t['ok'] -and $t['value'].Hour -eq 9) 'compact 14-digit stamp'
$t = ConvertTo-EbiDateTime -Text '2026-08-24T09:50:03'
Assert-True $t['ok'] 'ISO'
$t = ConvertTo-EbiDateTime -Text 'yesterday'
Assert-True (-not $t['ok']) 'junk is not ok'
$t = ConvertTo-EbiDateTime -Text '24.08.2026 9:50' -Formats @('dd.MM.yyyy H:mm')
Assert-True ($t['ok'] -and $t['format'] -eq 'dd.MM.yyyy H:mm') 'a grammar-supplied format is tried first'

# --- 2. grammars ------------------------------------------------------------------------
Write-Host '  -- grammars'
$delimText = "JobNo${tab}Project${tab}Folder${tab}Status`r`n100234${tab}ABC123${tab}Receive${tab}${jaNormal}${tab}2026/08/24 9:50:03${tab}1,204`r`n100235${tab}ABC124${tab}Receive${tab}${jaNormal}${tab}2026/08/24 10:02:11${tab}0`r`n`r`nTotal 2 rows`r`nsomething odd here"
$g = @{ parser = 'delimited'; delimiter = "`t"; rowWhen = @{ field = 0; matches = '^\d+$' }; fields = @('jobNo', 'key', 'folder', 'status', 'recvTime', 'count'); ignore = @('^Total ') }
$r = ConvertFrom-EbiGrammar -Text $delimText -Grammar $g
Assert-True ($r['ok'] -and @($r['records']).Count -eq 2 -and $r['records'][0]['key'] -eq 'ABC123' -and $r['records'][1]['recvTime'] -eq '2026/08/24 10:02:11' -and $r['records'][0]['_line'] -eq 2) 'delimited: data rows by rowWhen, fields named, line numbers kept'
Assert-True (@($r['unrecognized']).Count -eq 2 -and $r['unrecognized'][0]['line'] -eq 1 -and $r['unrecognized'][1]['text'] -eq 'something odd here') 'delimited: the header and the odd line are REPORTED; the ignored footer is not'
$g2 = @{ parser = 'delimited'; delimiter = "`t"; rowWhen = @{ field = 'jobNo'; matches = '^\d+$' }; fields = @('jobNo', 'key', 'folder', 'status', 'recvTime', 'count') }
$r = ConvertFrom-EbiGrammar -Text $delimText -Grammar $g2
Assert-True ($r['ok'] -and @($r['records']).Count -eq 2) 'delimited: rowWhen.field by name'
$r = ConvertFrom-EbiGrammar -Text "a b c`r`n1 2 3" -Grammar @{ parser = 'delimited'; delimiter = 'ws'; fields = @('x', 'y', 'z') }
Assert-True ($r['ok'] -and @($r['records']).Count -eq 2 -and $r['records'][1]['z'] -eq '3') 'delimited: whitespace delimiter, no rowWhen = every full line'
$r = ConvertFrom-EbiGrammar -Text 'x' -Grammar @{ parser = 'delimited' }
Assert-True (-not $r['ok']) 'delimited: fields required'

$labeledText = "Batch detail`r`n${jaStatus}: ${jaNormal}`r`n${jaEnd}   2026/08/24 9:51:02 (JST)`r`nfooter"
$g = @{ parser = 'labeled'; pairs = @{ status = @{ after = $jaStatus; take = 'line' }; endTime = @{ after = $jaEnd; take = 'line' }; endToken = @{ after = $jaEnd; take = 'token' }; nothing = @{ after = 'NOPE'; take = 'line' } } }
$r = ConvertFrom-EbiGrammar -Text $labeledText -Grammar $g
Assert-True ($r['ok'] -and @($r['records']).Count -eq 1 -and $r['records'][0]['status'] -eq $jaNormal) 'labeled: value after the label, colon stripped'
Assert-True ($r['records'][0]['endTime'] -eq '2026/08/24 9:51:02 (JST)' -and $r['records'][0]['endToken'] -eq '2026/08/24') 'labeled: line vs token'
Assert-True ($r['records'][0]['nothing'] -eq '' -and @($r['missing']) -contains 'nothing') 'labeled: a label not found -> blank + missing'
Assert-True (@($r['unrecognized']).Count -eq 2) 'labeled: lines no pair used are reported'
$r = ConvertFrom-EbiGrammar -Text "Label:`r`nvalue on next line" -Grammar @{ parser = 'labeled'; pairs = @{ v = @{ after = 'Label'; take = 'line' } } }
Assert-True ($r['ok'] -and $r['records'][0]['v'] -eq 'value on next line') 'labeled: an empty rest takes the next line'

$colText = "  ${jaFile}                          ${jaTime}            size`r`nABC123.260824.10515511.dat          2026/08/24 9:51:02   1.2 MB   ${jaRef}`r`nABC123.dat                          2026/08/24 8:12:00   1.2 MB   ${jaRef}`r`n--"
$g = @{ parser = 'columns'; headerLine = @{ contains = @($jaFile, $jaTime) }; columns = @{ name = @(0, 36); time = @(36, 55); size = @(57, 65) } }
$r = ConvertFrom-EbiGrammar -Text $colText -Grammar $g
Assert-True ($r['ok'] -and @($r['records']).Count -eq 2 -and $r['records'][0]['name'] -eq 'ABC123.260824.10515511.dat' -and $r['records'][0]['time'] -eq '2026/08/24 9:51:02' -and $r['records'][1]['size'] -eq '1.2 MB') 'columns: fixed offsets under the header'
Assert-True (@($r['unrecognized']).Count -eq 1 -and $r['unrecognized'][0]['text'] -eq '--') 'columns: the short ruler line is reported'
$r = ConvertFrom-EbiGrammar -Text 'no header here' -Grammar $g
Assert-True (-not $r['ok'] -and $r['message'] -like '*header*') 'columns: missing header is an error, not zero rows'

$g = @{ parser = 'regex'; pattern = '^(?<name>\S+)\s+(?<date>\d{4}/\d{2}/\d{2})\s+(?<time>\d{1,2}:\d{2}:\d{2})\s+(?<size>.+?)\s+' + $jaRef + '$' }
$r = ConvertFrom-EbiGrammar -Text $colText -Grammar $g
Assert-True ($r['ok'] -and @($r['records']).Count -eq 2 -and $r['records'][0]['time'] -eq '9:51:02' -and $r['records'][1]['name'] -eq 'ABC123.dat') 'regex: named groups are fields; the 9:51:02 row is NOT dropped'
Assert-True (@($r['unrecognized']).Count -eq 2) 'regex: header and ruler reported'
$r = ConvertFrom-EbiGrammar -Text 'x' -Grammar @{ parser = 'regex'; pattern = '(' }
Assert-True (-not $r['ok']) 'regex: a bad pattern is an error'
$r = ConvertFrom-EbiGrammar -Text 'x' -Grammar @{ parser = 'csv' }
Assert-True (-not $r['ok'] -and $r['message'] -like '*unknown parser*') 'unknown parser'

# --- 3. verify.parse_text step ------------------------------------------------------------
Write-Host '  -- verify.parse_text'
$g = @{ parser = 'delimited'; delimiter = "`t"; rowWhen = @{ field = 0; matches = '^\d+$' }; fields = @('jobNo', 'key', 'folder', 'status', 'recvTime', 'count') }
$ret = Step 'verify.parse_text' @{ text = $delimText; grammar = $g; maxWarnings = 20 } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['recordCount'] -eq 2 -and $ret['unrecognized'] -eq 3 -and (@($ret['names']) -contains 'ABC124')) 'parse_text: records, unrecognized count, names'
$w = @($ret['warnings'])
Assert-True ($w.Count -eq 3 -and $w[0]['code'] -eq 'unrecognized_line' -and $w[0]['data']['line'] -eq 1 -and $w[0]['message'] -like '*JobNo*') 'parse_text: every unrecognised line is a warning with line number and text'
$ret = Step 'verify.parse_text' @{ text = $delimText; grammar = $g; maxWarnings = 1 } (New-TestCtx)
Assert-True (@($ret['warnings']).Count -eq 2 -and $ret['warnings'][1]['code'] -eq 'unrecognized_lines' -and $ret['warnings'][1]['data']['total'] -eq 3) 'parse_text: beyond maxWarnings one summary carries the total'
$ret = Step 'verify.parse_text' @{ text = "header only`r`nstill loading"; grammar = $g; maxWarnings = 20 } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'no_records' -and @($ret['warnings']).Count -eq 2) 'parse_text: nothing recognised is no_records, with the lines still reported'
$ret = Step 'verify.parse_text' @{ text = 'x'; grammar = @{ parser = 'nope' }; maxWarnings = 20 } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'grammar_invalid') 'parse_text: bad grammar'
$ret = Step 'verify.parse_text' @{ text = $labeledText; grammar = @{ parser = 'labeled'; pairs = @{ status = @{ after = $jaStatus }; gone = @{ after = 'ZZZ' } } }; maxWarnings = 20 } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['recordCount'] -eq 1 -and (@($ret['warnings'] | Where-Object { $_['code'] -eq 'label_missing' }).Count -eq 1)) 'parse_text: labeled always one record; missing label is a warning'

# --- 4. verify.match_record --------------------------------------------------------------
Write-Host '  -- verify.match_record'
$recs = @(
    @{ key = 'ABC123.260824.10515511'; time = '2026/08/24 9:51:02'; _line = 2 }
    @{ key = 'DEF456'; time = '2026/08/24 9:52:00'; _line = 3 }
    @{ key = 'ABC123.260824.10515533'; time = '2026/08/24 9:53:40'; _line = 4 }
    @{ key = 'abc123'; time = ''; _line = 5 }
)
$ret = Step 'verify.match_record' @{ records = $recs; key = 'DEF456'; field = 'key'; tieBreak = 'newest'; timeField = 'time'; window = @{} } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['index'] -eq 2 -and $ret['found'] -eq 1 -and $ret['record']['_line'] -eq 3 -and $ret['matchedBy'] -eq 'exact') 'match: single exact hit, 1-based index'
$ret = Step 'verify.match_record' @{ records = $recs; key = 'ABC123'; field = 'key'; tieBreak = 'newest'; timeField = 'time'; window = @{} } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['found'] -eq 2 -and $ret['index'] -eq 3 -and (@($ret['warnings']).Count -eq 1)) 'match: two stamped reruns at the stripped tier; newest wins (the case variant is a worse tier, not counted); warned'
$ret = Step 'verify.match_record' @{ records = $recs; key = 'ABC123'; field = 'key'; tieBreak = 'newest'; timeField = 'time'; window = @{ from = '2026/08/24 9:50:00'; to = '2026/08/24 9:52:00' } } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['index'] -eq 1) 'match: newest INSIDE the run window beats a newer one outside'
$ret = Step 'verify.match_record' @{ records = $recs; key = 'ABC123'; field = 'key'; tieBreak = 'first'; timeField = 'time'; window = @{} } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['index'] -eq 1) 'match: first'
$ret = Step 'verify.match_record' @{ records = $recs; key = 'ABC123'; field = 'key'; tieBreak = 'none'; timeField = 'time'; window = @{} } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'ambiguous' -and @($ret['candidates']['candidates']).Count -eq 2 -and $ret['candidates']['candidates'][1]['evidence']['time'] -eq '2026/08/24 9:53:40') 'match: none -> ambiguous with the candidate shape and evidence'
$ret = Step 'verify.match_record' @{ records = $recs; key = 'ZZZ'; field = 'key'; tieBreak = 'newest'; timeField = 'time'; window = @{} } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'record_not_found') 'match: none'
$ret = Step 'verify.match_record' @{ records = @(@{ key = 'a'; time = 'x' }, @{ key = 'a'; time = 'y' }); key = 'a'; field = 'key'; tieBreak = 'newest'; timeField = 'time'; window = @{} } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['index'] -eq 1) 'match: no usable time -> first, not a crash'
$ret = Step 'verify.match_record' @{ records = @(@{ name = 'ABC123'; time = '2026/08/24 9:00:00' }, @{ name = 'ABC123'; time = '2026/08/24 9:05:00' }); key = 'ABC123'; field = 'name'; tieBreak = 'newest'; timeField = 'time'; window = @{} } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['index'] -eq 2) 'match: 9:05 beats 9:00 (single-digit hours compare as times, not as text)'

# --- 5. verify.assert ------------------------------------------------------------------------
Write-Host '  -- verify.assert'
$rec = @{ status = $jaNormal; rtncd = '0'; recvTime = '2026/08/24 9:50:03'; count = '1,204'; name = 'ABC123.dat'; blank = '' }
function Rule { param([string]$Field, [string]$Op, $Value = $null, [string]$Else = 'ng') $r = @{ field = $Field; op = $Op; else = $Else; message = ($Field + ' ' + $Op) }; if ($null -ne $Value) { $r['value'] = $Value }; return $r }
function Judge { param($Rules, $Record = $rec, [string]$Default = 'ok') return (Step 'verify.assert' @{ record = $Record; rules = @{ rules = @($Rules); default = $Default } } (New-TestCtx)) }
$ret = Judge @((Rule 'status' 'equals' $jaNormal), (Rule 'rtncd' 'equals' '0'), (Rule 'recvTime' 'within' @{ from = '2026/08/24 9:00:00'; to = '2026/08/24 10:00:00' }), (Rule 'count' 'present' $null 'unknown'))
Assert-True ($ret['ok'] -and $ret['code'] -eq 'ok' -and $ret['rule'] -eq 0 -and $ret['reason'] -eq '') 'assert: the section 5 example table passes'
$ret = Judge @((Rule 'status' 'equals' 'other'))
Assert-True ($ret['code'] -eq 'ng' -and $ret['rule'] -eq 1 -and $ret['field'] -eq 'status' -and $ret['reason'] -eq 'status equals') 'op equals fails -> else, reason = message'
Assert-Equal 'ok' (Judge @((Rule 'status' 'notEquals' 'other')))['code'] 'op notEquals'
Assert-Equal 'ng' (Judge @((Rule 'status' 'notEquals' $jaNormal)))['code'] 'op notEquals fails'
Assert-Equal 'ok' (Judge @((Rule 'rtncd' 'in' @('0', '4'))))['code'] 'op in'
Assert-Equal 'ng' (Judge @((Rule 'rtncd' 'in' @('8'))))['code'] 'op in fails'
Assert-Equal 'ok' (Judge @((Rule 'rtncd' 'notIn' @('8', '12'))))['code'] 'op notIn'
Assert-Equal 'ng' (Judge @((Rule 'rtncd' 'notIn' @('0'))))['code'] 'op notIn fails'
Assert-Equal 'ok' (Judge @((Rule 'name' 'matches' '^ABC\d+\.dat$')))['code'] 'op matches'
Assert-Equal 'ng' (Judge @((Rule 'name' 'matches' '^DEF')))['code'] 'op matches fails'
Assert-Equal 'ok' (Judge @((Rule 'count' 'present')))['code'] 'op present'
Assert-Equal 'unknown' (Judge @((Rule 'blank' 'present' $null 'unknown')))['code'] 'op present fails -> unknown'
Assert-Equal 'ok' (Judge @((Rule 'blank' 'empty')))['code'] 'op empty'
Assert-Equal 'ng' (Judge @((Rule 'count' 'empty')))['code'] 'op empty fails'
Assert-Equal 'ok' (Judge @((Rule 'recvTime' 'within' @{ from = '2026/08/24 9:50:03'; to = '2026/08/24 9:50:03' })))['code'] 'op within: endpoints inclusive'
Assert-Equal 'ng' (Judge @((Rule 'recvTime' 'within' @{ from = '2026/08/24 10:00:00'; to = '2026/08/24 11:00:00' })))['code'] 'op within fails'
Assert-Equal 'ng' (Judge @((Rule 'blank' 'within' @{ from = '2026/08/24 9:00:00'; to = '2026/08/24 11:00:00' })))['code'] 'op within: unreadable time never passes'
Assert-Equal 'ok' (Judge @((Rule 'count' 'gt' '1000')))['code'] 'op gt (comma-grouped number)'
Assert-Equal 'ng' (Judge @((Rule 'count' 'gt' '2000')))['code'] 'op gt fails'
Assert-Equal 'ok' (Judge @((Rule 'count' 'lt' '2000')))['code'] 'op lt'
Assert-Equal 'ok' (Judge @((Rule 'count' 'gte' '1204')))['code'] 'op gte'
Assert-Equal 'ok' (Judge @((Rule 'count' 'lte' '1204')))['code'] 'op lte'
Assert-Equal 'ng' (Judge @((Rule 'name' 'gt' '1')))['code'] 'numeric op on a non-number never passes'
Assert-Equal 'unknown' (Judge @() -Default 'unknown')['code'] 'default applies when no rule decides'
$ret = Judge @((Rule 'status' 'equals' 'x' 'ok'))
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'rules_invalid' -and $ret['message'] -like '*never ok*') 'else: ok is REFUSED'
$ret = Judge @((Rule 'status' 'startsWith' 'x'))
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'rules_invalid') 'a thirteenth op is refused'
$ret = Judge @(@{ field = 'status'; op = 'equals'; else = 'ng' })
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'rules_invalid') 'equals without a value is refused'
$ret = Step 'verify.assert' @{ record = $rec; rules = @{ rules = @(); default = 'maybe' } } (New-TestCtx)
Assert-True (-not $ret['ok']) 'a bad default is refused'
$ret = Judge @((Rule 'status' 'equals' 'x'), (Rule 'blank' 'present' $null 'unknown'))
Assert-True ($ret['code'] -eq 'ng' -and $ret['rule'] -eq 1) 'the FIRST failing rule decides'

# --- 6. human.* on Gate.ps1 -------------------------------------------------------------------
Write-Host '  -- human.gate / human.choose / human.prepare'
$d = HumanGate-Decide -Code 'ok' -AskWhen @('unknown') -Answer ''
Assert-True ($d['action'] -eq 'pass' -and $d['code'] -eq 'ok') 'gate decide: not in askWhen -> pass'
$d = HumanGate-Decide -Code 'unknown' -AskWhen @('unknown') -Answer 'o'
Assert-True ($d['action'] -eq 'ok' -and $d['code'] -eq 'ok') 'gate decide: o -> ok'
$d = HumanGate-Decide -Code 'unknown' -AskWhen @('unknown') -Answer 'n'
Assert-True ($d['action'] -eq 'ng' -and $d['code'] -eq 'ng') 'gate decide: n -> ng'
$d = HumanGate-Decide -Code 'unknown' -AskWhen @('unknown') -Answer 's'
Assert-True ($d['action'] -eq 'skip' -and $d['code'] -eq '') 'gate decide: s -> skip, code empty'
$d = HumanGate-Decide -Code 'ng' -AskWhen @('unknown', 'ng') -Answer 'k'
Assert-True ($d['action'] -eq 'keep' -and $d['code'] -eq 'ng') 'gate decide: k keeps'
$ret = Step 'human.gate' @{ code = 'ok'; askWhen = @('unknown'); reason = ''; evidence = ''; key = '' } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['action'] -eq 'pass' -and $ret['code'] -eq 'ok') 'gate step: ok passes through, no panel'
$ctx = New-TestCtx -DryRun $true
$ret = Step 'human.gate' @{ code = 'unknown'; askWhen = @('unknown'); reason = 'count unreadable'; evidence = @('a.png'); key = 'K1' } $ctx
Assert-True ($ret['ok'] -and $ret['action'] -eq 'keep' -and $ret['code'] -eq 'unknown' -and $ctx['Log'].Lines.Count -eq 1) 'gate step: dry run keeps the verdict and says so'
# a scripted reader: answers in order
$script:answers = New-Object System.Collections.ArrayList
function Set-Answers { param($List) $script:answers.Clear(); foreach ($a in @($List)) { [void]$script:answers.Add($a) } }
$reader = { if ($script:answers.Count -eq 0) { return 'q' }; $a = $script:answers[0]; $script:answers.RemoveAt(0); return $a }
Set-Answers @('m looks fine to me', '')
$r = Show-EbiGate -Title 'T' -What 'w' -Actions (HumanGate-Actions) -Default 'o' -Auto 'k' -Reader $reader
Assert-True ($r['action'] -eq 'm' -and $r['note'] -eq 'looks fine to me') 'gate panel: m <note> comes back as action m with the note'
$r = Show-EbiGate -Title 'T' -What 'w' -Actions (HumanGate-Actions) -Default 'o' -Auto 'k' -Reader $reader
Assert-True ($r['action'] -eq 'o' -and -not $r['auto']) 'gate panel: Enter is o'
Set-Answers @('3', 'y')
$shape = @{ candidates = @(@{ id = 'c1'; candidate = 'A.dat'; evidence = @{ time = '9:51' } }, @{ id = 'c2'; candidate = 'B.dat'; evidence = @{ time = '9:53' } }, @{ id = 'c3'; candidate = 'C.dat'; evidence = @{} }); suggestion = @{ id = 'c2'; reason = 'newest' }; doubts = 'close together' }
$lines = @(HumanChoose-Lines -Shape $shape)
Assert-True ($lines.Count -eq 3 -and $lines[0] -eq '#1 A.dat   [time=9:51]' -and $lines[2] -like '#3 C.dat*no evidence*') 'choose lines: every candidate with its evidence, missing evidence said so'
Assert-Equal 2 (HumanChoose-SuggestIndex -Shape $shape) 'choose: suggestion index'
$ctx = New-TestCtx -DryRun $true
$ret = Step 'human.choose' @{ candidates = $shape; question = 'which?'; key = 'K' } $ctx
Assert-True ($ret['ok'] -and $ret['action'] -eq 'chosen' -and $ret['index'] -eq 2 -and $ret['id'] -eq 'c2' -and $ret['candidate'] -eq 'B.dat' -and -not $ret['learn']) 'choose step: dry run takes the suggestion, never asks to learn'
$ret = Step 'human.choose' @{ candidates = @{ candidates = @(@{ id = 'c1'; candidate = 'x'; evidence = @{} }) }; question = 'q'; key = '' } (New-TestCtx -DryRun $true)
Assert-True ($ret['ok'] -and $ret['action'] -eq 'none') 'choose step: no suggestion, dry run -> none'
$ret = Step 'human.choose' @{ candidates = @{ candidates = @() }; question = 'q'; key = '' } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['action'] -eq 'none' -and @($ret['warnings']).Count -eq 1) 'choose step: empty list -> none with a warning'
$ret = Step 'human.choose' @{ candidates = 'nope'; question = 'q'; key = '' } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'input_invalid') 'choose step: not the shape'
$acts = @(1..3 | ForEach-Object { @{ key = [string]$_; label = ('#' + $_) } }) + @(@{ key = 'n'; label = 'none' }, @{ key = 's'; label = 'skip' }, @{ key = 'q'; label = 'quit' })
$r = Show-EbiGate -Title 'C' -What 'w' -Actions $acts -Default '' -Auto '2' -Reader $reader
Assert-True ($r['action'] -eq '3') 'choose panel: a number is accepted as an action'
$ret = Step 'human.prepare' @{ message = 'ready?'; url = 'http://x' } (New-TestCtx -DryRun $true)
Assert-True ($ret['ok'] -and $ret['action'] -eq 'enter') 'prepare: dry run answers Enter'
$r = Show-EbiGate -Title 'P' -What 'w' -Actions (HumanPrepare-Actions) -Default 'c' -Auto 'c' -Reader { 'q' }
Assert-True ($r['action'] -eq 'q') 'prepare panel: q is an answer'

$failed = Complete-Tests
exit $failed
