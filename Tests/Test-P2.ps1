# Test-P2.ps1 -- P2-07 human.input + run.timeWindow, P2-08 mask-lite,
# P2-10 verify.crosscheck. ASCII source; no param() block.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_TestCommon.ps1')
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'kernel/Registry.ps1')
. (Join-Path $repoRoot 'kernel/Runner.ps1')
. (Join-Path $repoRoot 'kernel/Mask.ps1')
. (Join-Path $repoRoot 'kernel/Table.ps1')

Reset-Tests 'P2'
$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-p2-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding($false)

function New-TestLog {
    $log = New-Object PSObject
    $log | Add-Member -MemberType NoteProperty -Name Lines -Value (New-Object System.Collections.ArrayList)
    $log | Add-Member -MemberType ScriptMethod -Name Info  -Value { param($m) [void]$this.Lines.Add('info:' + $m) }
    $log | Add-Member -MemberType ScriptMethod -Name Warn  -Value { param($m) [void]$this.Lines.Add('warn:' + $m) }
    $log | Add-Member -MemberType ScriptMethod -Name Debug -Value { param($m) }
    return $log
}
function New-TestCtx { param([bool]$DryRun = $false, $Reader = $null) return @{ WorkDir = $tmpRoot; RunId = 'r1'; Profile = @{}; Log = (New-TestLog); DryRun = $DryRun; Session = @{}; Item = $null; KeyColumns = @(); Run = @{ runId = 'r1'; timeWindow = $null }; Reader = $Reader } }
$reg = New-EbiRegistry -ModulesRoot (Join-Path $repoRoot 'modules')
foreach ($u in @('human.input', 'verify.crosscheck')) { $r = . Import-EbiStep -Registry $reg -Use $u; Assert-True $r['ok'] ('loads: ' + $u + ' ' + $r['message']) }
function Step { param([string]$Use, [hashtable]$With, $Ctx) return (& (Get-EbiStep -Registry $reg -Use $Use)['Invoke'] $With $Ctx) }

# --- 1. human.input ---------------------------------------------------------------------
Write-Host '  -- human.input'
$now = [datetime]'2026-06-12T10:00:00'
$p = HumanInput-Parse -Answer '9:00..12:00' -Kind 'timeWindow' -Format 'yyyy/MM/dd H:mm:ss' -Now $now
Assert-True ($p['ok'] -and $p['timeWindow']['from'] -eq '2026-06-12T09:00:00' -and $p['timeWindow']['to'] -eq '2026-06-12T12:00:00') 'input parse: "9:00..12:00" is today'
$p = HumanInput-Parse -Answer '2026/06/12 9:00:00..2026/06/12 8:00:00' -Kind 'timeWindow' -Format 'x' -Now $now
Assert-True (-not $p['ok'] -and $p['message'] -like '*before*') 'input parse: to before from is refused'
$p = HumanInput-Parse -Answer 'noon' -Kind 'timeWindow' -Format 'x' -Now $now
Assert-True (-not $p['ok']) 'input parse: junk window'
$p = HumanInput-Parse -Answer '9:50' -Kind 'time' -Format 'yyyy/MM/dd H:mm:ss' -Now $now
Assert-True ($p['ok'] -and $p['value'] -eq '2026/06/12 9:50:00') 'input parse: a bare time is today, stored in the format'
$p = HumanInput-Parse -Answer '  free text ' -Kind 'text' -Format 'x' -Now $now
Assert-True ($p['ok'] -and $p['value'] -eq 'free text') 'input parse: text'
Assert-Equal '2026/06/12 9:00:00..2026/06/12 10:00:00' (HumanInput-Default -Default '' -Kind 'timeWindow' -Now $now) 'input default: the last hour'
Assert-Equal 'given' (HumanInput-Default -Default 'given' -Kind 'text' -Now $now) 'input default: an explicit default wins'

$ctx = New-TestCtx -DryRun $true
$ret = Step 'human.input' @{ question = 'window?'; kind = 'timeWindow'; default = '2026/06/12 9:00..2026/06/12 12:00'; persistTo = ''; format = 'yyyy/MM/dd H:mm:ss' } $ctx
Assert-True ($ret['ok'] -and $ret['auto'] -and $ret['timeWindow']['from'] -eq '2026-06-12T09:00:00' -and $ctx['Run']['timeWindow']['to'] -eq '2026-06-12T12:00:00') 'input step: dry run takes the default and sets $Ctx.Run.timeWindow'
$script:answers = New-Object System.Collections.ArrayList
$reader = { if ($script:answers.Count -eq 0) { return 'q' }; $a = $script:answers[0]; $script:answers.RemoveAt(0); return $a }
[void]$script:answers.Add('nonsense'); [void]$script:answers.Add('9:00..11:00')
$ctx = New-TestCtx -Reader $reader
$ret = Step 'human.input' @{ question = 'window?'; kind = 'timeWindow'; default = ''; persistTo = ''; format = 'yyyy/MM/dd H:mm:ss' } $ctx
Assert-True ($ret['ok'] -and -not $ret['auto'] -and $ret['timeWindow']['from'] -like '*T09:00:00' -and $ret['timeWindow']['to'] -like '*T11:00:00') 'input step: a bad answer is asked again, the next good one is taken'
# run.timeWindow already set (a resume restored it, or the CLI gave -TimeWindow): kept, nobody asked
$script:asked = 0
$askReader = { $script:asked++; return '2026/06/12 13:00..2026/06/12 14:00' }
$ctx = New-TestCtx -Reader $askReader
$ctx['Run']['timeWindow'] = @{ from = '2026-06-12T08:00:00'; to = '2026-06-12T09:00:00' }
$ret = Step 'human.input' @{ question = 'window?'; kind = 'timeWindow'; default = ''; persistTo = ''; format = 'yyyy/MM/dd H:mm:ss' } $ctx
Assert-True ($ret['ok'] -and $ret['kept'] -and -not $ret['auto'] -and $script:asked -eq 0 -and $ret['timeWindow']['from'] -eq '2026-06-12T08:00:00' -and $ret['value'] -eq '2026-06-12T08:00:00..2026-06-12T09:00:00' -and $ctx['Run']['timeWindow']['to'] -eq '2026-06-12T09:00:00') 'input step: run.timeWindow already set -> kept as is, not asked, not overwritten'
$ctx = New-TestCtx -DryRun $true
$ctx['Run']['timeWindow'] = @{ from = '2026-06-12T08:00:00'; to = '2026-06-12T09:00:00' }
$ret = Step 'human.input' @{ question = 'window?'; kind = 'timeWindow'; default = ''; persistTo = ''; format = 'yyyy/MM/dd H:mm:ss' } $ctx
Assert-True ($ret['ok'] -and $ret['kept'] -and $ret['timeWindow']['from'] -eq '2026-06-12T08:00:00') 'input step: ... also with nobody to ask (the last-hour default does not replace it)'
$ret = Step 'human.input' @{ question = 'x'; kind = 'text'; default = 'dflt'; persistTo = ''; format = '' } $ctx
Assert-True ($ret['ok'] -and -not $ret['kept'] -and $ret['value'] -eq 'dflt') 'input step: a text question is not the window: still asked (dry run takes the default)'
$script:answers.Clear(); [void]$script:answers.Add('q')
$ret = Step 'human.input' @{ question = 'x'; kind = 'text'; default = ''; persistTo = ''; format = '' } (New-TestCtx -Reader $reader)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'operator_quit') 'input step: q quits'
$script:answers.Clear(); [void]$script:answers.Add('')
$ret = Step 'human.input' @{ question = 'x'; kind = 'text'; default = 'dflt'; persistTo = ''; format = '' } (New-TestCtx -Reader $reader)
Assert-True ($ret['ok'] -and $ret['value'] -eq 'dflt') 'input step: Enter takes the default'
# persistTo fills blanks only, and flushes
$csv = Join-Path $tmpRoot 'wl.csv'
[void](Write-EbiCsvAtomic -Path $csv -Columns @('K', 'Expected_Time') -Rows @(@{ K = 'a'; Expected_Time = '' }, @{ K = 'b'; Expected_Time = '2026/06/01 8:00:00' }))
$wl = New-EbiWorklist -Path $csv -Columns @('K', 'Expected_Time') -Rows (Read-EbiCsv -Path $csv)['rows']
$script:answers.Clear(); [void]$script:answers.Add('9:30')
$ret = Step 'human.input' @{ question = 'when?'; kind = 'time'; default = ''; persistTo = 'Expected_Time'; worklist = $wl; format = 'yyyy/MM/dd H:mm:ss' } (New-TestCtx -Reader $reader)
Assert-True ($ret['ok'] -and $ret['filled'] -eq 1 -and $wl['rows'][0]['Expected_Time'] -like '* 9:30:00' -and $wl['rows'][1]['Expected_Time'] -eq '2026/06/01 8:00:00') 'input step: persistTo fills only the blank cell'
Assert-True ((Read-EbiCsv -Path $csv)['rows'][0]['Expected_Time'] -like '* 9:30:00') 'input step: the table was flushed'

# the runner: -TimeWindow reaches run.timeWindow, run.json, and a resume keeps it
Write-Host '  -- run.timeWindow through the runner'
$modules = Join-Path $tmpRoot 'modules'
New-Item -ItemType Directory -Path (Join-Path $modules 'fake') -Force | Out-Null
[System.IO.File]::WriteAllText((Join-Path $modules 'fake/fake.echo.ps1'), @'
$Manifest = @{ id = 'fake.echo'; group = 'fake'; summary = 'fixture'; tier = 'core'; effects = 'pure'; needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs = @{ w = @{ type = 'any' } }; outputs = @{ w = @{ type = 'any' } }; failures = @( @{ id = 'never'; transient = $false } ); example = @{ use = 'fake.echo'; with = @{ w = '{{run.timeWindow}}' } } }
function Invoke-Step { param($In, $Ctx) return @{ ok = $true; w = $In['w'] } }
'@, $utf8)
$wfPath = Join-Path $tmpRoot 'tw.json'
[System.IO.File]::WriteAllText($wfPath, '{ "schema": 1, "id": "p2.tw", "setup": [ { "id": "e", "use": "fake.echo", "with": { "w": "{{run.timeWindow}}" } } ] }', $utf8)
$wd = Join-Path $tmpRoot 'twwork'
$res = Invoke-EbiWorkflow -Path $wfPath -WorkDir $wd -ModulesRoot $modules -RunId 'tw1' -TimeWindow @{ from = '2026-06-12T09:00:00'; to = '2026-06-12T12:00:00' }
Assert-True ($res['ok'] -and $res['steps'][0]['outputs']['w']['from'] -eq '2026-06-12T09:00:00') 'runner: -TimeWindow is run.timeWindow for the steps'
$runDoc = (Read-EbiRunFile -WorkDir $wd -RunId 'tw1')['value']
Assert-True ($runDoc['timeWindow']['to'] -eq '2026-06-12T12:00:00' -and $runDoc['args']['timeWindow']['from'] -eq '2026-06-12T09:00:00') 'runner: run.json carries the window'
$res = Invoke-EbiWorkflow -Path $wfPath -WorkDir $wd -ModulesRoot $modules -RunId 'tw1' -Resume
Assert-True ($res['ok'] -and $res['steps'][0]['outputs']['w']['from'] -eq '2026-06-12T09:00:00') 'runner: a resume restores the window from run.json without asking'
# human.input in setup persists the window before the next setup step sees it
[System.IO.File]::WriteAllText((Join-Path $tmpRoot 'tw2.json'), '{ "schema": 1, "id": "p2.tw2", "setup": [ { "id": "ask", "use": "human.input", "with": { "question": "w?", "kind": "timeWindow", "default": "2026/06/12 8:00..2026/06/12 9:00" } }, { "id": "e", "use": "fake.echo", "with": { "w": "{{run.timeWindow}}" } } ] }', $utf8)
# the fixture modules tree gets the real human.input step, and a copy of kernel/ beside it
# (the step dot-sources ..\..\kernel relative to its own file). The directory is created
# first: on PS 5.1 a Copy-Item into a missing directory is a TERMINATING error that
# -ErrorAction SilentlyContinue does not swallow (it did on pwsh 7, which hid this).
New-Item -ItemType Directory -Path (Join-Path $modules 'human') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $repoRoot 'modules/human/human.input.ps1') -Destination (Join-Path $modules 'human/human.input.ps1') -Force
Copy-Item -LiteralPath (Join-Path $repoRoot 'kernel') -Destination (Join-Path $tmpRoot 'kernel') -Recurse -Force
$res = Invoke-EbiWorkflow -Path (Join-Path $tmpRoot 'tw2.json') -WorkDir $wd -ModulesRoot $modules -RunId 'tw2' -DryRun
Assert-True ($res['ok'] -and $res['steps'][1]['outputs']['w']['from'] -eq '2026-06-12T08:00:00') 'runner: human.input in setup sets run.timeWindow for the next setup step (dry run takes the default)'
Assert-True ((Read-EbiRunFile -WorkDir $wd -RunId 'tw2')['value']['timeWindow']['from'] -eq '2026-06-12T08:00:00') 'runner: ... and run.json was rewritten at once'
# a resume runs setup again (the ledger holds only "each"), so human.input runs again: it must
# keep the 08:00..09:00 that run.json carries even though the workflow's default changed and
# nobody can answer (a dry run would otherwise take the new default)
[System.IO.File]::WriteAllText((Join-Path $tmpRoot 'tw2.json'), '{ "schema": 1, "id": "p2.tw2", "setup": [ { "id": "ask", "use": "human.input", "with": { "question": "w?", "kind": "timeWindow", "default": "2026/06/12 10:00..2026/06/12 11:00" } }, { "id": "e", "use": "fake.echo", "with": { "w": "{{run.timeWindow}}" } } ] }', $utf8)
$res = Invoke-EbiWorkflow -Path (Join-Path $tmpRoot 'tw2.json') -WorkDir $wd -ModulesRoot $modules -RunId 'tw2' -DryRun -Resume
Assert-True ($res['ok'] -and $res['steps'][0]['outputs']['kept'] -and $res['steps'][1]['outputs']['w']['from'] -eq '2026-06-12T08:00:00') 'runner: a resumed run keeps the window human.input persisted and does not ask again'
Assert-True ((Read-EbiRunFile -WorkDir $wd -RunId 'tw2')['value']['timeWindow']['from'] -eq '2026-06-12T08:00:00') 'runner: ... and run.json still carries it'
$res = Invoke-EbiWorkflow -Path (Join-Path $tmpRoot 'tw2.json') -WorkDir $wd -ModulesRoot $modules -RunId 'tw2' -DryRun -Resume -TimeWindow @{ from = '2026-06-12T07:00:00'; to = '2026-06-12T07:30:00' }
Assert-True ($res['ok'] -and $res['steps'][0]['outputs']['kept'] -and $res['steps'][1]['outputs']['w']['from'] -eq '2026-06-12T07:00:00') 'runner: ... unless the CLI gives -TimeWindow, which wins and is also kept'

# --- 2. verify.crosscheck ----------------------------------------------------------------
Write-Host '  -- verify.crosscheck'
function Cross { param($Readings, [string]$Compare = 'equal', [int]$Tol = 0, $Norm = @()) return (Step 'verify.crosscheck' @{ readings = $Readings; compare = $Compare; toleranceSec = $Tol; normalize = $Norm } (New-TestCtx)) }
$ret = Cross @(@{ source = 'page'; value = '00:00:01' }, @{ source = 'log'; value = '00:00:01' })
Assert-True ($ret['ok'] -and $ret['code'] -eq 'ok' -and @($ret['disagreements']).Count -eq 0 -and $ret['sources'] -eq 2) 'crosscheck: two agree'
$ret = Cross @(@{ source = 'page'; value = '00:00:01' }, @{ source = 'ocr'; value = '00:00:07' }, @{ source = 'log'; value = '00:00:01' })
Assert-True ($ret['code'] -eq 'unknown' -and @($ret['disagreements']).Count -eq 2 -and $ret['disagreements'][0]['b'] -eq 'ocr' -and $ret['disagreements'][0]['bValue'] -eq '00:00:07') 'crosscheck: one of three disagrees -> unknown, no majority vote, both pairs named'
$ret = Cross @(@{ source = 'page'; value = 'x' })
Assert-True ($ret['code'] -eq 'ok' -and @($ret['warnings']).Count -eq 1 -and $ret['warnings'][0]['code'] -eq 'single_source') 'crosscheck: a single source is ok with single_source'
$ret = Cross @(@{ source = 'a'; value = '1,204' }, @{ source = 'b'; value = '1204.0' }) 'numericEqual'
Assert-Equal 'ok' $ret['code'] 'crosscheck: numericEqual ignores grouping and .0'
$ret = Cross @(@{ source = 'a'; value = '1,204' }, @{ source = 'b'; value = 'n/a' }) 'numericEqual'
Assert-Equal 'unknown' $ret['code'] 'crosscheck: numericEqual with an unreadable value is a disagreement'
$ret = Cross @(@{ source = 'a'; value = '2026/06/12 9:50:03' }, @{ source = 'b'; value = '2026/06/12 09:50:40' }) 'timeWithinSec' 60
Assert-Equal 'ok' $ret['code'] 'crosscheck: timeWithinSec inside tolerance (single-digit hour too)'
$ret = Cross @(@{ source = 'a'; value = '2026/06/12 9:50:03' }, @{ source = 'b'; value = '2026/06/12 9:52:03' }) 'timeWithinSec' 60
Assert-Equal 'unknown' $ret['code'] 'crosscheck: timeWithinSec outside tolerance'
$fw = [string][char]0xFF11
$ret = Cross @(@{ source = 'a'; value = ' 1,204 ' }, @{ source = 'b'; value = ($fw + '204') }) 'equal' 0 @('trim', 'fullwidth', 'thousands')
Assert-Equal 'ok' $ret['code'] 'crosscheck: trim + fullwidth + thousands make the texts equal'
$ret = Cross @(@{ source = 'a'; value = ' 1,204 ' }, @{ source = 'b'; value = '1204' }) 'equal'
Assert-Equal 'unknown' $ret['code'] 'crosscheck: without normalize the same texts differ'
$ret = Step 'verify.crosscheck' @{ readings = @(); compare = 'equal'; toleranceSec = 0; normalize = @() } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'input_invalid') 'crosscheck: no readings'

# --- 3. mask-lite ------------------------------------------------------------------------
Write-Host '  -- mask check'
$bad = @"
employee JP123456 opened the file
mail me at taro.yamada@company.co.jp please
copy from \\fileserver01\share\evidence
saved under C:\Users\jp123456\Desktop
see http://portal.corp/apps/mq and http://10.20.30.40/x
also the code name PROJECT-FOXTROT appears here
clean line
"@
$hits = @(Find-EbiMaskHitsInText -Text $bad -Words @('PROJECT-FOXTROT'))
$rules = @($hits | ForEach-Object { $_['rule'] })
foreach ($r in @('employee_id', 'email', 'unc_path', 'user_home', 'intranet_url', 'private_ip', 'dictionary')) { Assert-True ($rules -contains $r) ('mask: rule ' + $r + ' fires') }
$emailHits = @($hits | Where-Object { $_['rule'] -eq 'email' })
Assert-True ($emailHits.Count -eq 1 -and $emailHits[0]['line'] -eq 2) 'mask: hits carry the line number'
$hits = @(Find-EbiMaskHitsInText -Text 'see http://portal.corp/apps' -Allow @('^http://portal\.corp'))
Assert-Equal 0 $hits.Count 'mask: an allow pattern silences a known-safe hit'
Assert-Equal 0 @(Find-EbiMaskHitsInText -Text "JIDSK05S`t2026/06/12 07:54:22`tA2009999A0000000001").Count 'mask: fixture-like text (correl ids, msg ids) is clean'
$badFile = Join-Path $tmpRoot 'leak.txt'; [System.IO.File]::WriteAllText($badFile, "contact JP654321`n", $utf8)
$okFile = Join-Path $tmpRoot 'fine.txt'; [System.IO.File]::WriteAllText($okFile, "nothing here`n", $utf8)
$res = Invoke-EbiMaskCheck -RepoRoot $repoRoot -Files @($badFile, $okFile) -DictionaryPath (Join-Path $tmpRoot 'nodict.json')
Assert-True (-not $res['ok'] -and @($res['hits']).Count -eq 1 -and $res['hits'][0]['file'] -eq $badFile -and $res['files'] -eq 2) 'mask check: one hit in one of two files fails the gate'
Assert-True ((@(Format-EbiMaskReport -Result $res -RepoRoot $tmpRoot) -join "`n") -like '*[[]MASK] leak.txt:1  employee_id  "JP654321"*FAIL*') 'mask report: file:line rule match'
$res = Invoke-EbiMaskCheck -RepoRoot $repoRoot
Assert-True ($res['ok'] -and $res['files'] -ge 15) ('mask check: the repo''s profiles and workflows are clean (' + $res['message'] + ')')
$dict = Read-EbiMaskDictionary -Path (Join-Path (Join-Path $repoRoot 'profiles') 'mask-dictionary.json')
Assert-True (@($dict['allow']).Count -ge 1) 'mask dictionary: loads'

Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
$failed = Complete-Tests
exit $failed
