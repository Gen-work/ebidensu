#Requires -Version 5.1
# Test-Dashboard.ps1 -- kernel/Dashboard.ps1: the day dashboard behind
# `ebi status` and the menu behind `ebi menu`. The pure half (phases,
# columns, model, totals, next, run line, rendering, menu answers) runs on
# hand-built fixtures; the loader, the one write (clear a verdict) and the
# menu loop run on a temp tool dir + work dir; ebi.ps1 status / menu run in
# a child PowerShell of the same executable (ebi.ps1 calls exit).

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Dashboard.ps1')

Reset-Tests 'Dashboard'

$utf8 = New-Object System.Text.UTF8Encoding($false)
$verdict = @{ ok = 'ok'; ng = 'ng'; unknown = ''; pending = '' }
$profile = @{
    name = 'demo'
    worklist = @{
        file = 'demo_{{run.operator}}.csv'
        key = @{ columns = @('Excel_NAME') }
        dashboard = @{ columns = @('JOB', 'GFIX_TIME', 'NotInCsv') }
        columns = @(
            @{ name = 'Excel_NAME'; role = 'key' },
            @{ name = 'JOB'; role = 'text' },
            @{ name = 'USER'; role = 'text' },
            @{ name = 'GFIX_TIME'; role = 'time' },
            @{ name = 'track'; role = 'verdict'; values = $verdict },
            @{ name = 'logs'; role = 'verdict'; values = $verdict },
            @{ name = 'evidence'; role = 'verdict'; values = $verdict },
            @{ name = 'note'; role = 'text' }
        )
    }
}
$csvColumns = @('Excel_NAME', 'JOB', 'USER', 'GFIX_TIME', 'track', 'logs', 'evidence', 'note')
function New-Wf { param([string]$Id, [string]$Field, [string]$Prof = 'demo', [string]$Title = 't')
    $v = @{ schema = 1; id = $Id; title = $Title; profile = $Prof }
    if ($Field -ne '') { $v['source'] = @{ table = 'wl'; select = @{ field = $Field; pendingWhen = '!= ok' } } }
    return @{ path = ('C:\tool\workflows\' + $Id + '.json'); value = $v }
}
$workflows = @(
    (New-Wf 'demo.evidence' 'evidence'),
    (New-Wf 'demo.track' 'track'),
    (New-Wf 'other.track' 'track' 'other'),
    (New-Wf 'demo.plan' ''),
    (New-Wf 'demo.logs' 'logs'),
    (New-Wf 'demo.extra' 'nosuchcolumn'),
    @{ path = 'C:\tool\workflows\broken.json'; value = @{ schema = 1 } }
)

# ---------------------------------------------------------------- phases
$phases = @(Get-EbiDayPhases -Workflows $workflows -ProfileName 'demo' -Columns $profile['worklist']['columns'])
Assert-Equal 'plan,track,logs,evidence,extra' (@($phases | ForEach-Object { $_['name'] }) -join ',') 'phases: plan (no source) first, then by the verdict column order, an undeclared field last'
Assert-Equal 'demo.plan' $phases[0]['id'] 'phases: id kept'
Assert-Equal 'demo.plan.json' $phases[0]['file'] 'phases: file = the leaf name'
Assert-Equal '' $phases[0]['field'] 'phases: plan has no field'
Assert-Equal '!= ok' $phases[1]['pendingWhen'] 'phases: pendingWhen kept'
Assert-True (-not (@($phases | ForEach-Object { $_['id'] }) -contains 'other.track')) 'phases: another profile''s workflow is left out'
$vp = @(Get-EbiDashboardVerdictPhases -Phases $phases)
Assert-Equal 'track,logs,evidence,nosuchcolumn' (@($vp | ForEach-Object { $_['field'] }) -join ',') 'verdict phases: the ones with a field'
$phases = @($phases | Where-Object { $_['name'] -ne 'extra' })   # the rest of the file uses the three real columns
$vp = @(Get-EbiDashboardVerdictPhases -Phases $phases)

# ---------------------------------------------------------------- columns
$show = @(Get-EbiDashboardColumns -Profile $profile -Columns $csvColumns -Exclude @('Excel_NAME', 'track', 'logs', 'evidence'))
Assert-Equal 'JOB,GFIX_TIME' ($show -join ',') 'columns: worklist.dashboard.columns, a name the CSV lacks dropped'
$noDash = @{ name = 'demo'; worklist = @{ key = $profile['worklist']['key']; columns = $profile['worklist']['columns'] } }
$show2 = @(Get-EbiDashboardColumns -Profile $noDash -Columns $csvColumns -Exclude @('Excel_NAME', 'track', 'logs', 'evidence'))
Assert-Equal 'JOB,USER,GFIX_TIME,note' ($show2 -join ',') 'columns: no dashboard section -> every non-key non-verdict column in declared order'
$show3 = @(Get-EbiDashboardColumns -Profile $null -Columns @('a', 'b', 'k') -Exclude @('k'))
Assert-Equal 'a,b' ($show3 -join ',') 'columns: no profile -> the CSV columns minus the excluded'

# ---------------------------------------------------------------- model / totals / next
function New-Row { param([string]$K, [string]$T, [string]$L, [string]$E) return @{ Excel_NAME = $K; JOB = ($K + 'J'); USER = 'u'; GFIX_TIME = '10:00'; track = $T; logs = $L; evidence = $E; note = '' } }
$rows = @((New-Row 'A1' 'ok' 'ok' 'ok'), (New-Row 'A2' 'ok' 'ok' 'ng'), (New-Row 'A3' 'ok' '' ''), (New-Row 'A4' '' '' ''), (New-Row 'A5' 'weird' '' ''))
$model = @(Get-EbiDashboardModel -Rows $rows -Phases $phases -Profile $profile -KeyColumns @('Excel_NAME') -ShowColumns $show)
Assert-Equal 5 $model.Count 'model: one record per row'
Assert-Equal 'A1' $model[0]['key'] 'model: key display'
Assert-Equal 'A1J,10:00' (@($model[0]['cells']) -join ',') 'model: cells follow the shown columns'
Assert-True ($null -eq $model[0]['next']) 'model: every verdict ok -> no next'
Assert-Equal 'ok' $model[0]['verdicts']['evidence'] 'model: ok'
Assert-Equal 'ng' $model[1]['verdicts']['evidence'] 'model: ng'
Assert-Equal 'clear' $model[1]['next']['action'] 'model: an ng cell -> clear and rerun'
Assert-Equal 'evidence' $model[1]['next']['field'] 'model: ... that column'
Assert-Equal 'run' $model[2]['next']['action'] 'model: a blank cell -> run'
Assert-Equal 'logs' $model[2]['next']['field'] 'model: the FIRST non-ok phase in day order'
Assert-Equal 'track' $model[3]['next']['field'] 'model: nothing done -> the first phase'
Assert-Equal 'pending' $model[4]['verdicts']['track'] 'model: a value no code matches counts as pending'
$totals = Get-EbiDashboardTotals -Model $model -Phases $phases
Assert-Equal 3 $totals['track']['ok'] 'totals: track ok'
Assert-Equal 2 $totals['track']['pending'] 'totals: track pending (blank + a value no code matches)'
Assert-Equal 1 $totals['evidence']['ng'] 'totals: evidence ng'
Assert-Equal 3 $totals['evidence']['pending'] 'totals: evidence pending'

$unf = @(@{ runId = 'r9'; startedAt = '2026-10-12T10:30:15'; workflow = @{ id = 'demo.track' } })
$next = Get-EbiDashboardNext -Model $model -Phases $phases -Unfinished $unf
Assert-Equal 'resume' $next['first']['kind'] 'next: an unfinished run comes first'
Assert-Equal 'r9' $next['first']['runId'] 'next: ... with its run id'
Assert-True ($next['lines'][0] -eq 'unfinished track run r9 (started 2026-10-12 10:30:15) -> -Resume') ('next: the resume line names the phase and the stamp: ' + $next['lines'][0])
$sameDay = Get-EbiDashboardNext -Model $model -Phases $phases -Unfinished @(@{ runId = '20261012-103015-ab12'; startedAt = '2026-10-12T10:30:15'; workflow = @{ id = 'demo.track' } })
Assert-True ($sameDay['lines'][0] -eq 'unfinished track run 20261012-103015-ab12 (started 10:30:15) -> -Resume') ('next: the date is left out when the run id carries it: ' + $sameDay['lines'][0])
$next = Get-EbiDashboardNext -Model $model -Phases $phases -Unfinished @()
Assert-Equal 'clear' $next['first']['kind'] 'next: no unfinished run -> the ng row first'
Assert-Equal 'A2' (@($next['first']['keys']) -join ',') 'next: ... naming it'
Assert-True (($next['lines'] -join "`n").Contains('evidence NG: A2 -> clear the cell and rerun evidence')) 'next: the ng line'
Assert-True (($next['lines'] -join "`n").Contains('track: 2 rows wait (A4, A5)')) 'next: rows waiting for track'
Assert-True (($next['lines'] -join "`n").Contains('logs: 1 row waits (A3)')) 'next: one row waits (singular)'
$okModel = @(Get-EbiDashboardModel -Rows @((New-Row 'A1' 'ok' 'ok' 'ok')) -Phases $phases -Profile $profile -KeyColumns @('Excel_NAME') -ShowColumns $show)
$next = Get-EbiDashboardNext -Model $okModel -Phases $phases -Unfinished @()
Assert-Equal 'done' $next['first']['kind'] 'next: everything ok -> done'
$next = Get-EbiDashboardNext -Model @() -Phases $phases -Unfinished @()
Assert-Equal 'plan' $next['first']['kind'] 'next: an empty worklist -> run the plan'
Assert-Equal 'demo.plan' $next['first']['phase']['id'] 'next: ... the no-source phase'
# 10/16: the menu's child now runs through Start-Process with ONE argument string
Assert-Equal 'run' (ConvertTo-EbiProcessArgument -Value 'run') 'arg: a plain word stays bare'
Assert-Equal '""' (ConvertTo-EbiProcessArgument -Value '') 'arg: empty -> ""'
Assert-Equal '"C:\work dir\x.json"' (ConvertTo-EbiProcessArgument -Value 'C:\work dir\x.json') 'arg: a space -> quoted, inner backslashes kept'
Assert-Equal '"C:\work dir\\"' (ConvertTo-EbiProcessArgument -Value 'C:\work dir\') 'arg: a trailing backslash is doubled before the closing quote'
Assert-Equal '"say \"hi\""' (ConvertTo-EbiProcessArgument -Value 'say "hi"') 'arg: inner quotes are escaped'
Assert-Equal 'C:\Users\JP\work-BIX' (ConvertTo-EbiProcessArgument -Value 'C:\Users\JP\work-BIX') 'arg: a path without spaces stays bare'
Assert-Equal 'KJODWSPP,RJODWSPP' (ConvertTo-EbiProcessArgument -Value 'KJODWSPP,RJODWSPP') 'arg: an -Only list stays one argument'
Assert-True ((Format-EbiMenuRunResult -Phase 'plan' -Code 0) -like 'plan: done*') 'run result: 0 -> done'
Assert-True ((Format-EbiMenuRunResult -Phase 'track' -Code 1) -like 'track: stopped with a failure*') 'run result: 1 -> failure'
Assert-Equal 'deliver: cancelled (q)' (Format-EbiMenuRunResult -Phase 'deliver' -Code 3) 'run result: 3 -> cancelled'
Assert-Equal 'logs: ended with exit code 7' (Format-EbiMenuRunResult -Phase 'logs' -Code 7) 'run result: anything else -> the code'
# 10/16: a two-digit total under a four-letter column ("ok 21" under "logs") was cut to "ok ~"
$bigRows = @(); for ($k = 1; $k -le 21; $k++) { $bigRows += (New-Row ('B' + $k) 'ok' 'ok' 'ok') }
$bigModel = @(Get-EbiDashboardModel -Rows $bigRows -Phases $phases -Profile $profile -KeyColumns @('Excel_NAME') -ShowColumns $show)
$bigLines = @(Format-EbiDashboard -Title 'demo' -WorkDir 'C:\w' -WorklistPath 'C:\w\x.csv' -Model $bigModel -Phases $phases -Totals (Get-EbiDashboardTotals -Model $bigModel -Phases $phases) -Next $null -ShowColumns $show -KeyHeader 'Excel_NAME')
$totLine = @($bigLines | Where-Object { $_ -like 'total 21*' })[0]
Assert-True (-not $totLine.Contains('~') -and ([regex]::Matches($totLine, 'ok 21')).Count -eq 3) ('render: two-digit totals are never cut: ' + $totLine)
Assert-Equal 0 @($bigLines | Where-Object { $_.Length -gt 80 }).Count 'render: ... and still no line wider than 80'
# 10/16: the earlier days' rows stay in the worklist; no row dated today -> plan first
$next = Get-EbiDashboardNext -Model $model -Phases $phases -Unfinished @() -TodayRows 0 -Today '2026-10-16'
Assert-Equal 'plan' $next['first']['kind'] 'next: no row dated today -> plan before the waiting / ng rows'
Assert-Equal 'no row is dated today (2026-10-16) -> run plan first' $next['lines'][0] 'next: ... and says why'
Assert-True (($next['lines'] -join "`n").Contains('track: 2 rows wait (A4, A5)')) 'next: ... the waiting rows are still listed'
$next = Get-EbiDashboardNext -Model $model -Phases $phases -Unfinished $unf -TodayRows 0
Assert-Equal 'resume' $next['first']['kind'] 'next: an unfinished run still comes before plan'
Assert-Equal 'clear' (Get-EbiDashboardNext -Model $model -Phases $phases -Unfinished @() -TodayRows 2)['first']['kind'] 'next: rows dated today -> as before'
Assert-Equal 'clear' (Get-EbiDashboardNext -Model $model -Phases $phases -Unfinished @() -TodayRows -1)['first']['kind'] 'next: no date column (-1) -> as before'
Assert-Equal '2026-10-09' (ConvertTo-EbiDashboardDay -Text '2026/10/9') 'day: 2026/10/9'
Assert-Equal '2026-10-14' (ConvertTo-EbiDashboardDay -Text '2026-10-14') 'day: 2026-10-14'
Assert-Equal '2026-10-14' (ConvertTo-EbiDashboardDay -Text '2026/10/14 10:15:00') 'day: a time after the date is ignored'
Assert-Equal '' (ConvertTo-EbiDashboardDay -Text '') 'day: blank'
Assert-Equal '' (ConvertTo-EbiDashboardDay -Text '2026/13/40') 'day: not a date'
$dProf = @{ worklist = @{ dashboard = @{ dateColumn = 'D' } } }
$dRows = @(@{ D = '2026/10/16' }, @{ D = '2026-10-16' }, @{ D = '2026-10-15' }, @{ D = '' })
Assert-Equal 2 (Get-EbiDashboardTodayRows -Rows $dRows -Profile $dProf -Columns @('D') -Today '2026-10-16') 'today rows: both date forms count'
Assert-Equal -1 (Get-EbiDashboardTodayRows -Rows $dRows -Profile @{ worklist = @{} } -Columns @('D') -Today '2026-10-16') 'today rows: no dateColumn -> -1'
Assert-Equal -1 (Get-EbiDashboardTodayRows -Rows $dRows -Profile $dProf -Columns @('X') -Today '2026-10-16') 'today rows: the CSV lacks the column -> -1'
$runOnly = @(Get-EbiDashboardModel -Rows @((New-Row 'A4' '' '' '')) -Phases $phases -Profile $profile -KeyColumns @('Excel_NAME') -ShowColumns $show)
$next = Get-EbiDashboardNext -Model $runOnly -Phases $phases -Unfinished @()
Assert-Equal 'run' $next['first']['kind'] 'next: rows waiting -> run'
Assert-Equal 'demo.track' $next['first']['phase']['id'] 'next: ... the earliest phase'

# ---------------------------------------------------------------- run line
$cmd = New-EbiRunCommand -Phase $phases[1] -ToolDir 'C:\tool' -WorkDir 'C:\work' -Only @('A1', ' A2 ')
Assert-Equal 'ebi.ps1 run workflows\demo.track.json -WorkDir C:\work -Only A1,A2' $cmd['text'] 'run line: relative workflow path, -Only joined'
Assert-Equal 'run|C:\tool\workflows\demo.track.json|-WorkDir|C:\work|-Only|A1,A2' (@($cmd['args']) -join '|') 'run line: the child argument list'
$cmd = New-EbiRunCommand -Phase $phases[1] -ToolDir 'C:\tool' -WorkDir 'C:\my work' -Resume $true -RunId 'r9'
Assert-Equal 'ebi.ps1 run workflows\demo.track.json -WorkDir "C:\my work" -Resume -RunId r9' $cmd['text'] 'run line: a work dir with a space is quoted; -Resume -RunId'
Assert-True (@($cmd['args']) -contains '-Resume' -and @($cmd['args']) -contains 'r9') 'run line: resume args'
$cmd = New-EbiRunCommand -Phase $phases[1] -ToolDir 'D:\elsewhere' -WorkDir ''
Assert-Equal 'ebi.ps1 run C:\tool\workflows\demo.track.json' $cmd['text'] 'run line: a workflow outside the tool dir keeps its full path; no -WorkDir'

# ---------------------------------------------------------------- rendering
$lines = @(Format-EbiDashboard -Title 'demo' -WorkDir 'C:\work' -WorklistPath 'C:\work\demo_op1.csv' -Model $model -Phases $phases -Totals $totals -Next (Get-EbiDashboardNext -Model $model -Phases $phases -Unfinished $unf) -ShowColumns $show -KeyHeader 'Excel_NAME' -RunLine 'next: ebi.ps1 run x' -Today '2026-10-12')
$tooWide = @($lines | Where-Object { $_.Length -gt 80 })
Assert-Equal 0 $tooWide.Count ('render: every line is at most 80 characters' + $(if ($tooWide.Count -gt 0) { ': ' + $tooWide[0] } else { '' }))
Assert-Equal 'ebi status: demo  2026-10-12  C:\work' $lines[0] 'render: the header'
Assert-Equal 'worklist: demo_op1.csv  5 rows' $lines[1] 'render: the worklist line'
Assert-True ($lines[3] -like 'Excel_NAME  JOB  GFIX_TIME  track  logs  evidence') ('render: the column header: ' + $lines[3])
Assert-True ($lines[5] -like 'A2*ok*ok*NG') ('render: an ng cell reads NG: ' + $lines[5])
Assert-True ($lines[7] -like 'A4*-*-*-') ('render: a blank cell reads -: ' + $lines[7])
Assert-True (($lines -join "`n").Contains('total 5')) 'render: the totals row'
Assert-True (($lines -join "`n").Contains("`nNEXT`n")) 'render: the NEXT section'
Assert-True ($lines[$lines.Count - 1] -eq '  next: ebi.ps1 run x') 'render: the run line last'
$wideShow = @(1..12 | ForEach-Object { 'column_' + $_ })
$wideModel = @(@{ key = 'K1'; cells = @(1..12 | ForEach-Object { 'value_number_' + $_ + '_long' }); verdicts = @{ track = 'ok'; logs = 'ok'; evidence = 'ok' }; next = $null })
$lines = @(Format-EbiDashboard -Title 'demo' -WorkDir '' -WorklistPath '' -Model $wideModel -Phases $phases -Totals (Get-EbiDashboardTotals -Model $wideModel -Phases $phases) -Next $null -ShowColumns $wideShow -KeyHeader 'key' -Today '')
$tooWide = @($lines | Where-Object { $_.Length -gt 80 })
Assert-Equal 0 $tooWide.Count 'render: too many columns -> dropped from the right, still 80'
Assert-True ($lines[3].StartsWith('key') -and $lines[3].EndsWith('evidence')) 'render: ... the key and the verdicts stay'
Assert-True ($lines[3].Contains('column_1') -and -not $lines[3].Contains('column_12')) 'render: ... the leftmost shown columns survive'
$empty = @(Format-EbiDashboard -Title 'demo' -WorkDir 'W' -WorklistPath '' -Model @() -Phases $phases -Totals @{} -Next (Get-EbiDashboardNext -Model @() -Phases $phases) -ShowColumns @() -KeyHeader 'k')
Assert-True (($empty -join "`n").Contains('worklist: (none)  0 rows') -and ($empty -join "`n").Contains('the worklist is empty -> run plan first')) 'render: no worklist yet'
Assert-Equal '2026-10-12 10:30:15' (ConvertTo-EbiDashboardStamp -Value ([datetime]'2026-10-12T10:30:15')) 'stamp: a DateTime (pwsh 7 JSON)'
Assert-Equal '2026-10-12 10:30:15' (ConvertTo-EbiDashboardStamp -Value '2026-10-12T10:30:15') 'stamp: an ISO string (5.1 JSON)'
Assert-Equal 'soon' (ConvertTo-EbiDashboardStamp -Value 'soon') 'stamp: anything else as-is'

$wide = ConvertTo-EbiHelpFitDashboard -Text ('abc ' + ([string][char]0x65E5 * 50)) -Width 20
Assert-True ($wide.EndsWith('...') -and $wide.Length -eq 13) ('fit: a full-width character counts as two console cells (4 + 6 x 2 + 3 = 19 cells): ' + $wide.Length)
Assert-Equal 'abc' (ConvertTo-EbiHelpFitDashboard -Text 'abc' -Width 20) 'fit: a short line is untouched'

# ---------------------------------------------------------------- menu (pure parts)
$choices = @(Format-EbiMenuChoices -Phases $phases -Unfinished $unf)
Assert-True ($choices[0] -like '  1  plan *' -and $choices[3] -like '  4  evidence *') 'menu: numbered phases'
Assert-True ($choices[4] -like '  r  resume*demo.track*r9*') ('menu: the resume line: ' + $choices[4])
Assert-True (($choices -join "`n").Contains('  c  clear') -and ($choices -join "`n").Contains('  q  quit')) 'menu: the fixed keys'
$a = Read-EbiMenuAnswer -Text '3' -PhaseCount 4 -HasUnfinished $false
Assert-True ($a['ok'] -and $a['action'] -eq 'phase' -and $a['index'] -eq 3) 'answer: a phase number'
Assert-True (-not (Read-EbiMenuAnswer -Text '5' -PhaseCount 4 -HasUnfinished $false)['ok']) 'answer: out of range'
Assert-True (-not (Read-EbiMenuAnswer -Text '' -PhaseCount 4 -HasUnfinished $false)['ok']) 'answer: Enter alone is nothing'
Assert-True ((Read-EbiMenuAnswer -Text ' R ' -PhaseCount 4 -HasUnfinished $true)['action'] -eq 'r') 'answer: r when there is a run to resume'
Assert-True (-not (Read-EbiMenuAnswer -Text 'r' -PhaseCount 4 -HasUnfinished $false)['ok']) 'answer: r without one is nothing'
Assert-True ((Read-EbiMenuAnswer -Text 'Q' -PhaseCount 4 -HasUnfinished $false)['action'] -eq 'q') 'answer: q'
$w = @(Get-EbiMenuPhaseWarnings -Model $model -Phases $phases -Phase $phases[3])
Assert-Equal 'A3: logs is not ok yet -> evidence will refuse it|A4: track is not ok yet -> evidence will refuse it|A5: track is not ok yet -> evidence will refuse it' ($w -join '|') 'warnings: evidence before logs / track, one line per row'
Assert-Equal 0 @(Get-EbiMenuPhaseWarnings -Model $model -Phases $phases -Phase $phases[1]).Count 'warnings: the first phase has no prerequisite'
Assert-Equal 'A4: track is not ok yet -> evidence will refuse it' (@(Get-EbiMenuPhaseWarnings -Model $model -Phases $phases -Phase $phases[3] -Only @('A4', 'A1')) -join '|') 'warnings: -Only narrows them'

# ---------------------------------------------------------------- the loader, the write, the loop (temp files)
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('ebidensu-dash-' + [guid]::NewGuid().ToString('N'))
$tool = Join-Path $tmp 'tool'; $work = Join-Path $tmp 'work'
New-Item -ItemType Directory -Path (Join-Path (Join-Path $tool 'profiles') 'demo') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tool 'workflows') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path (Join-Path $work 'run') 'r1') -Force | Out-Null
try {
    [void](Write-EbiJson -Path (Join-Path (Join-Path (Join-Path $tool 'profiles') 'demo') 'worklist.json') -Value $profile['worklist'])
    [void](Write-EbiJson -Path (Join-Path (Join-Path $tool 'workflows') 'demo.plan.json') -Value (New-Wf 'demo.plan' '')['value'])
    [void](Write-EbiJson -Path (Join-Path (Join-Path $tool 'workflows') 'demo.track.json') -Value (New-Wf 'demo.track' 'track')['value'])
    [void](Write-EbiJson -Path (Join-Path (Join-Path $tool 'workflows') 'demo.logs.json') -Value (New-Wf 'demo.logs' 'logs')['value'])
    [void](Write-EbiJson -Path (Join-Path (Join-Path $tool 'workflows') 'other.x.json') -Value (New-Wf 'other.x' 'track' 'other')['value'])
    [System.IO.File]::WriteAllText((Join-Path (Join-Path $tool 'workflows') 'bad.json'), '{ not json', $utf8)
    [void](Write-EbiJson -Path (Join-Path (Join-Path (Join-Path $work 'run') 'r1') 'run.json') -Value @{ runId = 'r1'; startedAt = '2026-10-12T09:00:00'; workflow = @{ id = 'demo.track' }; finished = $false })
    $csv = Join-Path $work 'demo_op1.csv'
    [System.IO.File]::WriteAllText($csv, "Excel_NAME,JOB,USER,GFIX_TIME,track,logs,evidence,note`r`nA1,A1J,u,10:00,ok,ok,,`r`nA2,A2J,u,11:00,ng,,,`r`n", (New-Object System.Text.UTF8Encoding($true)))

    $d = Read-EbiDashboard -ToolDir $tool -WorkDir $work -Operator 'op1' -Today '2026-10-12'
    Assert-True ($d['ok']) ('loader: ok: ' + $d['message'])
    Assert-Equal 'demo' $d['profileName'] 'loader: the one profile whose worklist is in the work dir is found without -Profile'
    Assert-Equal $csv $d['worklistPath'] 'loader: the worklist file template ({{run.operator}}) expanded'
    Assert-Equal 'plan,track,logs' (@($d['phases'] | ForEach-Object { $_['name'] }) -join ',') 'loader: phases from workflows/, the other profile and the bad file skipped'
    Assert-Equal 2 @($d['model']).Count 'loader: rows'
    Assert-Equal 1 @($d['unfinished']).Count 'loader: the unfinished run'
    Assert-True ($d['runLine'] -like 'next: ebi.ps1 run *demo.track.json -WorkDir * -Resume -RunId r1') ('loader: the recommended line resumes it: ' + $d['runLine'])
    Assert-True ($d['lines'][0] -like 'ebi status: demo  2026-10-12  *') ('loader: rendered header: ' + $d['lines'][0])
    $noWork = Join-Path $tmp 'nowork'; New-Item -ItemType Directory -Path $noWork -Force | Out-Null
    $d2 = Read-EbiDashboard -ToolDir $tool -WorkDir $noWork -Operator 'op1'
    Assert-True ($d2['ok'] -and $d2['profileName'] -eq 'demo' -and @($d2['model']).Count -eq 0 -and $d2['next']['first']['kind'] -eq 'plan') 'loader: a fresh work folder (no worklist yet) and ONE profile with a plan workflow -> that profile, an empty board that says run plan'
    Assert-Equal 'demo' (@(Get-EbiPlanProfiles -ToolDir $tool) -join ',') 'plan profiles: only profiles with a folder and a no-source workflow (other.x has a source and no folder)'
    $dt = Read-EbiDashboard -ToolDir $tool -WorkDir $tool -Operator 'op1'
    Assert-True (-not $dt['ok'] -and $dt['message'] -like 'the work folder is the tool folder*') ('loader: the tool folder as the work folder -> says so: ' + $dt['message'])
    $badWork = Join-Path $tmp 'badwork'; New-Item -ItemType Directory -Path $badWork -Force | Out-Null
    [System.IO.File]::WriteAllText((Join-Path $badWork 'ebi.local.json'), "{ `"paths`": { `"a`": `"x`"`r`n `"b`": `"y`" } }", $utf8)
    $db = Read-EbiDashboard -ToolDir $tool -WorkDir $badWork -Operator 'op1'
    Assert-True (-not $db['ok'] -and $db['message'] -like 'could not read the profile(s)*demo:*ebi.local.json*') ('loader: a broken ebi.local.json is reported as such, not as "no worklist": ' + $db['message'])
    New-Item -ItemType Directory -Path (Join-Path (Join-Path $tool 'profiles') 'second') -Force | Out-Null
    [void](Write-EbiJson -Path (Join-Path (Join-Path $tool 'workflows') 'second.plan.json') -Value (New-Wf 'second.plan' '' 'second')['value'])
    $d2b = Read-EbiDashboard -ToolDir $tool -WorkDir $noWork -Operator 'op1'
    Assert-True (-not $d2b['ok'] -and $d2b['message'] -like 'no profile has its worklist in *') 'loader: no worklist and TWO plan profiles -> says so (give -Profile)'
    Remove-Item -LiteralPath (Join-Path (Join-Path $tool 'workflows') 'second.plan.json') -Force
    Remove-Item -LiteralPath (Join-Path (Join-Path $tool 'profiles') 'second') -Recurse -Force
    $d3 = Read-EbiDashboard -ToolDir $tool -WorkDir $noWork -ProfileName 'demo' -Operator 'op1'
    Assert-True ($d3['ok'] -and @($d3['model']).Count -eq 0 -and $d3['next']['first']['kind'] -eq 'plan') 'loader: -Profile with no worklist yet -> an empty dashboard that says run plan'
    $d4 = Read-EbiDashboard -ToolDir $tool -WorkDir $work -ProfileName 'nosuch'
    Assert-True (-not $d4['ok'] -and $d4['message'] -like 'profile directory not found*') 'loader: an unknown profile'

    $c = Clear-EbiWorklistVerdict -Path $csv -Key 'A1' -Field 'track' -KeyColumns @('Excel_NAME') -DryRun $true
    Assert-True ($c['ok'] -and $c['before'] -eq 'ok' -and $c['message'] -like 'would clear*') 'clear: dry run says what it would do'
    Assert-Equal 'ok' (Read-EbiCsv -Path $csv)['rows'][0]['track'] 'clear: ... and writes nothing'
    $c = Clear-EbiWorklistVerdict -Path $csv -Key 'a2' -Field 'track' -KeyColumns @('Excel_NAME')
    Assert-True ($c['ok'] -and $c['before'] -eq 'ng') ('clear: a case-folded key still finds the row: ' + $c['message'])
    Assert-Equal '' (Read-EbiCsv -Path $csv)['rows'][1]['track'] 'clear: the cell is blank on disk'
    Assert-Equal 'ok' (Read-EbiCsv -Path $csv)['rows'][0]['track'] 'clear: the other row untouched'
    $c = Clear-EbiWorklistVerdict -Path $csv -Key 'Z9' -Field 'track' -KeyColumns @('Excel_NAME')
    Assert-True (-not $c['ok'] -and $c['message'] -like 'no row with key*') 'clear: an unknown key is refused'
    $c = Clear-EbiWorklistVerdict -Path $csv -Key 'A1' -Field 'nosuch' -KeyColumns @('Excel_NAME')
    Assert-True (-not $c['ok'] -and $c['message'] -like 'column "nosuch" is not in*') 'clear: an unknown column is refused'

    # the loop: run track for A1, clear A1's logs, redraw, quit -- answers scripted, the child run captured
    $answers = New-Object System.Collections.Queue
    foreach ($x in @('2', 'A1', 'c', 'A1', 'logs', '', 's', 'zz', 'q')) { $answers.Enqueue($x) }
    $script:ranArgs = New-Object System.Collections.ArrayList
    $m = Invoke-EbiMenu -ToolDir $tool -WorkDir $work -Operator 'op1' -Reader { $answers.Dequeue() } -Runner { param($CliArgs) [void]$script:ranArgs.Add(@($CliArgs) -join ' '); return 0 } 6>$null
    Assert-True ($m['ok'] -and $m['message'] -eq 'quit') ('menu: ran to q: ' + $m['message'])
    Assert-Equal 1 @($m['ran']).Count 'menu: one command ran'
    Assert-True ($m['ran'][0] -like 'ebi.ps1 run *demo.track.json -WorkDir * -Only A1') ('menu: the printed line: ' + $m['ran'][0])
    Assert-True ($script:ranArgs[0] -like 'run *demo.track.json -WorkDir * -Only A1') ('menu: the child got the same arguments: ' + $script:ranArgs[0])
    Assert-Equal 1 @($m['cleared']).Count 'menu: one verdict cleared'
    Assert-Equal '' (Read-EbiCsv -Path $csv)['rows'][0]['logs'] 'menu: ... A1 logs is blank on disk'
    $state = Read-EbiMenuState -ToolDir $tool
    Assert-True ($state['workDir'] -eq $work -and $state['profile'] -eq 'demo') 'menu: the work dir and profile are remembered (.ebi/menu.json)'
    $prevNoAsk = $env:EBI_NO_ASK
    $env:EBI_NO_ASK = '1'
    try {
        $m2 = Invoke-EbiMenu -ToolDir $tool -WorkDir $work -Operator 'op1' -Runner { param($CliArgs) return 0 } 6>$null
        Assert-True ($m2['ok'] -and $m2['message'] -eq 'no console' -and @($m2['ran']).Count -eq 0) 'menu: with nobody to answer it draws once and quits'
    } finally { $env:EBI_NO_ASK = $prevNoAsk }
    $m3 = Invoke-EbiMenu -ToolDir $tool -WorkDir $work -Operator 'op1' -Reader { 'q' } -Runner { param($CliArgs) return 0 } -DryRun $true 6>$null
    Assert-True ($m3['ok']) 'menu: dry run draws and quits'

    # ebi.ps1 status / menu in a child process, on the real gfix-recv profile
    $psExe = (Get-Process -Id $PID).Path
    $ebi = Join-Path $repoRoot 'ebi.ps1'
    $cliWork = Join-Path $tmp 'cliwork'; New-Item -ItemType Directory -Path $cliWork -Force | Out-Null
    [System.IO.File]::WriteAllText((Join-Path $cliWork 'gfixrecv.csv'), "Excel_NAME,JOB,USER,RCVFILE,GFIX_DATE,GFIX_TIME,FileCount,Owner,track,logs,evidence,note`r`nXDEMWM01,XDEMJM01,u,," + (Get-Date).ToString('yyyy/MM/dd') + ",09:31:05,2,op1,ok,,,`r`n", (New-Object System.Text.UTF8Encoding($true)))
    $out = '' | & $psExe -NoLogo -NoProfile -File $ebi status -WorkDir $cliWork 2>&1 | Out-String
    $code = $LASTEXITCODE
    Assert-True ($code -eq 0 -and $out.Contains('ebi status: gfix-recv') -and $out.Contains('XDEMWM01') -and $out.Contains('logs: 1 row waits (XDEMWM01)')) ('cli status: the gfix-recv dashboard, exit 0' + $(if ($code -ne 0) { "`n" + $out } else { '' }))
    Assert-True ($out.Contains('next: ebi.ps1 run ') -and $out.Contains('gfixRecv.logs.json')) 'cli status: recommends the logs run'
    $out = '' | & $psExe -NoLogo -NoProfile -File $ebi menu -WorkDir $cliWork 2>&1 | Out-String
    Assert-True ($LASTEXITCODE -eq 0 -and $out.Contains('ebi menu: gfix-recv') -and $out.Contains('  4  evidence') -and $out.Contains('(no console to ask: quit)')) 'cli menu: redirected stdin -> draws the menu and quits, exit 0'
    $out = '' | & $psExe -NoLogo -NoProfile -File $ebi status -WorkDir $noWork 2>&1 | Out-String
    Assert-True ($LASTEXITCODE -eq 0 -and $out.Contains('ebi status: gfix-recv') -and $out.Contains('run plan first')) ('cli status: a fresh work folder -> the one plan profile (gfix-recv), empty, "run plan first"' + "`n" + $out)
    $out = '' | & $psExe -NoLogo -NoProfile -File $ebi status -WorkDir (Join-Path $tmp 'nowhere') 2>&1 | Out-String
    Assert-True ($LASTEXITCODE -eq 2 -and $out.Contains('work dir not found')) 'cli status: a missing work dir -> usage (exit 2)'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    # the child `ebi.ps1 menu` remembered the temp work dir in the repo's .ebi/menu.json: forget it, but never an operator's own entry
    $stateFile = Get-EbiMenuStateFile -ToolDir $repoRoot
    if ((Test-Path -LiteralPath $stateFile) -and (Read-EbiMenuState -ToolDir $repoRoot)['workDir'] -eq (Join-Path $tmp 'cliwork')) { Remove-Item -LiteralPath $stateFile -Force -ErrorAction SilentlyContinue }
}

exit (Complete-Tests)
