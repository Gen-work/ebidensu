#Requires -Version 5.1
# Test-DashboardDay.ps1 -- the day-scoped board and menu (10/20): a worklist
# with a work-day column (role "day") and a hold column shows TODAY's rows
# only and sums up the rest (held, earlier days not finished, earlier days
# all ok, later days); the menu's Enter runs the next step, m marks a
# verdict ok by hand, h holds / releases rows, and before a phase runs the
# menu offers to mark rows whose phase file is already there, and asks what
# to do with earlier days' unfinished rows. Temp tool dir + work dir.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Dashboard.ps1')

Reset-Tests 'DashboardDay'

$today = (Get-Date).ToString('yyyy-MM-dd')
$yday = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd')
$tmrw = (Get-Date).AddDays(1).ToString('yyyy-MM-dd')
$verdict = @{ ok = 'ok'; ng = 'ng'; unknown = ''; pending = '' }
$profile = @{
    name = 'day'
    worklist = @{
        file = 'day.csv'
        key = @{ columns = @('Excel_NAME') }
        dashboard = @{ columns = @('GFIX_DATE'); dateColumn = 'WORKDAY' }
        columns = @(
            @{ name = 'Excel_NAME'; role = 'key' },
            @{ name = 'GFIX_DATE'; role = 'text' },
            @{ name = 'WORKDAY'; role = 'day'; fallback = 'GFIX_DATE' },
            @{ name = 'hold'; role = 'hold' },
            @{ name = 'track'; role = 'verdict'; values = $verdict; doneFile = 'capture/{{item.keySafe}}/track.json' },
            @{ name = 'logs'; role = 'verdict'; values = $verdict; doneFile = 'capture/{{item.keySafe}}/logs.json' }
        )
    }
}
$rows = @(
    @{ Excel_NAME = 'T1'; GFIX_DATE = $today; WORKDAY = $today; hold = ''; track = 'ok'; logs = '' },
    @{ Excel_NAME = 'T2'; GFIX_DATE = $today; WORKDAY = $today; hold = ''; track = ''; logs = '' },
    @{ Excel_NAME = 'T3'; GFIX_DATE = $today; WORKDAY = ''; hold = ''; track = ''; logs = '' },
    @{ Excel_NAME = 'E1DBC'; GFIX_DATE = $yday; WORKDAY = $yday; hold = ''; track = 'ok'; logs = '' },
    @{ Excel_NAME = 'E2'; GFIX_DATE = $yday; WORKDAY = $yday; hold = ''; track = 'ok'; logs = 'ok' },
    @{ Excel_NAME = 'H1DBC'; GFIX_DATE = $yday; WORKDAY = $yday; hold = 'HOST'; track = ''; logs = '' },
    @{ Excel_NAME = 'L1'; GFIX_DATE = ''; WORKDAY = $tmrw; hold = ''; track = ''; logs = '' }
)

# ---------------------------------------------------------------- pure
$sp = Split-EbiDashboardRows -Rows $rows -Profile $profile -Today $today -KeyColumns @('Excel_NAME') -VerdictFields @('track', 'logs')
Assert-Equal 'T1,T2,T3' (@($sp['today'] | ForEach-Object { $_['Excel_NAME'] }) -join ',') 'split: today''s rows (T3 by its GFIX date: no WORKDAY yet)'
Assert-Equal 'H1DBC (HOST)' (@($sp['held']) -join ',') 'split: the held row, with its reason'
Assert-Equal ('E1DBC (' + $yday.Substring(5).Replace('-', '/') + ': track ok)') (@($sp['open']) -join ',') 'split: an earlier day''s row not finished, with what is done'
Assert-Equal 1 $sp['doneCount'] 'split: an earlier day''s row that is all ok is only counted'
Assert-Equal 1 @($sp['later']).Count 'split: a row given a later day'
$aside = @(Format-EbiDashboardAside -Split $sp)
$asideText = $aside -join ' | '
Assert-True ($aside[0] -like 'on hold: H1DBC (HOST)*h releases' -and $asideText -like '*| earlier days, not finished: E1DBC*' -and $aside[-1] -like 'earlier days, all ok: 1 row(s)*' -and $asideText -like '*planned for a later day: L1*') ('aside: ' + $asideText)
Assert-True (@($aside | Where-Object { $_.Length -gt 80 }).Count -eq 0) 'aside: every line fits 80 columns (long lists wrap, indented)'
$sp0 = Split-EbiDashboardRows -Rows $rows -Profile @{ worklist = @{ key = @{ columns = @('Excel_NAME') }; columns = @() } } -Today $today -KeyColumns @('Excel_NAME')
Assert-True (-not $sp0['scoped'] -and @($sp0['today']).Count -eq 7) 'split: without a day column every row is the board'
$rk = Resolve-EbiMenuKeys -Text 'dbc, t1' -Keys @('T1', 'T2', 'E1DBC', 'H1DBC')
Assert-Equal 'E1DBC,H1DBC,T1' (@($rk['keys']) -join ',') 'menu keys: part of a name takes every match, any case'
Assert-True (-not (Resolve-EbiMenuKeys -Text 'T1 ZZ' -Keys @('T1'))['ok']) 'menu keys: an unknown one refuses the whole answer'
Assert-Equal 'logs' (Resolve-EbiMenuField -Text '2' -Fields @('track', 'logs')) 'menu field: by number'
Assert-Equal 'track' (Resolve-EbiMenuField -Text 'TRACK' -Fields @('track', 'logs')) 'menu field: by name'
Assert-Equal '' (Resolve-EbiMenuField -Text '3' -Fields @('track', 'logs')) 'menu field: out of range'
Assert-Equal 'next' (Read-EbiMenuAnswer -Text '' -PhaseCount 3 -HasUnfinished $false -HasNext $true)['action'] 'menu answer: Enter = next when there is one'
Assert-Equal 'm' (Read-EbiMenuAnswer -Text 'M' -PhaseCount 3 -HasUnfinished $false)['action'] 'menu answer: m'
Assert-Equal 'h' (Read-EbiMenuAnswer -Text ([string][char]0xFF48) -PhaseCount 3 -HasUnfinished $false)['action'] 'menu answer: a full-width h too'
$ph = @{ path = 'C:\t\workflows\day.track.json'; name = 'track' }
Assert-True ((New-EbiRunCommand -Phase $ph -ToolDir 'C:\t' -WorkDir 'W' -AllDays $true)['text'] -like '* -AllDays') 'run line: -AllDays'
Assert-True (-not ((New-EbiRunCommand -Phase $ph -ToolDir 'C:\t' -WorkDir 'W' -Only @('A') -AllDays $true)['text'] -like '*-AllDays*')) 'run line: -Only wins over -AllDays'

# ---------------------------------------------------------------- loader + menu on temp files
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('ebidensu-day-' + [guid]::NewGuid().ToString('N'))
$tool = Join-Path $tmp 'tool'; $work = Join-Path $tmp 'work'
New-Item -ItemType Directory -Path (Join-Path (Join-Path $tool 'profiles') 'day') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tool 'workflows') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path (Join-Path $work 'capture') 'T2') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path (Join-Path $work 'capture') 'T3') -Force | Out-Null
try {
    [void](Write-EbiJson -Path (Join-Path (Join-Path (Join-Path $tool 'profiles') 'day') 'worklist.json') -Value $profile['worklist'])
    foreach ($w in @(@('day.plan', ''), @('day.track', 'track'), @('day.logs', 'logs'))) {
        $v = @{ schema = 1; id = $w[0]; title = 't'; profile = 'day' }
        if ($w[1] -ne '') { $v['source'] = @{ table = 'wl'; select = @{ field = $w[1]; pendingWhen = '!= ok' } } }
        [void](Write-EbiJson -Path (Join-Path (Join-Path $tool 'workflows') ($w[0] + '.json')) -Value $v)
    }
    $csv = Join-Path $work 'day.csv'
    $cols = @('Excel_NAME', 'GFIX_DATE', 'WORKDAY', 'hold', 'track', 'logs')
    [void](Write-EbiCsvAtomic -Path $csv -Columns $cols -Rows $rows)
    [System.IO.File]::WriteAllText((Join-Path (Join-Path (Join-Path $work 'capture') 'T2') 'track.json'), '{}')
    $old = Join-Path (Join-Path (Join-Path $work 'capture') 'T3') 'track.json'
    [System.IO.File]::WriteAllText($old, '{}')
    (Get-Item -LiteralPath $old).LastWriteTime = (Get-Date).AddDays(-2)

    $d = Read-EbiDashboard -ToolDir $tool -WorkDir $work -Today $today
    Assert-True ($d['ok']) ('loader: ok: ' + $d['message'])
    Assert-Equal 'T1,T2,T3' (@($d['model'] | ForEach-Object { $_['key'] }) -join ',') 'loader: the board is today''s rows'
    Assert-True ((@($d['lines']) -join "`n") -like '*worklist: day.csv  7 rows, today 3*') 'loader: the worklist line says how many rows and how many today'
    Assert-True ((@($d['lines']) -join "`n") -like '*on hold: H1DBC (HOST)*') 'loader: the held row is under the table'
    Assert-Equal 'track' ([string]$d['next']['first']['phase']['name']) 'loader: next = track (T2, T3 wait)'
    $hits = @(Get-EbiDoneFileHits -Rows $d['split']['today'] -Profile $d['profile'] -Field 'track' -WorkDir $work -KeyColumns @('Excel_NAME'))
    Assert-Equal 'T2' (@($hits | ForEach-Object { $_['key'] }) -join ',') 'done files: T2''s track.json is today''s; T3''s is older than its day (an earlier run) and is not offered'

    $answers = New-Object System.Collections.Queue
    foreach ($x in @(
        '',            # Enter = next: track for today's rows
        '',            #   rows: Enter = today's
        '',            #   LOOKS DONE ALREADY T2: Enter = mark ok
        '3',           # logs
        '',            #   rows: today's
        't',           #   EARLIER DAYS, NOT FINISHED (E1DBC): take it today
        'm', 'T1', '2', '',        # mark T1 logs ok
        'h', 'L1', '', '',         # hold L1 (reason Enter), confirm
        'h', 'h1', '',             # release H1DBC (part of the name, any case)
        'q')) { $answers.Enqueue($x) }
    $script:ranArgs = New-Object System.Collections.ArrayList
    $m = Invoke-EbiMenu -ToolDir $tool -WorkDir $work -Reader { $answers.Dequeue() } -Runner { param($CliArgs) [void]$script:ranArgs.Add(@($CliArgs) -join ' '); return 0 } 6>$null
    Assert-True ($m['ok'] -and $m['message'] -eq 'quit') ('menu: ran to q: ' + $m['message'] + ' (answers left: ' + $answers.Count + ')')
    Assert-Equal 2 @($m['ran']).Count 'menu: two phases ran'
    Assert-True ($m['ran'][0] -like 'ebi.ps1 run *day.track.json -WorkDir *' -and $m['ran'][0] -notlike '*-Only*') ('menu: Enter ran track for today''s rows: ' + $m['ran'][0])
    Assert-True ($m['ran'][1] -like 'ebi.ps1 run *day.logs.json -WorkDir *') ('menu: then logs: ' + $m['ran'][1])
    $after = @{}; foreach ($r in @((Read-EbiCsv -Path $csv)['rows'])) { $after[[string]$r['Excel_NAME']] = $r }
    Assert-Equal 'ok' $after['T2']['track'] 'menu: T2 track marked ok from its track.json'
    Assert-Equal '' $after['T3']['track'] 'menu: T3 untouched (its file is older than its day)'
    Assert-Equal $today $after['E1DBC']['WORKDAY'] 'menu: t took the earlier day''s row into today'
    Assert-Equal $yday $after['E1DBC']['GFIX_DATE'] 'menu: ... keeping the day its GFIX job ran'
    Assert-Equal 'ok' $after['T1']['logs'] 'menu: m marked T1 logs ok'
    Assert-Equal 'waiting for HOST' $after['L1']['hold'] 'menu: h put L1 on hold with the default reason'
    Assert-Equal '' $after['H1DBC']['hold'] 'menu: h on a held row released it'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    $stateFile = Get-EbiMenuStateFile -ToolDir $tool
    if (Test-Path -LiteralPath $stateFile) { Remove-Item -LiteralPath $stateFile -Force -ErrorAction SilentlyContinue }
}

exit (Complete-Tests)
