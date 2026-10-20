#Requires -Version 5.1
# ============================================================
#  kernel/Dashboard.ps1
#
#  The one-screen day dashboard and the menu behind `ebi status` and
#  `ebi menu` (TODO card "gfix-recv: a dashboard and one entry point",
#  2026-10-12). Dot-source only (no param() block, ASCII source, no
#  class).
#
#  After the first full office day the operator typed four long
#  `ebi.ps1 run workflows\<x>.json -WorkDir ... [-Only ...] [-Resume]`
#  lines from the runbook and read progress by opening the worklist CSV
#  in Notepad. This file replaces both habits with one read-only screen
#  (80 columns, the Gate.ps1 style) and one menu that builds and prints
#  the very same run line, so the runbook commands stay the source of
#  truth.
#
#  Nothing project-specific lives here. The day's PHASES are the
#  workflows that name the profile; their ORDER is the order of their
#  verdict columns in the profile's worklist.json (a workflow without a
#  source -- the morning plan -- comes first); the COLUMNS shown come
#  from worklist.json too (an optional "dashboard": { "columns": [...] }
#  picks them, else every non-key non-verdict column in declared order,
#  as many as fit).
#
#  Split of concerns:
#      Get-EbiDayPhases           PURE: workflows + profile -> ordered phases
#      Get-EbiDashboardColumns    PURE: which worklist columns to show
#      Get-EbiDashboardModel      PURE: rows -> per-row verdicts + next phase
#      Get-EbiDashboardTotals     PURE: per phase ok / ng / pending counts
#      Get-EbiDashboardNext       PURE: the NEXT section + the one line to run
#      Format-EbiDashboard        PURE: model -> lines, each <= 80 characters
#      New-EbiRunCommand          PURE: phase + options -> the ebi.ps1 run line
#      Read-EbiDashboard          the loader (profile, worklist CSV, run/,
#                                 workflows/) feeding the pure functions
#      Clear-EbiWorklistVerdict   the ONE write: blank one row's verdict cell
#      Invoke-EbiMenu             the loop: dashboard, choices, read, run
#                                 (reader and runner injectable for tests;
#                                 self-answering with no console)
# ============================================================

. (Join-Path $PSScriptRoot 'Json.ps1')
. (Join-Path $PSScriptRoot 'Key.ps1')
. (Join-Path $PSScriptRoot 'Worklist.ps1')
. (Join-Path $PSScriptRoot 'Table.ps1')
. (Join-Path $PSScriptRoot 'Ledger.ps1')
. (Join-Path $PSScriptRoot 'Context.ps1')
. (Join-Path $PSScriptRoot 'Profile.ps1')
. (Join-Path $PSScriptRoot 'Gate.ps1')

function Get-EbiDashboardWidth { return 80 }

function Get-EbiDashboardLeaf {
    # PURE. The last segment of a path, whichever separator wrote it (a
    # Windows path in a fixture is still a path on the Linux CI).
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $parts = $Path.Split([char[]]@('\', '/'))
    return [string]$parts[$parts.Count - 1]
}

function Get-EbiDashboardDirName {
    # The repo-root folder the menu remembers its last answers in (ignored by git: /.ebi/).
    return '.ebi'
}

# ---------------------------------------------------------------- phases

function Get-EbiDayPhases {
    <#
      PURE. The day's phases for one profile, in the order the day runs them.
        Workflows   @( @{ path; value } ) -- every workflow JSON that was read
        ProfileName the profile the dashboard is about
        Columns     the profile's worklist column specs (worklist.columns)
      A phase: @{ id (workflow id); name (the part after the last '.');
      title; path; file (leaf name); field (its source.select.field, ''
      for a workflow without a source); pendingWhen }.
      Order: no-source workflows first (by file name), then by the index of
      the field among the worklist columns; a field the worklist does not
      declare goes last, by name.
    #>
    param($Workflows, [string]$ProfileName, $Columns = @())
    $colIndex = @{}
    $i = 0
    foreach ($c in @($Columns)) {
        if ($c -is [System.Collections.IDictionary] -and $c.Contains('name')) { $colIndex[[string]$c['name']] = $i }
        $i++
    }
    $list = New-Object System.Collections.ArrayList
    foreach ($w in @($Workflows)) {
        if ($null -eq $w -or -not ($w -is [System.Collections.IDictionary]) -or -not $w.Contains('value')) { continue }
        $v = $w['value']
        if (-not ($v -is [System.Collections.IDictionary]) -or -not $v.Contains('id')) { continue }
        $p = if ($v.Contains('profile') -and $null -ne $v['profile']) { [string]$v['profile'] } else { '' }
        if ($p -ne $ProfileName) { continue }
        $id = [string]$v['id']
        $name = $id
        $dot = $id.LastIndexOf('.')
        if ($dot -ge 0 -and $dot -lt ($id.Length - 1)) { $name = $id.Substring($dot + 1) }
        $field = ''; $pw = ''
        if ($v.Contains('source') -and ($v['source'] -is [System.Collections.IDictionary]) -and $v['source'].Contains('select') -and ($v['source']['select'] -is [System.Collections.IDictionary])) {
            $sel = $v['source']['select']
            if ($sel.Contains('field') -and $null -ne $sel['field']) { $field = [string]$sel['field'] }
            if ($sel.Contains('pendingWhen') -and $null -ne $sel['pendingWhen']) { $pw = [string]$sel['pendingWhen'] }
        }
        $path = if ($w.Contains('path')) { [string]$w['path'] } else { '' }
        $file = Get-EbiDashboardLeaf -Path $path
        $title = if ($v.Contains('title') -and $null -ne $v['title']) { [string]$v['title'] } else { '' }
        $rank = if ($field -eq '') { -1 } elseif ($colIndex.Contains($field)) { [int]$colIndex[$field] } else { 100000 }
        [void]$list.Add(@{ id = $id; name = $name; title = $title; path = $path; file = $file; field = $field; pendingWhen = $pw; rank = $rank })
    }
    $sorted = @($list.ToArray() | Sort-Object -Property @{ Expression = { [int]$_['rank'] } }, @{ Expression = { [string]$_['file'] } }, @{ Expression = { [string]$_['field'] } })
    $out = New-Object System.Collections.ArrayList
    foreach ($s in $sorted) { $s.Remove('rank'); [void]$out.Add($s) }
    return $out.ToArray()
}

function Get-EbiDashboardVerdictPhases {
    # PURE. The phases that own a verdict column (the ones shown as columns).
    param($Phases)
    $out = New-Object System.Collections.ArrayList
    foreach ($p in @($Phases)) { if ($null -ne $p -and [string]$p['field'] -ne '') { [void]$out.Add($p) } }
    return $out.ToArray()
}

# ---------------------------------------------------------------- columns

function ConvertTo-EbiDashboardDay {
    # PURE. A worklist date cell -> 'yyyy-MM-dd', '' when it is not one
    # (kernel/Worklist.ps1 ConvertTo-EbiWorklistDay: the runner reads days the same way).
    param([string]$Text)
    return (ConvertTo-EbiWorklistDay -Text $Text)
}

function Get-EbiDashboardTodayRows {
    <#
      PURE. How many rows carry today's date in the profile's
      worklist.dashboard.dateColumn; -1 when the profile names no date
      column (or the CSV does not have it): "unknown", NEXT ignores it.
    #>
    param($Rows, $Profile, $Columns = @(), [string]$Today)
    $col = ''
    if ($null -ne $Profile -and ($Profile -is [System.Collections.IDictionary]) -and $Profile.Contains('worklist') -and ($Profile['worklist'] -is [System.Collections.IDictionary])) {
        $wl = $Profile['worklist']
        if ($wl.Contains('dashboard') -and ($wl['dashboard'] -is [System.Collections.IDictionary]) -and $wl['dashboard'].Contains('dateColumn')) { $col = [string]$wl['dashboard']['dateColumn'] }
    }
    $daySpec = Get-EbiWorklistDaySpec -Profile $Profile
    if ($null -ne $daySpec -and [string]$daySpec['column'] -ne $col) { $daySpec = $null }
    $have = (@($Columns) -contains $col) -or ($null -ne $daySpec -and [string]$daySpec['fallback'] -ne '' -and (@($Columns) -contains [string]$daySpec['fallback']))
    if ($col -eq '' -or -not $have) { return -1 }
    $day = ConvertTo-EbiDashboardDay -Text $Today
    $n = 0
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $d = if ($null -ne $daySpec) { Get-EbiWorklistRowDay -Row $r -DaySpec $daySpec } elseif ($r.Contains($col)) { ConvertTo-EbiDashboardDay -Text ([string]$r[$col]) } else { '' }
        if ($d -eq $day) { $n++ }
    }
    return $n
}

function Split-EbiDashboardRows {
    <#
      PURE. Which worklist rows the board is about (10/20). With a work-day
      column (role "day"): today's rows, not on hold; the rest is summed up
      under the table -- held rows with their reason, earlier days' rows
      with a verdict not ok yet, and how many earlier rows are all ok. A
      worklist without a day column shows every row that is not held.
      Returns @{ today (rows); scoped; held (lines); heldKeys; open (lines);
      openKeys; doneCount; later (lines: rows given a later day) }.
    #>
    param($Rows, $Profile, [string]$Today, $KeyColumns = @(), $VerdictFields = @())
    $daySpec = Get-EbiWorklistDaySpec -Profile $Profile
    $holdCol = Get-EbiWorklistRoleColumn -Profile $Profile -Role 'hold'
    $out = @{ today = (New-Object System.Collections.ArrayList); scoped = ($null -ne $daySpec); held = (New-Object System.Collections.ArrayList); heldKeys = (New-Object System.Collections.ArrayList); open = (New-Object System.Collections.ArrayList); openKeys = (New-Object System.Collections.ArrayList); doneCount = 0; later = (New-Object System.Collections.ArrayList) }
    $day = ConvertTo-EbiDashboardDay -Text $Today
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $key = Get-EbiKeyDisplay -Item $r -KeyColumns $KeyColumns
        $h = if ($holdCol -ne '' -and $r.Contains($holdCol) -and $null -ne $r[$holdCol]) { ([string]$r[$holdCol]).Trim() } else { '' }
        if ($h -ne '') { [void]$out['held'].Add(($key + ' (' + $h + ')')); [void]$out['heldKeys'].Add($key); continue }
        if ($null -eq $daySpec) { [void]$out['today'].Add($r); continue }
        $d = Get-EbiWorklistRowDay -Row $r -DaySpec $daySpec
        if ($d -eq $day) { [void]$out['today'].Add($r); continue }
        $parts = New-Object System.Collections.ArrayList
        $allOk = $true
        foreach ($f in @($VerdictFields)) {
            $v = if ($r.Contains([string]$f) -and $null -ne $r[[string]$f]) { ([string]$r[[string]$f]).Trim() } else { '' }
            $l = Get-EbiDashboardLogical -Value $v -Spec (Get-EbiWorklistColumnSpec -Profile $Profile -Field ([string]$f))
            if ($l -ne 'ok') { $allOk = $false }
            if ($l -eq 'ok') { [void]$parts.Add(([string]$f + ' ok')) } elseif ($l -eq 'ng') { [void]$parts.Add(([string]$f + ' NG')) }
        }
        $shown = if ($d -eq '') { 'no day' } else { $d.Substring(5).Replace('-', '/') }
        if ($d -ne '' -and $d -gt $day) { [void]$out['later'].Add(($key + ' (' + $shown + ')')); continue }
        if ($allOk) { $out['doneCount'] = [int]$out['doneCount'] + 1; continue }
        [void]$out['open'].Add(($key + ' (' + $shown + $(if ($parts.Count -gt 0) { ': ' + ($parts.ToArray() -join ', ') } else { '' }) + ')'))
        [void]$out['openKeys'].Add($key)
    }
    foreach ($k in @('today', 'held', 'heldKeys', 'open', 'openKeys', 'later')) { $out[$k] = $out[$k].ToArray() }
    return $out
}

function Format-EbiDashboardAside {
    # PURE. The lines under the table about the rows it does not show.
    param($Split, [int]$Width = 0)
    if ($Width -le 0) { $Width = Get-EbiDashboardWidth }
    $L = New-Object System.Collections.ArrayList
    if ($null -eq $Split) { return $L.ToArray() }
    $add = {
        param([string]$Text)
        $first = $true
        foreach ($w in @(ConvertTo-EbiGateWrapped -Text $Text -Width ($Width - 4))) { [void]$L.Add($(if ($first) { $w } else { '    ' + $w })); $first = $false }
    }
    if (@($Split['held']).Count -gt 0) { & $add ('on hold: ' + (@($Split['held']) -join ', ') + '  -- h releases') }
    if (@($Split['open']).Count -gt 0) { & $add ('earlier days, not finished: ' + (@($Split['open']) -join ', ') + '  -- plan offers them as late') }
    if (@($Split['later']).Count -gt 0) { & $add ('planned for a later day: ' + (@($Split['later']) -join ', ')) }
    if ([int]$Split['doneCount'] -gt 0) { [void]$L.Add(('earlier days, all ok: ' + [int]$Split['doneCount'] + ' row(s) (not shown)')) }
    return $L.ToArray()
}

function Get-EbiDashboardColumns {
    <#
      PURE. Which worklist columns to show between the key and the verdicts.
        Profile   the profile hashtable (worklist.dashboard.columns wins)
        Columns   the CSV's column names (a configured column that is not
                  in the CSV is dropped)
        Exclude   key columns + verdict fields (never shown twice)
    #>
    param($Profile, $Columns, $Exclude = @())
    $have = @(@($Columns) | ForEach-Object { [string]$_ })
    $skip = @(@($Exclude) | ForEach-Object { [string]$_ })
    $wl = $null
    if ($null -ne $Profile -and ($Profile -is [System.Collections.IDictionary]) -and $Profile.Contains('worklist') -and ($Profile['worklist'] -is [System.Collections.IDictionary])) { $wl = $Profile['worklist'] }
    $out = New-Object System.Collections.ArrayList
    if ($null -ne $wl -and $wl.Contains('dashboard') -and ($wl['dashboard'] -is [System.Collections.IDictionary]) -and $wl['dashboard'].Contains('columns') -and $null -ne $wl['dashboard']['columns']) {
        foreach ($c in @($wl['dashboard']['columns'])) {
            $n = [string]$c
            if ($n -ne '' -and ($have -contains $n) -and -not ($skip -contains $n) -and -not ($out -contains $n)) { [void]$out.Add($n) }
        }
        return $out.ToArray()
    }
    if ($null -ne $wl -and $wl.Contains('columns') -and $null -ne $wl['columns']) {
        foreach ($c in @($wl['columns'])) {
            if (-not ($c -is [System.Collections.IDictionary]) -or -not $c.Contains('name')) { continue }
            $n = [string]$c['name']
            $role = if ($c.Contains('role')) { [string]$c['role'] } else { 'text' }
            if ($role -in @('key', 'verdict', 'bitmask')) { continue }
            if (($have -contains $n) -and -not ($skip -contains $n) -and -not ($out -contains $n)) { [void]$out.Add($n) }
        }
        return $out.ToArray()
    }
    foreach ($n in $have) { if (-not ($skip -contains $n) -and -not ($out -contains $n)) { [void]$out.Add($n) } }
    return $out.ToArray()
}

# ---------------------------------------------------------------- model

function Get-EbiDashboardLogical {
    # PURE. One verdict cell -> 'ok' / 'ng' / 'pending' (blank, unknown, or anything else).
    param($Value, $Spec)
    $l = ConvertFrom-EbiStoredValue -Value $Value -Spec $Spec
    if ($l -eq 'ok') { return 'ok' }
    if ($l -eq 'ng') { return 'ng' }
    return 'pending'
}

function Get-EbiDashboardModel {
    <#
      PURE. Worklist rows -> one record per row:
        @{ key; cells (ordered values of ShowColumns); verdicts (field ->
           'ok' | 'ng' | 'pending'); next (@{ phase; field; action = 'run' |
           'clear' } or $null when every verdict is ok) }
      next = the FIRST verdict phase (day order) whose cell is not ok; an
      ng cell means "clear it and rerun" (action 'clear'), a blank or
      unknown one "run the phase" (action 'run').
    #>
    param($Rows, $Phases, $Profile, $KeyColumns = @(), $ShowColumns = @())
    $vp = @(Get-EbiDashboardVerdictPhases -Phases $Phases)
    $specs = @{}
    foreach ($p in $vp) { $specs[[string]$p['field']] = Get-EbiWorklistColumnSpec -Profile $Profile -Field ([string]$p['field']) }
    $out = New-Object System.Collections.ArrayList
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $key = if (@($KeyColumns).Count -gt 0) { Get-EbiKeyDisplay -Item $row -KeyColumns $KeyColumns } else { '' }
        $cells = New-Object System.Collections.ArrayList
        foreach ($c in @($ShowColumns)) { [void]$cells.Add($(if ($row.Contains($c) -and $null -ne $row[$c]) { [string]$row[$c] } else { '' })) }
        $verdicts = @{}
        $next = $null
        foreach ($p in $vp) {
            $f = [string]$p['field']
            $v = if ($row.Contains($f)) { $row[$f] } else { '' }
            $l = Get-EbiDashboardLogical -Value $v -Spec $specs[$f]
            $verdicts[$f] = $l
            if ($null -eq $next -and $l -ne 'ok') { $next = @{ phase = $p; field = $f; action = $(if ($l -eq 'ng') { 'clear' } else { 'run' }) } }
        }
        [void]$out.Add(@{ key = $key; cells = $cells.ToArray(); verdicts = $verdicts; next = $next })
    }
    return $out.ToArray()
}

function Get-EbiDashboardTotals {
    # PURE. model -> field -> @{ ok; ng; pending }
    param($Model, $Phases)
    $t = @{}
    foreach ($p in @(Get-EbiDashboardVerdictPhases -Phases $Phases)) { $t[[string]$p['field']] = @{ ok = 0; ng = 0; pending = 0 } }
    foreach ($m in @($Model)) {
        foreach ($f in @($m['verdicts'].Keys)) {
            if (-not $t.Contains($f)) { $t[$f] = @{ ok = 0; ng = 0; pending = 0 } }
            $l = [string]$m['verdicts'][$f]
            $t[$f][$l] = [int]$t[$f][$l] + 1
        }
    }
    return $t
}

function Get-EbiDashboardNext {
    <#
      PURE. What to do next, as lines, plus the one phase to run first.
        Model       Get-EbiDashboardModel
        Phases      Get-EbiDayPhases (the first no-source phase is "plan")
        Unfinished  Find-EbiUnfinishedRuns (run.json hashtables, newest first)
      Returns @{ lines; first = @{ kind = 'resume' | 'clear' | 'run' | 'plan'
      | 'done'; phase; runId; keys } }.
      Priority: an unfinished run (resume it) > no row dated today (plan;
      only when TodayRows is 0, -1 = no date column) > ng rows (clear and
      rerun) > the earliest phase with rows waiting > an empty worklist
      (plan) > nothing left.
    #>
    param($Model, $Phases, $Unfinished = @(), [int]$TodayRows = -1, [string]$Today = '', [int]$OtherRows = 0)
    $lines = New-Object System.Collections.ArrayList
    $first = $null
    $byId = @{}
    foreach ($p in @($Phases)) { $byId[[string]$p['id']] = $p }
    foreach ($u in @($Unfinished)) {
        if ($null -eq $u) { continue }
        $wid = if ($u.Contains('workflow') -and ($u['workflow'] -is [System.Collections.IDictionary]) -and $u['workflow'].Contains('id')) { [string]$u['workflow']['id'] } else { '' }
        $rid = if ($u.Contains('runId')) { [string]$u['runId'] } else { '' }
        $started = if ($u.Contains('startedAt')) { ConvertTo-EbiDashboardStamp -Value $u['startedAt'] } else { '' }
        $phase = if ($byId.Contains($wid)) { $byId[$wid] } else { $null }
        $label = if ($null -ne $phase) { [string]$phase['name'] } else { $wid }
        # the run id already carries the date: show only the clock when they agree
        $clock = $started
        if ($started -match '^(\d{4})-(\d{2})-(\d{2}) (\d{2}:\d{2}:\d{2})$' -and $rid.StartsWith($Matches[1] + $Matches[2] + $Matches[3])) { $clock = $Matches[4] }
        [void]$lines.Add(('unfinished ' + $label + ' run ' + $rid + ' (started ' + $clock + ') -> -Resume'))
        if ($null -eq $first) { $first = @{ kind = 'resume'; phase = $phase; runId = $rid; keys = @() } }
    }
    # 10/16: the worklist kept the earlier days' rows and NEXT said "track"
    # before plan had run that morning. With a date column (worklist.
    # dashboard.dateColumn) and no row dated today, plan comes first.
    if ($TodayRows -eq 0 -and @($Model).Count -gt 0) {
        $plan0 = $null
        foreach ($p in @($Phases)) { if ([string]$p['field'] -eq '') { $plan0 = $p; break } }
        if ($null -ne $plan0) {
            [void]$lines.Add(('no row is dated today' + $(if ($Today -ne '') { ' (' + $Today + ')' } else { '' }) + ' -> run ' + [string]$plan0['name'] + ' first'))
            if ($null -eq $first) { $first = @{ kind = 'plan'; phase = $plan0; runId = ''; keys = @() } }
        }
    }
    $vp = @(Get-EbiDashboardVerdictPhases -Phases $Phases)
    # ng rows, by phase, in day order
    foreach ($p in $vp) {
        $f = [string]$p['field']
        $keys = @(@($Model) | Where-Object { $null -ne $_['next'] -and [string]$_['next']['field'] -eq $f -and [string]$_['next']['action'] -eq 'clear' } | ForEach-Object { [string]$_['key'] })
        if ($keys.Count -eq 0) { continue }
        [void]$lines.Add(($f + ' NG: ' + ($keys -join ', ') + ' -> clear the cell and rerun ' + [string]$p['name']))
        if ($null -eq $first) { $first = @{ kind = 'clear'; phase = $p; runId = ''; keys = $keys } }
    }
    foreach ($p in $vp) {
        $f = [string]$p['field']
        $keys = @(@($Model) | Where-Object { $null -ne $_['next'] -and [string]$_['next']['field'] -eq $f -and [string]$_['next']['action'] -eq 'run' } | ForEach-Object { [string]$_['key'] })
        if ($keys.Count -eq 0) { continue }
        $n = $keys.Count
        [void]$lines.Add(([string]$p['name'] + ': ' + $n + ' row' + $(if ($n -ne 1) { 's' } else { '' }) + ' wait' + $(if ($n -eq 1) { 's' } else { '' }) + ' (' + ($keys -join ', ') + ')'))
        if ($null -eq $first) { $first = @{ kind = 'run'; phase = $p; runId = ''; keys = $keys } }
    }
    if (@($Model).Count -eq 0) {
        $plan = $null
        foreach ($p in @($Phases)) { if ([string]$p['field'] -eq '') { $plan = $p; break } }
        if ($null -ne $plan -and $OtherRows -gt 0) {
            # 10/20: the board shows today's rows only; the worklist itself is not empty
            [void]$lines.Add(('nothing is planned for today' + $(if ($Today -ne '') { ' (' + $Today + ')' } else { '' }) + ' -> run ' + [string]$plan['name'] + ' first (it also offers the late and next-day jobs)'))
            if ($null -eq $first) { $first = @{ kind = 'plan'; phase = $plan; runId = ''; keys = @() } }
        } elseif ($null -ne $plan) {
            [void]$lines.Add(('the worklist is empty -> run ' + [string]$plan['name'] + ' first'))
            if ($null -eq $first) { $first = @{ kind = 'plan'; phase = $plan; runId = ''; keys = @() } }
        } else {
            [void]$lines.Add('the worklist is empty')
        }
    }
    if ($null -eq $first) {
        [void]$lines.Add('nothing left: every row is ok in every phase')
        $first = @{ kind = 'done'; phase = $null; runId = ''; keys = @() }
    }
    return @{ lines = $lines.ToArray(); first = $first }
}

function ConvertTo-EbiDashboardStamp {
    # PURE. A run.json timestamp -> 'yyyy-MM-dd HH:mm:ss' (pwsh 7's JSON reader hands back a DateTime, 5.1 a string).
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss') }
    $t = [string]$Value
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($t, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt.ToString('yyyy-MM-dd HH:mm:ss') }
    return $t
}

# ---------------------------------------------------------------- run line

function New-EbiRunCommand {
    <#
      PURE. The ebi.ps1 run line for one phase, as the operator would type
      it (text) and as the argument list a child PowerShell gets (args).
        Phase     a Get-EbiDayPhases record (path = the workflow file)
        ToolDir   where ebi.ps1 lives (the workflow path is shown relative
                  to it when it is under it)
        WorkDir   -WorkDir
        Only      key display strings for -Only (empty = every pending row)
        Resume    $true -> -Resume [-RunId <RunId>]
    #>
    param($Phase, [string]$ToolDir, [string]$WorkDir, $Only = @(), [bool]$Resume = $false, [string]$RunId = '', [bool]$AllDays = $false)
    $wf = [string]$Phase['path']
    $shown = $wf
    if ($ToolDir -ne '' -and $wf.StartsWith($ToolDir, [System.StringComparison]::OrdinalIgnoreCase)) {
        $shown = $wf.Substring($ToolDir.Length).TrimStart('\', '/')
    }
    $argList = New-Object System.Collections.ArrayList
    [void]$argList.Add('run'); [void]$argList.Add($wf)
    $text = 'ebi.ps1 run ' + $shown
    if ($WorkDir -ne '') { [void]$argList.Add('-WorkDir'); [void]$argList.Add($WorkDir); $text += ' -WorkDir ' + $(if ($WorkDir.Contains(' ')) { '"' + $WorkDir + '"' } else { $WorkDir }) }
    $only = @(@($Only) | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -ne '' })
    if ($only.Count -gt 0) { [void]$argList.Add('-Only'); [void]$argList.Add(($only -join ',')); $text += ' -Only ' + ($only -join ',') }
    if ($AllDays -and $only.Count -eq 0) { [void]$argList.Add('-AllDays'); $text += ' -AllDays' }
    if ($Resume) {
        [void]$argList.Add('-Resume'); $text += ' -Resume'
        if ($RunId -ne '') { [void]$argList.Add('-RunId'); [void]$argList.Add($RunId); $text += ' -RunId ' + $RunId }
    }
    return @{ text = $text; args = $argList.ToArray() }
}

# ---------------------------------------------------------------- render

function ConvertTo-EbiDashboardCell {
    # PURE. Fit one value into a column: cut with '~' when too long.
    param([string]$Text, [int]$Width)
    if ($null -eq $Text) { $Text = '' }
    if ($Text.Length -le $Width) { return $Text.PadRight($Width) }
    if ($Width -le 1) { return $Text.Substring(0, $Width) }
    return ($Text.Substring(0, $Width - 1) + '~')
}

function Get-EbiDashboardVerdictText {
    # PURE. 'ok' -> 'ok', 'ng' -> 'NG', pending -> '-'
    param([string]$Logical)
    if ($Logical -eq 'ok') { return 'ok' }
    if ($Logical -eq 'ng') { return 'NG' }
    return '-'
}

function Format-EbiDashboard {
    <#
      PURE. Everything -> lines, each at most Width characters.
        Title        'gfix-recv' (profile name)
        WorkDir      shown in the header
        WorklistPath shown in the header (leaf name)
        Model / Phases / Totals / Next / Unfinished  the pure results
        ShowColumns  the column names between the key and the verdicts
        KeyHeader    the key column header (first key column name)
        RunLine      the recommended command (one line) or ''
      Layout: header, table (key, shown columns, one column per verdict
      phase), totals, UNFINISHED RUNS / NEXT. When the table would be
      wider than Width, shown columns are dropped from the right; the
      key and the verdicts always stay.
    #>
    param([string]$Title, [string]$WorkDir, [string]$WorklistPath, $Model, $Phases, $Totals, $Next, $ShowColumns = @(), [string]$KeyHeader = 'key', [string]$RunLine = '', [string]$Today = '', [int]$Width = 0, $Aside = @(), [int]$AllRows = -1)
    if ($Width -le 0) { $Width = Get-EbiDashboardWidth }
    $L = New-Object System.Collections.ArrayList
    $rows = @($Model)
    $vp = @(Get-EbiDashboardVerdictPhases -Phases $Phases)
    $leaf = Get-EbiDashboardLeaf -Path $WorklistPath

    $head = 'ebi status: ' + $Title
    if ($Today -ne '') { $head += '  ' + $Today }
    if ($WorkDir -ne '') { $head += '  ' + $WorkDir }
    [void]$L.Add((ConvertTo-EbiHelpFitDashboard -Text $head -Width $Width))
    $wlLine = 'worklist: ' + $(if ($leaf -ne '') { $leaf } else { '(none)' }) + '  ' + $(if ($AllRows -ge 0) { ('' + $AllRows + ' row' + $(if ($AllRows -ne 1) { 's' } else { '' }) + ', today ' + $rows.Count) } else { ('' + $rows.Count + ' row' + $(if ($rows.Count -ne 1) { 's' } else { '' })) })
    [void]$L.Add((ConvertTo-EbiHelpFitDashboard -Text $wlLine -Width $Width))
    [void]$L.Add('-' * $Width)

    # column widths
    $keyW = $KeyHeader.Length
    foreach ($m in $rows) { if (([string]$m['key']).Length -gt $keyW) { $keyW = ([string]$m['key']).Length } }
    $vW = @()
    foreach ($p in $vp) {
        $w = ([string]$p['field']).Length; if ($w -lt 4) { $w = 4 }
        # wide enough for its totals too (10/16: "ok 21" under "logs" showed as "ok ~")
        $f = [string]$p['field']
        foreach ($kind in @('ok', 'ng', 'pending')) {
            $label = if ($kind -eq 'ok') { 'ok' } elseif ($kind -eq 'ng') { 'NG' } else { '-' }
            $n = if ($null -ne $Totals -and $Totals.Contains($f)) { [int]$Totals[$f][$kind] } else { 0 }
            $tl = ($label + ' ' + $n).Length
            if ($tl -gt $w) { $w = $tl }
        }
        $vW += $w
    }
    $show = @(@($ShowColumns) | ForEach-Object { [string]$_ })
    $sW = @()
    for ($i = 0; $i -lt $show.Count; $i++) {
        $w = $show[$i].Length
        foreach ($m in $rows) { $c = [string]$m['cells'][$i]; if ($c.Length -gt $w) { $w = $c.Length } }
        if ($w -gt 20) { $w = 20 }
        $sW += $w
    }
    $fixed = $keyW
    foreach ($w in $vW) { $fixed += 2 + $w }
    while ($show.Count -gt 0) {
        $total = $fixed
        for ($i = 0; $i -lt $show.Count; $i++) { $total += 2 + $sW[$i] }
        if ($total -le $Width) { break }
        $show = @($show | Select-Object -First ($show.Count - 1))
        $sW = @($sW | Select-Object -First $show.Count)
    }
    $line = $KeyHeader.PadRight($keyW)
    for ($i = 0; $i -lt $show.Count; $i++) { $line += '  ' + (ConvertTo-EbiDashboardCell -Text $show[$i] -Width $sW[$i]) }
    for ($i = 0; $i -lt $vp.Count; $i++) { $line += '  ' + (ConvertTo-EbiDashboardCell -Text ([string]$vp[$i]['field']) -Width $vW[$i]) }
    [void]$L.Add($line.TrimEnd())
    foreach ($m in $rows) {
        $line = ([string]$m['key']).PadRight($keyW)
        for ($i = 0; $i -lt $show.Count; $i++) { $line += '  ' + (ConvertTo-EbiDashboardCell -Text ([string]$m['cells'][$i]) -Width $sW[$i]) }
        for ($i = 0; $i -lt $vp.Count; $i++) { $line += '  ' + (ConvertTo-EbiDashboardCell -Text (Get-EbiDashboardVerdictText -Logical ([string]$m['verdicts'][[string]$vp[$i]['field']])) -Width $vW[$i]) }
        [void]$L.Add($line.TrimEnd())
    }
    [void]$L.Add('-' * $Width)
    if ($vp.Count -gt 0) {
        $pad = $keyW
        for ($i = 0; $i -lt $show.Count; $i++) { $pad += 2 + $sW[$i] }
        foreach ($kind in @('ok', 'ng', 'pending')) {
            $label = if ($kind -eq 'ok') { 'ok' } elseif ($kind -eq 'ng') { 'NG' } else { '-' }
            $line = $(if ($kind -eq 'ok') { ('total ' + $rows.Count) } else { '' }).PadRight($pad)
            for ($i = 0; $i -lt $vp.Count; $i++) {
                $f = [string]$vp[$i]['field']
                $n = if ($null -ne $Totals -and $Totals.Contains($f)) { [int]$Totals[$f][$kind] } else { 0 }
                $line += '  ' + (ConvertTo-EbiDashboardCell -Text (($label + ' ' + $n)) -Width $vW[$i])
            }
            [void]$L.Add($line.TrimEnd())
        }
    }
    $aside = @(@($Aside) | Where-Object { [string]$_ -ne '' })
    if ($aside.Count -gt 0) {
        [void]$L.Add('')
        foreach ($a in $aside) { [void]$L.Add((ConvertTo-EbiHelpFitDashboard -Text ([string]$a) -Width $Width)) }
    }
    if ($null -ne $Next -and @($Next['lines']).Count -gt 0) {
        [void]$L.Add('')
        [void]$L.Add('NEXT')
        foreach ($t in @($Next['lines'])) {
            foreach ($w in @(ConvertTo-EbiGateWrapped -Text $t -Width ($Width - 2))) { [void]$L.Add('  ' + $w) }
        }
    }
    if ($RunLine -ne '') {
        [void]$L.Add('')
        foreach ($w in @(ConvertTo-EbiGateWrapped -Text $RunLine -Width ($Width - 2))) { [void]$L.Add('  ' + $w) }
    }
    return $L.ToArray()
}

function Get-EbiDashboardCharWidth {
    # PURE. Console cells one character takes: 2 for East Asian wide / full-width
    # characters (the workflow titles are Japanese), else 1.
    param([char]$C)
    $n = [int]$C
    if (($n -ge 0x1100 -and $n -le 0x115F) -or ($n -ge 0x2E80 -and $n -le 0xA4CF) -or ($n -ge 0xAC00 -and $n -le 0xD7A3) -or ($n -ge 0xF900 -and $n -le 0xFAFF) -or ($n -ge 0xFE30 -and $n -le 0xFE4F) -or ($n -ge 0xFF00 -and $n -le 0xFF60) -or ($n -ge 0xFFE0 -and $n -le 0xFFE6)) { return 2 }
    return 1
}

function ConvertTo-EbiHelpFitDashboard {
    # PURE. One line, cut to the width (in console cells, wide characters
    # count twice) with '...'.
    param([string]$Text, [int]$Width)
    if ($null -eq $Text) { return '' }
    $cells = 0
    foreach ($ch in $Text.ToCharArray()) { $cells += Get-EbiDashboardCharWidth -C $ch }
    if ($cells -le $Width) { return $Text }
    $sb = New-Object System.Text.StringBuilder
    $used = 0
    foreach ($ch in $Text.ToCharArray()) {
        $w = Get-EbiDashboardCharWidth -C $ch
        if (($used + $w) -gt ($Width - 3)) { break }
        [void]$sb.Append($ch); $used += $w
    }
    return ($sb.ToString() + '...')
}

# ---------------------------------------------------------------- loader

function Get-EbiDayWorkflows {
    # workflows/*.json -> @( @{ path; value } ), unreadable files skipped (listed in 'bad').
    param([string]$Dir)
    $out = New-Object System.Collections.ArrayList
    $bad = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($Dir) -or -not (Test-Path -LiteralPath $Dir -PathType Container)) { return @{ workflows = $out.ToArray(); bad = $bad.ToArray() } }
    foreach ($f in @(Get-ChildItem -LiteralPath $Dir -Filter '*.json' -File | Sort-Object Name)) {
        $r = Read-EbiJson -Path $f.FullName
        if (-not $r['ok'] -or -not ($r['value'] -is [System.Collections.IDictionary])) { [void]$bad.Add($f.FullName); continue }
        [void]$out.Add(@{ path = $f.FullName; value = $r['value'] })
    }
    return @{ workflows = $out.ToArray(); bad = $bad.ToArray() }
}

function Resolve-EbiWorklistFile {
    # profile.worklist.file ("mapping_{{run.operator}}.csv") -> absolute path under WorkDir.
    param($Profile, [string]$WorkDir, [string]$Operator = '')
    $file = ''
    if ($null -ne $Profile -and ($Profile -is [System.Collections.IDictionary]) -and $Profile.Contains('worklist') -and ($Profile['worklist'] -is [System.Collections.IDictionary]) -and $Profile['worklist'].Contains('file')) { $file = [string]$Profile['worklist']['file'] }
    if ($file -eq '') { return '' }
    $op = if ($Operator -ne '') { $Operator } elseif ($env:USERNAME) { [string]$env:USERNAME } else { 'operator' }
    $scope = New-EbiTemplateScope -Profile $(if ($Profile -is [hashtable]) { $Profile } else { @{} }) -Run @{ operator = $op }
    $x = Expand-EbiTemplate -Value $file -Scope $scope
    if ($x['ok'] -and $null -ne $x['value']) { $file = [string]$x['value'] }
    if ([System.IO.Path]::IsPathRooted($file)) { return $file }
    return (Join-Path $WorkDir $file)
}

function Get-EbiPlanProfiles {
    <#
      The profiles (names) that have a profiles/<name>/ folder AND a
      workflow without a row source naming them (the "plan" that builds
      the worklist) -- the ones a fresh work folder can start a day with.
    #>
    param([string]$ToolDir)
    $names = New-Object System.Collections.ArrayList
    $wfs = Get-EbiDayWorkflows -Dir (Join-Path $ToolDir 'workflows')
    foreach ($w in @($wfs['workflows'])) {
        $v = $w['value']
        if (-not $v.Contains('profile') -or $null -eq $v['profile'] -or [string]$v['profile'] -eq '') { continue }
        $hasSource = $v.Contains('source') -and ($v['source'] -is [System.Collections.IDictionary]) -and $v['source'].Contains('select')
        if ($hasSource) { continue }
        $n = [string]$v['profile']
        if (-not (Test-Path -LiteralPath (Join-Path (Join-Path $ToolDir 'profiles') $n) -PathType Container)) { continue }
        if (-not ($names -contains $n)) { [void]$names.Add($n) }
    }
    return @($names.ToArray() | Sort-Object)
}

function Find-EbiDayProfiles {
    # Every profile under ProfilesRoot whose worklist file exists in WorkDir
    # (or that has a worklist at all when -All), plus every profile that
    # could not be READ (error = the reason, worklist = ''): the caller
    # tells those apart from "found" by the error field.
    param([string]$ProfilesRoot, [string]$WorkDir, [string]$Operator = '', [switch]$All)
    $out = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $ProfilesRoot -PathType Container)) { return $out.ToArray() }
    foreach ($d in @(Get-ChildItem -LiteralPath $ProfilesRoot -Directory | Sort-Object Name)) {
        $pr = Read-EbiProfile -Dir $d.FullName -WorkDir $WorkDir
        # 10/16: a broken <WorkDir>\ebi.local.json (one missing comma) made
        # every profile fail here, and the caller said "no profile has its
        # worklist" -- keep the reason so it can say what is wrong instead
        if (-not $pr['ok']) { [void]$out.Add(@{ name = $d.Name; dir = $d.FullName; worklist = ''; error = [string]$pr['message'] }); continue }
        $path = Resolve-EbiWorklistFile -Profile $pr['value'] -WorkDir $WorkDir -Operator $Operator
        if ($path -eq '') { continue }
        if ($All -or (Test-Path -LiteralPath $path -PathType Leaf)) { [void]$out.Add(@{ name = $d.Name; dir = $d.FullName; worklist = $path; error = '' }) }
    }
    return $out.ToArray()
}

function Read-EbiDashboard {
    <#
      Load everything the dashboard needs and run the pure half.
        ToolDir      the repo root (workflows/ and profiles/ under it)
        WorkDir      the work folder (worklist CSV, run/)
        ProfileName  '' = the one profile whose worklist exists in WorkDir
      Returns @{ ok; message; profile; profileName; worklistPath; phases;
      model; totals; next; unfinished; showColumns; keyColumns; lines;
      runLine; runCommand }.
    #>
    param([string]$ToolDir, [string]$WorkDir, [string]$ProfileName = '', [string]$Operator = '', [string]$Today = '')
    $res = @{ ok = $false; message = ''; profile = $null; profileName = ''; worklistPath = ''; phases = @(); model = @(); totals = @{}; next = $null; unfinished = @(); showColumns = @(); keyColumns = @(); lines = @(); runLine = ''; runCommand = $null; split = $null; rows = @(); columns = @(); allKeys = @(); today = '' }
    $profilesRoot = Join-Path $ToolDir 'profiles'
    $dir = ''
    # 10/16: a bare double-click on ebi.cmd with nothing remembered ran in
    # the tool folder itself; say so instead of "no profile has its worklist"
    try {
        $wdFull = [System.IO.Path]::GetFullPath($WorkDir).TrimEnd('\', '/')
        $tdFull = [System.IO.Path]::GetFullPath($ToolDir).TrimEnd('\', '/')
        if ($wdFull -ieq $tdFull) { $res['message'] = ('the work folder is the tool folder (' + $ToolDir + '): give the work folder -- ebi.cmd <work folder>, drag the folder onto ebi.cmd, or -WorkDir <dir>'); return $res }
    } catch { }
    if (-not [string]::IsNullOrWhiteSpace($ProfileName)) {
        $dir = Resolve-EbiProfileDir -NameOrPath $ProfileName -ProfilesRoot $profilesRoot
        if ($dir -eq '' -or -not (Test-Path -LiteralPath $dir -PathType Container)) { $res['message'] = ('profile directory not found: ' + $ProfileName); return $res }
    } else {
        $all = @(Find-EbiDayProfiles -ProfilesRoot $profilesRoot -WorkDir $WorkDir -Operator $Operator)
        $found = @($all | Where-Object { [string]$_['error'] -eq '' })
        $broken = @($all | Where-Object { [string]$_['error'] -ne '' })
        if ($found.Count -eq 1) { $dir = [string]$found[0]['dir'] }
        elseif ($found.Count -eq 0 -and $broken.Count -gt 0) { $res['message'] = ('could not read the profile(s) -- fix this first: ' + ((@($broken | ForEach-Object { [string]$_['name'] + ': ' + [string]$_['error'] })) -join ' | ')); return $res }
        elseif ($found.Count -eq 0 -and @(Get-EbiPlanProfiles -ToolDir $ToolDir).Count -eq 1) {
            # a fresh work folder: no worklist yet, but only one profile starts
            # its day with a plan workflow -> that one (an empty board: run plan)
            $dir = Resolve-EbiProfileDir -NameOrPath ([string]@(Get-EbiPlanProfiles -ToolDir $ToolDir)[0]) -ProfilesRoot $profilesRoot
        }
        elseif ($found.Count -eq 0) { $res['message'] = ('no profile has its worklist in ' + $WorkDir + ' (is this the work folder? ebi.cmd <work folder>, or drag the folder onto ebi.cmd); or say which profile with -Profile <name>'); return $res }
        else { $res['message'] = ('several profiles have a worklist in ' + $WorkDir + ' (' + (@($found | ForEach-Object { [string]$_['name'] }) -join ', ') + '); say which one with -Profile <name>'); return $res }
    }
    $pr = Read-EbiProfile -Dir $dir -WorkDir $WorkDir
    if (-not $pr['ok']) { $res['message'] = $pr['message']; return $res }
    $profile = $pr['value']
    $res['profile'] = $profile; $res['profileName'] = [string]$pr['name']
    $wlPath = Resolve-EbiWorklistFile -Profile $profile -WorkDir $WorkDir -Operator $Operator
    $res['worklistPath'] = $wlPath
    $columns = @(); $rows = @()
    if ($wlPath -ne '') {
        $read = Read-EbiCsv -Path $wlPath
        if ($read['exists'] -and -not $read['ok']) { $res['message'] = $read['message']; return $res }
        if ($read['ok']) { $columns = @($read['columns']); $rows = @($read['rows']) }
    }
    $wfs = Get-EbiDayWorkflows -Dir (Join-Path $ToolDir 'workflows')
    $colSpecs = @()
    if ($profile.Contains('worklist') -and ($profile['worklist'] -is [System.Collections.IDictionary]) -and $profile['worklist'].Contains('columns')) { $colSpecs = @($profile['worklist']['columns']) }
    $phases = @(Get-EbiDayPhases -Workflows $wfs['workflows'] -ProfileName $res['profileName'] -Columns $colSpecs)
    $keyColumns = @(Get-EbiWorklistKeyColumns -Profile $profile)
    $exclude = @($keyColumns) + @(@(Get-EbiDashboardVerdictPhases -Phases $phases) | ForEach-Object { [string]$_['field'] })
    # a verdict column the CSV does not have yet still shows (blank = pending)
    $allColumns = @($columns)
    foreach ($p in @(Get-EbiDashboardVerdictPhases -Phases $phases)) { if (-not ($allColumns -contains [string]$p['field'])) { $allColumns += [string]$p['field'] } }
    $show = @(Get-EbiDashboardColumns -Profile $profile -Columns $allColumns -Exclude $exclude)
    $today = if ($Today -ne '') { $Today } else { (Get-Date).ToString('yyyy-MM-dd') }
    $verdictFields = @(@(Get-EbiDashboardVerdictPhases -Phases $phases) | ForEach-Object { [string]$_['field'] })
    $split = Split-EbiDashboardRows -Rows $rows -Profile $profile -Today $today -KeyColumns $keyColumns -VerdictFields $verdictFields
    $boardRows = @($split['today'])
    $model = @(Get-EbiDashboardModel -Rows $boardRows -Phases $phases -Profile $profile -KeyColumns $keyColumns -ShowColumns $show)
    $totals = Get-EbiDashboardTotals -Model $model -Phases $phases
    $unfinished = @(Find-EbiUnfinishedRuns -WorkDir $WorkDir)
    $todayRows = Get-EbiDashboardTodayRows -Rows $rows -Profile $profile -Columns $columns -Today $today
    if ([bool]$split['scoped']) { $todayRows = $boardRows.Count }
    $next = Get-EbiDashboardNext -Model $model -Phases $phases -Unfinished $unfinished -TodayRows $todayRows -Today $today -OtherRows ($rows.Count - $boardRows.Count)
    $runLine = ''; $runCommand = $null
    $first = $next['first']
    if ($null -ne $first -and $null -ne $first['phase']) {
        if ($first['kind'] -eq 'resume') { $runCommand = New-EbiRunCommand -Phase $first['phase'] -ToolDir $ToolDir -WorkDir $WorkDir -Resume $true -RunId ([string]$first['runId']) }
        elseif ($first['kind'] -eq 'run' -or $first['kind'] -eq 'plan') { $runCommand = New-EbiRunCommand -Phase $first['phase'] -ToolDir $ToolDir -WorkDir $WorkDir }
        if ($null -ne $runCommand) { $runLine = 'next: ' + $runCommand['text'] }
    }
    $keyHeader = if ($keyColumns.Count -gt 0) { [string]$keyColumns[0] } else { 'key' }
    $res['phases'] = $phases; $res['model'] = $model; $res['totals'] = $totals; $res['next'] = $next; $res['unfinished'] = $unfinished
    $res['showColumns'] = $show; $res['keyColumns'] = $keyColumns; $res['runLine'] = $runLine; $res['runCommand'] = $runCommand
    $res['split'] = $split; $res['rows'] = $rows; $res['columns'] = $columns; $res['today'] = $today
    $res['allKeys'] = @(foreach ($r in $rows) { if ($null -ne $r) { Get-EbiKeyDisplay -Item $r -KeyColumns $keyColumns } })
    $aside = @(Format-EbiDashboardAside -Split $split)
    $allRows = if ([bool]$split['scoped'] -or @($split['held']).Count -gt 0) { $rows.Count } else { -1 }
    $res['lines'] = @(Format-EbiDashboard -Title $res['profileName'] -WorkDir $WorkDir -WorklistPath $wlPath -Model $model -Phases $phases -Totals $totals -Next $next -ShowColumns $show -KeyHeader $keyHeader -RunLine $runLine -Today $today -Aside $aside -AllRows $allRows)
    $res['ok'] = $true
    return $res
}

function Write-EbiDashboardLines {
    # Print the lines with the few colours that help (NG red, ok green, NEXT yellow).
    param($Lines)
    foreach ($l in @($Lines)) {
        $color = 'Gray'
        if ($l -eq 'NEXT' -or $l -like '  next: *' -or $l -like '  unfinished run*') { $color = 'Yellow' }
        elseif ($l -match '\bNG\b(?! 0\b)') { $color = 'Red' }      # a row with an NG cell, a non-zero NG total, an NG line in NEXT
        elseif ($l -like 'ebi status:*' -or $l -like 'ebi menu:*') { $color = 'Cyan' }
        elseif ($l -like 'total *' -or $l -like 'worklist:*' -or $l -like 'earlier days, all ok*') { $color = 'DarkGray' }
        elseif ($l -like 'on hold:*' -or $l -like 'earlier days, not finished*') { $color = 'DarkYellow' }
        Write-Host ('  ' + $l) -ForegroundColor $color
    }
}

# ---------------------------------------------------------------- the one write

function Clear-EbiWorklistVerdict {
    <#
      Blank one row's verdict cell so the phase runs it again ("clear the
      cell in Notepad" without Notepad). Atomic write (Write-EbiCsvAtomic).
      Returns @{ ok; message; before }.
    #>
    param([string]$Path, [string]$Key, [string]$Field, $KeyColumns = @(), [bool]$DryRun = $false)
    $read = Read-EbiCsv -Path $Path
    if (-not $read['ok']) { return @{ ok = $false; message = $read['message']; before = '' } }
    if (-not (@($read['columns']) -contains $Field)) { return @{ ok = $false; message = ('column "' + $Field + '" is not in ' + $Path); before = '' } }
    $hit = $null
    foreach ($row in @($read['rows'])) {
        if ($null -eq $row) { continue }
        if ((Get-EbiKeyDisplay -Item $row -KeyColumns $KeyColumns) -eq $Key.Trim()) { $hit = $row; break }
    }
    if ($null -eq $hit) {
        # second chance: the folded key forms (full-width, case)
        $m = Find-EbiKeyMatches -Records @($read['rows']) -Key $Key.Trim() -KeyColumns $KeyColumns
        if (@($m['records']).Count -eq 1) { $hit = $m['records'][0] }
    }
    if ($null -eq $hit) { return @{ ok = $false; message = ('no row with key "' + $Key + '" in ' + $Path); before = '' } }
    $before = if ($hit.Contains($Field) -and $null -ne $hit[$Field]) { [string]$hit[$Field] } else { '' }
    if ($DryRun) { return @{ ok = $true; message = ('would clear ' + $Key + ': ' + $Field + ' (was "' + $before + '")'); before = $before } }
    $hit[$Field] = ''
    $w = Write-EbiCsvAtomic -Path $Path -Columns $read['columns'] -Rows $read['rows']
    if (-not $w['ok']) { return @{ ok = $false; message = $w['message']; before = $before } }
    return @{ ok = $true; message = ('cleared ' + $Key + ': ' + $Field + ' (was "' + $before + '")'); before = $before }
}

function Set-EbiWorklistCells {
    <#
      Write one value into one column of several rows (menu m = ok by hand,
      h = hold reason / release, t = WORKDAY today). Atomic write. A column
      the CSV does not have yet is added when AddColumn (hold / WORKDAY in
      an older worklist). Returns @{ ok; message; changed ("KEY: a -> b") }.
    #>
    param([string]$Path, $Keys, [string]$Field, [string]$Value, $KeyColumns = @(), [bool]$AddColumn = $false, [bool]$DryRun = $false)
    $read = Read-EbiCsv -Path $Path
    if (-not $read['ok']) { return @{ ok = $false; message = $read['message']; changed = @() } }
    $columns = @($read['columns'])
    if (-not ($columns -contains $Field)) {
        if (-not $AddColumn) { return @{ ok = $false; message = ('column "' + $Field + '" is not in ' + $Path); changed = @() } }
        $columns += $Field
        foreach ($row in @($read['rows'])) { if ($null -ne $row -and -not $row.Contains($Field)) { $row[$Field] = '' } }
    }
    $changed = New-Object System.Collections.ArrayList
    foreach ($k in @($Keys)) {
        $want = ([string]$k).Trim()
        $hit = $null
        foreach ($row in @($read['rows'])) { if ($null -ne $row -and (Get-EbiKeyDisplay -Item $row -KeyColumns $KeyColumns) -eq $want) { $hit = $row; break } }
        if ($null -eq $hit) {
            $m = Find-EbiKeyMatches -Records @($read['rows']) -Key $want -KeyColumns $KeyColumns
            if (@($m['records']).Count -eq 1) { $hit = $m['records'][0] }
        }
        if ($null -eq $hit) { return @{ ok = $false; message = ('no row with key "' + $want + '" in ' + $Path); changed = @() } }
        $before = if ($hit.Contains($Field) -and $null -ne $hit[$Field]) { [string]$hit[$Field] } else { '' }
        if ($before -eq $Value) { continue }
        [void]$changed.Add(($want + ': ' + $Field + ' "' + $before + '" -> "' + $Value + '"'))
        if (-not $DryRun) { $hit[$Field] = $Value }
    }
    if ($DryRun -or $changed.Count -eq 0) { return @{ ok = $true; message = $(if ($changed.Count -eq 0) { 'nothing to change' } else { 'would change ' + $changed.Count + ' row(s)' }); changed = $changed.ToArray() } }
    $w = Write-EbiCsvAtomic -Path $Path -Columns $columns -Rows $read['rows']
    if (-not $w['ok']) { return @{ ok = $false; message = $w['message']; changed = @() } }
    return @{ ok = $true; message = ('changed ' + $changed.Count + ' row(s)'); changed = $changed.ToArray() }
}

function Resolve-EbiMenuKeys {
    <#
      PURE. What the operator typed for "which rows?" -> @{ ok; keys; message }.
      Tokens split on spaces / commas; each is a whole key (any case,
      full-width folded) or a PART of keys ("DBC" = every key holding DBC).
      An unknown token is refused, nothing is half-taken.
    #>
    param([string]$Text, $Keys)
    $t = (ConvertTo-EbiHalfWidth -Value $(if ($null -eq $Text) { '' } else { $Text })).Trim()
    if ($t -eq '') { return @{ ok = $true; keys = @(); message = '' } }
    $pool = @(@($Keys) | ForEach-Object { [string]$_ } | Where-Object { $_ -ne '' })
    $norm = { param($v) (ConvertTo-EbiHalfWidth -Value ([string]$v)).Trim().ToUpperInvariant() }
    $chosen = New-Object System.Collections.ArrayList
    $bad = New-Object System.Collections.ArrayList
    foreach ($tok in @($t -split '[\s,;]+' | Where-Object { $_ -ne '' })) {
        $n = & $norm $tok
        $hits = @($pool | Where-Object { (& $norm $_) -eq $n })
        if ($hits.Count -eq 0) { $hits = @($pool | Where-Object { (& $norm $_).Contains($n) }) }
        if ($hits.Count -eq 0) { [void]$bad.Add($tok); continue }
        foreach ($h in $hits) { if (-not $chosen.Contains($h)) { [void]$chosen.Add($h) } }
    }
    if ($bad.Count -gt 0) { return @{ ok = $false; keys = @(); message = ('not in the worklist: ' + ($bad.ToArray() -join ', ')) } }
    return @{ ok = $true; keys = $chosen.ToArray(); message = '' }
}

function Resolve-EbiMenuField {
    # PURE. "which column?" -> the verdict field: its number in the list or its name ('' = not one).
    param([string]$Text, $Fields)
    $t = if ($null -eq $Text) { '' } else { $Text.Trim() }
    $list = @(@($Fields) | ForEach-Object { [string]$_ })
    if ($t -match '^\d+$') { $n = [int]$t; if ($n -ge 1 -and $n -le $list.Count) { return $list[$n - 1] }; return '' }
    foreach ($f in $list) { if ($f -ieq $t) { return $f } }
    return ''
}

function Get-EbiDoneFileHits {
    <#
      Rows whose verdict for a phase is still blank although that phase's
      own record file is there and was written on or after the row's day
      (worklist.json verdict column "doneFile", e.g. capture/gfix/
      {{item.keySafe}}/track.json): the phase most likely finished and the
      cell was never written (a run cut off between the two, 10/20). An ng
      cell is not offered (it was judged). Returns @( @{ key; file; at } ).
    #>
    param($Rows, $Profile, [string]$Field, [string]$WorkDir, $KeyColumns = @())
    $out = New-Object System.Collections.ArrayList
    $spec = Get-EbiWorklistColumnSpec -Profile $Profile -Field $Field
    if ($null -eq $spec -or -not $spec.Contains('doneFile') -or [string]$spec['doneFile'] -eq '') { return $out.ToArray() }
    $daySpec = Get-EbiWorklistDaySpec -Profile $Profile
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $v = if ($r.Contains($Field) -and $null -ne $r[$Field]) { ([string]$r[$Field]).Trim() } else { '' }
        if ($v -ne '') { continue }
        $scope = New-EbiTemplateScope -Profile $(if ($Profile -is [hashtable]) { $Profile } else { @{} }) -Item $r -KeyColumns $KeyColumns
        $x = Expand-EbiTemplate -Value ([string]$spec['doneFile']) -Scope $scope
        if (-not $x['ok'] -or $null -eq $x['value']) { continue }
        $rel = [string]$x['value']
        $path = if ([System.IO.Path]::IsPathRooted($rel)) { $rel } else { Join-Path $WorkDir $rel }
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $at = (Get-Item -LiteralPath $path).LastWriteTime
        $day = Get-EbiWorklistRowDay -Row $r -DaySpec $daySpec
        if ($day -ne '' -and $at.ToString('yyyy-MM-dd') -lt $day) { continue }   # an older run's file (the row was planned again since)
        [void]$out.Add(@{ key = (Get-EbiKeyDisplay -Item $r -KeyColumns $KeyColumns); file = $rel; at = $at.ToString('MM/dd HH:mm') })
    }
    return $out.ToArray()
}

# ---------------------------------------------------------------- menu

function Format-EbiMenuRunResult {
    # PURE. The menu's one line after a phase: words, not an exit-code legend.
    param([string]$Phase, $Code)
    $c = [string]$Code
    $word = switch ($c) { '0' { 'done (see the run summary above; the board below is refreshed)' } '1' { 'stopped with a failure (the FAILED lines above say where)' } '2' { 'did not start: the command was wrong (the message above says why)' } '3' { 'cancelled (q)' } default { 'ended with exit code ' + $c } }
    return ($Phase + ': ' + $word)
}

function ConvertTo-EbiProcessArgument {
    <#
      PURE. One argument for a Windows command line (Start-Process
      -ArgumentList as ONE string, which PowerShell 5.1 passes as is):
      quoted when it holds a space, a tab or a quote, with inner quotes
      escaped and the backslashes before them / before the closing quote
      doubled (the CommandLineToArgvW rules). '' -> "".
    #>
    param([string]$Value)
    if ($Value -eq '') { return '""' }
    if ($Value -notmatch '[\s"]') { return $Value }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('"')
    $bs = 0
    foreach ($ch in $Value.ToCharArray()) {
        if ($ch -eq [char]92) { $bs++; continue }
        if ($ch -eq [char]34) { [void]$sb.Append(([string][char]92) * ($bs * 2 + 1)); [void]$sb.Append('"'); $bs = 0; continue }
        if ($bs -gt 0) { [void]$sb.Append(([string][char]92) * $bs); $bs = 0 }
        [void]$sb.Append($ch)
    }
    if ($bs -gt 0) { [void]$sb.Append(([string][char]92) * ($bs * 2)) }
    [void]$sb.Append('"')
    return $sb.ToString()
}

function Get-EbiMenuStateFile {
    param([string]$ToolDir)
    return (Join-Path (Join-Path $ToolDir (Get-EbiDashboardDirName)) 'menu.json')
}

function Read-EbiMenuState {
    # The last WorkDir / profile the menu used (so a double-click asks nothing). @{ workDir; profile }.
    param([string]$ToolDir)
    $r = Read-EbiJson -Path (Get-EbiMenuStateFile -ToolDir $ToolDir)
    $s = @{ workDir = ''; profile = '' }
    if ($r['ok'] -and ($r['value'] -is [System.Collections.IDictionary])) {
        if ($r['value'].Contains('workDir')) { $s['workDir'] = [string]$r['value']['workDir'] }
        if ($r['value'].Contains('profile')) { $s['profile'] = [string]$r['value']['profile'] }
    }
    return $s
}

function Write-EbiMenuState {
    param([string]$ToolDir, [string]$WorkDir, [string]$ProfileName)
    $path = Get-EbiMenuStateFile -ToolDir $ToolDir
    $dir = Split-Path $path -Parent
    try { if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null } } catch { return @{ ok = $false; message = $_.Exception.Message } }
    return (Write-EbiJson -Path $path -Value @{ workDir = $WorkDir; profile = $ProfileName; updatedAt = (Get-Date).ToString('o') })
}

function Format-EbiMenuChoices {
    # PURE. The numbered phase list + the fixed keys, as lines. NextText =
    # what Enter runs (the board's "next:" line), '' = Enter is nothing.
    param($Phases, $Unfinished = @(), [int]$Width = 0, [string]$NextText = '', [bool]$HasHold = $false)
    if ($Width -le 0) { $Width = Get-EbiDashboardWidth }
    $L = New-Object System.Collections.ArrayList
    $i = 0
    $nameW = 6
    foreach ($p in @($Phases)) { if (([string]$p['name']).Length -gt $nameW) { $nameW = ([string]$p['name']).Length } }
    if ($NextText -ne '') { [void]$L.Add((ConvertTo-EbiHelpFitDashboard -Text ('Enter  ' + $NextText) -Width $Width)) }
    foreach ($p in @($Phases)) {
        $i++
        $t = ([string]$i).PadLeft(3) + '  ' + ([string]$p['name']).PadRight($nameW) + '  ' + [string]$p['title']
        [void]$L.Add((ConvertTo-EbiHelpFitDashboard -Text $t -Width $Width))
    }
    $n = @($Unfinished).Count
    if ($n -gt 0) {
        $u = $Unfinished[0]
        $wid = if ($u.Contains('workflow') -and ($u['workflow'] -is [System.Collections.IDictionary]) -and $u['workflow'].Contains('id')) { [string]$u['workflow']['id'] } else { '?' }
        [void]$L.Add((ConvertTo-EbiHelpFitDashboard -Text ('  r  ' + 'resume'.PadRight($nameW) + '  the unfinished run of ' + $wid + ' (' + [string]$u['runId'] + ')' + $(if ($n -gt 1) { ' +' + ($n - 1) + ' more' } else { '' })) -Width $Width))
    }
    [void]$L.Add('  m  ' + 'mark'.PadRight($nameW) + '  set a verdict to ok by hand (done outside the tool)')
    [void]$L.Add('  c  ' + 'clear'.PadRight($nameW) + '  blank one row''s verdict so the phase redoes it')
    [void]$L.Add('  h  ' + 'hold'.PadRight($nameW) + '  put rows on hold (waiting for the HOST side)' + $(if ($HasHold) { ' / release' } else { '' }))
    [void]$L.Add('  s  ' + 'status'.PadRight($nameW) + '  redraw the dashboard')
    [void]$L.Add('  d  ' + 'doctor'.PadRight($nameW) + '  environment self-check')
    [void]$L.Add('  q  ' + 'quit'.PadRight($nameW))
    [void]$L.Add('     (a row can be typed by part of its name: DBC = every row with DBC)')
    return $L.ToArray()
}

function Read-EbiMenuAnswer {
    # PURE. Typed text -> @{ ok; action ('phase' | 'next' | 'r' | 'c' | 'm' | 'h' | 's' | 'd' | 'q'); index }.
    param([string]$Text, [int]$PhaseCount, [bool]$HasUnfinished, [bool]$HasNext = $false)
    $t = if ($null -eq $Text) { '' } else { (ConvertTo-EbiHalfWidth -Value $Text).Trim().ToLowerInvariant() }
    if ($t -eq '') {
        if ($HasNext) { return @{ ok = $true; action = 'next'; index = 0 } }
        return @{ ok = $false; action = ''; index = 0 }
    }
    if ($t -match '^\d+$') {
        $n = [int]$t
        if ($n -ge 1 -and $n -le $PhaseCount) { return @{ ok = $true; action = 'phase'; index = $n } }
        return @{ ok = $false; action = ''; index = 0 }
    }
    if ($t -eq 'r' -and $HasUnfinished) { return @{ ok = $true; action = 'r'; index = 0 } }
    if ($t -in @('c', 'm', 'h', 's', 'd', 'q')) { return @{ ok = $true; action = $t; index = 0 } }
    return @{ ok = $false; action = ''; index = 0 }
}

function Get-EbiMenuPhaseWarnings {
    <#
      PURE. Before a phase runs: which of the rows it would take have an
      EARLIER verdict still not ok (evidence before logs). The workflow
      refuses them itself (file_not_found); this only says so first.
      Returns lines.
    #>
    param($Model, $Phases, $Phase, $Only = @())
    $L = New-Object System.Collections.ArrayList
    $vp = @(Get-EbiDashboardVerdictPhases -Phases $Phases)
    $idx = -1
    for ($i = 0; $i -lt $vp.Count; $i++) { if ([string]$vp[$i]['id'] -eq [string]$Phase['id']) { $idx = $i } }
    if ($idx -le 0) { return $L.ToArray() }
    $only = @(@($Only) | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -ne '' })
    foreach ($m in @($Model)) {
        if ($only.Count -gt 0 -and -not ($only -contains [string]$m['key'])) { continue }
        if ([string]$m['verdicts'][[string]$Phase['field']] -eq 'ok') { continue }
        for ($i = 0; $i -lt $idx; $i++) {
            $f = [string]$vp[$i]['field']
            if ([string]$m['verdicts'][$f] -ne 'ok') { [void]$L.Add(([string]$m['key'] + ': ' + $f + ' is not ok yet -> ' + [string]$Phase['name'] + ' will refuse it')); break }
        }
    }
    return $L.ToArray()
}

function Invoke-EbiMenu {
    <#
      The entry point loop. Draw the dashboard, offer the phases, run the
      chosen one through a CHILD PowerShell (ebi.ps1 calls exit), come
      back to the menu.
        ToolDir / WorkDir / ProfileName / Operator  as Read-EbiDashboard
        Reader    scriptblock returning the typed line (tests); default Read-Host
        Runner    scriptblock taking the argument list (tests); default:
                  the same PowerShell executable, -File ebi.ps1 @args
        Auto      with no console (redirected stdin, EBI_NO_ASK=1, DryRun)
                  the menu draws once and quits
      Returns @{ ok; message; ran (the command texts run); cleared }.
    #>
    param([string]$ToolDir, [string]$WorkDir, [string]$ProfileName = '', [string]$Operator = '', [scriptblock]$Reader = $null, [scriptblock]$Runner = $null, [bool]$DryRun = $false)
    $ran = New-Object System.Collections.ArrayList
    $cleared = New-Object System.Collections.ArrayList
    $noConsole = $DryRun
    if (-not $noConsole -and $null -eq $Reader) {
        try { if ([Console]::IsInputRedirected) { $noConsole = $true } } catch { }
        if (-not $noConsole -and [string]$env:EBI_NO_ASK -eq '1') { $noConsole = $true }
    }
    $read = if ($null -ne $Reader) { $Reader } else { { Read-Host } }
    $psExe = ''
    try { $psExe = (Get-Process -Id $PID).Path } catch { }
    if ([string]::IsNullOrWhiteSpace($psExe)) { $psExe = 'powershell.exe' }
    $ebi = Join-Path $ToolDir 'ebi.ps1'
    $run = if ($null -ne $Runner) { $Runner } else {
        {
            param($CliArgs)
            # 10/16: `& $psExe ...` inside this block put the child's whole
            # output into the caller's $code -- nothing showed while the
            # phase ran, then it printed as ONE line, and a gate's question
            # could never be seen. Start-Process -NoNewWindow shares this
            # console (live output, the gate reads the keyboard) and hands
            # back only the exit code.
            $all = @('-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ebi) + @($CliArgs)
            $line = (@($all | ForEach-Object { ConvertTo-EbiProcessArgument -Value ([string]$_) })) -join ' '
            $proc = Start-Process -FilePath $psExe -ArgumentList $line -NoNewWindow -Wait -PassThru
            return [int]$proc.ExitCode
        }
    }
    $tries = 0
    while ($true) {
        $d = Read-EbiDashboard -ToolDir $ToolDir -WorkDir $WorkDir -ProfileName $ProfileName -Operator $Operator
        if (-not $d['ok']) { return @{ ok = $false; message = $d['message']; ran = $ran.ToArray(); cleared = $cleared.ToArray() } }
        $ProfileName = [string]$d['profileName']
        [void](Write-EbiMenuState -ToolDir $ToolDir -WorkDir $WorkDir -ProfileName $ProfileName)
        $lines = @($d['lines'])
        if ($lines.Count -gt 0) { $lines[0] = $lines[0] -replace '^ebi status:', 'ebi menu:' }
        Write-Host ''
        Write-EbiDashboardLines -Lines $lines
        Write-Host ''
        $phases = @($d['phases'])
        $split = $d['split']
        $scoped = ($null -ne $split -and [bool]$split['scoped'])
        $keyName = if (@($d['keyColumns']).Count -gt 0) { [string]$d['keyColumns'][0] } else { 'key' }
        $first = if ($null -ne $d['next']) { $d['next']['first'] } else { $null }
        $nextText = ''
        if ($null -ne $first -and $null -ne $first['phase'] -and $null -ne $d['runCommand']) {
            $nextText = switch ([string]$first['kind']) { 'resume' { 'resume the unfinished ' + [string]$first['phase']['name'] + ' run' } 'plan' { 'run ' + [string]$first['phase']['name'] + ' (pick today''s jobs)' } default { 'run ' + [string]$first['phase']['name'] + ' for today''s rows' } }
        }
        $hasHold = ($null -ne $split -and @($split['held']).Count -gt 0)
        foreach ($l in @(Format-EbiMenuChoices -Phases $phases -Unfinished $d['unfinished'] -NextText $nextText -HasHold $hasHold)) { Write-Host ('  ' + $l) }
        if ($noConsole) { Write-Host '  (no console to ask: quit)' -ForegroundColor DarkGray; return @{ ok = $true; message = 'no console'; ran = $ran.ToArray(); cleared = $cleared.ToArray() } }
        $ans = $null
        while ($true) {
            Write-Host '  > ' -ForegroundColor Magenta -NoNewline
            $typed = [string](& $read)
            $ans = Read-EbiMenuAnswer -Text $typed -PhaseCount $phases.Count -HasUnfinished (@($d['unfinished']).Count -gt 0) -HasNext ($nextText -ne '')
            if ($ans['ok']) { break }
            $tries++
            if ($tries -ge 20) { return @{ ok = $true; message = 'no valid answer after 20 tries'; ran = $ran.ToArray(); cleared = $cleared.ToArray() } }
            Write-Host ('  not an option: "' + $typed + '"') -ForegroundColor DarkYellow
        }
        $a = [string]$ans['action']
        if ($a -eq 'next') {
            # Enter = the board's next step: a resume, or that phase for today's rows
            if ([string]$first['kind'] -eq 'resume') { $a = 'r' }
            else { $a = 'phase'; for ($i = 0; $i -lt $phases.Count; $i++) { if ([string]$phases[$i]['id'] -eq [string]$first['phase']['id']) { $ans['index'] = $i + 1 } } }
        }
        if ($a -eq 'q') { return @{ ok = $true; message = 'quit'; ran = $ran.ToArray(); cleared = $cleared.ToArray() } }
        if ($a -eq 's') { continue }
        if ($a -eq 'd') {
            Write-Host ''
            $code = & $run @('doctor', '-WorkDir', $WorkDir)
            [void]$ran.Add('ebi.ps1 doctor -WorkDir ' + $WorkDir)
            continue
        }
        if ($a -in @('c', 'm', 'h')) {
            $vp = @(Get-EbiDashboardVerdictPhases -Phases $phases)
            $fields = @($vp | ForEach-Object { [string]$_['field'] })
            if ($a -ne 'h' -and $vp.Count -eq 0) { Write-Host '  no verdict columns' -ForegroundColor DarkYellow; continue }
            Write-Host ('  which row(s)? (' + $keyName + ', or part of one; a,b = several; Enter = back) ') -NoNewline
            $rk = Resolve-EbiMenuKeys -Text ([string](& $read)) -Keys $d['allKeys']
            if (-not $rk['ok']) { Write-Host ('  ' + $rk['message']) -ForegroundColor DarkYellow; continue }
            $keys = @($rk['keys'])
            if ($keys.Count -eq 0) { continue }
            $leaf = Get-EbiDashboardLeaf -Path ([string]$d['worklistPath'])
            if ($a -eq 'h') {
                $holdCol = Get-EbiWorklistRoleColumn -Profile $d['profile'] -Role 'hold'
                if ($holdCol -eq '') { Write-Host '  this worklist has no hold column' -ForegroundColor DarkYellow; continue }
                $heldKeys = if ($null -ne $split) { @($split['heldKeys']) } else { @() }
                $release = (@($keys | Where-Object { -not ($heldKeys -contains $_) }).Count -eq 0)
                $reason = ''
                if (-not $release) {
                    Write-Host '  why? (Enter = waiting for HOST) ' -NoNewline
                    $reason = ([string](& $read)).Trim()
                    if ($reason -eq '') { $reason = 'waiting for HOST' }
                }
                $what = if ($release) { ('release ' + ($keys -join ', ') + ' (the phases take them again)') } else { ('put on hold: ' + ($keys -join ', ') + ' -- ' + $reason) }
                $g = Show-EbiGate -Title $(if ($release) { 'RELEASE' } else { 'HOLD' }) -What @($what, ('in ' + $leaf)) -Next @('Enter / y: write it', 'n: leave it') -Actions @(@{ key = 'y'; label = 'write' }, @{ key = 'n'; label = 'leave' }) -Default 'y' -Auto 'n' -Reader $read -DryRun:$DryRun
                if ($g['action'] -ne 'y') { continue }
                $w = Set-EbiWorklistCells -Path ([string]$d['worklistPath']) -Keys $keys -Field $holdCol -Value $reason -KeyColumns @($d['keyColumns']) -AddColumn $true -DryRun $DryRun
                Write-Host ('  ' + $w['message']) -ForegroundColor $(if ($w['ok']) { 'Green' } else { 'Red' })
                continue
            }
            $fl = @(for ($i = 0; $i -lt $fields.Count; $i++) { ('' + ($i + 1) + ' ' + $fields[$i]) })
            Write-Host ('  which column? (' + ($fl -join '  ') + '; Enter = back) ') -NoNewline
            $field = Resolve-EbiMenuField -Text ([string](& $read)) -Fields $fields
            if ($field -eq '') { continue }
            if ($a -eq 'm') {
                $g = Show-EbiGate -Title 'MARK OK BY HAND' -What @(('set ' + $field + ' = ok for: ' + ($keys -join ', ')), ('in ' + $leaf)) -Next @('Enter / y: write it (the ' + $field + ' phase skips these rows from now on)', 'n: leave it') -Actions @(@{ key = 'y'; label = 'write' }, @{ key = 'n'; label = 'leave' }) -Default 'y' -Auto 'n' -Reader $read -DryRun:$DryRun
                if ($g['action'] -ne 'y') { continue }
                $okValue = ConvertTo-EbiStoredValue -Logical 'ok' -Spec (Get-EbiWorklistColumnSpec -Profile $d['profile'] -Field $field)
                $w = Set-EbiWorklistCells -Path ([string]$d['worklistPath']) -Keys $keys -Field $field -Value $okValue -KeyColumns @($d['keyColumns']) -DryRun $DryRun
                Write-Host ('  ' + $w['message']) -ForegroundColor $(if ($w['ok']) { 'Green' } else { 'Red' })
                continue
            }
            $g = Show-EbiGate -Title 'CLEAR A VERDICT' -What @(('blank ' + ($keys -join ', ') + ': ' + $field + ' in ' + $leaf)) -Next @('Enter / y: clear it (the phase runs this row again)', 'n: leave it') -Actions @(@{ key = 'y'; label = 'clear' }, @{ key = 'n'; label = 'leave' }) -Default 'y' -Auto 'n' -Reader $read -DryRun:$DryRun
            if ($g['action'] -ne 'y') { continue }
            foreach ($key in $keys) {
                $c = Clear-EbiWorklistVerdict -Path ([string]$d['worklistPath']) -Key $key -Field $field -KeyColumns @($d['keyColumns']) -DryRun $DryRun
                Write-Host ('  ' + $c['message']) -ForegroundColor $(if ($c['ok']) { 'Green' } else { 'Red' })
                if ($c['ok']) { [void]$cleared.Add(@{ key = $key; field = $field }) }
            }
            continue
        }
        $phase = $null; $resume = $false; $runId = ''; $only = @(); $allDays = $false
        if ($a -eq 'r') {
            $u = $d['unfinished'][0]
            $wid = if ($u.Contains('workflow') -and ($u['workflow'] -is [System.Collections.IDictionary]) -and $u['workflow'].Contains('id')) { [string]$u['workflow']['id'] } else { '' }
            foreach ($p in $phases) { if ([string]$p['id'] -eq $wid) { $phase = $p } }
            if ($null -eq $phase) { Write-Host ('  the unfinished run is of "' + $wid + '", not one of this profile''s phases') -ForegroundColor DarkYellow; continue }
            $resume = $true; $runId = [string]$u['runId']
        } else {
            $phase = $phases[$ans['index'] - 1]
            $field = [string]$phase['field']
            if ($field -ne '') {
                $nToday = @($d['model']).Count
                $prompt = if ($scoped) { ('  ' + $keyName + ' to run: Enter = today''s ' + $nToday + ' row' + $(if ($nToday -ne 1) { 's' } else { '' }) + ', or names / part of a name, a = every day: ') } else { ('  ' + $keyName + ' to run (Enter = every pending row, a,b = several): ') }
                Write-Host $prompt -NoNewline
                $typed = ([string](& $read)).Trim()
                if ($typed -ieq 'a') { $allDays = $true }
                elseif ($typed -ne '') {
                    $rk = Resolve-EbiMenuKeys -Text $typed -Keys $d['allKeys']
                    if (-not $rk['ok']) { Write-Host ('  ' + $rk['message']) -ForegroundColor DarkYellow; continue }
                    $only = @($rk['keys'])
                    if (($only -join ',') -ne (@($typed -split '[\s,;]+' | Where-Object { $_ -ne '' }) -join ',')) { Write-Host ('  -> ' + ($only -join ', ')) -ForegroundColor Green }
                }
                # rows that look done already: the phase's own file is there, the cell is blank
                $scopeRows = if ($only.Count -gt 0) { @(@($d['rows']) | Where-Object { $null -ne $_ -and $only -contains (Get-EbiKeyDisplay -Item $_ -KeyColumns @($d['keyColumns'])) }) } elseif ($null -ne $split) { @($split['today']) } else { @($d['rows']) }
                $hits = @(Get-EbiDoneFileHits -Rows $scopeRows -Profile $d['profile'] -Field $field -WorkDir $WorkDir -KeyColumns @($d['keyColumns']))
                if ($hits.Count -gt 0) {
                    $g = Show-EbiGate -Title 'LOOKS DONE ALREADY' -What @(foreach ($h in $hits) { ([string]$h['key'] + ': ' + $field + ' is blank, but ' + [string]$h['file'] + ' was written ' + [string]$h['at']) }) -Next @('Enter / y: mark them ' + $field + ' = ok (this run skips them)', 'n: run them again') -Actions @(@{ key = 'y'; label = 'mark ok' }, @{ key = 'n'; label = 'run again' }) -Default 'y' -Auto 'n' -Reader $read -DryRun:$DryRun
                    if ($g['action'] -eq 'y') {
                        $hk = @($hits | ForEach-Object { [string]$_['key'] })
                        $okValue = ConvertTo-EbiStoredValue -Logical 'ok' -Spec (Get-EbiWorklistColumnSpec -Profile $d['profile'] -Field $field)
                        $w = Set-EbiWorklistCells -Path ([string]$d['worklistPath']) -Keys $hk -Field $field -Value $okValue -KeyColumns @($d['keyColumns']) -DryRun $DryRun
                        Write-Host ('  ' + $w['message']) -ForegroundColor $(if ($w['ok']) { 'Green' } else { 'Red' })
                        if ($only.Count -gt 0) {
                            $only = @($only | Where-Object { -not ($hk -contains $_) })
                            if ($only.Count -eq 0) { Write-Host '  nothing left to run' -ForegroundColor DarkGray; continue }
                        }
                    }
                }
                # earlier days' rows this phase has not finished (a rerun of today: 10/20)
                if ($scoped -and $only.Count -eq 0 -and -not $allDays) {
                    $spec = Get-EbiWorklistColumnSpec -Profile $d['profile'] -Field $field
                    $late = @(foreach ($r in @($d['rows'])) {
                        if ($null -eq $r) { continue }
                        $k = Get-EbiKeyDisplay -Item $r -KeyColumns @($d['keyColumns'])
                        if (-not (@($split['openKeys']) -contains $k)) { continue }
                        if ((Get-EbiDashboardLogical -Value $(if ($r.Contains($field)) { $r[$field] } else { '' }) -Spec $spec) -eq 'ok') { continue }
                        $k
                    })
                    if ($late.Count -gt 0) {
                        $lines = @(@($split['open']) | Where-Object { $late -contains (([string]$_ -split ' ')[0]) })
                        $g = Show-EbiGate -Title 'EARLIER DAYS, NOT FINISHED' -What $lines -Next @('t: take them today too (WORKDAY = today; the day their GFIX job ran is kept)', 'h: put them on hold (waiting for HOST)', 'Enter / n: leave them (this run is today''s rows only)') -Actions @(@{ key = 't'; label = 'today too' }, @{ key = 'h'; label = 'hold' }, @{ key = 'n'; label = 'leave' }) -Default 'n' -Auto 'n' -Reader $read -DryRun:$DryRun
                        $dayCol = ''; $ds = Get-EbiWorklistDaySpec -Profile $d['profile']; if ($null -ne $ds) { $dayCol = [string]$ds['column'] }
                        if ($g['action'] -eq 't' -and $dayCol -ne '') {
                            $w = Set-EbiWorklistCells -Path ([string]$d['worklistPath']) -Keys $late -Field $dayCol -Value ([string]$d['today']) -KeyColumns @($d['keyColumns']) -AddColumn $true -DryRun $DryRun
                            Write-Host ('  ' + $w['message']) -ForegroundColor $(if ($w['ok']) { 'Green' } else { 'Red' })
                        } elseif ($g['action'] -eq 'h') {
                            $holdCol = Get-EbiWorklistRoleColumn -Profile $d['profile'] -Role 'hold'
                            if ($holdCol -ne '') {
                                $w = Set-EbiWorklistCells -Path ([string]$d['worklistPath']) -Keys $late -Field $holdCol -Value 'waiting for HOST' -KeyColumns @($d['keyColumns']) -AddColumn $true -DryRun $DryRun
                                Write-Host ('  ' + $w['message']) -ForegroundColor $(if ($w['ok']) { 'Green' } else { 'Red' })
                            }
                        }
                    }
                }
                foreach ($w in @(Get-EbiMenuPhaseWarnings -Model $d['model'] -Phases $phases -Phase $phase -Only $only)) { Write-Host ('  [WARN] ' + $w) -ForegroundColor Yellow }
            }
        }
        $cmd = New-EbiRunCommand -Phase $phase -ToolDir $ToolDir -WorkDir $WorkDir -Only $only -Resume $resume -RunId $runId -AllDays $allDays
        Write-Host ''
        Write-Host ('  ' + $cmd['text']) -ForegroundColor Cyan
        Write-Host ''
        [void]$ran.Add($cmd['text'])
        if ($DryRun) { Write-Host '  (dry run: not started)' -ForegroundColor DarkGray; continue }
        $code = & $run $cmd['args']
        Write-Host ''
        Write-Host ('  ' + (Format-EbiMenuRunResult -Phase ([string]$phase['name']) -Code $code)) -ForegroundColor $(if ([string]$code -eq '0') { 'Green' } else { 'Yellow' })
    }
}
