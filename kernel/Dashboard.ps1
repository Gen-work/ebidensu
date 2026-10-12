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
      Priority: an unfinished run (resume it) > ng rows (clear and rerun)
      > the earliest phase with rows waiting > an empty worklist (plan) >
      nothing left.
    #>
    param($Model, $Phases, $Unfinished = @())
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
        if ($null -ne $plan) {
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
    param($Phase, [string]$ToolDir, [string]$WorkDir, $Only = @(), [bool]$Resume = $false, [string]$RunId = '')
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
    param([string]$Title, [string]$WorkDir, [string]$WorklistPath, $Model, $Phases, $Totals, $Next, $ShowColumns = @(), [string]$KeyHeader = 'key', [string]$RunLine = '', [string]$Today = '', [int]$Width = 0)
    if ($Width -le 0) { $Width = Get-EbiDashboardWidth }
    $L = New-Object System.Collections.ArrayList
    $rows = @($Model)
    $vp = @(Get-EbiDashboardVerdictPhases -Phases $Phases)
    $leaf = Get-EbiDashboardLeaf -Path $WorklistPath

    $head = 'ebi status: ' + $Title
    if ($Today -ne '') { $head += '  ' + $Today }
    if ($WorkDir -ne '') { $head += '  ' + $WorkDir }
    [void]$L.Add((ConvertTo-EbiHelpFitDashboard -Text $head -Width $Width))
    $wlLine = 'worklist: ' + $(if ($leaf -ne '') { $leaf } else { '(none)' }) + '  ' + $rows.Count + ' row' + $(if ($rows.Count -ne 1) { 's' } else { '' })
    [void]$L.Add((ConvertTo-EbiHelpFitDashboard -Text $wlLine -Width $Width))
    [void]$L.Add('-' * $Width)

    # column widths
    $keyW = $KeyHeader.Length
    foreach ($m in $rows) { if (([string]$m['key']).Length -gt $keyW) { $keyW = ([string]$m['key']).Length } }
    $vW = @()
    foreach ($p in $vp) { $w = ([string]$p['field']).Length; if ($w -lt 4) { $w = 4 }; $vW += $w }
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

function Find-EbiDayProfiles {
    # Every profile under ProfilesRoot whose worklist file exists in WorkDir (or that has a worklist at all when -All).
    param([string]$ProfilesRoot, [string]$WorkDir, [string]$Operator = '', [switch]$All)
    $out = New-Object System.Collections.ArrayList
    if (-not (Test-Path -LiteralPath $ProfilesRoot -PathType Container)) { return $out.ToArray() }
    foreach ($d in @(Get-ChildItem -LiteralPath $ProfilesRoot -Directory | Sort-Object Name)) {
        $pr = Read-EbiProfile -Dir $d.FullName -WorkDir $WorkDir
        if (-not $pr['ok']) { continue }
        $path = Resolve-EbiWorklistFile -Profile $pr['value'] -WorkDir $WorkDir -Operator $Operator
        if ($path -eq '') { continue }
        if ($All -or (Test-Path -LiteralPath $path -PathType Leaf)) { [void]$out.Add(@{ name = $d.Name; dir = $d.FullName; worklist = $path }) }
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
    $res = @{ ok = $false; message = ''; profile = $null; profileName = ''; worklistPath = ''; phases = @(); model = @(); totals = @{}; next = $null; unfinished = @(); showColumns = @(); keyColumns = @(); lines = @(); runLine = ''; runCommand = $null }
    $profilesRoot = Join-Path $ToolDir 'profiles'
    $dir = ''
    if (-not [string]::IsNullOrWhiteSpace($ProfileName)) {
        $dir = Resolve-EbiProfileDir -NameOrPath $ProfileName -ProfilesRoot $profilesRoot
        if ($dir -eq '' -or -not (Test-Path -LiteralPath $dir -PathType Container)) { $res['message'] = ('profile directory not found: ' + $ProfileName); return $res }
    } else {
        $found = @(Find-EbiDayProfiles -ProfilesRoot $profilesRoot -WorkDir $WorkDir -Operator $Operator)
        if ($found.Count -eq 1) { $dir = [string]$found[0]['dir'] }
        elseif ($found.Count -eq 0) { $res['message'] = ('no profile has its worklist in ' + $WorkDir + '; say which one with -Profile <name>'); return $res }
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
    $model = @(Get-EbiDashboardModel -Rows $rows -Phases $phases -Profile $profile -KeyColumns $keyColumns -ShowColumns $show)
    $totals = Get-EbiDashboardTotals -Model $model -Phases $phases
    $unfinished = @(Find-EbiUnfinishedRuns -WorkDir $WorkDir)
    $next = Get-EbiDashboardNext -Model $model -Phases $phases -Unfinished $unfinished
    $runLine = ''; $runCommand = $null
    $first = $next['first']
    if ($null -ne $first -and $null -ne $first['phase']) {
        if ($first['kind'] -eq 'resume') { $runCommand = New-EbiRunCommand -Phase $first['phase'] -ToolDir $ToolDir -WorkDir $WorkDir -Resume $true -RunId ([string]$first['runId']) }
        elseif ($first['kind'] -eq 'run' -or $first['kind'] -eq 'plan') { $runCommand = New-EbiRunCommand -Phase $first['phase'] -ToolDir $ToolDir -WorkDir $WorkDir }
        if ($null -ne $runCommand) { $runLine = 'next: ' + $runCommand['text'] }
    }
    $keyHeader = if ($keyColumns.Count -gt 0) { [string]$keyColumns[0] } else { 'key' }
    $today = if ($Today -ne '') { $Today } else { (Get-Date).ToString('yyyy-MM-dd') }
    $res['phases'] = $phases; $res['model'] = $model; $res['totals'] = $totals; $res['next'] = $next; $res['unfinished'] = $unfinished
    $res['showColumns'] = $show; $res['keyColumns'] = $keyColumns; $res['runLine'] = $runLine; $res['runCommand'] = $runCommand
    $res['lines'] = @(Format-EbiDashboard -Title $res['profileName'] -WorkDir $WorkDir -WorklistPath $wlPath -Model $model -Phases $phases -Totals $totals -Next $next -ShowColumns $show -KeyHeader $keyHeader -RunLine $runLine -Today $today)
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
        elseif ($l -like 'total *' -or $l -like 'worklist:*') { $color = 'DarkGray' }
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

# ---------------------------------------------------------------- menu

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
    # PURE. The numbered phase list + the fixed keys, as lines.
    param($Phases, $Unfinished = @(), [int]$Width = 0)
    if ($Width -le 0) { $Width = Get-EbiDashboardWidth }
    $L = New-Object System.Collections.ArrayList
    $i = 0
    $nameW = 4
    foreach ($p in @($Phases)) { if (([string]$p['name']).Length -gt $nameW) { $nameW = ([string]$p['name']).Length } }
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
    [void]$L.Add('  c  ' + 'clear'.PadRight($nameW) + '  blank one row''s verdict so the phase redoes it')
    [void]$L.Add('  s  ' + 'status'.PadRight($nameW) + '  redraw the dashboard')
    [void]$L.Add('  d  ' + 'doctor'.PadRight($nameW) + '  environment self-check')
    [void]$L.Add('  q  ' + 'quit'.PadRight($nameW))
    return $L.ToArray()
}

function Read-EbiMenuAnswer {
    # PURE. Typed text -> @{ ok; action ('phase' | 'r' | 'c' | 's' | 'd' | 'q'); index }.
    param([string]$Text, [int]$PhaseCount, [bool]$HasUnfinished)
    $t = if ($null -eq $Text) { '' } else { $Text.Trim().ToLowerInvariant() }
    if ($t -eq '') { return @{ ok = $false; action = ''; index = 0 } }
    if ($t -match '^\d+$') {
        $n = [int]$t
        if ($n -ge 1 -and $n -le $PhaseCount) { return @{ ok = $true; action = 'phase'; index = $n } }
        return @{ ok = $false; action = ''; index = 0 }
    }
    if ($t -eq 'r' -and $HasUnfinished) { return @{ ok = $true; action = 'r'; index = 0 } }
    if ($t -in @('c', 's', 'd', 'q')) { return @{ ok = $true; action = $t; index = 0 } }
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
    $ebi = Join-Path $ToolDir 'ebi.ps1'
    $run = if ($null -ne $Runner) { $Runner } else {
        {
            param($CliArgs)
            & $psExe -NoLogo -NoProfile -ExecutionPolicy Bypass -File $ebi @CliArgs
            return $LASTEXITCODE
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
        foreach ($l in @(Format-EbiMenuChoices -Phases $phases -Unfinished $d['unfinished'])) { Write-Host ('  ' + $l) }
        if ($noConsole) { Write-Host '  (no console to ask: quit)' -ForegroundColor DarkGray; return @{ ok = $true; message = 'no console'; ran = $ran.ToArray(); cleared = $cleared.ToArray() } }
        $ans = $null
        while ($true) {
            Write-Host '  > ' -ForegroundColor Magenta -NoNewline
            $typed = [string](& $read)
            $ans = Read-EbiMenuAnswer -Text $typed -PhaseCount $phases.Count -HasUnfinished (@($d['unfinished']).Count -gt 0)
            if ($ans['ok']) { break }
            $tries++
            if ($tries -ge 20) { return @{ ok = $true; message = 'no valid answer after 20 tries'; ran = $ran.ToArray(); cleared = $cleared.ToArray() } }
            Write-Host ('  not an option: "' + $typed + '"') -ForegroundColor DarkYellow
        }
        $a = [string]$ans['action']
        if ($a -eq 'q') { return @{ ok = $true; message = 'quit'; ran = $ran.ToArray(); cleared = $cleared.ToArray() } }
        if ($a -eq 's') { continue }
        if ($a -eq 'd') {
            Write-Host ''
            $code = & $run @('doctor', '-WorkDir', $WorkDir)
            [void]$ran.Add('ebi.ps1 doctor -WorkDir ' + $WorkDir)
            continue
        }
        if ($a -eq 'c') {
            $vp = @(Get-EbiDashboardVerdictPhases -Phases $phases)
            if ($vp.Count -eq 0) { Write-Host '  no verdict columns to clear' -ForegroundColor DarkYellow; continue }
            Write-Host ('  which row? (' + $(if (@($d['keyColumns']).Count -gt 0) { [string]$d['keyColumns'][0] } else { 'key' }) + ', Enter = back) ') -NoNewline
            $key = ([string](& $read)).Trim()
            if ($key -eq '') { continue }
            Write-Host ('  which column? (' + (@($vp | ForEach-Object { [string]$_['field'] }) -join ' / ') + ', Enter = back) ') -NoNewline
            $field = ([string](& $read)).Trim()
            if ($field -eq '') { continue }
            if (-not (@($vp | ForEach-Object { [string]$_['field'] }) -contains $field)) { Write-Host ('  not a verdict column: "' + $field + '"') -ForegroundColor DarkYellow; continue }
            $g = Show-EbiGate -Title 'CLEAR A VERDICT' -What @(('blank ' + $key + ': ' + $field + ' in ' + (Get-EbiDashboardLeaf -Path ([string]$d['worklistPath'])))) -Next @('Enter / y: clear it (the phase runs this row again)', 'n: leave it') -Actions @(@{ key = 'y'; label = 'clear' }, @{ key = 'n'; label = 'leave' }) -Default 'y' -Auto 'n' -Reader $read -DryRun:$DryRun
            if ($g['action'] -ne 'y') { continue }
            $c = Clear-EbiWorklistVerdict -Path ([string]$d['worklistPath']) -Key $key -Field $field -KeyColumns @($d['keyColumns']) -DryRun $DryRun
            Write-Host ('  ' + $c['message']) -ForegroundColor $(if ($c['ok']) { 'Green' } else { 'Red' })
            if ($c['ok']) { [void]$cleared.Add(@{ key = $key; field = $field }) }
            continue
        }
        $phase = $null; $resume = $false; $runId = ''; $only = @()
        if ($a -eq 'r') {
            $u = $d['unfinished'][0]
            $wid = if ($u.Contains('workflow') -and ($u['workflow'] -is [System.Collections.IDictionary]) -and $u['workflow'].Contains('id')) { [string]$u['workflow']['id'] } else { '' }
            foreach ($p in $phases) { if ([string]$p['id'] -eq $wid) { $phase = $p } }
            if ($null -eq $phase) { Write-Host ('  the unfinished run is of "' + $wid + '", not one of this profile''s phases') -ForegroundColor DarkYellow; continue }
            $resume = $true; $runId = [string]$u['runId']
        } else {
            $phase = $phases[$ans['index'] - 1]
            if ([string]$phase['field'] -ne '') {
                Write-Host ('  ' + $(if (@($d['keyColumns']).Count -gt 0) { [string]$d['keyColumns'][0] } else { 'key' }) + ' to run (Enter = every pending row, a,b = several): ') -NoNewline
                $typed = ([string](& $read)).Trim()
                if ($typed -ne '') { $only = @($typed -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) }
                foreach ($w in @(Get-EbiMenuPhaseWarnings -Model $d['model'] -Phases $phases -Phase $phase -Only $only)) { Write-Host ('  [WARN] ' + $w) -ForegroundColor Yellow }
            }
        }
        $cmd = New-EbiRunCommand -Phase $phase -ToolDir $ToolDir -WorkDir $WorkDir -Only $only -Resume $resume -RunId $runId
        Write-Host ''
        Write-Host ('  ' + $cmd['text']) -ForegroundColor Cyan
        Write-Host ''
        [void]$ran.Add($cmd['text'])
        if ($DryRun) { Write-Host '  (dry run: not started)' -ForegroundColor DarkGray; continue }
        $code = & $run $cmd['args']
        Write-Host ''
        Write-Host ('  ' + [string]$phase['name'] + ' finished with exit code ' + [string]$code + '  (0 ok, 1 failed, 3 cancelled)') -ForegroundColor $(if ([string]$code -eq '0') { 'Green' } else { 'Yellow' })
    }
}
