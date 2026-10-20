# modules/table/table.day_sets.ps1
# Sort the plan sheet's rows into the sets an operator picks today's work
# from (10/20, DETOUR-GFIX-001.10). The WBS keeps the PLANNED start (AF)
# and end (AG) and, once someone has done the job, the ACTUAL start (AJ).
# Matching "planned start = today" alone ran 10/20's jobs that were done
# early on 10/16 again and never showed the DBC jobs stuck since 10/15.
# Here every row is put in one place:
#     today     planned today (start <= today <= end), not done
#     late      planned before today, not done, not older than backDays
#     ahead     planned for one of the next aheadDays working days
#     others    someone else's today (only when an owner pattern is set)
#     mapping   in mapping.xlsx today but in none of the sets above
#     hold      the worklist row is on hold (never pickable)
# and done rows (WBS actual start, or every verdict ok in the worklist)
# become info lines instead of work. The worklist tells what this tool
# already did: a row whose first verdict (track) is ok keeps its GFIX date
# when it is picked again (the receive happened that day; logs / evidence
# continue it), any other row is planned for today.

. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Worklist.ps1')

$Manifest = @{
  id         = 'table.day_sets'
  group      = 'table'
  summary    = 'Sort plan rows into today / late / ahead / others / mapping / hold sets against the worklist'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    worklist     = @{ type='session'; sessionKind='worklist'; required=$true; desc='what this tool already did (verdicts, hold, work day)' }
    records      = @{ type='list';   required=$true; desc='plan rows (excel.read_rows -> verify.derive_fields), one or more per key' }
    today        = @{ type='string'; required=$true; desc='yyyy-MM-dd ({{run.date}})' }
    keyField     = @{ type='string'; default='Excel_NAME'; desc='record field holding the worklist key' }
    ownerField   = @{ type='string'; default='owner' }
    ownerPattern = @{ type='string'; default=''; desc='regex an owner cell of yours matches (verify.owner_pattern); blank = every row is yours' }
    startField   = @{ type='string'; default='start'; desc='planned start date' }
    endField     = @{ type='string'; default='end'; desc='planned end date (blank = the start)' }
    doneField    = @{ type='string'; default='done'; desc='actual start date: set = somebody did the job' }
    extra        = @{ type='list';   default=@(); desc='records of another list for today (mapping.xlsx), same keyField' }
    extraNote    = @{ type='string'; default='GFIX_TIME'; desc='extra field shown next to the key' }
    aheadDays    = @{ type='int';    default=2; desc='working days (Mon-Fri) looked ahead' }
    backDays     = @{ type='int';    default=21; desc='calendar days looked back for late rows' }
    dateColumn   = @{ type='string'; default='GFIX_DATE'; desc='worklist column the job''s run date lives in' }
    carryField   = @{ type='string'; default=''; desc='verdict column whose ok fixes the run date (blank = the first verdict column)' }
  }
  outputs    = @{
    sets        = @{ type='list'; desc='@{ n (1.. pickable, 0 not); id; label; names; notes (name -> text); default }' }
    info        = @{ type='list'; desc='lines about rows that are done or have no date' }
    records     = @{ type='list'; desc='one record per pickable key: the first plan row + set, the run date column, WORKDAY' }
    defaultKeys = @{ type='list'; desc='keys of the default sets (today; mapping too when no owner pattern)' }
    presets     = @{ type='list'; desc='the recommended combinations for human.pick: today, today + late, today + the next working day ahead' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'table.day_sets'; with = @{ worklist = 'wl'; records = '{{steps.derive.out.records}}'; today = '{{run.date}}'; ownerPattern = '{{steps.ownerPat.out.pattern}}'; extra = '{{steps.mine.out.records}}' } }
  notes      = 'Pure: reads the worklist, writes nothing. A key whose rows are all done in the plan sheet is done; one of them blank keeps it open (two files of one job). A row picked earlier today (worklist WORKDAY = today) stays in today whatever the sheet says. Dates are compared as yyyy-MM-dd text.'
}

function TableDaySets-Norm {
    param($Value)
    if ($null -eq $Value) { return '' }
    return (ConvertTo-EbiHalfWidth -Value ([string]$Value)).Trim().ToUpperInvariant()
}

function TableDaySets-Short {
    # 'yyyy-MM-dd' -> 'MM/dd'
    param([string]$Day)
    if ($Day.Length -ge 10) { return $Day.Substring(5, 5).Replace('-', '/') }
    return $Day
}

function TableDaySets-AddWorkdays {
    # PURE. The day N working days (Mon-Fri) after Day, as yyyy-MM-dd.
    param([string]$Day, [int]$N)
    $d = [datetime]::ParseExact($Day, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    $left = $N
    while ($left -gt 0) {
        $d = $d.AddDays(1)
        if ($d.DayOfWeek -ne [DayOfWeek]::Saturday -and $d.DayOfWeek -ne [DayOfWeek]::Sunday) { $left-- }
    }
    return $d.ToString('yyyy-MM-dd')
}

function TableDaySets-WorklistState {
    <#
      PURE. Worklist rows -> normalized key -> @{ verdicts (field -> logical);
      okAll; progress (verdicts written, "track ok, logs ng"); carry; date;
      day; hold }.
    #>
    param($Worklist, $Profile, [string]$DateColumn, [string]$CarryField)
    $keyCols = @(Get-EbiWorklistKeyColumns -Profile $Profile)
    $verdictCols = New-Object System.Collections.ArrayList
    if ($null -ne $Profile -and ($Profile -is [System.Collections.IDictionary]) -and $Profile.Contains('worklist') -and ($Profile['worklist'] -is [System.Collections.IDictionary]) -and $Profile['worklist'].Contains('columns')) {
        foreach ($c in @($Profile['worklist']['columns'])) { if ($c -is [System.Collections.IDictionary] -and [string]$c['role'] -eq 'verdict') { [void]$verdictCols.Add([string]$c['name']) } }
    }
    $carry = if ($CarryField -ne '') { $CarryField } elseif ($verdictCols.Count -gt 0) { [string]$verdictCols[0] } else { '' }
    $holdCol = Get-EbiWorklistRoleColumn -Profile $Profile -Role 'hold'
    $daySpec = Get-EbiWorklistDaySpec -Profile $Profile
    $state = @{}
    $rows = @()
    if ($null -ne $Worklist -and ($Worklist -is [System.Collections.IDictionary]) -and $Worklist.Contains('rows')) { $rows = @($Worklist['rows']) }
    foreach ($row in $rows) {
        if ($null -eq $row) { continue }
        $k = TableDaySets-Norm (Get-EbiKeyDisplay -Item $row -KeyColumns $keyCols)
        if ($k -eq '') { continue }
        $v = @{}; $parts = New-Object System.Collections.ArrayList; $okAll = ($verdictCols.Count -gt 0)
        foreach ($f in $verdictCols) {
            $raw = if ($row.Contains($f)) { $row[$f] } else { '' }
            $l = ConvertFrom-EbiStoredValue -Value $raw -Spec (Get-EbiWorklistColumnSpec -Profile $Profile -Field $f)
            $v[$f] = $l
            if ($l -ne 'ok') { $okAll = $false }
            if ($l -ne '') { [void]$parts.Add(($f + ' ' + $l)) }
        }
        $date = if ($DateColumn -ne '' -and $row.Contains($DateColumn)) { ConvertTo-EbiWorklistDay -Text ([string]$row[$DateColumn]) } else { '' }
        $state[$k] = @{
            shown = (Get-EbiKeyDisplay -Item $row -KeyColumns $keyCols)
            verdicts = $v; okAll = $okAll; progress = ($parts.ToArray() -join ', ')
            carry = ($carry -ne '' -and $v.Contains($carry) -and [string]$v[$carry] -eq 'ok')
            date = $date
            day = $(if ($null -ne $daySpec -and $row.Contains([string]$daySpec['column'])) { ConvertTo-EbiWorklistDay -Text ([string]$row[[string]$daySpec['column']]) } else { '' })
            hold = $(if ($holdCol -ne '' -and $row.Contains($holdCol) -and $null -ne $row[$holdCol]) { ([string]$row[$holdCol]).Trim() } else { '' })
        }
    }
    return $state
}

function TableDaySets-Classify {
    <#
      PURE. The whole decision, without $Ctx (unit-tested directly).
      Returns @{ sets; info; records; defaultKeys }.
    #>
    param($Records, $State, [string]$Today, [string]$KeyField = 'Excel_NAME', [string]$OwnerField = 'owner', [string]$OwnerPattern = '', [string]$StartField = 'start', [string]$EndField = 'end', [string]$DoneField = 'done', $Extra = @(), [string]$ExtraNote = 'GFIX_TIME', [int]$AheadDays = 2, [int]$BackDays = 21, [string]$DateColumn = 'GFIX_DATE')
    $todayD = [datetime]::ParseExact($Today, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    $backLimit = $todayD.AddDays(-1 * $BackDays).ToString('yyyy-MM-dd')
    $aheadLimit = TableDaySets-AddWorkdays -Day $Today -N $AheadDays
    if ($null -eq $State) { $State = @{} }

    # one entry per key, in sheet order
    $order = New-Object System.Collections.ArrayList
    $agg = @{}
    foreach ($r in @($Records)) {
        if (-not ($r -is [System.Collections.IDictionary])) { continue }
        $shown = if ($r.Contains($KeyField) -and $null -ne $r[$KeyField]) { ([string]$r[$KeyField]).Trim() } else { '' }
        $k = TableDaySets-Norm $shown
        if ($k -eq '') { continue }
        $start = ConvertTo-EbiWorklistDay -Text ([string]$r[$StartField])
        $end = ConvertTo-EbiWorklistDay -Text ([string]$r[$EndField])
        if ($end -eq '' -or ($start -ne '' -and $end -lt $start)) { $end = $start }
        $done = ConvertTo-EbiWorklistDay -Text ([string]$r[$DoneField])
        $owner = if ($r.Contains($OwnerField) -and $null -ne $r[$OwnerField]) { ([string]$r[$OwnerField]).Trim() } else { '' }
        $mine = ($OwnerPattern -eq '' -or ($owner -ne '' -and [regex]::IsMatch($owner, $OwnerPattern)))
        if (-not $agg.Contains($k)) {
            $agg[$k] = @{ key = $shown; record = $r; start = $start; end = $end; allDone = ($done -ne ''); done = $done; mine = $mine; owner = $owner }
            [void]$order.Add($k)
            continue
        }
        $a = $agg[$k]
        if ($start -ne '' -and ($a['start'] -eq '' -or $start -lt $a['start'])) { $a['start'] = $start }
        if ($end -ne '' -and ($a['end'] -eq '' -or $end -gt $a['end'])) { $a['end'] = $end }
        if ($done -eq '') { $a['allDone'] = $false } elseif ($done -gt $a['done']) { $a['done'] = $done }
        if ($mine) { $a['mine'] = $true }
        if ($a['owner'] -eq '') { $a['owner'] = $owner }
    }

    $sets = [ordered]@{ today = (New-Object System.Collections.ArrayList); late = (New-Object System.Collections.ArrayList); others = (New-Object System.Collections.ArrayList); mapping = (New-Object System.Collections.ArrayList); hold = (New-Object System.Collections.ArrayList) }
    $ahead = [ordered]@{}
    $notes = @{}
    $where = @{}           # norm key -> set id
    $doneWbs = New-Object System.Collections.ArrayList
    $doneHere = New-Object System.Collections.ArrayList
    $noDate = 0
    foreach ($k in $order) {
        $a = $agg[$k]
        $st = if ($State.Contains($k)) { $State[$k] } else { $null }
        $inWindow = ($a['start'] -ne '' -and $a['end'] -ge $backLimit -and $a['start'] -le $aheadLimit)
        if (-not $a['mine']) {
            if ($null -ne $st -and [string]$st['day'] -eq $Today) { [void]$sets['today'].Add($k); $where[$k] = 'today'; $notes[$k] = ('picked earlier today; ' + $a['owner']); continue }
            if ($a['allDone'] -or $a['start'] -eq '' -or $a['start'] -gt $Today -or $a['end'] -lt $Today) { continue }
            if ($null -ne $st -and [string]$st['hold'] -ne '') { continue }
            [void]$sets['others'].Add($k); $where[$k] = 'others'; $notes[$k] = $a['owner']
            continue
        }
        if ($null -ne $st -and [string]$st['hold'] -ne '') { [void]$sets['hold'].Add($k); $where[$k] = 'hold'; $notes[$k] = [string]$st['hold']; continue }
        if ($null -ne $st -and [string]$st['day'] -eq $Today) {
            [void]$sets['today'].Add($k); $where[$k] = 'today'
            $notes[$k] = $(if ($a['start'] -ne '' -and $a['start'] -le $Today -and $a['end'] -ge $Today) { [string]$st['progress'] } else { 'picked earlier today' + $(if ([string]$st['progress'] -ne '') { '; ' + [string]$st['progress'] } else { '' }) })
            continue
        }
        if ($a['allDone']) {
            # only the ones the plan would have offered: planned today or later, or done today
            if ($a['start'] -ge $Today -and $a['start'] -le $aheadLimit -or $a['done'] -eq $Today) { [void]$doneWbs.Add(($a['key'] + ' ' + (TableDaySets-Short $a['done']) + $(if ($a['done'] -lt $a['start']) { ' (planned ' + (TableDaySets-Short $a['start']) + ')' } else { '' }))) }
            continue
        }
        if ($null -ne $st -and [bool]$st['okAll']) {
            if ($inWindow) { [void]$doneHere.Add(($a['key'] + $(if ([string]$st['date'] -ne '') { ' ' + (TableDaySets-Short ([string]$st['date'])) } else { '' }))) }
            continue
        }
        if ($a['start'] -eq '') { $noDate++; continue }
        $prog = if ($null -ne $st) { [string]$st['progress'] } else { '' }
        if ($a['start'] -le $Today -and $a['end'] -ge $Today) {
            [void]$sets['today'].Add($k); $where[$k] = 'today'
            $n = New-Object System.Collections.ArrayList
            if ($a['end'] -gt $Today) { [void]$n.Add(('until ' + (TableDaySets-Short $a['end']))) }
            if ($prog -ne '') { [void]$n.Add($prog) }
            $notes[$k] = ($n.ToArray() -join '; ')
        }
        elseif ($a['end'] -lt $Today -and $a['end'] -ge $backLimit) {
            [void]$sets['late'].Add($k); $where[$k] = 'late'
            $notes[$k] = ('since ' + (TableDaySets-Short $a['end']) + $(if ($prog -ne '') { '; ' + $prog } else { '' }))
        }
        elseif ($a['start'] -gt $Today -and $a['start'] -le $aheadLimit) {
            $d = $a['start']
            if (-not $ahead.Contains($d)) { $ahead[$d] = New-Object System.Collections.ArrayList }
            [void]$ahead[$d].Add($k); $where[$k] = ('ahead:' + $d)
            $notes[$k] = $prog
        }
    }
    # rows picked earlier today that the sheet did not give (mapping.xlsx
    # only, or taken in from an earlier day): still today's
    foreach ($k in @($State.Keys | Sort-Object)) {
        if ($agg.Contains($k) -or [string]$State[$k]['day'] -ne $Today -or [string]$State[$k]['hold'] -ne '') { continue }
        $agg[$k] = @{ key = [string]$State[$k]['shown']; record = $null }
        [void]$sets['today'].Add($k); $where[$k] = 'today'
        $notes[$k] = ('picked earlier today' + $(if ([string]$State[$k]['progress'] -ne '') { '; ' + [string]$State[$k]['progress'] } else { '' }))
    }
    # the other list (mapping.xlsx today): a note on keys already placed,
    # its own set for the rest
    foreach ($r in @($Extra)) {
        if (-not ($r -is [System.Collections.IDictionary])) { continue }
        $shown = if ($r.Contains($KeyField) -and $null -ne $r[$KeyField]) { ([string]$r[$KeyField]).Trim() } else { '' }
        $k = TableDaySets-Norm $shown
        if ($k -eq '') { continue }
        $t = if ($ExtraNote -ne '' -and $r.Contains($ExtraNote) -and $null -ne $r[$ExtraNote]) { ([string]$r[$ExtraNote]).Trim() } else { '' }
        $tag = 'mapping' + $(if ($t -ne '') { ' ' + $t } else { '' })
        if ($where.Contains($k)) {
            if ([string]$notes[$k] -notlike ('*' + $tag + '*')) { $notes[$k] = $(if ([string]$notes[$k] -ne '') { [string]$notes[$k] + '; ' + $tag } else { $tag }) }
            continue
        }
        if (-not $agg.Contains($k)) { $agg[$k] = @{ key = $shown; record = $null } }
        elseif ($null -ne $State -and $State.Contains($k) -and [bool]$State[$k]['okAll']) { continue }
        [void]$sets['mapping'].Add($k); $where[$k] = 'mapping'; $notes[$k] = $(if ($t -ne '') { $t } else { '' })
    }

    # sets in screen order, numbered when pickable
    $out = New-Object System.Collections.ArrayList
    $n = 0
    $defaultKeys = New-Object System.Collections.ArrayList
    function TableDaySets-AddSet {
        param([string]$Id, [string]$Label, $Keys, [bool]$Pick, [bool]$Default, [bool]$ShowEmpty = $false)
        if (@($Keys).Count -eq 0 -and -not $ShowEmpty) { return }
        $num = 0
        if ($Pick) { $counter['n'] = [int]$counter['n'] + 1; $num = [int]$counter['n'] }
        $shownKeys = @(foreach ($k in @($Keys)) { [string]$agg[$k]['key'] })
        $nm = @{}
        foreach ($k in @($Keys)) { if ($notes.Contains($k) -and [string]$notes[$k] -ne '') { $nm[[string]$agg[$k]['key']] = [string]$notes[$k] } }
        [void]$out.Add(@{ n = $num; id = $Id; label = $Label; names = $shownKeys; notes = $nm; default = $Default })
        if ($Default) { foreach ($s in $shownKeys) { [void]$defaultKeys.Add($s) } }
    }
    $counter = @{ n = 0 }
    TableDaySets-AddSet -Id 'today' -Label ('today ' + (TableDaySets-Short $Today)) -Keys $sets['today'] -Pick $true -Default $true -ShowEmpty $true
    TableDaySets-AddSet -Id 'late' -Label 'late, not done' -Keys $sets['late'] -Pick $true -Default $false
    foreach ($d in @($ahead.Keys | Sort-Object)) { TableDaySets-AddSet -Id ('ahead:' + $d) -Label ('ahead ' + (TableDaySets-Short $d)) -Keys $ahead[$d] -Pick $true -Default $false }
    TableDaySets-AddSet -Id 'others' -Label 'others, today' -Keys $sets['others'] -Pick $true -Default $false
    TableDaySets-AddSet -Id 'mapping' -Label 'mapping.xlsx only' -Keys $sets['mapping'] -Pick $true -Default ($OwnerPattern -eq '')
    TableDaySets-AddSet -Id 'hold' -Label 'on hold' -Keys $sets['hold'] -Pick $false -Default $false

    $info = New-Object System.Collections.ArrayList
    if ($doneWbs.Count -gt 0) { [void]$info.Add(('done already (WBS actual start): ' + ($doneWbs.ToArray() -join ', '))) }
    if ($doneHere.Count -gt 0) { [void]$info.Add(('all ok here, sheet not filled yet: ' + ($doneHere.ToArray() -join ', '))) }
    if ($noDate -gt 0) { [void]$info.Add(('' + $noDate + ' of your row(s) have no planned date')) }

    # one record per pickable key, with its run date and work day
    $recs = New-Object System.Collections.ArrayList
    foreach ($k in $order) {
        if (-not $where.Contains($k) -or [string]$where[$k] -eq 'hold') { continue }
        $a = $agg[$k]
        if ($null -eq $a['record']) { continue }
        $c = @{}
        foreach ($f in $a['record'].Keys) { $c[[string]$f] = $a['record'][$f] }
        $st = if ($State.Contains($k)) { $State[$k] } else { $null }
        $c['set'] = [string]$where[$k]
        $c[$DateColumn] = $(if ($null -ne $st -and [bool]$st['carry'] -and [string]$st['date'] -ne '') { [string]$st['date'] } else { $Today })
        $c['WORKDAY'] = $Today
        [void]$recs.Add($c)
    }
    # the recommended combinations (human.pick's first screen)
    $presets = New-Object System.Collections.ArrayList
    $base = @($defaultKeys.ToArray())
    [void]$presets.Add(@{ label = ('today ' + (TableDaySets-Short $Today) + ' (recommended)'); names = $base; hint = '' })
    $lateKeys = @(foreach ($k in $sets['late']) { [string]$agg[$k]['key'] })
    if ($lateKeys.Count -gt 0) { [void]$presets.Add(@{ label = 'today + late (planned before today, not done)'; names = @($base + $lateKeys); hint = '' }) }
    foreach ($d in @($ahead.Keys | Sort-Object)) {
        $ak = @(foreach ($k in $ahead[$d]) { [string]$agg[$k]['key'] })
        [void]$presets.Add(@{ label = ('today + ahead ' + (TableDaySets-Short $d) + ' (do the next day early)'); names = @($base + $ak); hint = '' })
        break
    }
    return @{ sets = $out.ToArray(); info = $info.ToArray(); records = $recs.ToArray(); defaultKeys = $defaultKeys.ToArray(); presets = $presets.ToArray() }
}

function Invoke-Step {
    param($In, $Ctx)
    $today = ([string]$In['today']).Trim()
    if ($Ctx['DryRun'] -and $today -notmatch '^\d{4}-\d{2}-\d{2}$') { $today = (Get-Date).ToString('yyyy-MM-dd') }
    if ($today -notmatch '^\d{4}-\d{2}-\d{2}$') { return @{ ok = $false; failure = 'input_invalid'; message = ('today must be yyyy-MM-dd, got "' + $today + '"'); sets = @(); info = @(); records = @(); defaultKeys = @(); presets = @() } }
    if (-not [string]::IsNullOrWhiteSpace([string]$In['ownerPattern'])) {
        try { [void][regex]::new([string]$In['ownerPattern']) } catch { return @{ ok = $false; failure = 'input_invalid'; message = ('ownerPattern is not a regex: ' + $_.Exception.Message); sets = @(); info = @(); records = @(); defaultKeys = @(); presets = @() } }
    }
    $state = TableDaySets-WorklistState -Worklist $In['worklist'] -Profile $Ctx['Profile'] -DateColumn ([string]$In['dateColumn']) -CarryField ([string]$In['carryField'])
    $r = TableDaySets-Classify -Records $In['records'] -State $state -Today $today -KeyField ([string]$In['keyField']) -OwnerField ([string]$In['ownerField']) -OwnerPattern ([string]$In['ownerPattern']) -StartField ([string]$In['startField']) -EndField ([string]$In['endField']) -DoneField ([string]$In['doneField']) -Extra $In['extra'] -ExtraNote ([string]$In['extraNote']) -AheadDays ([int]$In['aheadDays']) -BackDays ([int]$In['backDays']) -DateColumn ([string]$In['dateColumn'])
    foreach ($s in @($r['sets'])) { $Ctx.Log.Info(('{0}: {1}' -f $s['label'], @($s['names']).Count)) }
    return @{ ok = $true; sets = $r['sets']; info = $r['info']; records = $r['records']; defaultKeys = $r['defaultKeys']; presets = $r['presets'] }
}
