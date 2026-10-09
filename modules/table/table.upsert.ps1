# modules/table/table.upsert.ps1
# Add rows to the worklist, or refresh the given columns of rows already
# there (matched by the profile's key columns), and flush. Several input
# rows with the same key merge into one worklist row, and how many there
# were can be stored (one deliverable, several files). Progress columns a
# row already has are never touched unless named in overwrite: re-running
# a plan in the afternoon must not reset the morning's work.

. (Join-Path $PSScriptRoot '..\..\kernel\Table.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Worklist.ps1')

$Manifest = @{
  id         = 'table.upsert'
  group      = 'table'
  summary    = 'Insert or refresh worklist rows by key; merge duplicates and count them'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    worklist  = @{ type='session'; sessionKind='worklist'; required=$true }
    rows      = @{ type='list';   required=$true; desc='records (excel.read_rows / verify.filter_records output)' }
    fields    = @{ type='map';    required=$true; desc='worklist column -> record field; the key columns must be among them' }
    countAs   = @{ type='string'; default=''; desc='worklist column that receives how many input rows had the key' }
    overwrite = @{ type='list';   default=@(); desc='columns refreshed on rows that already exist (others are only filled when blank)' }
    resetOnChange = @{ type='map'; default=@{}; desc='@{ watch = <column>; clear = @(<columns>) }: when an existing row''s watch value changes, the clear columns are emptied (the same deliverable planned again on another day starts over)' }
  }
  outputs    = @{
    added   = @{ type='int' }
    updated = @{ type='int' }
    keyList = @{ type='list'; desc='key display of every row touched, in input order' }
  }
  failures   = @(
    @{ id = 'key_column_missing'; transient = $false }
    @{ id = 'write_failed';       transient = $true  }
  )
  example    = @{ use = 'table.upsert'; with = @{ worklist = 'wl'; rows = '{{steps.today.out.records}}'; fields = @{ Correl_ID_S = 'id'; JOB_NAME = 'job' }; countAs = 'note' } }
  notes      = 'A column named in fields that the worklist lacks is added (and gets the profile default on the other rows). Rows whose key is blank are skipped with a warning.'
}

function TableUpsert-Default {
    param($Profile, [string]$Column)
    $spec = Get-EbiWorklistColumnSpec -Profile $Profile -Field $Column
    if ($null -ne $spec -and $spec.Contains('default') -and $null -ne $spec['default']) { return [string]$spec['default'] }
    return ''
}

function Invoke-Step {
    param($In, $Ctx)
    $wl = $In['worklist']
    $map = $In['fields']
    $keyCols = @($Ctx['KeyColumns']); if ($keyCols.Count -eq 0) { $keyCols = @(Get-EbiWorklistKeyColumns -Profile $Ctx['Profile']) }
    foreach ($k in $keyCols) { if (-not $map.Contains([string]$k)) { return @{ ok = $false; failure = 'key_column_missing'; message = ('fields must map the key column "' + $k + '"'); added = 0; updated = 0; keyList = @() } } }
    $over = @(@($In['overwrite']) | ForEach-Object { [string]$_ })
    $countAs = [string]$In['countAs']
    # merge input rows by key
    $order = New-Object System.Collections.ArrayList
    $merged = @{}
    $warn = New-Object System.Collections.ArrayList
    foreach ($r in @($In['rows'])) {
        if (-not ($r -is [System.Collections.IDictionary])) { continue }
        $vals = @{}
        foreach ($col in $map.Keys) { $f = [string]$map[$col]; $vals[[string]$col] = $(if ($r.Contains($f) -and $null -ne $r[$f]) { ([string]$r[$f]).Trim() } else { '' }) }
        $kd = Get-EbiKeyDisplay -Item $vals -KeyColumns $keyCols
        if ([string]::IsNullOrWhiteSpace($kd.Replace('/', ''))) { [void]$warn.Add(@{ code = 'blank_key'; message = 'an input row has a blank key and was skipped' }); continue }
        if ($merged.Contains($kd)) { $merged[$kd]['_n'] = [int]$merged[$kd]['_n'] + 1; continue }
        $vals['_n'] = 1
        $merged[$kd] = $vals
        [void]$order.Add($kd)
    }
    # columns
    $columns = New-Object System.Collections.ArrayList
    foreach ($c in @($wl['columns'])) { [void]$columns.Add([string]$c) }
    $newCols = @(@($map.Keys) + @(if ($countAs -ne '') { $countAs }) | Where-Object { -not ($columns -contains [string]$_) } | Select-Object -Unique)
    # existing rows by key
    $byKey = @{}
    foreach ($row in @($wl['rows'])) { $byKey[(Get-EbiKeyDisplay -Item $row -KeyColumns $keyCols)] = $row }
    $added = 0; $updated = 0
    $plan = New-Object System.Collections.ArrayList
    foreach ($kd in $order) {
        $vals = $merged[$kd]
        if ($byKey.Contains($kd)) {
            $row = $byKey[$kd]; $changed = $false
            $rc = $In['resetOnChange']
            if ($rc -is [System.Collections.IDictionary] -and $rc.Contains('watch') -and $map.Contains([string]$rc['watch'])) {
                $wc = [string]$rc['watch']
                $old = if ($row.Contains($wc) -and $null -ne $row[$wc]) { [string]$row[$wc] } else { '' }
                if ($old -ne '' -and $old -ne [string]$vals[$wc]) {
                    foreach ($cc in @($rc['clear'])) { $cur = if ($row.Contains([string]$cc)) { [string]$row[[string]$cc] } else { '' }; if ($cur -ne '') { [void]$plan.Add(@{ row = $row; col = [string]$cc; value = '' }); $changed = $true } }
                }
            }
            foreach ($col in $map.Keys) {
                $cur = if ($row.Contains($col) -and $null -ne $row[$col]) { [string]$row[$col] } else { '' }
                if (($over -contains [string]$col) -or $cur -eq '') { if ($cur -ne [string]$vals[$col]) { [void]$plan.Add(@{ row = $row; col = [string]$col; value = [string]$vals[$col] }); $changed = $true } }
            }
            if ($countAs -ne '') { $cur = if ($row.Contains($countAs)) { [string]$row[$countAs] } else { '' }; if ($cur -ne [string]$vals['_n']) { [void]$plan.Add(@{ row = $row; col = $countAs; value = [string]$vals['_n'] }); $changed = $true } }
            if ($changed) { $updated++ }
        } else {
            $added++
        }
    }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would add {0} row(s) and refresh {1} to the worklist' -f $added, $updated)); return @{ ok = $true; added = $added; updated = $updated; keyList = $order.ToArray(); warnings = $warn.ToArray() } }
    $before = @{ columns = @($wl['columns']); rows = @($wl['rows']) }
    $snap = @(foreach ($p in $plan) { @{ row = $p['row']; col = $p['col']; old = $(if ($p['row'].Contains($p['col'])) { $p['row'][$p['col']] } else { $null }) } })
    foreach ($c in $newCols) { [void]$columns.Add([string]$c); foreach ($row in @($wl['rows'])) { if (-not $row.Contains([string]$c)) { $row[[string]$c] = TableUpsert-Default -Profile $Ctx['Profile'] -Column ([string]$c) } } }
    foreach ($p in $plan) { $p['row'][$p['col']] = $p['value'] }
    $rows = New-Object System.Collections.ArrayList
    foreach ($row in @($wl['rows'])) { [void]$rows.Add($row) }
    foreach ($kd in $order) {
        if ($byKey.Contains($kd)) { continue }
        $vals = $merged[$kd]
        $new = @{}
        foreach ($c in $columns) { $new[[string]$c] = TableUpsert-Default -Profile $Ctx['Profile'] -Column ([string]$c) }
        foreach ($col in $map.Keys) { $new[[string]$col] = [string]$vals[$col] }
        if ($countAs -ne '') { $new[$countAs] = [string]$vals['_n'] }
        [void]$rows.Add($new)
    }
    $wl['columns'] = $columns.ToArray()
    $wl['rows'] = $rows.ToArray()
    if (-not [string]::IsNullOrWhiteSpace([string]$wl['path'])) {
        $s = Save-EbiWorklist -Worklist $wl
        if (-not $s['ok']) {
            foreach ($x in $snap) { if ($null -eq $x['old']) { [void]$x['row'].Remove($x['col']) } else { $x['row'][$x['col']] = $x['old'] } }
            $wl['columns'] = $before['columns']; $wl['rows'] = $before['rows']
            return @{ ok = $false; failure = 'write_failed'; message = $s['message']; added = 0; updated = 0; keyList = @() }
        }
    }
    $Ctx.Log.Info(('worklist: {0} row(s) added, {1} refreshed' -f $added, $updated))
    return @{ ok = $true; added = $added; updated = $updated; keyList = $order.ToArray(); warnings = $warn.ToArray() }
}
