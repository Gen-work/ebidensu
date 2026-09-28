# modules/table/table.load.ps1
# Load the worklist CSV into memory and register it as a 'worklist' session
# resource (P1-24, P0-R11). Ported from MappingStore.ps1 Import-Mapping.
# The whole table never appears in outputs; later steps reach it through a
# type='session' input and the runner iterates the same instance.
#
# Two checks nobody may skip: every key column exists, and no two rows
# share a keySafe (PROFILE-SCHEMA 6.6 a) -- a collision would overwrite
# captures silently, so it is refused here, not discovered mid-run.
# Columns the profile declares but the file lacks are added with their
# defaults (Ensure-MappingColumns' behaviour, PROFILE-SCHEMA 6.5).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Table.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Worklist.ps1')

$Manifest = @{
  id         = 'table.load'
  group      = 'table'
  summary    = 'Load the worklist CSV and register it as a worklist resource'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @('worklist')
  releases   = @()
  idempotent = $true
  inputs     = @{
    path       = @{ type='path'; required=$true; desc='the CSV (UTF-8 with BOM); relative paths resolve under the work dir' }
    keyColumns = @{ type='list'; default=@(); desc='override of profile.worklist.key.columns' }
    mustExist  = @{ type='bool'; default=$true; desc='false: a missing file loads as an empty table with the profile columns' }
  }
  outputs    = @{
    path     = @{ type='path' }
    rowCount = @{ type='int' }
    columns  = @{ type='list'; desc='column names in file order, profile columns appended' }
  }
  failures   = @(
    @{ id = 'file_not_found';     transient = $false }
    @{ id = 'csv_invalid';        transient = $false }
    @{ id = 'key_column_missing'; transient = $false }
    @{ id = 'key_collision';      transient = $false }
  )
  example    = @{ use = 'table.load'; with = @{ path = '{{profile.worklist.file}}'; as = 'wl' } }
  notes      = 'provides is non-empty, so a resumed run loads the table again and sees the disk. A dry run reads the file too (read-only); a file that is not there dry-runs as an empty table with a warning.'
}

function TableLoad-ProfileColumns {
    # PURE. The profile's declared columns as @(@{ name; default }).
    param($Profile)
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $Profile -or -not ($Profile -is [System.Collections.IDictionary]) -or -not $Profile.Contains('worklist') -or -not ($Profile['worklist'] -is [System.Collections.IDictionary])) { return $out.ToArray() }
    $wl = $Profile['worklist']
    if (-not $wl.Contains('columns') -or $null -eq $wl['columns']) { return $out.ToArray() }
    foreach ($c in $wl['columns']) {
        if (-not ($c -is [System.Collections.IDictionary]) -or -not $c.Contains('name')) { continue }
        $d = if ($c.Contains('default') -and $null -ne $c['default']) { [string]$c['default'] } else { '' }
        [void]$out.Add(@{ name = [string]$c['name']; default = $d })
    }
    return $out.ToArray()
}

function TableLoad-AddColumns {
    # PURE (mutates the given rows). Adds missing columns with defaults;
    # returns the names added.
    param($Columns, $Rows, $Declared)
    $cols = New-Object System.Collections.ArrayList
    foreach ($c in @($Columns)) { [void]$cols.Add([string]$c) }
    $added = New-Object System.Collections.ArrayList
    foreach ($d in @($Declared)) {
        if ($cols -contains $d['name']) { continue }
        [void]$cols.Add($d['name']); [void]$added.Add($d['name'])
        foreach ($r in @($Rows)) { if ($null -ne $r -and -not $r.Contains($d['name'])) { $r[$d['name']] = $d['default'] } }
    }
    return @{ columns = $cols.ToArray(); added = $added.ToArray() }
}

function Invoke-Step {
    param($In, $Ctx)
    $path = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    $keyColumns = @(@($In['keyColumns']) | ForEach-Object { [string]$_ } | Where-Object { $_ -ne '' })
    if ($keyColumns.Count -eq 0) { $keyColumns = @(Get-EbiWorklistKeyColumns -Profile $Ctx['Profile']) }
    $declared = @(TableLoad-ProfileColumns -Profile $Ctx['Profile'])
    $warnings = New-Object System.Collections.ArrayList

    $read = Read-EbiCsv -Path $path
    if (-not $read['exists']) {
        if ([bool]$In['mustExist'] -and -not $Ctx['DryRun']) { return @{ ok = $false; failure = 'file_not_found'; message = $path; path = $path; rowCount = 0; columns = @() } }
        [void]$warnings.Add(@{ code = 'not_found'; message = ('no worklist at ' + $path + '; loaded as an empty table'); data = @{} })
        $read = @{ ok = $true; columns = @(); rows = @(); message = ''; exists = $false }
    } elseif (-not $read['ok']) {
        return @{ ok = $false; failure = 'csv_invalid'; message = $read['message']; path = $path; rowCount = 0; columns = @() }
    }
    $rows = @($read['rows'])
    $fix = TableLoad-AddColumns -Columns $read['columns'] -Rows $rows -Declared $declared
    $columns = @($fix['columns'])
    if (@($fix['added']).Count -gt 0) { [void]$warnings.Add(@{ code = 'columns_added'; message = ('added profile column(s): ' + (@($fix['added']) -join ', ')); data = @{ added = @($fix['added']) } }) }
    foreach ($kc in $keyColumns) {
        if (-not ($columns -contains $kc)) { return @{ ok = $false; failure = 'key_column_missing'; message = ('key column "' + $kc + '" is not in ' + $path + ' (columns: ' + ($columns -join ', ') + ')'); path = $path; rowCount = $rows.Count; columns = $columns } }
    }
    if ($keyColumns.Count -gt 0) {
        $coll = Find-EbiKeySafeCollisions -Rows $rows -KeyColumns $keyColumns
        if (-not $coll['ok']) {
            $lines = @(@($coll['collisions']) | ForEach-Object { $_['keySafe'] + ' <- ' + (@($_['keys']) -join ' | ') })
            return @{ ok = $false; failure = 'key_collision'; message = ('' + @($coll['collisions']).Count + ' keySafe collision(s); files would overwrite each other: ' + ($lines -join '; ')); path = $path; rowCount = $rows.Count; columns = $columns }
        }
    } else {
        [void]$warnings.Add(@{ code = 'no_key_columns'; message = 'no key columns declared (profile.worklist.key.columns or keyColumns); rows have no key'; data = @{} })
    }
    $wl = New-EbiWorklist -Path $path -Columns $columns -Rows $rows
    $Ctx.Log.Info(('loaded {0} row(s), {1} column(s) from {2}' -f $rows.Count, $columns.Count, $path))
    return @{ ok = $true; resource = $wl; path = $path; rowCount = $rows.Count; columns = $columns; warnings = $warnings.ToArray() }
}
