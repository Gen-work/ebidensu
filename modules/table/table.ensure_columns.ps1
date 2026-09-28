# modules/table/table.ensure_columns.ps1
# Add columns the table lacks (P1-25, ported from MappingStore.ps1
# Ensure-MappingColumns). With no `columns` input the profile's declared
# columns are the schema; a workflow can add its own on top. Flushes to
# disk when anything changed (STEP-CONTRACT 3.2).

. (Join-Path $PSScriptRoot '..\..\kernel\Table.ps1')

$Manifest = @{
  id         = 'table.ensure_columns'
  group      = 'table'
  summary    = 'Add missing columns (profile schema plus extras) with defaults'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    worklist = @{ type='session'; sessionKind='worklist'; required=$true }
    columns  = @{ type='map'; default=@{}; desc='extra columns: name -> default value; profile columns are always included' }
  }
  outputs    = @{
    added   = @{ type='list'; desc='names that were missing and got added' }
    columns = @{ type='list'; desc='all column names afterwards' }
  }
  failures   = @(
    @{ id = 'write_failed'; transient = $true }
  )
  example    = @{ use = 'table.ensure_columns'; with = @{ worklist = 'wl'; columns = @{ note = '' } } }
}

function TableEnsureColumns-Wanted {
    # PURE. profile columns + extras -> ordered @(@{ name; default }).
    param($Profile, $Extra)
    $out = New-Object System.Collections.ArrayList
    $seen = @{}
    if ($null -ne $Profile -and ($Profile -is [System.Collections.IDictionary]) -and $Profile.Contains('worklist') -and ($Profile['worklist'] -is [System.Collections.IDictionary]) -and $Profile['worklist'].Contains('columns') -and $null -ne $Profile['worklist']['columns']) {
        foreach ($c in $Profile['worklist']['columns']) {
            if (-not ($c -is [System.Collections.IDictionary]) -or -not $c.Contains('name')) { continue }
            $n = [string]$c['name']; if ($seen.Contains($n)) { continue }; $seen[$n] = $true
            [void]$out.Add(@{ name = $n; default = $(if ($c.Contains('default') -and $null -ne $c['default']) { [string]$c['default'] } else { '' }) })
        }
    }
    if ($null -ne $Extra -and ($Extra -is [System.Collections.IDictionary])) {
        foreach ($k in ($Extra.Keys | Sort-Object)) {
            $n = [string]$k; if ($seen.Contains($n)) { continue }; $seen[$n] = $true
            [void]$out.Add(@{ name = $n; default = $(if ($null -eq $Extra[$k]) { '' } else { [string]$Extra[$k] }) })
        }
    }
    return $out.ToArray()
}

function Invoke-Step {
    param($In, $Ctx)
    $wl = $In['worklist']
    $wanted = @(TableEnsureColumns-Wanted -Profile $Ctx['Profile'] -Extra $In['columns'])
    $cols = New-Object System.Collections.ArrayList
    foreach ($c in @($wl['columns'])) { [void]$cols.Add([string]$c) }
    $added = New-Object System.Collections.ArrayList
    foreach ($w in $wanted) { if (-not ($cols -contains $w['name'])) { [void]$added.Add($w['name']) } }
    if ($added.Count -eq 0) { if ($Ctx['DryRun']) { $Ctx.Log.Info('nothing to add: every wanted column exists') }; return @{ ok = $true; added = @(); columns = $cols.ToArray() } }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would add column(s): ' + ($added.ToArray() -join ', '))); return @{ ok = $true; added = $added.ToArray(); columns = @($cols.ToArray() + $added.ToArray()) } }
    foreach ($w in $wanted) {
        if ($cols -contains $w['name']) { continue }
        [void]$cols.Add($w['name'])
        foreach ($r in @($wl['rows'])) { if ($null -ne $r -and -not $r.Contains($w['name'])) { $r[$w['name']] = $w['default'] } }
    }
    $wl['columns'] = $cols.ToArray()
    if (-not [string]::IsNullOrWhiteSpace([string]$wl['path'])) {
        $s = Save-EbiWorklist -Worklist $wl
        if (-not $s['ok']) { return @{ ok = $false; failure = 'write_failed'; message = $s['message']; added = $added.ToArray(); columns = $cols.ToArray() } }
    }
    return @{ ok = $true; added = $added.ToArray(); columns = $cols.ToArray() }
}
