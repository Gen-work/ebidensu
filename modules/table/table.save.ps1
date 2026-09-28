# modules/table/table.save.ps1
# Write the worklist back to disk atomically (P1-24, ported from
# MappingStore.ps1 Export-MappingAtomic). Every writing step already flushes
# before it returns; this is for a workflow that wants a copy elsewhere or
# a flush after an in-memory-only change.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Table.ps1')

$Manifest = @{
  id         = 'table.save'
  group      = 'table'
  summary    = 'Write the worklist to CSV (UTF-8 with BOM, atomic)'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    worklist = @{ type='session'; sessionKind='worklist'; required=$true; desc='the loaded table' }
    path     = @{ type='path'; default=''; desc='write here instead of the path it was loaded from' }
  }
  outputs    = @{
    path     = @{ type='path' }
    rowCount = @{ type='int' }
  }
  failures   = @(
    @{ id = 'write_failed'; transient = $true  }
    @{ id = 'no_path';      transient = $false }
  )
  example    = @{ use = 'table.save'; with = @{ worklist = 'wl' } }
}

function Invoke-Step {
    param($In, $Ctx)
    $wl = $In['worklist']
    $path = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    if ($path -eq '' -and $null -ne $wl -and ($wl -is [System.Collections.IDictionary]) -and $wl.Contains('path')) { $path = [string]$wl['path'] }
    $n = if ($null -ne $wl -and ($wl -is [System.Collections.IDictionary]) -and $wl.Contains('rows')) { @($wl['rows']).Count } else { 0 }
    if ([string]::IsNullOrWhiteSpace($path)) { return @{ ok = $false; failure = 'no_path'; message = 'the worklist has no path and none was given'; path = ''; rowCount = $n } }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would write {0} row(s) to {1}' -f $n, $path)); return @{ ok = $true; path = $path; rowCount = $n } }
    $w = Save-EbiWorklist -Worklist $wl -Path $path
    if (-not $w['ok']) { return @{ ok = $false; failure = 'write_failed'; message = $w['message']; path = $path; rowCount = $n } }
    return @{ ok = $true; path = $path; rowCount = $n }
}
