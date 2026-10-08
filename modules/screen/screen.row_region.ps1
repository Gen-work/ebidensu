# modules/screen/screen.row_region.ps1
# The screen rectangle that covers a list's header band plus a run of its
# rows (kernel/Layout.ps1 Get-EbiRowRegion): which rows comes from
# verify.filter_records (first, matched), the fixed geometry from the
# profile. Coordinates are relative to the window when a window is given
# (its current screen position is added), else screen coordinates.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Layout.ps1')

$Manifest = @{
  id         = 'screen.row_region'
  group      = 'screen'
  summary    = 'Compute the rectangle covering a list header plus a run of rows'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    window      = @{ type='session'; sessionKind='window'; desc='make the geometry relative to this window''s top-left' }
    insetX      = @{ type='int'; default=0; desc='added to the window''s left (a maximized window''s frame sits off-screen)' }
    insetY      = @{ type='int'; default=0 }
    left        = @{ type='int'; required=$true }
    right       = @{ type='int'; required=$true }
    top         = @{ type='int'; required=$true; desc='top of the region (the panel title)' }
    firstRowTop = @{ type='int'; required=$true; desc='top edge of list row 1' }
    rowHeight   = @{ type='int'; required=$true }
    first       = @{ type='int'; default=1; desc='1-based list row to start at' }
    rows        = @{ type='int'; default=1; desc='how many rows' }
    bottomPad   = @{ type='int'; default=2 }
  }
  outputs    = @{
    x = @{ type='int' }
    y = @{ type='int' }
    width = @{ type='int' }
    height = @{ type='int' }
    skippedRows = @{ type='int'; desc='rows above first that the region also shows (they cannot be cut out)' }
  }
  failures   = @(
    @{ id = 'window_gone'; transient = $true }
  )
  example    = @{ use = 'screen.row_region'; with = @{ left = 590; right = 1560; top = 146; firstRowTop = 268; rowHeight = 31; first = '{{steps.mine.out.first}}'; rows = '{{steps.mine.out.matched}}' } }
  notes      = 'skippedRows > 0 is reported as a warning: the picture then shows newer rows above the ones meant.'
}

function Invoke-Step {
    param($In, $Ctx)
    $ox = 0; $oy = 0
    if ($In.Contains('window') -and $null -ne $In['window'] -and -not $Ctx['DryRun']) {
        $r = Get-EbiWindowRect -HWnd (ConvertTo-EbiHandle $In['window'])
        if (-not $r['ok']) { return @{ ok = $false; failure = 'window_gone'; message = 'the window has no rectangle'; x = 0; y = 0; width = 0; height = 0; skippedRows = 0 } }
        $ox = [int]$r['X'] + [int]$In['insetX']; $oy = [int]$r['Y'] + [int]$In['insetY']
    }
    $g = Get-EbiRowRegion -OriginX $ox -OriginY $oy -Left ([int]$In['left']) -Right ([int]$In['right']) -Top ([int]$In['top']) -FirstRowTop ([int]$In['firstRowTop']) -RowHeight ([double]$In['rowHeight']) -FirstIndex ([int]$In['first']) -Count ([int]$In['rows']) -BottomPad ([int]$In['bottomPad'])
    $out = @{ ok = $true; x = [int]$g['x']; y = [int]$g['y']; width = [int]$g['width']; height = [int]$g['height']; skippedRows = [int]$g['skippedRows'] }
    if ([int]$g['skippedRows'] -gt 0) { $out['warnings'] = @(@{ code = 'rows_skipped'; message = ('the region also shows ' + $g['skippedRows'] + ' newer row(s) above the chosen ones') }) }
    return $out
}
