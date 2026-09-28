# modules/screen/screen.capture_region.ps1
# Screenshot a screen rectangle to a PNG. The rectangle is clamped into the
# virtual screen and the clamp is REPORTED, never silent (P1-18, ported from
# ScreenRegion.ps1 Resolve-ScreenRegion + DfSnap's region capture).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Image.ps1')

$Manifest = @{
  id         = 'screen.capture_region'
  group      = 'screen'
  summary    = 'Save a PNG of a screen rectangle, clamped to the screen'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    x      = @{ type='int';  required=$true; desc='left edge, screen pixels' }
    y      = @{ type='int';  required=$true; desc='top edge, screen pixels' }
    width  = @{ type='int';  required=$true }
    height = @{ type='int';  required=$true }
    saveAs = @{ type='path'; required=$true; desc='PNG path; relative paths resolve under the work dir' }
  }
  outputs    = @{
    path         = @{ type='path' }
    width        = @{ type='int';    desc='pixels actually captured (after the clamp)' }
    height       = @{ type='int' }
    clamped      = @{ type='bool';   desc='the rectangle was cut down to the screen' }
    clampedEdges = @{ type='string'; desc='which of x,y,width,height moved; empty when none' }
  }
  failures   = @(
    @{ id = 'region_empty'; transient = $false }
    @{ id = 'save_failed';  transient = $true  }
  )
  example    = @{ use = 'screen.capture_region'; with = @{ x = 100; y = 200; width = 800; height = 300; saveAs = 'capture/df/{{item.keySafe}}__result.png' } }
  notes      = 'A clamp is also a warning (region_clamped) so it shows in the run log; a rectangle entirely off screen is region_empty.'
}

function Invoke-Step {
    param($In, $Ctx)
    $dest = Resolve-EbiWorkPath -PathValue ([string]$In['saveAs']) -WorkDir ([string]$Ctx['WorkDir'])
    $x = [int]$In['x']; $y = [int]$In['y']; $w = [int]$In['width']; $h = [int]$In['height']
    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would capture {0},{1} {2}x{3} to {4}' -f $x, $y, $w, $h, $dest))
        return @{ ok = $true; path = $dest; width = $w; height = $h; clamped = $false; clampedEdges = '' }
    }
    $screen = Get-EbiVirtualScreen
    $r = Resolve-EbiScreenRegion -X $x -Y $y -W $w -H $h -BoundsX $screen['X'] -BoundsY $screen['Y'] -BoundsW $screen['W'] -BoundsH $screen['H']
    if ($r['w'] -le 0 -or $r['h'] -le 0) {
        return @{ ok = $false; failure = 'region_empty'; message = ('nothing of {0},{1} {2}x{3} lies on the screen' -f $x, $y, $w, $h); path = $dest; width = 0; height = 0; clamped = $r['clamped']; clampedEdges = $r['edges'] }
    }
    $warnings = @()
    if ($r['clamped']) { $warnings = @( @{ code = 'region_clamped'; message = ('rectangle clamped on ' + $r['edges']); data = @{ x = $r['x']; y = $r['y']; width = $r['w']; height = $r['h'] } } ) }
    $grab = Save-EbiScreenRegionPng -X $r['x'] -Y $r['y'] -W $r['w'] -H $r['h'] -Dest $dest
    if (-not $grab['ok']) { return @{ ok = $false; failure = 'save_failed'; message = $grab['message']; path = $dest; width = $r['w']; height = $r['h']; clamped = $r['clamped']; clampedEdges = $r['edges'] } }
    return @{ ok = $true; path = $dest; width = $r['w']; height = $r['h']; clamped = $r['clamped']; clampedEdges = $r['edges']; warnings = $warnings }
}
