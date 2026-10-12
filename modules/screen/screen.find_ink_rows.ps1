# modules/screen/screen.find_ink_rows.ps1
# Which pixel rows of a vertical strip of a PNG carry ink, grouped into
# bands (one band = one text line): the "where is the last line" and
# "where is each list row" question answered from the picture itself
# rather than from a fixed coordinate that moves with every window size.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Image.ps1')     # Get-EbiInkCounts
. (Join-Path $PSScriptRoot '..\..\kernel\Layout.ps1')    # Get-EbiInkBands

$Manifest = @{
  id         = 'screen.find_ink_rows'
  group      = 'screen'
  summary    = 'Find the bands of pixel rows carrying ink in a vertical strip of a PNG'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path     = @{ type='path';   required=$true }
    x        = @{ type='int';    required=$true; desc='strip left edge, px' }
    width    = @{ type='int';    required=$true }
    yFrom    = @{ type='int';    default=0 }
    yTo      = @{ type='int';    default=0; desc='0 = the bottom; negative = that many px above the bottom' }
    ink      = @{ type='string'; default='dark'; enum=@('dark', 'blue') }
    maxLuma  = @{ type='int';    default=160; desc='dark: pixels darker than this' }
    minInk   = @{ type='int';    default=2; desc='a row needs this many ink pixels' }
    mergeGap = @{ type='int';    default=1; desc='blank rows a band may contain' }
    minHeight = @{ type='int';   default=4; desc='thinner bands are dropped (underlines, borders)' }
  }
  outputs    = @{
    bands = @{ type='list'; desc='@{ top; bottom; height; center } in image px, top to bottom' }
    last  = @{ type='map';  desc='the bottom band, or empty' }
    lines = @{ type='int' }
  }
  failures   = @(
    @{ id = 'file_not_found';   transient = $true  }
    @{ id = 'image_read_error'; transient = $true  }
  )
  example    = @{ use = 'screen.find_ink_rows'; with = @{ path = '{{steps.shot.out.path}}'; x = 46; width = 38; yFrom = 105; yTo = -35 } }
}

function Invoke-Step {
    param($In, $Ctx)
    $p = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    if ($Ctx['DryRun'] -and -not (Test-Path -LiteralPath $p -PathType Leaf)) { $Ctx.Log.Info(('would scan {0}' -f $p)); return @{ ok = $true; bands = @(); last = @{}; lines = 0 } }
    $c = Get-EbiInkCounts -Path $p -X ([int]$In['x']) -Width ([int]$In['width']) -YFrom ([int]$In['yFrom']) -YTo ([int]$In['yTo']) -Ink ([string]$In['ink']) -MaxLuma ([int]$In['maxLuma'])
    if (-not $c['ok']) { return @{ ok = $false; failure = $c['failure']; message = $c['message']; bands = @(); last = @{}; lines = 0 } }
    $y0 = [int]$c['yFrom']
    $bands = @(foreach ($b in @(Get-EbiInkBands -Counts ([int[]]$c['counts']) -MinInk ([int]$In['minInk']) -MergeGap ([int]$In['mergeGap']))) {
        if ([int]$b['height'] -lt [int]$In['minHeight']) { continue }
        @{ top = $y0 + [int]$b['top']; bottom = $y0 + [int]$b['bottom']; height = [int]$b['height']; center = ($y0 + ([int]$b['top'] + [int]$b['bottom']) / 2.0) }
    })
    $last = if ($bands.Count -gt 0) { $bands[$bands.Count - 1] } else { @{} }
    return @{ ok = $true; bands = $bands; last = $last; lines = $bands.Count }
}
