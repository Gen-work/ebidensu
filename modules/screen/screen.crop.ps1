# modules/screen/screen.crop.ps1
# Crop a PNG by per-side pixel amounts. Pure image I/O, no window involved.
# STEP-CONTRACT.md section 8 is this file's spec; the crop itself is
# kernel/Image.ps1's Invoke-EbiCropPng, which also replaced the four
# Invoke-CropPng copies the old tool carried (P1-20). Per-page amounts live
# in the profile and reach this step as inputs.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Image.ps1')

$Manifest = @{
  id         = 'screen.crop'
  group      = 'screen'
  summary    = 'Crop a PNG by per-side pixel amounts and write the result'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path   = @{ type='path'; required=$true; desc='PNG to crop, modified in place unless out is given' }
    out    = @{ type='path'; default=''; desc='write here instead of overwriting path' }
    left   = @{ type='int'; default=0 }
    top    = @{ type='int'; default=0 }
    right  = @{ type='int'; default=0 }
    bottom = @{ type='int'; default=0 }
  }
  outputs    = @{
    path   = @{ type='path'; desc='the file that was written' }
    width  = @{ type='int' }
    height = @{ type='int' }
  }
  failures   = @(
    @{ id = 'file_not_found';     transient = $false }
    @{ id = 'crop_exceeds_image'; transient = $false }
    @{ id = 'image_read_error';   transient = $true  }
  )
  example    = @{
    use  = 'screen.crop'
    with = @{ path = '{{steps.shot.out.path}}'; left = 6; top = 6; right = 6; bottom = 6 }
  }
  notes      = 'Idempotent only in the "out" form: cropping in place twice takes the border off twice. Zero on all four sides copies (or leaves) the file untouched.'
}

function Invoke-Step {
    param($In, $Ctx)
    $src = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    $out = Resolve-EbiWorkPath -PathValue ([string]$In['out']) -WorkDir ([string]$Ctx['WorkDir'])
    $dest = if ($out -ne '') { $out } else { $src }
    $l = [int]$In['left']; $t = [int]$In['top']; $r = [int]$In['right']; $b = [int]$In['bottom']
    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would crop {0} by L{1} T{2} R{3} B{4} -> {5}' -f $src, $l, $t, $r, $b, $dest))
        return @{ ok = $true; path = $dest; width = 0; height = 0 }
    }
    if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = $src; path = $dest; width = 0; height = 0 } }
    $res = Invoke-EbiCropPng -Path $src -Out $out -Left $l -Top $t -Right $r -Bottom $b
    if (-not $res['ok']) { return @{ ok = $false; failure = $res['failure']; message = $res['message']; path = $dest; width = $res['width']; height = $res['height'] } }
    return @{ ok = $true; path = $res['path']; width = $res['width']; height = $res['height'] }
}
