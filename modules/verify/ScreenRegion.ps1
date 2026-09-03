# ============================================================
#  ScreenRegion.ps1
#
#  PURE screen-region clamping math -- no System.Windows.Forms / Graphics,
#  so it is unit-testable without a desktop (Tests\Test-ScreenRegion.ps1).
#  DfSnap.ps1 feeds it the real primary-screen bounds.
#  Dot-source only (no param() block).
# ============================================================

# Clamp a requested capture rectangle into the given screen bounds.
# Returns X/Y/W/H plus Clamped (bool) and Warn (csv of clamped edges).
function Resolve-ScreenRegion {
    param(
        [int]$X, [int]$Y, [int]$W, [int]$H,
        [int]$BoundsX, [int]$BoundsY, [int]$BoundsW, [int]$BoundsH
    )
    $right  = $BoundsX + $BoundsW
    $bottom = $BoundsY + $BoundsH
    $warn = [System.Collections.Generic.List[string]]::new()

    if ($X -lt $BoundsX) { $X = $BoundsX; $warn.Add('x') }
    if ($Y -lt $BoundsY) { $Y = $BoundsY; $warn.Add('y') }
    if (($X + $W) -gt $right)  { $W = $right  - $X; $warn.Add('width') }
    if (($Y + $H) -gt $bottom) { $H = $bottom - $Y; $warn.Add('height') }
    if ($W -lt 0) { $W = 0 }
    if ($H -lt 0) { $H = 0 }

    return [pscustomobject]@{
        X = $X; Y = $Y; W = $W; H = $H
        Clamped = ($warn.Count -gt 0)
        Warn    = ($warn -join ',')
    }
}

# Resolve the four per-side crop amounts (Left/Top/Right/Bottom, px) for a
# snap screenshot from: a legacy uniform CropPx, optional global per-side
# overrides (-1 = inherit CropPx for that side), and an optional per-folder
# override hashtable (VerifyConfig.psd1 Window.CropByFolder.<folder>, e.g.
# @{ Top = 10 }). Precedence per side: FolderOverride[<Side>] (if the key is
# present) > global Crop<Side> (when >= 0) > CropPx. Used by VerifyTool.ps1
# to turn Window.CropPx/CropLeft/CropTop/CropRight/CropBottom/CropByFolder
# into the concrete non-negative ints threaded to HmSnap/MqSnap/JenkinsSnap/
# Crop-Snap's -CropLeft/-CropTop/-CropRight/-CropBottom params.
function Resolve-DirectionalCrop {
    param(
        [int]$CropPx              = 0,
        [int]$CropLeft            = -1,
        [int]$CropTop             = -1,
        [int]$CropRight           = -1,
        [int]$CropBottom          = -1,
        [hashtable]$FolderOverride = $null
    )
    $left   = if ($CropLeft   -ge 0) { $CropLeft }   else { $CropPx }
    $top    = if ($CropTop    -ge 0) { $CropTop }    else { $CropPx }
    $right  = if ($CropRight  -ge 0) { $CropRight }  else { $CropPx }
    $bottom = if ($CropBottom -ge 0) { $CropBottom } else { $CropPx }

    if ($FolderOverride) {
        if ($FolderOverride.ContainsKey('Left')   -and $null -ne $FolderOverride['Left'])   { $left   = [int]$FolderOverride['Left'] }
        if ($FolderOverride.ContainsKey('Top')    -and $null -ne $FolderOverride['Top'])    { $top    = [int]$FolderOverride['Top'] }
        if ($FolderOverride.ContainsKey('Right')  -and $null -ne $FolderOverride['Right'])  { $right  = [int]$FolderOverride['Right'] }
        if ($FolderOverride.ContainsKey('Bottom') -and $null -ne $FolderOverride['Bottom']) { $bottom = [int]$FolderOverride['Bottom'] }
    }

    if ($left   -lt 0) { $left   = 0 }
    if ($top    -lt 0) { $top    = 0 }
    if ($right  -lt 0) { $right  = 0 }
    if ($bottom -lt 0) { $bottom = 0 }

    return [pscustomobject]@{ Left = $left; Top = $top; Right = $right; Bottom = $bottom }
}
