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
