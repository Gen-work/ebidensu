#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'ScreenRegion.ps1')

Reset-Tests 'ScreenRegion'

# Target screen ~1980x1020; recommended region (120,280,1250,657) fits.
$r = Resolve-ScreenRegion -X 120 -Y 280 -W 1250 -H 657 -BoundsX 0 -BoundsY 0 -BoundsW 1980 -BoundsH 1020
Assert-Equal 120  $r.X 'in-bounds X unchanged'
Assert-Equal 280  $r.Y 'in-bounds Y unchanged'
Assert-Equal 1250 $r.W 'in-bounds W unchanged'
Assert-Equal 657  $r.H 'in-bounds H unchanged'
Assert-True (-not $r.Clamped) 'in-bounds: not clamped'

# Width overflow -> clamped to screen right edge.
$rw = Resolve-ScreenRegion -X 1900 -Y 280 -W 1250 -H 657 -BoundsX 0 -BoundsY 0 -BoundsW 1980 -BoundsH 1020
Assert-Equal 80 $rw.W 'width clamped to 1980-1900=80'
Assert-True $rw.Clamped 'width overflow: clamped flag'
Assert-True ($rw.Warn -match 'width') 'width overflow: warn mentions width'

# Height overflow.
$rh = Resolve-ScreenRegion -X 120 -Y 1000 -W 1250 -H 657 -BoundsX 0 -BoundsY 0 -BoundsW 1980 -BoundsH 1020
Assert-Equal 20 $rh.H 'height clamped to 1020-1000=20'
Assert-True ($rh.Warn -match 'height') 'height overflow: warn mentions height'

# Negative origin -> clamped to 0.
$rn = Resolve-ScreenRegion -X -10 -Y -5 -W 100 -H 100 -BoundsX 0 -BoundsY 0 -BoundsW 1980 -BoundsH 1020
Assert-Equal 0 $rn.X 'negative X clamped to 0'
Assert-Equal 0 $rn.Y 'negative Y clamped to 0'
Assert-True ($rn.Warn -match 'x' -and $rn.Warn -match 'y') 'negative origin: warns x and y'

exit (Complete-Tests)
