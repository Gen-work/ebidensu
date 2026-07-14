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


# ---- Resolve-DirectionalCrop ----

# All sentinels (-1) -> uniform CropPx on every side (legacy behavior).
$dcUniform = Resolve-DirectionalCrop -CropPx 6
Assert-Equal 6 $dcUniform.Left   'uniform: Left = CropPx'
Assert-Equal 6 $dcUniform.Top    'uniform: Top = CropPx'
Assert-Equal 6 $dcUniform.Right  'uniform: Right = CropPx'
Assert-Equal 6 $dcUniform.Bottom 'uniform: Bottom = CropPx'

# CropPx = 0, all sides sentinel -> no-op crop (all sides 0).
$dcZero = Resolve-DirectionalCrop -CropPx 0
Assert-Equal 0 $dcZero.Left   'CropPx=0: Left = 0'
Assert-Equal 0 $dcZero.Bottom 'CropPx=0: Bottom = 0'

# Global per-side override wins over CropPx for that side only.
$dcGlobal = Resolve-DirectionalCrop -CropPx 6 -CropTop 10
Assert-Equal 6  $dcGlobal.Left   'global override: Left still CropPx'
Assert-Equal 10 $dcGlobal.Top    'global override: Top uses CropTop'
Assert-Equal 6  $dcGlobal.Right  'global override: Right still CropPx'
Assert-Equal 6  $dcGlobal.Bottom 'global override: Bottom still CropPx'

# All four global sides set explicitly.
$dcAllGlobal = Resolve-DirectionalCrop -CropPx 6 -CropLeft 1 -CropTop 2 -CropRight 3 -CropBottom 4
Assert-Equal 1 $dcAllGlobal.Left   'all-global: Left'
Assert-Equal 2 $dcAllGlobal.Top    'all-global: Top'
Assert-Equal 3 $dcAllGlobal.Right  'all-global: Right'
Assert-Equal 4 $dcAllGlobal.Bottom 'all-global: Bottom'

# Per-folder override wins over the global value, only for the keys it sets.
$dcFolder = Resolve-DirectionalCrop -CropPx 6 -CropLeft 1 -CropTop 2 -CropRight 3 -CropBottom 4 `
    -FolderOverride @{ Top = 20 }
Assert-Equal 1  $dcFolder.Left   'folder override: Left untouched (global)'
Assert-Equal 20 $dcFolder.Top    'folder override: Top from folder'
Assert-Equal 3  $dcFolder.Right  'folder override: Right untouched (global)'
Assert-Equal 4  $dcFolder.Bottom 'folder override: Bottom untouched (global)'

# Per-folder override can set all four sides, ignoring CropPx/global entirely.
$dcFolderAll = Resolve-DirectionalCrop -CropPx 6 -FolderOverride @{ Left = 1; Top = 2; Right = 3; Bottom = 4 }
Assert-Equal 1 $dcFolderAll.Left   'folder-all: Left'
Assert-Equal 2 $dcFolderAll.Top    'folder-all: Top'
Assert-Equal 3 $dcFolderAll.Right  'folder-all: Right'
Assert-Equal 4 $dcFolderAll.Bottom 'folder-all: Bottom'

# Negative CropPx floors at 0 (defensive; should never happen via VerifyTool).
$dcNeg = Resolve-DirectionalCrop -CropPx -5
Assert-Equal 0 $dcNeg.Left 'negative CropPx floors at 0'

exit (Complete-Tests)
