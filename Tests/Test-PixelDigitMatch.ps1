#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'PixelDigitMatch.ps1')

Reset-Tests 'PixelDigitMatch'

# Build an ink image (@{Ink;W;H}) from a row-major 0/1 matrix (1 = black ink).
function New-InkImage {
    param([int[][]]$Rows)
    $h = $Rows.Count; $w = $Rows[0].Count
    $ink = New-Object 'double[]' ($w * $h)
    for ($y = 0; $y -lt $h; $y++) {
        for ($x = 0; $x -lt $w; $x++) { $ink[$y * $w + $x] = [double]$Rows[$y][$x] }
    }
    return @{ Ink = $ink; W = $w; H = $h }
}

# --- ConvertTo-DigitInk ---------------------------------------------------
$ink = ConvertTo-DigitInk ([double[]]@(0, 255, 128))
Assert-Equal 1 ([int][Math]::Round($ink[0])) 'black (0) -> ink 1'
Assert-Equal 0 ([int][Math]::Round($ink[1])) 'white (255) -> ink 0'
Assert-True ([Math]::Abs($ink[2] - 0.498) -lt 0.01) 'mid-gray (128) -> ink ~0.498'

# --- Get-DigitInkBBox -----------------------------------------------------
$boxImg = New-InkImage @(
    @(0, 0, 0, 0),
    @(0, 1, 1, 0),
    @(0, 1, 1, 0),
    @(0, 0, 0, 0)
)
$box = Get-DigitInkBBox -Ink $boxImg.Ink -Width $boxImg.W -Height $boxImg.H
Assert-Equal 1 $box.X0 'bbox X0'
Assert-Equal 1 $box.Y0 'bbox Y0'
Assert-Equal 3 $box.X1 'bbox X1 (exclusive)'
Assert-Equal 3 $box.Y1 'bbox Y1 (exclusive)'
$blank = New-InkImage @(@(0, 0), @(0, 0))
Assert-Equal $null (Get-DigitInkBBox -Ink $blank.Ink -Width $blank.W -Height $blank.H) 'blank image -> null bbox'

# --- ConvertTo-DigitNormalizedGrid : a solid box -> full-ink grid ---------
$solid = New-InkImage @(
    @(0, 0, 0, 0, 0),
    @(0, 1, 1, 1, 0),
    @(0, 1, 1, 1, 0),
    @(0, 0, 0, 0, 0)
)
$sbox = Get-DigitInkBBox -Ink $solid.Ink -Width $solid.W -Height $solid.H
$grid = ConvertTo-DigitNormalizedGrid -Ink $solid.Ink -Width $solid.W -Height $solid.H -Box $sbox -GridW 3 -GridH 2
$allFull = $true
foreach ($v in $grid) { if ($v -le 0.9) { $allFull = $false } }
Assert-True $allFull 'a solid box average-pools to an all-ink normalized grid'

# --- Get-DigitNcc ---------------------------------------------------------
Assert-True ([Math]::Abs((Get-DigitNcc -A ([double[]]@(1,0,1,0)) -B ([double[]]@(1,0,1,0))) - 1) -lt 1e-9) 'identical vectors -> 1'
Assert-True ((Get-DigitNcc -A ([double[]]@(1,0,1,0)) -B ([double[]]@(0,1,0,1))) -lt -0.9) 'opposite pattern -> negative'
Assert-Equal 0 (Get-DigitNcc -A ([double[]]@(1,1,1,1)) -B ([double[]]@(1,0,1,0))) 'flat vs varying -> 0'
Assert-Equal 1 (Get-DigitNcc -A ([double[]]@(1,1,1)) -B ([double[]]@(1,1,1))) 'two flat -> 1'

# --- Get-DigitSimilarity : same glyph (scaled) beats a different glyph -----
$L = New-InkImage @(
    @(1, 0, 0),
    @(1, 0, 0),
    @(1, 1, 1)
)
$Lbig = New-InkImage @(
    @(1, 1, 0, 0, 0, 0),
    @(1, 1, 0, 0, 0, 0),
    @(1, 1, 0, 0, 0, 0),
    @(1, 1, 0, 0, 0, 0),
    @(1, 1, 1, 1, 1, 1),
    @(1, 1, 1, 1, 1, 1)
)
$T = New-InkImage @(
    @(1, 1, 1),
    @(0, 1, 0),
    @(0, 1, 0)
)
$sSame = Get-DigitSimilarity -InkA $L.Ink -WA $L.W -HA $L.H -InkB $Lbig.Ink -WB $Lbig.W -HB $Lbig.H -GridW 6 -GridH 6
$sDiff = Get-DigitSimilarity -InkA $L.Ink -WA $L.W -HA $L.H -InkB $T.Ink -WB $T.W -HB $T.H -GridW 6 -GridH 6
Assert-True ($sSame -gt $sDiff) 'scaled same-glyph similarity beats cross-glyph'
Assert-True ($sSame -gt 0.8) 'scaled same glyph scores high'

# --- Compare-DigitCandidate + Get-DigitPixelVerdict : the 3/9 decision -----
# Stylized 5-row 3 (open left) and 9 (closed top-left loop).
$three = New-InkImage @(
    @(1, 1, 1),
    @(0, 0, 1),
    @(1, 1, 1),
    @(0, 0, 1),
    @(1, 1, 1)
)
$nine = New-InkImage @(
    @(1, 1, 1),
    @(1, 0, 1),
    @(1, 1, 1),
    @(0, 0, 1),
    @(1, 1, 1)
)
$threeNoisy = New-InkImage @(
    @(1, 1, 1),
    @(0, 0, 1),
    @(1, 1, 1),
    @(0, 0, 1),
    @(1, 1, 0)
)
$cmp = Compare-DigitCandidate -CandInk $threeNoisy.Ink -CandW $threeNoisy.W -CandH $threeNoisy.H `
    -TemplateAInk $three.Ink -TemplateAW $three.W -TemplateAH $three.H `
    -TemplateBInk $nine.Ink -TemplateBW $nine.W -TemplateBH $nine.H -GridW 5 -GridH 6
Assert-Equal 'a' $cmp.Pick 'noisy 3 classifies as the 3 template (a)'
Assert-True ($cmp.Margin -gt 0) 'a positive decision margin'

# Verdict: OCR read '3', image agrees -> ok (using a low MinMargin for the
# tiny synthetic glyphs).
Assert-Equal 'ok' (Get-DigitPixelVerdict -Compare $cmp -Expected '3' -MinMargin 0.001) 'image matches OCR digit 3 -> ok'
# Same image but OCR claimed '9' -> the image says 3 -> ng.
Assert-Equal 'ng' (Get-DigitPixelVerdict -Compare $cmp -Expected '9' -MinMargin 0.001) 'image contradicts OCR digit 9 -> ng'
# Below the confidence margin -> unknown ('').
Assert-Equal '' (Get-DigitPixelVerdict -Compare $cmp -Expected '3' -MinMargin 0.999) 'sub-threshold margin -> unknown'
Assert-Equal '' (Get-DigitPixelVerdict -Compare $null -Expected '3') 'null compare -> unknown'

# --- Merge-DigitPixelVerdicts : conservative row aggregation --------------
Assert-Equal 'ok' (Merge-DigitPixelVerdicts ([string[]]@('ok','ok','ok'))) 'all ok -> ok'
Assert-Equal 'ng' (Merge-DigitPixelVerdicts ([string[]]@('ok','ng','ok'))) 'any ng -> ng'
Assert-Equal 'ng' (Merge-DigitPixelVerdicts ([string[]]@('','ng'))) 'ng wins over unknown'
Assert-Equal '' (Merge-DigitPixelVerdicts ([string[]]@('ok','','ok'))) 'any unknown (no ng) -> unknown'
Assert-Equal '' (Merge-DigitPixelVerdicts ([string[]]@())) 'no digits checked -> unknown'

exit (Complete-Tests)
