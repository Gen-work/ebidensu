# ============================================================
#  PixelDigitMatch.ps1
#
#  PURE library for D2 per-digit 3/9 discrimination (ProcessTime old-snap
#  9->3 hand-verification, docs/ProcessTime-OldSnap-Verify-Plan.md section
#  4.2). NO Excel COM, NO GDI/System.Drawing, NO file I/O. Dot-source only
#  (no param() block); ASCII source.
#
#  This is the PowerShell port of the Phase-0-proven metric prototyped in
#  mock-page/pixeldiff.mjs (GO decision recorded 2026-07-24): grayscale ink
#  -> binarize -> trim to the ink bounding box -> average-pool to a fixed
#  normalized grid -> normalized cross-correlation (NCC) between grids. The
#  metric compares SHAPE (mean-subtracted, norm-normalized), so it is
#  invariant to overall ink amount/brightness and tolerant to anti-aliasing
#  and small offsets, which is exactly what separates a 3 from a 9.
#
#  The COM/GDI glue that turns a snap PNG + the OCR'd digit into these
#  inputs (crop the digit, render MS Gothic 3/9 templates) lives in the
#  non-pure OldSnapPixelVerify.ps1 and is static-checked only. THIS file is
#  the reusable, unit-tested core (Tests\Test-PixelDigitMatch.ps1).
#
#  Images are passed as a flat row-major [double[]] "ink" array (0 = white
#  paper, 1 = full black ink) plus Width/Height. ConvertTo-DigitInk builds
#  that from a 0..255 grayscale array (as a GDI Bitmap yields).
#
#  Convention: functions return plain values -- never return ,@(...).
# ============================================================

# ---------------------------------------------------------------------------
# ConvertTo-DigitInk
#   0..255 grayscale (row-major, 0=black) -> ink [double[]] (1=black ink).
# ---------------------------------------------------------------------------
function ConvertTo-DigitInk {
    param([double[]]$Gray)
    $n = $Gray.Length
    $ink = New-Object 'double[]' $n
    for ($i = 0; $i -lt $n; $i++) { $ink[$i] = (255.0 - $Gray[$i]) / 255.0 }
    return $ink
}

# ---------------------------------------------------------------------------
# Get-DigitInkBBox
#   Tight bounding box of pixels whose ink exceeds -Threshold (default 0.5).
#   Returns @{ X0; Y0; X1; Y1 } (inclusive-min / exclusive-max) or $null when
#   the image is blank.
# ---------------------------------------------------------------------------
function Get-DigitInkBBox {
    param([double[]]$Ink, [int]$Width, [int]$Height, [double]$Threshold = 0.5)
    $x0 = $Width; $y0 = $Height; $x1 = 0; $y1 = 0; $any = $false
    for ($y = 0; $y -lt $Height; $y++) {
        $rowbase = $y * $Width
        for ($x = 0; $x -lt $Width; $x++) {
            if ($Ink[$rowbase + $x] -gt $Threshold) {
                $any = $true
                if ($x -lt $x0) { $x0 = $x }
                if ($y -lt $y0) { $y0 = $y }
                if (($x + 1) -gt $x1) { $x1 = $x + 1 }
                if (($y + 1) -gt $y1) { $y1 = $y + 1 }
            }
        }
    }
    if (-not $any) { return $null }
    return @{ X0 = $x0; Y0 = $y0; X1 = $x1; Y1 = $y1 }
}

# ---------------------------------------------------------------------------
# ConvertTo-DigitNormalizedGrid
#   Average-pool the ink inside -Box into a GridW x GridH normalized grid
#   (row-major [double[]]). Average pooling (not nearest-neighbor) keeps
#   anti-aliased edge coverage -> AA tolerance. A blank image (Box $null)
#   yields an all-zero grid.
# ---------------------------------------------------------------------------
function ConvertTo-DigitNormalizedGrid {
    param([double[]]$Ink, [int]$Width, [int]$Height, $Box, [int]$GridW = 16, [int]$GridH = 24)
    $out = New-Object 'double[]' ($GridW * $GridH)
    if ($null -eq $Box) { return $out }
    $bw = $Box.X1 - $Box.X0
    $bh = $Box.Y1 - $Box.Y0
    if ($bw -le 0 -or $bh -le 0) { return $out }
    for ($cy = 0; $cy -lt $GridH; $cy++) {
        $sy0 = $Box.Y0 + ($cy / [double]$GridH) * $bh
        $sy1 = $Box.Y0 + (($cy + 1) / [double]$GridH) * $bh
        $iy0 = [int][Math]::Floor($sy0)
        $iy1 = [Math]::Max($iy0 + 1, [int][Math]::Ceiling($sy1))
        for ($cx = 0; $cx -lt $GridW; $cx++) {
            $sx0 = $Box.X0 + ($cx / [double]$GridW) * $bw
            $sx1 = $Box.X0 + (($cx + 1) / [double]$GridW) * $bw
            $ix0 = [int][Math]::Floor($sx0)
            $ix1 = [Math]::Max($ix0 + 1, [int][Math]::Ceiling($sx1))
            $sum = 0.0; $cnt = 0
            for ($y = $iy0; $y -lt $iy1 -and $y -lt $Height; $y++) {
                $rowbase = $y * $Width
                for ($x = $ix0; $x -lt $ix1 -and $x -lt $Width; $x++) {
                    $sum += $Ink[$rowbase + $x]; $cnt++
                }
            }
            $out[$cy * $GridW + $cx] = if ($cnt -gt 0) { $sum / $cnt } else { 0.0 }
        }
    }
    return $out
}

# ---------------------------------------------------------------------------
# Get-DigitNcc
#   Normalized cross-correlation of two equal-length vectors, in [-1, 1].
#   Two flat (zero-variance) vectors -> 1 (identical); one flat vs a varying
#   one -> 0 (unrelated).
# ---------------------------------------------------------------------------
function Get-DigitNcc {
    param([double[]]$A, [double[]]$B)
    $n = $A.Length
    if ($n -eq 0 -or $B.Length -ne $n) { return 0.0 }
    $ma = 0.0; $mb = 0.0
    for ($i = 0; $i -lt $n; $i++) { $ma += $A[$i]; $mb += $B[$i] }
    $ma /= $n; $mb /= $n
    $num = 0.0; $da = 0.0; $db = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $va = $A[$i] - $ma; $vb = $B[$i] - $mb
        $num += $va * $vb; $da += $va * $va; $db += $vb * $vb
    }
    if ($da -eq 0 -and $db -eq 0) { return 1.0 }
    if ($da -eq 0 -or $db -eq 0) { return 0.0 }
    return $num / [Math]::Sqrt($da * $db)
}

# ---------------------------------------------------------------------------
# Get-DigitSimilarity
#   Shape similarity of two ink images in [-1, 1]: normalize both to the same
#   grid, then NCC. Higher = more alike.
# ---------------------------------------------------------------------------
function Get-DigitSimilarity {
    param(
        [double[]]$InkA, [int]$WA, [int]$HA,
        [double[]]$InkB, [int]$WB, [int]$HB,
        [int]$GridW = 16, [int]$GridH = 24, [double]$Threshold = 0.5
    )
    $ga = ConvertTo-DigitNormalizedGrid -Ink $InkA -Width $WA -Height $HA -Box (Get-DigitInkBBox -Ink $InkA -Width $WA -Height $HA -Threshold $Threshold) -GridW $GridW -GridH $GridH
    $gb = ConvertTo-DigitNormalizedGrid -Ink $InkB -Width $WB -Height $HB -Box (Get-DigitInkBBox -Ink $InkB -Width $WB -Height $HB -Threshold $Threshold) -GridW $GridW -GridH $GridH
    return (Get-DigitNcc -A $ga -B $gb)
}

# ---------------------------------------------------------------------------
# Compare-DigitCandidate
#   Classify a candidate digit crop between two templates (A and B). Returns
#   @{ Pick = 'a'|'b'; Sa; Sb; Margin } where Margin = |Sa - Sb| (a small
#   margin means the templates are nearly indistinguishable for this
#   candidate -> the caller must refuse to auto-confirm).
# ---------------------------------------------------------------------------
function Compare-DigitCandidate {
    param(
        [double[]]$CandInk, [int]$CandW, [int]$CandH,
        [double[]]$TemplateAInk, [int]$TemplateAW, [int]$TemplateAH,
        [double[]]$TemplateBInk, [int]$TemplateBW, [int]$TemplateBH,
        [int]$GridW = 16, [int]$GridH = 24, [double]$Threshold = 0.5
    )
    $sa = Get-DigitSimilarity -InkA $CandInk -WA $CandW -HA $CandH -InkB $TemplateAInk -WB $TemplateAW -HB $TemplateAH -GridW $GridW -GridH $GridH -Threshold $Threshold
    $sb = Get-DigitSimilarity -InkA $CandInk -WA $CandW -HA $CandH -InkB $TemplateBInk -WB $TemplateBW -HB $TemplateBH -GridW $GridW -GridH $GridH -Threshold $Threshold
    $pick = if ($sa -ge $sb) { 'a' } else { 'b' }
    return @{ Pick = $pick; Sa = $sa; Sb = $sb; Margin = [Math]::Abs($sa - $sb) }
}

# ---------------------------------------------------------------------------
# Get-DigitPixelVerdict
#   Turns one Compare-DigitCandidate result into a per-digit verdict for the
#   3/9 check. -Expected is the OCR'd digit ('3' or '9'); template 'a' is the
#   3 template and 'b' is the 9 template (the caller renders them in that
#   order). -MinMargin (from the Phase-0/office-PC threshold) is the smallest
#   |Sa-Sb| that counts as a confident decision.
#     'ok'  the image matches the OCR'd digit, confidently.
#     'ng'  the image matches the OTHER digit better, confidently (misread).
#     ''    ambiguous (margin below MinMargin) -> conservative: caller flags.
# ---------------------------------------------------------------------------
function Get-DigitPixelVerdict {
    param([hashtable]$Compare, [string]$Expected, [double]$MinMargin = 0.04)
    if ($null -eq $Compare) { return '' }
    if ([double]$Compare.Margin -lt $MinMargin) { return '' }   # not confident
    $imgDigit = if ($Compare.Pick -eq 'a') { '3' } else { '9' }
    if ($imgDigit -eq ([string]$Expected)) { return 'ok' }
    return 'ng'
}

# ---------------------------------------------------------------------------
# Merge-DigitPixelVerdicts
#   Row-level aggregation of the per-digit verdicts for all the 3/9 digits in
#   one output row's time fields. Conservative: any 'ng' -> 'ng'; else any ''
#   (ambiguous) -> '' (the row cannot be auto-confirmed on the image); else
#   'ok'. An empty input (no 3/9 digits to check) -> '' (nothing decided by
#   the image; the deterministic checks stand alone).
# ---------------------------------------------------------------------------
function Merge-DigitPixelVerdicts {
    param([string[]]$Verdicts)
    if ($null -eq $Verdicts -or $Verdicts.Count -eq 0) { return '' }
    $sawUnknown = $false
    foreach ($v in $Verdicts) {
        if ($v -eq 'ng') { return 'ng' }
        if ($v -ne 'ok') { $sawUnknown = $true }
    }
    if ($sawUnknown) { return '' }
    return 'ok'
}
