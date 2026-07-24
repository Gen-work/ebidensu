# ============================================================
#  OldSnapPixelVerify.ps1
#
#  NON-PURE GDI+/System.Drawing glue for D2 per-digit 3/9 discrimination
#  (ProcessTime old-snap 9->3 hand-verification, docs/ProcessTime-OldSnap-
#  Verify-Plan.md section 4.2). NO Excel COM, NO param() block (safe to
#  dot-source), ASCII source.
#
#  STATIC-CHECKED ONLY: there is no System.Drawing, no MS Gothic, and no real
#  HM snap in the Linux/CI dev environment, so every function here is
#  confirmed by parse-check + the PURE core's unit tests, and must be
#  validated on an office PC before D2 is trusted. The Phase-0 GO decision
#  (2026-07-24, mock-page/pixeldiff.mjs) proved the *comparison method*; the
#  crop GEOMETRY per snap-window size is the remaining office-PC calibration
#  (plan section 9), so this ships DISABLED by default
#  (ProcessTime.OldSnapVerify.PixelDiff.Enabled = $false).
#
#  Split of responsibilities:
#    * PixelDigitMatch.ps1  -- PURE, unit-tested scoring (grid NCC classify).
#    * THIS file            -- render MS Gothic 3/9 templates, crop a digit
#                              box out of the snap bitmap, and drive the pure
#                              scorer to a per-digit / per-row verdict.
#
#  Every entry point SWALLOWS errors and returns '' (unknown) on any failure
#  -- an image check can never block the write or mislead the verdict; ''
#  makes Get-OldSnapVerifyVerdict fall back to the conservative flag.
# ============================================================

# ---------------------------------------------------------------------------
# New-DigitTemplateGray
#   Renders a single digit char in -FontName at -FontPx onto a white -Width x
#   -Height bitmap (black text, centered) and returns @{ Gray; W; H } where
#   Gray is a row-major 0..255 luminance array (the shape PixelDigitMatch's
#   ConvertTo-DigitInk expects). GDI+ rasterises differently from the Edge
#   snap (DirectWrite) -- the pure metric is AA-tolerant by design (Phase 0),
#   but if the office-PC margins are marginal, render the reference glyphs via
#   a headless Edge/WebBrowser instead (plan section 4.2 note).
# ---------------------------------------------------------------------------
function New-DigitTemplateGray {
    param([string]$Char, [int]$Width, [int]$Height, [string]$FontName = 'MS Gothic', [double]$FontPx = 13)
    $bmp = $null; $g = $null; $font = $null
    try {
        $bmp = New-Object System.Drawing.Bitmap ($Width, $Height, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
        $font = New-Object System.Drawing.Font ($FontName, [single]$FontPx, [System.Drawing.GraphicsUnit]::Pixel)
        $fmt = New-Object System.Drawing.StringFormat
        $fmt.Alignment = [System.Drawing.StringAlignment]::Center
        $fmt.LineAlignment = [System.Drawing.StringAlignment]::Center
        $rect = New-Object System.Drawing.RectangleF (0, 0, $Width, $Height)
        $g.DrawString([string]$Char, $font, [System.Drawing.Brushes]::Black, $rect, $fmt)
        return (Get-BitmapGrayRegion -Bitmap $bmp -X 0 -Y 0 -W $Width -H $Height)
    } catch {
        return $null
    } finally {
        if ($null -ne $font) { try { $font.Dispose() } catch {} }
        if ($null -ne $g)    { try { $g.Dispose() } catch {} }
        if ($null -ne $bmp)  { try { $bmp.Dispose() } catch {} }
    }
}

# ---------------------------------------------------------------------------
# Get-BitmapGrayRegion
#   Crops a -W x -H rectangle at (-X,-Y) out of -Bitmap and returns
#   @{ Gray; W; H } (row-major 0..255 luminance, Rec.601). Uses GetPixel: the
#   digit crops are tiny (tens of pixels), so LockBits' complexity is not
#   worth it here. Clamps to the bitmap bounds; out-of-range pixels read as
#   white (255) so a slightly oversized rect never throws.
# ---------------------------------------------------------------------------
function Get-BitmapGrayRegion {
    param($Bitmap, [int]$X, [int]$Y, [int]$W, [int]$H)
    $gray = New-Object 'double[]' ($W * $H)
    $bw = [int]$Bitmap.Width; $bh = [int]$Bitmap.Height
    for ($j = 0; $j -lt $H; $j++) {
        $py = $Y + $j
        for ($i = 0; $i -lt $W; $i++) {
            $px = $X + $i
            if ($px -ge 0 -and $py -ge 0 -and $px -lt $bw -and $py -lt $bh) {
                $c = $Bitmap.GetPixel($px, $py)
                $lum = 0.299 * $c.R + 0.587 * $c.G + 0.114 * $c.B
            } else {
                $lum = 255.0
            }
            $gray[$j * $W + $i] = $lum
        }
    }
    return @{ Gray = $gray; W = $W; H = $H }
}

# ---------------------------------------------------------------------------
# Get-OldSnapDigitVerdict
#   One digit's 3/9 verdict from an already-open snap bitmap and a calibrated
#   pixel rectangle for that digit. Renders 3 and 9 templates at the crop
#   size, classifies via PixelDigitMatch, and returns 'ok' / 'ng' / '' (the
#   pure Get-DigitPixelVerdict semantics). -Expected is the OCR'd digit
#   ('3' or '9'). Returns '' on any failure.
# ---------------------------------------------------------------------------
function Get-OldSnapDigitVerdict {
    param($Bitmap, $Rect, [string]$Expected, [string]$FontName = 'MS Gothic', [double]$MinMargin = 0.04)
    try {
        $x = [int]$Rect.X; $y = [int]$Rect.Y; $w = [int]$Rect.W; $h = [int]$Rect.H
        if ($w -le 0 -or $h -le 0) { return '' }
        $crop = Get-BitmapGrayRegion -Bitmap $Bitmap -X $x -Y $y -W $w -H $h
        $fontPx = [double]($h * 0.8)
        $tpl3 = New-DigitTemplateGray -Char '3' -Width $w -Height $h -FontName $FontName -FontPx $fontPx
        $tpl9 = New-DigitTemplateGray -Char '9' -Width $w -Height $h -FontName $FontName -FontPx $fontPx
        if ($null -eq $tpl3 -or $null -eq $tpl9) { return '' }
        $candInk = ConvertTo-DigitInk $crop.Gray
        $ink3    = ConvertTo-DigitInk $tpl3.Gray
        $ink9    = ConvertTo-DigitInk $tpl9.Gray
        $cmp = Compare-DigitCandidate -CandInk $candInk -CandW $crop.W -CandH $crop.H `
            -TemplateAInk $ink3 -TemplateAW $tpl3.W -TemplateAH $tpl3.H `
            -TemplateBInk $ink9 -TemplateBW $tpl9.W -TemplateBH $tpl9.H
        return (Get-DigitPixelVerdict -Compare $cmp -Expected $Expected -MinMargin $MinMargin)
    } catch {
        return ''
    }
}

# ---------------------------------------------------------------------------
# Resolve-OldSnapTimeDigitRects
#   Computes the per-digit pixel rectangles for the '3'/'9' characters of one
#   time field, from the OCR'd text and a CALIBRATED field cell geometry.
#   MS Gothic on the HM page is effectively monospaced, so digit x-positions
#   are cell-left + index * digit-width (plan section 4.2). Returns a list of
#   @{ Rect = @{X;Y;W;H}; Expected = '3'|'9' } -- only the 3/9 positions.
#
#   -Geometry is the office-PC-calibrated cell rect + digit metrics for the
#   field, e.g. @{ CellX; CellY; CellH; DigitW; DigitPad }. Until that
#   calibration exists (plan section 9) -Geometry is $null/empty and this
#   returns an EMPTY list, so the row verdict is '' (conservative). This is
#   the ONE piece D2 cannot compute without office-PC snap samples.
# ---------------------------------------------------------------------------
function Resolve-OldSnapTimeDigitRects {
    param([string]$Text, $Geometry)
    $rects = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Geometry) { return $rects }
    foreach ($k in @('CellX', 'CellY', 'CellH', 'DigitW')) {
        if (-not ($Geometry -is [hashtable]) -or -not $Geometry.ContainsKey($k)) { return $rects }
    }
    $cellX  = [double]$Geometry.CellX
    $cellY  = [double]$Geometry.CellY
    $cellH  = [double]$Geometry.CellH
    $digitW = [double]$Geometry.DigitW
    $pad    = if ($Geometry.ContainsKey('DigitPad')) { [double]$Geometry.DigitPad } else { 0.0 }
    if ($digitW -le 0 -or $cellH -le 0) { return $rects }
    $t = [string]$Text
    for ($i = 0; $i -lt $t.Length; $i++) {
        $ch = $t[$i]
        if ($ch -ne '3' -and $ch -ne '9') { continue }
        $x = [int][Math]::Round($cellX + $i * $digitW + $pad)
        $y = [int][Math]::Round($cellY)
        $w = [int][Math]::Round($digitW)
        $h = [int][Math]::Round($cellH)
        $rects.Add(@{ Rect = @{ X = $x; Y = $y; W = $w; H = $h }; Expected = [string]$ch })
    }
    return $rects
}

# ---------------------------------------------------------------------------
# Get-OldSnapRowPixelVerdict
#   Row-level D2 verdict: opens the snap PNG once and checks every 3/9 digit
#   across the row's time fields, merging the per-digit verdicts (any 'ng' ->
#   'ng'; any unknown -> ''; else 'ok'). -Fields is a list of
#   @{ Text; Geometry } (one per time field: start/end/duration, each with
#   its calibrated cell geometry). Returns '' (unknown -> conservative flag)
#   whenever the snap is missing, geometry is un-calibrated, or anything
#   throws. Static-checked only; confirm on an office PC once the geometry is
#   calibrated (plan section 9).
# ---------------------------------------------------------------------------
function Get-OldSnapRowPixelVerdict {
    param([string]$SnapPath, [object[]]$Fields, [string]$FontName = 'MS Gothic', [double]$MinMargin = 0.04)
    if ([string]::IsNullOrWhiteSpace($SnapPath) -or -not (Test-Path -LiteralPath $SnapPath)) { return '' }
    $bmp = $null
    try {
        $bytes = [System.IO.File]::ReadAllBytes($SnapPath)
        $ms = New-Object System.IO.MemoryStream (, $bytes)
        $bmp = [System.Drawing.Bitmap]::FromStream($ms)
        $verdicts = New-Object System.Collections.Generic.List[string]
        foreach ($f in @($Fields)) {
            $rects = Resolve-OldSnapTimeDigitRects -Text ([string]$f.Text) -Geometry $f.Geometry
            foreach ($dr in $rects) {
                $verdicts.Add((Get-OldSnapDigitVerdict -Bitmap $bmp -Rect $dr.Rect -Expected $dr.Expected -FontName $FontName -MinMargin $MinMargin))
            }
        }
        return (Merge-DigitPixelVerdicts ([string[]]$verdicts.ToArray()))
    } catch {
        return ''
    } finally {
        if ($null -ne $bmp) { try { $bmp.Dispose() } catch {} }
    }
}
