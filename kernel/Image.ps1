#Requires -Version 5.1
# ============================================================
#  kernel/Image.ps1
#
#  The one place GDI+ (System.Drawing) is touched for PNG work: screen
#  region grab, per-side crop, size read. Dot-source only (no param()
#  block, ASCII source, no class). Steps under modules/screen dot-source
#  it with  . (Join-Path $PSScriptRoot '..\..\kernel\Image.ps1')  and the
#  legacy snap scripts (HmSnap / MqSnap / JenkinsSnap / Crop-Snap) call
#  Invoke-EbiCropPng in place of the four Invoke-CropPng copies they
#  used to carry (P1-20).
#
#  Nothing here loads System.Drawing until Get-EbiDrawing is called, so a
#  DryRun path (and the Linux CI that dry-runs every step) never reaches
#  GDI+. The pure geometry helpers are unit-tested without it.
#
#  Every impure entry point returns a record @{ ok; failure; message; ... }
#  and never throws (STEP-CONTRACT 3.1): the caller turns it into a step
#  failure, or throws itself where the old code expected a throw.
# ============================================================

function Get-EbiDrawing {
    Add-Type -AssemblyName System.Drawing
    return $true
}

function Get-EbiCropGeometry {
    <#
      PURE. Image size + per-side amounts -> @{ ok; width; height; message;
      noop }. Negative amounts read as 0 (nothing to take off that side);
      ok=$false when the crop would leave no pixels.
    #>
    param([int]$Width, [int]$Height, [int]$Left = 0, [int]$Top = 0, [int]$Right = 0, [int]$Bottom = 0)
    if ($Left -lt 0) { $Left = 0 }; if ($Top -lt 0) { $Top = 0 }; if ($Right -lt 0) { $Right = 0 }; if ($Bottom -lt 0) { $Bottom = 0 }
    $w = $Width - $Left - $Right
    $h = $Height - $Top - $Bottom
    if ($Width -le 0 -or $Height -le 0) { return @{ ok = $false; width = 0; height = 0; noop = $false; message = ('image is empty ({0}x{1})' -f $Width, $Height) } }
    if ($w -le 0 -or $h -le 0) {
        return @{ ok = $false; width = 0; height = 0; noop = $false; message = ('image too small ({0}x{1}) to crop L{2}/T{3}/R{4}/B{5} px' -f $Width, $Height, $Left, $Top, $Right, $Bottom) }
    }
    return @{ ok = $true; width = $w; height = $h; noop = ($Left -eq 0 -and $Top -eq 0 -and $Right -eq 0 -and $Bottom -eq 0); message = '' }
}

function Resolve-EbiCropSides {
    # PURE. The legacy "-1 = inherit CropPx" convention -> four ints >= 0.
    param([int]$CropPx = 0, [int]$Left = -1, [int]$Top = -1, [int]$Right = -1, [int]$Bottom = -1)
    $l = if ($Left -ge 0) { $Left } else { $CropPx }
    $t = if ($Top -ge 0) { $Top } else { $CropPx }
    $r = if ($Right -ge 0) { $Right } else { $CropPx }
    $b = if ($Bottom -ge 0) { $Bottom } else { $CropPx }
    if ($l -lt 0) { $l = 0 }; if ($t -lt 0) { $t = 0 }; if ($r -lt 0) { $r = 0 }; if ($b -lt 0) { $b = 0 }
    return @{ left = $l; top = $t; right = $r; bottom = $b }
}

function Resolve-EbiScreenRegion {
    <#
      PURE. Clamp a requested rectangle into bounds -> @{ x; y; w; h;
      clamped; edges } (edges = 'x,y,width,height' subset, '' when nothing
      moved). ScreenRegion.ps1's Resolve-ScreenRegion as a hashtable.
    #>
    param([int]$X, [int]$Y, [int]$W, [int]$H, [int]$BoundsX, [int]$BoundsY, [int]$BoundsW, [int]$BoundsH)
    $right = $BoundsX + $BoundsW; $bottom = $BoundsY + $BoundsH
    $edges = New-Object System.Collections.ArrayList
    if ($X -lt $BoundsX) { $W = $W - ($BoundsX - $X); $X = $BoundsX; [void]$edges.Add('x') }
    if ($Y -lt $BoundsY) { $H = $H - ($BoundsY - $Y); $Y = $BoundsY; [void]$edges.Add('y') }
    if (($X + $W) -gt $right) { $W = $right - $X; [void]$edges.Add('width') }
    if (($Y + $H) -gt $bottom) { $H = $bottom - $Y; [void]$edges.Add('height') }
    if ($W -lt 0) { $W = 0 }
    if ($H -lt 0) { $H = 0 }
    return @{ x = $X; y = $Y; w = $W; h = $H; clamped = ($edges.Count -gt 0); edges = ($edges.ToArray() -join ',') }
}

function New-EbiParentDirectory {
    param([string]$Path)
    $dir = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
}

function Save-EbiScreenRegionPng {
    # Grab a screen rectangle to a PNG. @{ ok; failure; message }.
    param([int]$X, [int]$Y, [int]$W, [int]$H, [string]$Dest)
    if ($W -le 0 -or $H -le 0) { return @{ ok = $false; failure = 'region_empty'; message = ('region is empty ({0}x{1})' -f $W, $H) } }
    [void](Get-EbiDrawing)
    $bmp = $null; $gfx = $null
    try {
        New-EbiParentDirectory -Path $Dest
        $bmp = New-Object System.Drawing.Bitmap($W, $H)
        $gfx = [System.Drawing.Graphics]::FromImage($bmp)
        $gfx.CopyFromScreen($X, $Y, 0, 0, (New-Object System.Drawing.Size($W, $H)))
        $bmp.Save($Dest, [System.Drawing.Imaging.ImageFormat]::Png)
    } catch {
        return @{ ok = $false; failure = 'save_failed'; message = $_.Exception.Message }
    } finally {
        if ($null -ne $gfx) { $gfx.Dispose() }
        if ($null -ne $bmp) { $bmp.Dispose() }
    }
    return @{ ok = $true; failure = ''; message = '' }
}

function Get-EbiPngSize {
    # @{ ok; width; height; message }. Reads through a byte copy so the
    # file is never locked by GDI+ (the old Invoke-CropPng did the same).
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{ ok = $false; width = 0; height = 0; message = ('file not found: ' + $Path) } }
    return (Get-EbiPngSizeCore -Path $Path)
}

function Get-EbiPngSizeCore {
    # The GDI+ half of Get-EbiPngSize. Names System.Drawing types, so it
    # is only ever CALLED after the pure checks passed: on Linux pwsh the
    # first call of any function that mentions System.Drawing throws
    # (PlatformNotSupported) before a single statement runs.
    param([string]$Path)
    [void](Get-EbiDrawing)
    $ms = $null; $img = $null
    try {
        $ms = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($Path))
        $img = [System.Drawing.Image]::FromStream($ms)
        return @{ ok = $true; width = [int]$img.Width; height = [int]$img.Height; message = '' }
    } catch {
        return @{ ok = $false; width = 0; height = 0; message = $_.Exception.Message }
    } finally {
        if ($null -ne $img) { $img.Dispose() }
        if ($null -ne $ms) { $ms.Dispose() }
    }
}

function Invoke-EbiCropPng {
    <#
      Crop a PNG by per-side pixel amounts, in place (Out empty) or to Out.
      -CropPx + -1 sides is the legacy uniform convention (Resolve-EbiCrop
      Sides); a step passes the four sides directly. The write is atomic:
      a temp file next to the destination, then a move.
      -> @{ ok; failure; message; width; height; path }
         failures: file_not_found | crop_exceeds_image | image_read_error
    #>
    param([string]$Path, [string]$Out = '', [int]$CropPx = 0, [int]$Left = -1, [int]$Top = -1, [int]$Right = -1, [int]$Bottom = -1)
    $sides = Resolve-EbiCropSides -CropPx $CropPx -Left $Left -Top $Top -Right $Right -Bottom $Bottom
    $dest = if ([string]::IsNullOrWhiteSpace($Out)) { $Path } else { $Out }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = ('file not found: ' + $Path); width = 0; height = 0; path = $dest } }
    return (Invoke-EbiCropPngCore -Path $Path -Dest $dest -Sides $sides)
}

function Invoke-EbiCropPngCore {
    # The GDI+ half of Invoke-EbiCropPng (see Get-EbiPngSizeCore for why
    # it is a separate function).
    param([string]$Path, [string]$Dest, [hashtable]$Sides)
    $sides = $Sides; $dest = $Dest
    [void](Get-EbiDrawing)
    $ms = $null; $orig = $null; $bmp = $null; $gfx = $null
    $tmp = $dest + '.crop.tmp'
    try {
        $ms = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($Path))
        $orig = [System.Drawing.Image]::FromStream($ms)
    } catch {
        if ($null -ne $ms) { $ms.Dispose() }
        return @{ ok = $false; failure = 'image_read_error'; message = $_.Exception.Message; width = 0; height = 0; path = $dest }
    }
    try {
        $geo = Get-EbiCropGeometry -Width $orig.Width -Height $orig.Height -Left $sides['left'] -Top $sides['top'] -Right $sides['right'] -Bottom $sides['bottom']
        if (-not $geo['ok']) { return @{ ok = $false; failure = 'crop_exceeds_image'; message = $geo['message']; width = [int]$orig.Width; height = [int]$orig.Height; path = $dest } }
        if ($geo['noop']) {
            if ($dest -ne $Path) { New-EbiParentDirectory -Path $dest; Copy-Item -LiteralPath $Path -Destination $dest -Force }
            return @{ ok = $true; failure = ''; message = 'nothing to crop'; width = $geo['width']; height = $geo['height']; path = $dest }
        }
        New-EbiParentDirectory -Path $dest
        $bmp = New-Object System.Drawing.Bitmap([int]$geo['width'], [int]$geo['height'])
        $gfx = [System.Drawing.Graphics]::FromImage($bmp)
        $src = New-Object System.Drawing.Rectangle([int]$sides['left'], [int]$sides['top'], [int]$geo['width'], [int]$geo['height'])
        $dst = New-Object System.Drawing.Rectangle(0, 0, [int]$geo['width'], [int]$geo['height'])
        $gfx.DrawImage($orig, $dst, $src, [System.Drawing.GraphicsUnit]::Pixel)
        $bmp.Save($tmp, [System.Drawing.Imaging.ImageFormat]::Png)
    } catch {
        return @{ ok = $false; failure = 'image_read_error'; message = $_.Exception.Message; width = 0; height = 0; path = $dest }
    } finally {
        if ($null -ne $gfx) { $gfx.Dispose() }
        if ($null -ne $bmp) { $bmp.Dispose() }
        if ($null -ne $orig) { $orig.Dispose() }
        if ($null -ne $ms) { $ms.Dispose() }
    }
    try {
        Move-Item -LiteralPath $tmp -Destination $dest -Force
    } catch {
        return @{ ok = $false; failure = 'image_read_error'; message = ('cannot replace ' + $dest + ': ' + $_.Exception.Message); width = $geo['width']; height = $geo['height']; path = $dest }
    }
    return @{ ok = $true; failure = ''; message = ''; width = $geo['width']; height = $geo['height']; path = $dest }
}

function Invoke-EbiInvertPng {
    <#
      A negative, grey copy of a PNG at Out (10/21: the HOST emulator shots
      are light text on black; Windows OCR reads dark text on a light
      ground far better). -> @{ ok; failure; message; path }
         failures: file_not_found | image_read_error
    #>
    param([string]$Path, [string]$Out)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = ('file not found: ' + $Path); path = $Out } }
    return (Invoke-EbiInvertPngCore -Path $Path -Out $Out)
}

function Invoke-EbiInvertPngCore {
    # The GDI+ half of Invoke-EbiInvertPng (see Get-EbiPngSizeCore).
    param([string]$Path, [string]$Out)
    [void](Get-EbiDrawing)
    $ms = $null; $img = $null; $bmp = $null; $gfx = $null; $ia = $null
    try {
        $ms = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($Path))
        $img = [System.Drawing.Image]::FromStream($ms)
        $bmp = New-Object System.Drawing.Bitmap([int]$img.Width, [int]$img.Height)
        $gfx = [System.Drawing.Graphics]::FromImage($bmp)
        # grey = luminance, then 1 - grey
        $g = -0.333
        $cm = New-Object System.Drawing.Imaging.ColorMatrix(, [single[][]]@(
            [single[]]@($g, $g, $g, 0, 0),
            [single[]]@($g, $g, $g, 0, 0),
            [single[]]@($g, $g, $g, 0, 0),
            [single[]]@(0, 0, 0, 1, 0),
            [single[]]@(1, 1, 1, 0, 1)))
        $ia = New-Object System.Drawing.Imaging.ImageAttributes
        $ia.SetColorMatrix($cm)
        $r = New-Object System.Drawing.Rectangle(0, 0, [int]$img.Width, [int]$img.Height)
        $gfx.DrawImage($img, $r, 0, 0, [int]$img.Width, [int]$img.Height, [System.Drawing.GraphicsUnit]::Pixel, $ia)
        New-EbiParentDirectory -Path $Out
        $bmp.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
        return @{ ok = $true; failure = ''; message = ''; path = $Out }
    } catch {
        return @{ ok = $false; failure = 'image_read_error'; message = $_.Exception.Message; path = $Out }
    } finally {
        if ($null -ne $ia) { $ia.Dispose() }
        if ($null -ne $gfx) { $gfx.Dispose() }
        if ($null -ne $bmp) { $bmp.Dispose() }
        if ($null -ne $img) { $img.Dispose() }
        if ($null -ne $ms) { $ms.Dispose() }
    }
}

function Get-EbiStripGeometry {
    <#
      PURE. Two pictures of the same rows, B scaled to A's height: the strip
      FromFrac..ToFrac of A's width, in px, and whether B (scaled) is wide
      enough to hold it. -> @{ ok; x0; x1; scaledWidthB; message }
    #>
    param([int]$WidthA, [int]$HeightA, [int]$WidthB, [int]$HeightB, [double]$FromFrac, [double]$ToFrac)
    if ($WidthA -le 0 -or $HeightA -le 0 -or $WidthB -le 0 -or $HeightB -le 0) { return @{ ok = $false; x0 = 0; x1 = 0; scaledWidthB = 0; message = 'an empty picture' } }
    if ($FromFrac -lt 0 -or $ToFrac -gt 1.0001 -or $ToFrac -le $FromFrac) { return @{ ok = $false; x0 = 0; x1 = 0; scaledWidthB = 0; message = ('no strip {0:0.000}..{1:0.000}' -f $FromFrac, $ToFrac) } }
    $sw = [int][Math]::Round($WidthB * ([double]$HeightA / $HeightB))
    $x0 = [int][Math]::Floor($FromFrac * $WidthA); $x1 = [int][Math]::Ceiling([Math]::Min(1.0, $ToFrac) * $WidthA)
    if ($x1 - $x0 -lt 4) { return @{ ok = $false; x0 = $x0; x1 = $x1; scaledWidthB = $sw; message = 'the strip is narrower than 4 px' } }
    if ($x1 -gt $sw) { return @{ ok = $false; x0 = $x0; x1 = $x1; scaledWidthB = $sw; message = ('the other picture is too narrow for the strip ({0} px, the strip ends at {1})' -f $sw, $x1) } }
    return @{ ok = $true; x0 = $x0; x1 = $x1; scaledWidthB = $sw; message = '' }
}

function Get-EbiOcrFrame {
    <#
      PURE. The part of a W x H picture handed to OCR: the LEFT MaxWidth px
      (the line numbers sit at the left; a shot of four emulator windows is
      far wider than OCR takes) scaled so neither side passes MaxDim (the
      Windows recognizer refuses bigger images). -> @{ cropW; scale; outW; outH }
    #>
    param([int]$Width, [int]$Height, [int]$MaxWidth = 1400, [int]$MaxDim = 2500)
    $cw = [Math]::Min($Width, $MaxWidth)
    $sc = [Math]::Min(1.0, [Math]::Min([double]$MaxDim / [Math]::Max(1, $cw), [double]$MaxDim / [Math]::Max(1, $Height)))
    return @{ cropW = $cw; scale = $sc; outW = [int][Math]::Max(1, [Math]::Round($cw * $sc)); outH = [int][Math]::Max(1, [Math]::Round($Height * $sc)) }
}

function New-EbiOcrVariants {
    <#
      Three pictures for OCR to try (10/21 FJODWWBV: none of the HOST shots
      gave a line number): <Base>.ocr.png (the left part, sized for the
      recognizer), <Base>.bin.png (the same in black text on white: dark-
      ground shots are thresholded on brightness) and <Base>.neg.png (the
      negative). -> @{ ok; failure; message; paths }
    #>
    param([string]$Path, [string]$Base, [int]$MaxWidth = 1400, [int]$MaxDim = 2500)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = ('file not found: ' + $Path); paths = @() } }
    return (New-EbiOcrVariantsCore -Path $Path -Base $Base -MaxWidth $MaxWidth -MaxDim $MaxDim)
}

function New-EbiOcrVariantsCore {
    # The GDI+ half of New-EbiOcrVariants (see Get-EbiPngSizeCore).
    param([string]$Path, [string]$Base, [int]$MaxWidth, [int]$MaxDim)
    [void](Get-EbiDrawing)
    $ms = $null; $img = $null; $bmp = $null; $gfx = $null
    try {
        $ms = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($Path))
        $img = [System.Drawing.Image]::FromStream($ms)
        $f = Get-EbiOcrFrame -Width $img.Width -Height $img.Height -MaxWidth $MaxWidth -MaxDim $MaxDim
        $bmp = New-Object System.Drawing.Bitmap([int]$f['outW'], [int]$f['outH'], [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $gfx = [System.Drawing.Graphics]::FromImage($bmp)
        $gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $gfx.DrawImage($img, (New-Object System.Drawing.Rectangle(0, 0, [int]$f['outW'], [int]$f['outH'])), 0, 0, [int]$f['cropW'], [int]$img.Height, [System.Drawing.GraphicsUnit]::Pixel)
        $p1 = $Base + '.ocr.png'; New-EbiParentDirectory -Path $p1
        $bmp.Save($p1, [System.Drawing.Imaging.ImageFormat]::Png)
        # threshold: dark ground (mean < 110) -> bright pixels are text
        $w = [int]$f['outW']; $h = [int]$f['outH']
        $data = $bmp.LockBits((New-Object System.Drawing.Rectangle(0, 0, $w, $h)), [System.Drawing.Imaging.ImageLockMode]::ReadWrite, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $st = [int]$data.Stride
        $buf = New-Object byte[] ($st * $h)
        [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $buf, 0, $buf.Length)
        [void](Get-EbiImageMath)
        $gy = [EbiImageMath]::Gray($buf, $st, $w, $h)
        $out = [EbiImageMath]::Threshold($buf, $st, $w, $h, $gy)
        [System.Runtime.InteropServices.Marshal]::Copy($out, 0, $data.Scan0, $out.Length)
        $bmp.UnlockBits($data)
        $p2 = $Base + '.bin.png'
        $bmp.Save($p2, [System.Drawing.Imaging.ImageFormat]::Png)
        $p3 = $Base + '.neg.png'
        $v = Invoke-EbiInvertPng -Path $p1 -Out $p3
        $paths = @($p1, $p2); if ($v['ok']) { $paths += $p3 }
        return @{ ok = $true; failure = ''; message = ''; paths = $paths }
    } catch {
        return @{ ok = $false; failure = 'image_read_error'; message = $_.Exception.Message; paths = @() }
    } finally {
        if ($null -ne $gfx) { $gfx.Dispose() }
        if ($null -ne $bmp) { $bmp.Dispose() }
        if ($null -ne $img) { $img.Dispose() }
        if ($null -ne $ms) { $ms.Dispose() }
    }
}

function Get-EbiImageMath {
    <#
      Compile the grey-image helpers once per process (plain C#, no
      System.Drawing: unit-tested on Linux too). PowerShell loops are far
      too slow for a sliding-window search.
        Gray(buf, stride, w, h)                 24bpp rows -> grey bytes
        DarkShare(g)                            share of pixels < 128
        BestDiff(a, aw, ah, b, bw, bh, cx, cy, rx, ry)
          a placed at every (cx+ox, cy+oy), |ox| <= rx, |oy| <= ry, inside
          b -> double[] { least mean |a - b|, x, y }; { -1, 0, 0 } when it
          never fits.
    #>
    if (-not ('EbiImageMath' -as [type])) {
        Add-Type -TypeDefinition @"
public static class EbiImageMath {
    public static byte[] Gray(byte[] buf, int stride, int w, int h) {
        byte[] g = new byte[w * h];
        for (int y = 0; y < h; y++) { int row = y * stride; for (int x = 0; x < w; x++) { int o = row + x * 3; g[y * w + x] = (byte)((buf[o] + buf[o + 1] + buf[o + 2]) / 3); } }
        return g;
    }
    public static byte[] Threshold(byte[] buf, int stride, int w, int h, byte[] g) {
        long sum = 0; for (int i = 0; i < g.Length; i++) sum += g[i];
        bool darkGround = g.Length > 0 && (sum / g.Length) < 110;
        int cut = 0;
        if (darkGround) { cut = 80; } else { cut = 140; }
        byte[] o = new byte[buf.Length];
        for (int y = 0; y < h; y++) { int row = y * stride; for (int x = 0; x < w; x++) {
            bool text = darkGround ? g[y * w + x] > cut : g[y * w + x] < cut;
            byte v = text ? (byte)0 : (byte)255; int k = row + x * 3; o[k] = v; o[k + 1] = v; o[k + 2] = v; } }
        return o;
    }
    public static byte[] GrayToRgb(byte[] g, int w, int h, int stride) {
        byte[] o = new byte[stride * h];
        for (int y = 0; y < h; y++) for (int x = 0; x < w; x++) { byte v = g[y * w + x]; int k = y * stride + x * 3; o[k] = v; o[k + 1] = v; o[k + 2] = v; }
        return o;
    }
    public static byte[] Crop(byte[] g, int w, int x0, int y0, int cw, int ch) {
        byte[] o = new byte[cw * ch];
        for (int y = 0; y < ch; y++) System.Array.Copy(g, (y0 + y) * w + x0, o, y * cw, cw);
        return o;
    }
    public static double DarkShare(byte[] g) {
        if (g.Length == 0) return 0; int n = 0;
        for (int i = 0; i < g.Length; i++) if (g[i] < 128) n++;
        return (double)n / g.Length;
    }
    public static double[] BestDiff(byte[] a, int aw, int ah, byte[] b, int bw, int bh, int cx, int cy, int rx, int ry) {
        double best = -1; int bx = 0, by = 0;
        for (int oy = -ry; oy <= ry; oy++) {
            int py = cy + oy; if (py < 0 || py + ah > bh) continue;
            for (int ox = -rx; ox <= rx; ox++) {
                int px = cx + ox; if (px < 0 || px + aw > bw) continue;
                long sum = 0;
                for (int y = 0; y < ah; y++) { int ai = y * aw; int bi = (py + y) * bw + px; for (int x = 0; x < aw; x++) { int d = a[ai + x] - b[bi + x]; sum += d < 0 ? -d : d; } }
                double m = (double)sum / (aw * ah);
                if (best < 0 || m < best) { best = m; bx = px; by = py; }
            }
        }
        return new double[] { best, bx, by };
    }
    public static double[] BandStats(byte[] buf, int stride, int w, int h, int[] tops, int bandH) {
        // ground = the commonest colour of the strip (16 levels a channel);
        // ink = a pixel whose channel differs from the ground by more than 60.
        // Per band: { ink pixels, mean R, mean G, mean B of the ink }.
        int[] hist = new int[4096];
        for (int y = 0; y < h; y++) { int row = y * stride; for (int x = 0; x < w; x++) { int o = row + x * 3; hist[((buf[o + 2] >> 4) << 8) | ((buf[o + 1] >> 4) << 4) | (buf[o] >> 4)]++; } }
        int mode = 0; for (int i = 1; i < 4096; i++) if (hist[i] > hist[mode]) mode = i;
        int gr = ((mode >> 8) & 15) * 16 + 8, gg = ((mode >> 4) & 15) * 16 + 8, gb = (mode & 15) * 16 + 8;
        double[] o2 = new double[tops.Length * 4];
        for (int k = 0; k < tops.Length; k++) {
            long n = 0, sr = 0, sg = 0, sb = 0;
            for (int y = tops[k]; y < tops[k] + bandH; y++) {
                if (y < 0 || y >= h) continue;
                int row = y * stride;
                for (int x = 0; x < w; x++) {
                    int o = row + x * 3; int b = buf[o], g = buf[o + 1], r = buf[o + 2];
                    if (System.Math.Abs(r - gr) > 60 || System.Math.Abs(g - gg) > 60 || System.Math.Abs(b - gb) > 60) { n++; sr += r; sg += g; sb += b; }
                }
            }
            o2[k * 4] = n;
            if (n > 0) { o2[k * 4 + 1] = (double)sr / n; o2[k * 4 + 2] = (double)sg / n; o2[k * 4 + 3] = (double)sb / n; }
        }
        return o2;
    }
}
"@
    }
    return [type]'EbiImageMath'
}

function Compare-EbiPngStrip {
    <#
      How different is the same cell strip of two pictures of the same
      rows? The strip FromFrac..ToFrac of A's width, rows BodyFrom..1 of
      its height (the data under the headers), is SEARCHED for in B: B
      stretched to A's exact width and height (+-2 / 3 %), the strip slid
      +-12 px and +-RangeY rows; then B at A's height in its own
      proportions, the strip slid across every column (10/21: the same job
      read 52 at a fixed place, 46 at A's height alone -- one picture was
      drawn 5 % taller). SaveBase: <SaveBase>.a.png / .b.png = A's strip and
      the best place in B, for a person. Both at half size. Score = the
      least mean grey difference found (0 = the same, 255 = black vs
      white); ink = the share of dark pixels in A's strip.
      -> @{ ok; failure; message; score; ink; x0; x1 }
         failures: file_not_found | image_read_error | geometry
    #>
    param([string]$PathA, [string]$PathB, [double]$FromFrac, [double]$ToFrac, [double]$BodyFrom = 0, [int]$RangeY = 8, [string]$SaveBase = '')
    foreach ($p in @($PathA, $PathB)) { if ([string]::IsNullOrWhiteSpace($p) -or -not (Test-Path -LiteralPath $p -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = ('no picture at ' + $p); score = -1; ink = 0; x0 = 0; x1 = 0 } } }
    return (Compare-EbiPngStripCore -PathA $PathA -PathB $PathB -FromFrac $FromFrac -ToFrac $ToFrac -BodyFrom $BodyFrom -RangeY $RangeY -SaveBase $SaveBase)
}

function Get-EbiGrayHalf {
    # GDI+: an image drawn at (ScaleX, ScaleY) x its size, halved, cropped to
    # (sx, sy, w, h) of that, as grey bytes. -> @{ g; w; h }
    param($Image, [double]$ScaleX, [double]$ScaleY = 0, [int]$SX = 0, [int]$SY = 0, [int]$W = 0, [int]$H = 0)
    if ($ScaleY -le 0) { $ScaleY = $ScaleX }
    $fw = [int][Math]::Max(1, [Math]::Round($Image.Width * $ScaleX / 2.0)); $fh = [int][Math]::Max(1, [Math]::Round($Image.Height * $ScaleY / 2.0))
    if ($W -le 0) { $W = $fw - $SX }; if ($H -le 0) { $H = $fh - $SY }
    $W = [Math]::Max(1, $W); $H = [Math]::Max(1, $H)
    $bmp = New-Object System.Drawing.Bitmap($W, $H, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $gfx.Clear([System.Drawing.Color]::White)
        $gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $gfx.DrawImage($Image, (New-Object System.Drawing.RectangleF([single](-$SX), [single](-$SY), [single]$fw, [single]$fh)))
        $data = $bmp.LockBits((New-Object System.Drawing.Rectangle(0, 0, $W, $H)), [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $buf = New-Object byte[] ([int]$data.Stride * $H)
        [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $buf, 0, $buf.Length)
        $st = [int]$data.Stride
        $bmp.UnlockBits($data)
        return @{ g = [EbiImageMath]::Gray($buf, $st, $W, $H); w = $W; h = $H }
    } finally { $gfx.Dispose(); $bmp.Dispose() }
}

function Save-EbiGrayPng {
    # GDI+: grey bytes (w x h) -> a PNG, for a person to see what was compared.
    param([byte[]]$G, [int]$W, [int]$H, [string]$Path)
    $bmp = New-Object System.Drawing.Bitmap($W, $H, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    try {
        $data = $bmp.LockBits((New-Object System.Drawing.Rectangle(0, 0, $W, $H)), [System.Drawing.Imaging.ImageLockMode]::WriteOnly, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $st = [int]$data.Stride
        $buf = [EbiImageMath]::GrayToRgb($G, $W, $H, $st)
        [System.Runtime.InteropServices.Marshal]::Copy($buf, 0, $data.Scan0, $buf.Length)
        $bmp.UnlockBits($data)
        New-EbiParentDirectory -Path $Path
        $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally { $bmp.Dispose() }
}

function Get-EbiStripCandidates {
    <#
      PURE. The sizes B is tried at: @(@(scaleX, scaleY, kind)), first B
      stretched to A's width and height (x +-2 %, y +-3 %), then A's height
      in B's own proportions. Each element is parenthesised: in PowerShell
      the comma binds tighter than *, and @($a * $b, ...) once multiplied
      by an array (10/21 SJODWDBC: "op_Multiply" on Object[]).
    #>
    param([double]$ScaleW, [double]$ScaleH)
    $out = New-Object System.Collections.ArrayList
    foreach ($kx in @(1.0, 0.98, 1.02)) { foreach ($ky in @(1.0, 0.97, 1.03)) { [void]$out.Add(@(($ScaleW * $kx), ($ScaleH * $ky), 'fit')) } }
    foreach ($k in @(1.0, 0.97, 1.03)) { [void]$out.Add(@(($ScaleH * $k), ($ScaleH * $k), 'height')) }
    return $out.ToArray()
}

function Compare-EbiPngStripCore {
    # The GDI+ half of Compare-EbiPngStrip (see Get-EbiPngSizeCore).
    param([string]$PathA, [string]$PathB, [double]$FromFrac, [double]$ToFrac, [double]$BodyFrom, [int]$RangeY, [string]$SaveBase = '')
    [void](Get-EbiDrawing); [void](Get-EbiImageMath)
    $objs = New-Object System.Collections.ArrayList
    try {
        $msA = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($PathA)); [void]$objs.Add($msA)
        $imA = [System.Drawing.Image]::FromStream($msA); [void]$objs.Add($imA)
        $msB = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($PathB)); [void]$objs.Add($msB)
        $imB = [System.Drawing.Image]::FromStream($msB); [void]$objs.Add($imB)
        $g = Get-EbiStripGeometry -WidthA $imA.Width -HeightA $imA.Height -WidthB $imB.Width -HeightB $imB.Height -FromFrac $FromFrac -ToFrac $ToFrac
        if (-not $g['ok'] -and $g['message'] -notmatch 'too narrow') { return @{ ok = $false; failure = 'geometry'; message = $g['message']; score = -1; ink = 0; x0 = $g['x0']; x1 = $g['x1']; how = '' } }
        $y0 = [int][Math]::Floor([Math]::Max(0.0, [Math]::Min(0.9, $BodyFrom)) * $imA.Height)
        $a = Get-EbiGrayHalf -Image $imA -ScaleX 1.0 -SX ([int]($g['x0'] / 2)) -SY ([int]($y0 / 2)) -W ([int](($g['x1'] - $g['x0']) / 2)) -H ([int](($imA.Height - $y0) / 2))
        $ink = [EbiImageMath]::DarkShare($a['g'])
        # candidate sizes for B: (1) stretched to A's exact width and height --
        # the same range snapped twice, one picture drawn taller (10/21
        # MJODWWBV: the same job read 46 at A's height alone) -- then
        # (2) A's height, B's own proportions (another right edge)
        $sxW = [double]$imA.Width / $imB.Width; $syH = [double]$imA.Height / $imB.Height
        $cands = @(Get-EbiStripCandidates -ScaleW $sxW -ScaleH $syH)
        $best = -1.0; $how = ''; $bestB = $null; $bx = 0; $by = 0
        foreach ($c in $cands) {
            $b = Get-EbiGrayHalf -Image $imB -ScaleX ([double]$c[0]) -ScaleY ([double]$c[1])
            if ($b['w'] -lt $a['w'] -or $b['h'] -lt $a['h']) { continue }
            $cy = [int]($y0 / 2 * ([double]$c[1] / $syH))
            if ([string]$c[2] -eq 'fit') { $cx = [int]($g['x0'] / 2 * ([double]$c[0] / $sxW)); $rx = 12 }
            else { $cx = [int](($b['w'] - $a['w']) / 2); $rx = $cx + 1 }
            $r = [EbiImageMath]::BestDiff($a['g'], $a['w'], $a['h'], $b['g'], $b['w'], $b['h'], $cx, $cy, $rx, $RangeY)
            if ([double]$r[0] -ge 0 -and ($best -lt 0 -or [double]$r[0] -lt $best)) { $best = [double]$r[0]; $how = ('{0} x{1:0.00}/y{2:0.00}' -f $c[2], $c[0], $c[1]); $bestB = $b; $bx = [int]$r[1]; $by = [int]$r[2] }
            if ($best -ge 0 -and $best -le 2) { break }
        }
        if ($best -lt 0) { return @{ ok = $false; failure = 'geometry'; message = 'the strip never fits inside the other picture'; score = -1; ink = $ink; x0 = $g['x0']; x1 = $g['x1']; how = '' } }
        if ($SaveBase -ne '') {
            # what was compared, for a person: A's strip and the best place in B
            try {
                Save-EbiGrayPng -G $a['g'] -W $a['w'] -H $a['h'] -Path ($SaveBase + '.a.png')
                Save-EbiGrayPng -G ([EbiImageMath]::Crop($bestB['g'], $bestB['w'], $bx, $by, $a['w'], $a['h'])) -W $a['w'] -H $a['h'] -Path ($SaveBase + '.b.png')
            } catch { }
        }
        return @{ ok = $true; failure = ''; message = ''; score = [Math]::Round($best, 2); ink = [Math]::Round($ink, 4); x0 = $g['x0']; x1 = $g['x1']; how = $how }
    } catch {
        return @{ ok = $false; failure = 'image_read_error'; message = $_.Exception.Message; score = -1; ink = 0; x0 = 0; x1 = 0; how = '' }
    } finally {
        for ($i = $objs.Count - 1; $i -ge 0; $i--) { try { $objs[$i].Dispose() } catch { } }
    }
}

function Get-EbiBandColors {
    <#
      Per band of the vertical strip X..X+W of a PNG: how much ink and its
      mean colour (EbiImageMath.BandStats; ground = the strip's commonest
      colour). -Tops: each band's top row, -Height: band height.
      -> @{ ok; failure; message; bands = @(@{ top; ink; r; g; b }) }
    #>
    param([string]$Path, [int]$X, [int]$W, [int[]]$Tops, [int]$Height)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = ('file not found: ' + $Path); bands = @() } }
    return (Get-EbiBandColorsCore -Path $Path -X $X -W $W -Tops $Tops -Height $Height)
}

function Get-EbiBandColorsCore {
    # The GDI+ half of Get-EbiBandColors (see Get-EbiPngSizeCore).
    param([string]$Path, [int]$X, [int]$W, [int[]]$Tops, [int]$Height)
    [void](Get-EbiDrawing); [void](Get-EbiImageMath)
    $ms = $null; $img = $null; $bmp = $null
    try {
        $ms = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($Path))
        $img = [System.Drawing.Image]::FromStream($ms)
        $x0 = [Math]::Max(0, $X); $x1 = [Math]::Min([int]$img.Width, $X + $W); $ih = [int]$img.Height
        if ($x1 -le $x0) { return @{ ok = $false; failure = 'geometry'; message = 'the strip is outside the picture'; bands = @() } }
        $bmp = New-Object System.Drawing.Bitmap($img)
        $data = $bmp.LockBits((New-Object System.Drawing.Rectangle($x0, 0, ($x1 - $x0), $ih)), [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $st = [int]$data.Stride
        $buf = New-Object byte[] ($st * $ih)
        [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $buf, 0, $buf.Length)
        $bmp.UnlockBits($data)
        $v = [EbiImageMath]::BandStats($buf, $st, ($x1 - $x0), $ih, [int[]]$Tops, $Height)
        $bands = New-Object System.Collections.ArrayList
        for ($k = 0; $k -lt @($Tops).Count; $k++) { [void]$bands.Add(@{ top = [int]$Tops[$k]; ink = [double]$v[$k * 4]; r = [double]$v[$k * 4 + 1]; g = [double]$v[$k * 4 + 2]; b = [double]$v[$k * 4 + 3] }) }
        return @{ ok = $true; failure = ''; message = ''; bands = $bands.ToArray() }
    } catch {
        return @{ ok = $false; failure = 'image_read_error'; message = $_.Exception.Message; bands = @() }
    } finally {
        if ($null -ne $bmp) { $bmp.Dispose() }
        if ($null -ne $img) { $img.Dispose() }
        if ($null -ne $ms) { $ms.Dispose() }
    }
}

function Get-EbiGrayRegion {
    <#
      Grey bytes of a rectangle of a PNG (clipped to the picture).
      -> @{ ok; failure; message; g; w; h; x0; y0 }
    #>
    param([string]$Path, [int]$X, [int]$Y, [int]$W, [int]$H)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = ('file not found: ' + $Path); g = $null; w = 0; h = 0; x0 = 0; y0 = 0 } }
    return (Get-EbiGrayRegionCore -Path $Path -X $X -Y $Y -W $W -H $H)
}

function Get-EbiGrayRegionCore {
    # The GDI+ half of Get-EbiGrayRegion (see Get-EbiPngSizeCore).
    param([string]$Path, [int]$X, [int]$Y, [int]$W, [int]$H)
    [void](Get-EbiDrawing); [void](Get-EbiImageMath)
    $ms = $null; $img = $null; $bmp = $null
    try {
        $ms = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($Path))
        $img = [System.Drawing.Image]::FromStream($ms)
        $x0 = [Math]::Max(0, $X); $y0 = [Math]::Max(0, $Y)
        $x1 = [Math]::Min([int]$img.Width, $X + $W); $y1 = [Math]::Min([int]$img.Height, $Y + $H)
        if ($x1 -le $x0 -or $y1 -le $y0) { return @{ ok = $false; failure = 'geometry'; message = 'the region is outside the picture'; g = $null; w = 0; h = 0; x0 = $x0; y0 = $y0 } }
        $bmp = New-Object System.Drawing.Bitmap($img)
        $rw = $x1 - $x0; $rh = $y1 - $y0
        $data = $bmp.LockBits((New-Object System.Drawing.Rectangle($x0, $y0, $rw, $rh)), [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $st = [int]$data.Stride
        $buf = New-Object byte[] ($st * $rh)
        [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $buf, 0, $buf.Length)
        $bmp.UnlockBits($data)
        return @{ ok = $true; failure = ''; message = ''; g = [EbiImageMath]::Gray($buf, $st, $rw, $rh); w = $rw; h = $rh; x0 = $x0; y0 = $y0 }
    } catch {
        return @{ ok = $false; failure = 'image_read_error'; message = $_.Exception.Message; g = $null; w = 0; h = 0; x0 = 0; y0 = 0 }
    } finally {
        if ($null -ne $bmp) { $bmp.Dispose() }
        if ($null -ne $img) { $img.Dispose() }
        if ($null -ne $ms) { $ms.Dispose() }
    }
}

function Get-EbiInkCounts {
    <#
      Per pixel row, how many pixels of a vertical strip are "ink":
        dark  luminance (R+G+B)/3 below MaxLuma
        blue  B > 140 and R < 120 and G < 140 (link text)
      Strip: columns [X, X+Width), rows [YFrom, YTo) -- a negative YTo
      counts from the bottom edge (-35 = stop 35 px above it). The result
      is relative to YFrom: counts[0] is row YFrom.
      -> @{ ok; failure; message; counts; width; height; yFrom }
         failures: file_not_found | image_read_error
    #>
    param([string]$Path, [int]$X, [int]$Width, [int]$YFrom = 0, [int]$YTo = 0, [string]$Ink = 'dark', [int]$MaxLuma = 160)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = ('file not found: ' + $Path); counts = @(); width = 0; height = 0; yFrom = $YFrom } }
    return (Get-EbiInkCountsCore -Path $Path -X $X -Width $Width -YFrom $YFrom -YTo $YTo -Ink $Ink -MaxLuma $MaxLuma)
}

function Get-EbiInkCountsCore {
    # The GDI+ half of Get-EbiInkCounts (see Get-EbiPngSizeCore). LockBits
    # once and walk a byte copy -- GetPixel per pixel is far too slow in PS.
    param([string]$Path, [int]$X, [int]$Width, [int]$YFrom, [int]$YTo, [string]$Ink, [int]$MaxLuma)
    [void](Get-EbiDrawing)
    $ms = $null; $img = $null; $bmp = $null
    try {
        $ms = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($Path))
        $img = [System.Drawing.Image]::FromStream($ms)
        $bmp = New-Object System.Drawing.Bitmap($img)
        $w = [int]$bmp.Width; $h = [int]$bmp.Height
        $y0 = [Math]::Max(0, $YFrom)
        $y1 = if ($YTo -le 0) { $h + $YTo } else { [Math]::Min($h, $YTo) }
        $x0 = [Math]::Max(0, $X); $x1 = [Math]::Min($w, $X + $Width)
        if ($y1 -le $y0 -or $x1 -le $x0) { return @{ ok = $true; failure = ''; message = 'strip outside the image'; counts = @(); width = $w; height = $h; yFrom = $y0 } }
        $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
        $data = $bmp.LockBits($rect, [System.Drawing.Imaging.ImageLockMode]::ReadOnly, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $stride = [int]$data.Stride
        $buf = New-Object byte[] ($stride * $h)
        [System.Runtime.InteropServices.Marshal]::Copy($data.Scan0, $buf, 0, $buf.Length)
        $bmp.UnlockBits($data)
        $counts = New-Object int[] ($y1 - $y0)
        $blue = ($Ink -eq 'blue')
        for ($y = $y0; $y -lt $y1; $y++) {
            $row = $y * $stride
            $n = 0
            for ($x = $x0; $x -lt $x1; $x++) {
                $o = $row + $x * 3
                $b = [int]$buf[$o]; $g = [int]$buf[$o + 1]; $r = [int]$buf[$o + 2]
                if ($blue) { if ($b -gt 140 -and $r -lt 120 -and $g -lt 140) { $n++ } }
                elseif ((($r + $g + $b) / 3) -lt $MaxLuma) { $n++ }
            }
            $counts[$y - $y0] = $n
        }
        return @{ ok = $true; failure = ''; message = ''; counts = $counts; width = $w; height = $h; yFrom = $y0 }
    } catch {
        return @{ ok = $false; failure = 'image_read_error'; message = $_.Exception.Message; counts = @(); width = 0; height = 0; yFrom = $YFrom }
    } finally {
        if ($null -ne $bmp) { $bmp.Dispose() }
        if ($null -ne $img) { $img.Dispose() }
        if ($null -ne $ms) { $ms.Dispose() }
    }
}
