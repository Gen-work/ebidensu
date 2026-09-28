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
