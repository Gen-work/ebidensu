# ============================================================
#  Crop-Snap.ps1
#
#  Crops N pixels off all four sides of PNG file(s).
#  Removes DWM shadow / border from window screenshots.
#
#  Modes:
#    -Path xxx.png         : crop single file in place. No marker.
#    -Dir  folder          : batch crop all PNGs. Marks done files with
#                            hidden ".cropped" sidecar; skips already-marked
#                            unless -Force.
#
#  As library:
#    . .\Crop-Snap.ps1
#    Invoke-CropPng -path "x.png" -cropPx 15
#
#  Usage examples:
#    .\Crop-Snap.ps1 -Path "snap\GIFT_HM\JIDSL48S.png"
#    .\Crop-Snap.ps1 -Dir  "snap\GIFT_HM"
#    .\Crop-Snap.ps1 -Dir  "snap\GIFT_HM" -CropPx 20 -Force
#    .\Crop-Snap.ps1 -Dir  "snap" -Recurse        # all subfolders
#
#  Save as UTF-8 with BOM, CRLF.
# ============================================================

param(
    [string]$Path    = "",
    [string]$Dir     = "",
    [int]$CropPx     = 15,
    [switch]$Force,
    [switch]$Recurse
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

# ============================================================
# Core: crop one PNG file in-place
# ============================================================
function Invoke-CropPng {
    param(
        [Parameter(Mandatory=$true)][string]$path,
        [int]$cropPx = 15
    )

    if (-not (Test-Path -LiteralPath $path)) {
        throw "File not found: $path"
    }

    # Read bytes first so we don't hold a file lock on save
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $ms    = New-Object System.IO.MemoryStream(, $bytes)
    $tmpPath = "$path.crop.tmp"

    try {
        $orig = [System.Drawing.Image]::FromStream($ms)
        try {
            $newW = $orig.Width  - 2 * $cropPx
            $newH = $orig.Height - 2 * $cropPx
            if ($newW -le 0 -or $newH -le 0) {
                throw ("Image too small ({0}x{1}) to crop {2} px per side" -f $orig.Width, $orig.Height, $cropPx)
            }

            $bmp = New-Object System.Drawing.Bitmap($newW, $newH)
            try {
                $gfx = [System.Drawing.Graphics]::FromImage($bmp)
                try {
                    $srcRect = New-Object System.Drawing.Rectangle($cropPx, $cropPx, $newW, $newH)
                    $dstRect = New-Object System.Drawing.Rectangle(0, 0, $newW, $newH)
                    $gfx.DrawImage($orig, $dstRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
                } finally {
                    $gfx.Dispose()
                }
                $bmp.Save($tmpPath, [System.Drawing.Imaging.ImageFormat]::Png)
            } finally {
                $bmp.Dispose()
            }
        } finally {
            $orig.Dispose()
        }
    } finally {
        $ms.Dispose()
    }

    # Replace original atomically-ish
    Move-Item -LiteralPath $tmpPath -Destination $path -Force
}

# ============================================================
# Batch: walk a directory
# ============================================================
function Invoke-CropDir {
    param(
        [Parameter(Mandatory=$true)][string]$dir,
        [int]$cropPx     = 15,
        [switch]$Recurse,
        [switch]$Force
    )

    if (-not (Test-Path -LiteralPath $dir)) {
        throw "Directory not found: $dir"
    }

    $gciArgs = @{ LiteralPath = $dir; Filter = "*.png"; File = $true }
    if ($Recurse) { $gciArgs.Recurse = $true }
    $files = @(Get-ChildItem @gciArgs)

    Write-Host ("Found {0} PNG file(s) in {1}{2}" -f $files.Count, $dir, $(if ($Recurse) { " (recursive)" } else { "" }))

    $done = 0; $skipped = 0; $failed = 0
    foreach ($f in $files) {
        $marker = "$($f.FullName).cropped"
        if ((Test-Path -LiteralPath $marker) -and -not $Force) {
            $skipped++
            continue
        }
        try {
            Invoke-CropPng -path $f.FullName -cropPx $cropPx
            # Create hidden marker
            "" | Out-File -LiteralPath $marker -Encoding ASCII -NoNewline
            try {
                $mi = Get-Item -LiteralPath $marker -Force
                $mi.Attributes = ([System.IO.FileAttributes]::Hidden -bor [System.IO.FileAttributes]::Archive)
            } catch {}
            $done++
            Write-Host ("  [OK]   {0}" -f $f.Name) -ForegroundColor Green
        } catch {
            $failed++
            Write-Host ("  [FAIL] {0} - {1}" -f $f.Name, $_.Exception.Message) -ForegroundColor Red
        }
    }
    Write-Host ""
    Write-Host ("Cropped: {0}, Skipped (already done): {1}, Failed: {2}" -f $done, $skipped, $failed) -ForegroundColor Cyan
}

# ============================================================
# CLI dispatch (no-op if dot-sourced with no args)
# ============================================================
if (-not [string]::IsNullOrWhiteSpace($Path)) {
    Write-Host ("Cropping (single): {0}  [{1} px]" -f $Path, $CropPx)
    Invoke-CropPng -path $Path -cropPx $CropPx
    Write-Host "[OK] done." -ForegroundColor Green
} elseif (-not [string]::IsNullOrWhiteSpace($Dir)) {
    Write-Host ("Cropping (batch): {0}  [{1} px]" -f $Dir, $CropPx)
    Invoke-CropDir -dir $Dir -cropPx $CropPx -Recurse:$Recurse -Force:$Force
}
`
``n
--- File: ExcelHelpers.ps1 ---