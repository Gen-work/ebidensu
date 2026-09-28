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
#  As library: dot-source kernel\Image.ps1 instead and call
#    Invoke-EbiCropPng -Path "x.png" -CropPx 15        (P1-20)
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
    # Per-side overrides in px. -1 (default) = inherit CropPx for that side
    # (uniform crop; existing -CropPx-only usage is unchanged).
    [int]$CropLeft   = -1,
    [int]$CropTop    = -1,
    [int]$CropRight  = -1,
    [int]$CropBottom = -1,
    [switch]$Force,
    [switch]$Recurse
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

# ============================================================
# Core: crop one PNG file in-place
# ============================================================
. (Join-Path $PSScriptRoot 'kernel/Image.ps1')   # Invoke-EbiCropPng (P1-20)

# ============================================================
# Batch: walk a directory
# ============================================================
function Invoke-CropDir {
    param(
        [Parameter(Mandatory=$true)][string]$dir,
        [int]$cropPx     = 15,
        # -1 (default) = inherit cropPx for that side (uniform crop).
        [int]$cropLeft   = -1,
        [int]$cropTop    = -1,
        [int]$cropRight  = -1,
        [int]$cropBottom = -1,
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
            $cropResult = Invoke-EbiCropPng -Path $f.FullName -CropPx $cropPx `
                -Left $cropLeft -Top $cropTop -Right $cropRight -Bottom $cropBottom
            if (-not $cropResult.ok) { throw $cropResult.message }
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
    $cropResult = Invoke-EbiCropPng -Path $Path -CropPx $CropPx `
        -Left $CropLeft -Top $CropTop -Right $CropRight -Bottom $CropBottom
    if (-not $cropResult.ok) { throw $cropResult.message }
    Write-Host "[OK] done." -ForegroundColor Green
} elseif (-not [string]::IsNullOrWhiteSpace($Dir)) {
    Write-Host ("Cropping (batch): {0}  [{1} px]" -f $Dir, $CropPx)
    Invoke-CropDir -dir $Dir -cropPx $CropPx `
        -cropLeft $CropLeft -cropTop $CropTop -cropRight $CropRight -cropBottom $CropBottom `
        -Recurse:$Recurse -Force:$Force
}
