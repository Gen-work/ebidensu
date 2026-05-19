# ============================================================
# Find-Abend.ps1 v2
#   4 パラメータに簡略化:
#     Row1Top / RowHeight / StatusColLeft / StatusColWidth
#   ROI 計算: roiTop = Row1Top + (RecordIndex - 1) * RowHeight
# ============================================================
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SnapPath,
    [int]$RecordIndex = 1,
    [Parameter(Mandatory)][string]$TemplateGreen,
    [Parameter(Mandatory)][string]$TemplateWhite,
    [Parameter(Mandatory)][int]$Row1Top,
    [Parameter(Mandatory)][int]$RowHeight,
    [Parameter(Mandatory)][int]$StatusColLeft,
    [Parameter(Mandatory)][int]$StatusColWidth,
    [double]$MatchThreshold = 0.80,
    [int]$ColorTolerance = 35,
    [switch]$SaveRoi
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function Get-BestMatchScore {
    param(
        [System.Drawing.Bitmap]$Target,
        [System.Drawing.Bitmap]$Template,
        [int]$Tolerance
    )
    $tw = $Template.Width; $th = $Template.Height
    $aw = $Target.Width;   $ah = $Target.Height
    if ($tw -gt $aw -or $th -gt $ah) { return 0.0 }

    $tplData = $Template.LockBits(
        (New-Object System.Drawing.Rectangle(0,0,$tw,$th)),
        [System.Drawing.Imaging.ImageLockMode]::ReadOnly,
        [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $tgtData = $Target.LockBits(
        (New-Object System.Drawing.Rectangle(0,0,$aw,$ah)),
        [System.Drawing.Imaging.ImageLockMode]::ReadOnly,
        [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    try {
        $tplStride = $tplData.Stride
        $tgtStride = $tgtData.Stride
        $tplBytes = New-Object byte[] ($tplStride * $th)
        $tgtBytes = New-Object byte[] ($tgtStride * $ah)
        [System.Runtime.InteropServices.Marshal]::Copy($tplData.Scan0, $tplBytes, 0, $tplBytes.Length)
        [System.Runtime.InteropServices.Marshal]::Copy($tgtData.Scan0, $tgtBytes, 0, $tgtBytes.Length)

        $totalPx = $tw * $th
        $best = 0.0
        for ($oy = 0; $oy -le ($ah - $th); $oy++) {
            for ($ox = 0; $ox -le ($aw - $tw); $ox++) {
                $hits = 0
                for ($y = 0; $y -lt $th; $y++) {
                    $tplRow = $y * $tplStride
                    $tgtRow = ($oy + $y) * $tgtStride + $ox * 3
                    for ($x = 0; $x -lt $tw; $x++) {
                        $to = $tplRow + $x * 3
                        $go = $tgtRow + $x * 3
                        $db = [Math]::Abs($tplBytes[$to]   - $tgtBytes[$go])
                        $dg = [Math]::Abs($tplBytes[$to+1] - $tgtBytes[$go+1])
                        $dr = [Math]::Abs($tplBytes[$to+2] - $tgtBytes[$go+2])
                        if ($db -le $Tolerance -and $dg -le $Tolerance -and $dr -le $Tolerance) {
                            $hits++
                        }
                    }
                }
                $score = $hits / [double]$totalPx
                if ($score -gt $best) { $best = $score }
            }
        }
        return $best
    } finally {
        $Template.UnlockBits($tplData)
        $Target.UnlockBits($tgtData)
    }
}

# ROI: 目標行の「処理状態」セル
$roiTop = $Row1Top + ($RecordIndex - 1) * $RowHeight
$roiRect = New-Object System.Drawing.Rectangle($StatusColLeft, $roiTop, $StatusColWidth, $RowHeight)

$snap = [System.Drawing.Bitmap]::FromFile($SnapPath)
try {
    if ($roiRect.Right -gt $snap.Width -or $roiRect.Bottom -gt $snap.Height -or $roiRect.X -lt 0 -or $roiRect.Y -lt 0) {
        throw ("ROI out of snap bounds. snap={0}x{1} roi=({2},{3},{4},{5})" -f `
            $snap.Width, $snap.Height, $roiRect.X, $roiRect.Y, $roiRect.Width, $roiRect.Height)
    }
    $roi = $snap.Clone($roiRect, $snap.PixelFormat)
} finally { $snap.Dispose() }

try {
    if ($SaveRoi) {
        $dbgPath = [System.IO.Path]::ChangeExtension($SnapPath, ".roi_r$RecordIndex.png")
        $roi.Save($dbgPath, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Host ("  [DEBUG] ROI saved: {0}" -f $dbgPath) -ForegroundColor DarkGray
    }
    $bestScore = 0.0
    $bestKind  = "none"
    foreach ($pair in @(
        @{Path=$TemplateGreen; Kind="green"},
        @{Path=$TemplateWhite; Kind="white"}
    )) {
        $tpl = [System.Drawing.Bitmap]::FromFile($pair.Path)
        try {
            $score = Get-BestMatchScore -Target $roi -Template $tpl -Tolerance $ColorTolerance
            Write-Host ("  template={0,-5} score={1:F3}" -f $pair.Kind, $score) -ForegroundColor DarkGray
            if ($score -gt $bestScore) {
                $bestScore = $score
                $bestKind  = $pair.Kind
            }
        } finally { $tpl.Dispose() }
    }
    [PSCustomObject]@{
        IsAbend         = ($bestScore -ge $MatchThreshold)
        MatchScore      = [Math]::Round($bestScore, 3)
        MatchedTemplate = if ($bestScore -ge $MatchThreshold) { $bestKind } else { "none" }
        RoiRect         = ("{0},{1},{2},{3}" -f $roiRect.X, $roiRect.Y, $roiRect.Width, $roiRect.Height)
    }
} finally { $roi.Dispose() }