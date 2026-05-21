# Sample-HighlightColor.ps1
# 高亮行の中心1ピクセルを取って実RGBを表示する
param([string]$ImagePath, [int]$X, [int]$Y)
Add-Type -AssemblyName System.Drawing
$bmp = [System.Drawing.Bitmap]::FromFile("C:\workspace\VerifyToolsHO\inActive_yellow.png")
try {
    $p = $bmp.GetPixel($X, $Y)
    Write-Host ("RGB at ({0},{1}) = R{2} G{3} B{4}  #{5:X2}{6:X2}{7:X2}" -f `
        $X, $Y, $p.R, $p.G, $p.B, $p.R, $p.G, $p.B)
} finally { $bmp.Dispose() }
