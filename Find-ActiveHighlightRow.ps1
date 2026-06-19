function Find-ActiveHighlightRow {
    <#
    Detect the Edge Ctrl+F active-match (orange) row.
    Inactive matches (pale yellow) are ignored.
    .OUTPUTS @{ Top=int; Bottom=int; Score=int } or $null
    #>
    param(
        [Parameter(Mandatory=$true)][string]$ImagePath,
        # Override with values measured on the office PC (step 1).
        [int]$ActiveR = 255, [int]$ActiveG = 150, [int]$ActiveB = 50, #FF9632
        [int]$InactiveR = 255, [int]$InactiveG = 255, [int]$InactiveB = 0, #FFFF00
        [int]$Tolerance = 25,
        [int]$MinPixelsPerRow = 30
    )
    Add-Type -AssemblyName System.Drawing
    $bmp = [System.Drawing.Bitmap]::FromFile($ImagePath)
    try {
        $w = $bmp.Width; $h = $bmp.Height
        $rowCount = New-Object int[] $h
        for ($y = 0; $y -lt $h; $y++) {
            $cnt = 0
            for ($x = 0; $x -lt $w; $x++) {
                $px = $bmp.GetPixel($x, $y)
                if ([Math]::Abs($px.R - $ActiveR) -le $Tolerance -and
                    [Math]::Abs($px.G - $ActiveG) -le $Tolerance -and
                    [Math]::Abs($px.B - $ActiveB) -le $Tolerance) { $cnt++ }
            }
            $rowCount[$y] = $cnt
        }
        # Extract the densest contiguous cluster of high-count rows.
        $best = $null; $start = -1
        for ($y = 0; $y -lt $h; $y++) {
            if ($rowCount[$y] -ge $MinPixelsPerRow) {
                if ($start -lt 0) { $start = $y }
            } elseif ($start -ge 0) {
                $end = $y - 1
                $score = ($rowCount[$start..$end] | Measure-Object -Sum).Sum
                if (-not $best -or $score -gt $best.Score) {
                    $best = @{ Top = $start; Bottom = $end; Score = $score }
                }
                $start = -1
            }
        }
        return $best
    } finally { $bmp.Dispose() }
}
