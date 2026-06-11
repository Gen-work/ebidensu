# ============================================================
#  EvidenceImageExport.ps1
#
#  Excel COM helpers to export the screenshot pictures embedded in an
#  evidence workbook sheet to PNG files, so OcrWindows.ps1 can read
#  them. Dot-source only (no param() block).
#
#  Export trick (Excel has no direct Shape->file API): copy the shape
#  to the clipboard, paste it into a temporary ChartObject sized to the
#  shape, Chart.Export PNG, delete the chart. The chart frame may add a
#  ~1px border; harmless for OCR. NOTE: this clobbers the clipboard.
# ============================================================

# Picture shapes on a sheet, top-to-bottom then left-to-right, skipping
# the verifyMark_* red rectangles. msoPicture=13, msoLinkedPicture=11.
# Ctrl+G groups (msoGroup=6) are flattened: each child picture is returned
# on its own. GroupItems report sheet-absolute Top/Left, so ordering and
# region filtering keep working, and per-child export avoids one giant
# composite PNG that would exceed the Windows OCR max image dimension.
function Get-EvidencePictureShapes {
    param($Worksheet)
    $found = @()
    $queue = New-Object System.Collections.Generic.Queue[object]
    foreach ($sp in $Worksheet.Shapes) { $queue.Enqueue($sp) }
    while ($queue.Count -gt 0) {
        $sp = $queue.Dequeue()
        $t = 0
        try { $t = [int]$sp.Type } catch {}
        if ($t -eq 6) {
            try { foreach ($child in $sp.GroupItems) { $queue.Enqueue($child) } } catch {}
            continue
        }
        if ($t -ne 13 -and $t -ne 11) { continue }
        $nm = ''
        try { $nm = [string]$sp.Name } catch {}
        if ($nm -like 'verifyMark_*') { continue }
        $found += $sp
    }
    # Round Top so the pictures of one horizontal strip (near-equal tops)
    # sort left-to-right in capture order.
    return ,@($found | Sort-Object @{Expression = { [Math]::Round([double]$_.Top, 0) }}, @{Expression = { [double]$_.Left }})
}

# Exports one shape to a PNG file. Returns $true when the file exists.
function Export-ShapeToPng {
    param($Worksheet, $Shape, [string]$OutPath)
    $dir = Split-Path -Parent $OutPath
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $chartObj = $null
    try {
        $w = [double]$Shape.Width
        $h = [double]$Shape.Height
        $Shape.Copy() | Out-Null
        Start-Sleep -Milliseconds 100
        $chartObj = $Worksheet.ChartObjects().Add(0, 0, $w, $h)
        $chartObj.Activate() | Out-Null
        $chartObj.Chart.Paste() | Out-Null
        Start-Sleep -Milliseconds 100
        [void]$chartObj.Chart.Export($OutPath, 'PNG')
        return (Test-Path -LiteralPath $OutPath)
    } catch {
        Write-Host ("  [WARN] shape export failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        return $false
    } finally {
        if ($null -ne $chartObj) {
            try { $chartObj.Delete() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($chartObj) } catch {}
        }
    }
}

# Exports every picture on the named sheet as <BaseName>_NN.png under
# OutDir. Returns the array of created file paths (empty when the sheet
# is missing or carries no pictures). Optional TopMin/TopMax (sheet
# points, -1 = unbounded) limit the export to one vertical section, e.g.
# the pictures between two correl-id labels in column A.
function Export-SheetPicturesToPng {
    param($Workbook, [string]$SheetName, [string]$OutDir, [string]$BaseName,
          [double]$TopMin = -1.0, [double]$TopMax = -1.0)
    $ws = $null
    foreach ($s in $Workbook.Worksheets) {
        if ([string]$s.Name -eq $SheetName) { $ws = $s; break }
    }
    if ($null -eq $ws) {
        Write-Host ("  [WARN] sheet not found for picture export: {0}" -f $SheetName) -ForegroundColor Yellow
        return ,@()
    }
    try { $ws.Visible = -1 } catch {}
    try { $ws.Activate() | Out-Null } catch {}

    $shapes = @(Get-EvidencePictureShapes $ws)
    if ($TopMin -ge 0 -or $TopMax -ge 0) {
        $shapes = @($shapes | Where-Object {
            $st = 0.0
            try { $st = [double]$_.Top } catch {}
            (($TopMin -lt 0) -or ($st -ge $TopMin)) -and (($TopMax -lt 0) -or ($st -lt $TopMax))
        })
    }
    if ($shapes.Count -eq 0) { return ,@() }
    if (-not (Test-Path -LiteralPath $OutDir)) {
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
    }

    $paths = @()
    $i = 0
    foreach ($sp in $shapes) {
        $i++
        $png = Join-Path $OutDir ('{0}_{1:D2}.png' -f $BaseName, $i)
        if (Export-ShapeToPng $ws $sp $png) {
            $paths += $png
        } else {
            Write-Host ("  [WARN] picture {0} on sheet '{1}' was not exported." -f $i, $SheetName) -ForegroundColor Yellow
        }
    }
    return ,@($paths)
}
