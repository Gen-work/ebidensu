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

# Picture shapes on a sheet, skipping the verifyMark_* red rectangles.
# msoPicture=13, msoLinkedPicture=11. Ctrl+G groups (msoGroup=6) are
# flattened: each child picture is returned on its own (a grouped strip
# exported as one composite PNG could exceed the Windows OCR max image
# dimension). IMPORTANT: child shapes inside a group can report
# group-RELATIVE Top/Left (observed Top=0 in production workbooks), so
# each entry carries the TOP-LEVEL shape's sheet coordinates for section
# filtering and coarse ordering; the child's own Top/Left only orders
# pictures within their group.
# Returns hashtable entries: @{ Shape; TopLevelTop; TopLevelLeft; SubTop; SubLeft }
function Get-EvidencePictureEntries {
    param($Worksheet)
    $entries = @()
    foreach ($sp in $Worksheet.Shapes) {
        $t = 0
        try { $t = [int]$sp.Type } catch {}
        $nm = ''
        try { $nm = [string]$sp.Name } catch {}
        if ($nm -like 'verifyMark_*') { continue }
        $top = 0.0; $left = 0.0
        try { $top = [double]$sp.Top } catch {}
        try { $left = [double]$sp.Left } catch {}
        if ($t -eq 6) {
            foreach ($child in @(Get-GroupPictureShapes $sp)) {
                $st = 0.0; $sl = 0.0
                try { $st = [double]$child.Top } catch {}
                try { $sl = [double]$child.Left } catch {}
                $entries += @{ Shape = $child; TopLevelTop = $top; TopLevelLeft = $left; SubTop = $st; SubLeft = $sl }
            }
        } elseif ($t -eq 13 -or $t -eq 11) {
            $entries += @{ Shape = $sp; TopLevelTop = $top; TopLevelLeft = $left; SubTop = 0.0; SubLeft = 0.0 }
        }
    }
    # Round tops so the pictures of one horizontal strip (near-equal tops)
    # sort left-to-right in capture order.
    # NOTE: no comma protection on the returns in this file -- every caller
    # wraps the call in @(...), and @( ,@($arr) ) yields a NESTED array in
    # PS 5.1: member enumeration then returns Object[] and [double] casts
    # explode ('cannot convert System.Object[] to System.Double').
    return @($entries | Sort-Object `
        @{Expression = { [Math]::Round([double]$_.TopLevelTop, 0) }}, `
        @{Expression = { [double]$_.TopLevelLeft }}, `
        @{Expression = { [Math]::Round([double]$_.SubTop, 0) }}, `
        @{Expression = { [double]$_.SubLeft }})
}

# Recursively collects the picture shapes inside one group. A GroupItems
# read failure is reported (not swallowed): losing a whole strip of
# evidence silently is exactly the bug this diagnostic exists for.
function Get-GroupPictureShapes {
    param($GroupShape)
    $found = @()
    $queue = New-Object System.Collections.Generic.Queue[object]
    $queue.Enqueue($GroupShape)
    while ($queue.Count -gt 0) {
        $sp = $queue.Dequeue()
        $t = 0
        try { $t = [int]$sp.Type } catch {}
        if ($t -eq 6) {
            try {
                foreach ($child in $sp.GroupItems) { $queue.Enqueue($child) }
            } catch {
                $gnm = ''
                try { $gnm = [string]$sp.Name } catch {}
                Write-Host ("  [WARN] cannot enumerate group '{0}' children: {1}" -f $gnm, $_.Exception.Message) -ForegroundColor Yellow
            }
            continue
        }
        if ($t -ne 13 -and $t -ne 11) { continue }
        $nm = ''
        try { $nm = [string]$sp.Name } catch {}
        if ($nm -like 'verifyMark_*') { continue }
        $found += $sp
    }
    return $found
}

# Back-compat wrapper: bare picture shapes in display order.
function Get-EvidencePictureShapes {
    param($Worksheet)
    return @(Get-EvidencePictureEntries $Worksheet) | ForEach-Object { $_.Shape }
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
        return @()
    }
    try { $ws.Visible = -1 } catch {}
    try { $ws.Activate() | Out-Null } catch {}

    $allEntries = @(Get-EvidencePictureEntries $ws)
    $entries = $allEntries
    if ($TopMin -ge 0 -or $TopMax -ge 0) {
        # Section membership uses the TOP-LEVEL shape's Top: pictures inside
        # a Ctrl+G group can report group-relative Top (e.g. 0), so the
        # group's own sheet position decides which correl section they
        # belong to.
        $entries = @($entries | Where-Object {
            $st = [double]$_.TopLevelTop
            (($TopMin -lt 0) -or ($st -ge $TopMin)) -and (($TopMax -lt 0) -or ($st -lt $TopMax))
        })
    }
    if ($entries.Count -eq 0) {
        # Diagnostics: say WHY nothing was exported so the operator can tell
        # a shape-type problem from a section-bounds problem at a glance.
        if ($allEntries.Count -gt 0) {
            $tops = @($allEntries | ForEach-Object {
                [Math]::Round([double]$_.TopLevelTop, 0)
            }) -join ', '
            Write-Host ("  [DIAG] sheet '{0}' has {1} picture(s) but none inside section Top range [{2}, {3}); top-level picture Tops: {4}" -f `
                $SheetName, $allEntries.Count, [Math]::Round($TopMin, 0), [Math]::Round($TopMax, 0), $tops) -ForegroundColor Yellow
        } else {
            $typeCounts = @{}
            try {
                foreach ($sp in $ws.Shapes) {
                    $t = -1
                    try { $t = [int]$sp.Type } catch {}
                    $typeCounts[$t] = 1 + [int]$typeCounts[$t]
                }
            } catch {}
            $desc = @($typeCounts.GetEnumerator() | Sort-Object Name | ForEach-Object { ('msoType{0} x{1}' -f $_.Key, $_.Value) }) -join ', '
            if ([string]::IsNullOrWhiteSpace($desc)) { $desc = 'no shapes at all' }
            Write-Host ("  [DIAG] no picture shapes (msoPicture=13 / msoLinkedPicture=11) on sheet '{0}'; found: {1}" -f $SheetName, $desc) -ForegroundColor Yellow
        }
        return @()
    }
    if (-not (Test-Path -LiteralPath $OutDir)) {
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
    }

    $paths = @()
    $i = 0
    foreach ($e in $entries) {
        $i++
        $png = Join-Path $OutDir ('{0}_{1:D2}.png' -f $BaseName, $i)
        if (Export-ShapeToPng $ws $e.Shape $png) {
            $paths += $png
        } else {
            Write-Host ("  [WARN] picture {0} on sheet '{1}' was not exported." -f $i, $SheetName) -ForegroundColor Yellow
        }
    }
    return $paths
}
