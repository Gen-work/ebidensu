# ============================================================
#  ExcelHelpers.ps1
#
#  Shared Excel COM helpers for Clone.ps1 / ReplaceEvidence.ps1
#  and (later) Mark.ps1.
#
#  Dot-source from caller. Pure function file:
#    - no param() block (safe under dot-source)
#    - no script-level mutable state
#    - no Add-Type (caller handles Common.ps1 first)
# ============================================================

# -- Excel application lifecycle -----------------------------

function New-ExcelApp {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.ScreenUpdating = $false } catch {}
    return $excel
}

function Close-ExcelApp($excel) {
    if ($null -eq $excel) { return }
    try { $excel.DisplayAlerts = $true } catch {}
    try { $excel.ScreenUpdating = $true } catch {}
    try { $excel.Quit() } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Open-Workbook($excel, [string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Workbook not found: $path" }
    return $excel.Workbooks.Open($path)
}

function Close-Workbook($wb, [bool]$save = $false) {
    if ($null -eq $wb) { return }
    try { $wb.Close([bool]$save) } catch {}
}

# -- Sheet operations ----------------------------------------

function Get-SheetByName($wb, [string]$name) {
    foreach ($ws in $wb.Worksheets) {
        if ($ws.Name -eq $name) { return $ws }
    }
    return $null
}

function Unhide-AllSheets($wb) {
    foreach ($ws in $wb.Worksheets) {
        try { $ws.Visible = -1 } catch {}  # xlSheetVisible
    }
}

# -- Cleanup helpers -----------------------------------------

function Reset-SheetBelowRow($ws, [int]$startRow, [int]$endCol = 20) {
    <#
    Wipe everything visually below $startRow:
      1. Delete all shapes whose top edge is at or below row $startRow's top.
      2. Clear values, fonts, fills, and highlights in A$startRow:T<last>.
    Row $startRow itself is preserved as the first anchor row.
    #>
    if ($null -eq $ws) { return }
    $rowTop = 0.0
    try { $rowTop = [double]$ws.Rows.Item($startRow).Top } catch {}

    # 1) shapes
    $toDel = New-Object System.Collections.Generic.List[string]
    foreach ($s in $ws.Shapes) {
        try {
            if ([double]$s.Top -ge ($rowTop - 1.0)) { $toDel.Add($s.Name) }
        } catch {}
    }
    foreach ($n in $toDel) {
        try { $ws.Shapes.Item($n).Delete() | Out-Null } catch {}
    }

    # 2) range values + formatting
    $xlUp = -4162
    $lastRow = 0
    try { $lastRow = [int]$ws.Cells.Item($ws.Rows.Count, 1).End($xlUp).Row } catch { $lastRow = 0 }
    if ($lastRow -lt $startRow) {
        try {
            $used = $ws.UsedRange
            $lastRow = [int]($used.Row + $used.Rows.Count - 1)
        } catch { $lastRow = $startRow }
    }
    if ($lastRow -ge $startRow) {
        $range = $ws.Range($ws.Cells.Item($startRow, 1), $ws.Cells.Item($lastRow, $endCol))

        # A merged "header" cell (e.g. a colored banner merged across N1:U1...)
        # whose anchor sits ABOVE $startRow can still dip into this clear range.
        # Range.Clear() would wipe the ENTIRE merged cell -- fill and all --
        # which is what silently dropped the N1:U1 header fill after a Replace.
        # Such a merge must cross row $startRow, so snapshot every above-anchored
        # merge found on that single boundary row, clear, then restore it
        # (re-merging defensively in case Clear unmerged it). This protects the
        # header for all three Replace modes, since they share this helper.
        $preserved = New-Object System.Collections.Generic.List[object]
        $seen = @{}
        for ($c = 1; $c -le $endCol; $c++) {
            try {
                $cell = $ws.Cells.Item($startRow, $c)
                if ($cell.MergeCells) {
                    $area = $cell.MergeArea
                    if ([int]$area.Row -lt $startRow) {
                        $addr = [string]$area.Address($true, $true)
                        if (-not $seen.ContainsKey($addr)) {
                            $seen[$addr] = $true
                            $snap = @{ AreaAddr = $addr; Color = $null; ColorIndex = $null; Value = $null }
                            try { $snap.ColorIndex = [int]$area.Cells.Item(1, 1).Interior.ColorIndex } catch {}
                            try { $snap.Color      = [long]$area.Cells.Item(1, 1).Interior.Color } catch {}
                            try { $snap.Value      = $area.Cells.Item(1, 1).Value2 } catch {}
                            $preserved.Add([pscustomobject]$snap)
                        }
                    }
                }
            } catch {}
        }

        try { $range.Clear() | Out-Null } catch {}

        foreach ($p in $preserved) {
            try {
                $mrng = $ws.Range($p.AreaAddr)
                try { if (-not $mrng.MergeCells) { $mrng.Merge() | Out-Null } } catch {}
                if ($null -ne $p.Value) { try { $mrng.Cells.Item(1, 1).Value2 = $p.Value } catch {} }
                if ($null -ne $p.ColorIndex -and $p.ColorIndex -ne -4142 -and $null -ne $p.Color) {
                    try { $mrng.Interior.Color = $p.Color } catch {}
                }
            } catch {}
        }
    }
}

# -- Anchor row math -----------------------------------------

function Get-MaxSheetRow($ws) {
    <#
    Excel's hard row limit for this worksheet (1,048,576 on .xlsx/.xlsm,
    65,536 on legacy .xls). Used as the anchor-scan ceiling so the row math
    never caps prematurely. An evidence sheet with 10+ correl sections pushes
    the trailing NoGfix block well past a few thousand rows; a fixed 2000-row
    ceiling made every picture beyond that collapse onto the same anchor
    (overlapping images + overwritten id labels).
    #>
    $max = 1048576
    try {
        $c = [int]$ws.Rows.Count
        if ($c -gt 0) { $max = $c }
    } catch {}
    return $max
}

function Get-RowAtOrBelow($ws, [double]$targetTop, [int]$startRow = 1, [int]$maxScanRows = 0) {
    <#
    Linear scan: returns the first row r in [$startRow, ceiling] such that
    ws.Cells(r,1).Top >= $targetTop. Cells(r,1).Top is monotonic with r, so
    walking is safe. $maxScanRows <= 0 means "use the worksheet row limit"
    (Get-MaxSheetRow) so tall sheets are never capped; callers approximate
    $startRow from shape.Top/15, so the walk stays a handful of rows long
    even when the target is deep in the sheet.
    #>
    $ceiling = if ($maxScanRows -gt 0) { $maxScanRows } else { Get-MaxSheetRow $ws }
    $r = [Math]::Max(1, $startRow)
    if ($r -gt $ceiling) { return $ceiling }
    while ($r -le $ceiling) {
        $t = 0.0
        try { $t = [double]$ws.Cells.Item($r, 1).Top } catch { return $ceiling }
        if ($t -ge $targetTop) { return $r }
        $r++
    }
    return $ceiling
}

function Get-NextAnchorRow($ws, $shape, [int]$blankRows = 1, [int]$maxScanRows = 0) {
    <#
    Given an inserted Shape, return the row index for the next anchor
    that sits $blankRows rows below the shape's bottom edge.
    #>
    if ($null -eq $shape) { return 1 }
    $bottom = 0.0
    try { $bottom = [double]($shape.Top + $shape.Height) } catch { return 1 }
    # Speed: start scan from an approximated row instead of row 1.
    $startScan = 1
    try { $startScan = [Math]::Max(1, [int]([Math]::Floor([double]$shape.Top / 15.0))) } catch {}
    $rowAfter = Get-RowAtOrBelow $ws $bottom $startScan $maxScanRows
    return ($rowAfter + [Math]::Max(0, $blankRows))
}

function Get-PictureBottomRow($ws, $shape, [int]$maxScanRows = 0) {
    <#
    Returns the last row index that the shape's bottom edge falls within.
    Get-RowAtOrBelow returns the first row R where Top(R) >= picture bottom,
    meaning R is the row AFTER the picture; the last occupied row is R - 1.
    #>
    if ($null -eq $shape) { return 1 }
    $bottom = 0.0
    try { $bottom = [double]($shape.Top + $shape.Height) } catch { return 1 }
    $startScan = 1
    try { $startScan = [Math]::Max(1, [int]([Math]::Floor([double]$shape.Top / 15.0))) } catch {}
    $rowAfter = Get-RowAtOrBelow $ws $bottom $startScan $maxScanRows
    return [Math]::Max(1, $rowAfter - 1)
}

function Get-CellRangeRect($ws, [string]$colRange, [int]$startRow, [int]$endRow) {
    <#
    Returns a hashtable { Left; Top; Width; Height } (all in points) for the
    cell area spanning $colRange (e.g. "AW:BC") and rows $startRow..$endRow.
    Used by Mark.ps1 to place DF red rectangles by cell address rather than
    pixel offsets from the picture corner.
    #>
    $parts    = $colRange -split ':'
    $colStart = $parts[0].Trim()
    $colEnd   = if ($parts.Count -gt 1) { $parts[1].Trim() } else { $colStart }

    $left  = [double]$ws.Columns($colStart).Left
    $right = [double]$ws.Columns($colEnd).Left + [double]$ws.Columns($colEnd).Width
    $top   = [double]$ws.Rows($startRow).Top
    $bot   = [double]$ws.Rows($endRow).Top + [double]$ws.Rows($endRow).Height

    return @{ Left = $left; Top = $top; Width = ($right - $left); Height = ($bot - $top) }
}

# -- Inserts -------------------------------------------------

function Insert-PictureSendToBack($ws, [int]$row, [int]$col, [string]$imgPath) {
    <#
    Insert an image at the top-left of (row, col), native size.
    ZOrder = msoSendToBack (1) so later marks/rectangles stay on top.
    Returns the Shape object.
    #>
    if (-not (Test-Path -LiteralPath $imgPath)) {
        throw ("Image not found: {0}" -f $imgPath)
    }
    $left = [double]$ws.Cells.Item($row, $col).Left
    $top  = [double]$ws.Cells.Item($row, $col).Top
    # AddPicture(Filename, LinkToFile=0, SaveWithDoc=-1, Left, Top, Width=-1, Height=-1)
    $pic = $ws.Shapes.AddPicture($imgPath, 0, -1, $left, $top, -1, -1)
    try { $pic.ZOrder(1) | Out-Null } catch {}  # msoSendToBack
    return $pic
}

function Write-PlainText($ws, [int]$row, [int]$col, [string]$text) {
    <#
    Write a plain-text label without bold / color / fill.
    Used for the NoGfix label and
    log-section labels.
    #>
    $cell = $ws.Cells.Item($row, $col)
    $cell.Value2 = $text
    try {
        $cell.Font.Bold = $false
        $cell.Font.ColorIndex = 1
        $cell.Interior.ColorIndex = -4142  # xlColorIndexNone
    } catch {}
}

function Write-LogLines($ws, [int]$startRow, [int]$col, [string[]]$lines) {
    <#
    Paste each line in $lines into separate rows starting at $startRow.
    Returns the row index immediately after the last pasted line.
    #>
    if ($null -eq $lines -or $lines.Count -eq 0) { return $startRow }
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $cell = $ws.Cells.Item($startRow + $i, $col)
        $cell.Value2 = $lines[$i]
        try { $cell.Font.Bold = $false } catch {}
    }
    return ($startRow + $lines.Count)
}

# -- Bitmask helpers (mapping-side; used by callers) ---------

function Get-BitValue($row, [string]$field) {
    if ($null -eq $row) { return 0 }
    if (-not ($row.PSObject.Properties.Name -contains $field)) { return 0 }
    $v = 0
    try { $v = [int]$row.$field } catch { $v = 0 }
    return $v
}

function Set-BitValue($row, [string]$field, [int]$bit) {
    if ($null -eq $row) { return }
    if (-not ($row.PSObject.Properties.Name -contains $field)) {
        $row | Add-Member -NotePropertyName $field -NotePropertyValue '0' -Force
    }
    $cur = Get-BitValue $row $field
    $new = $cur -bor $bit
    $row.$field = [string]$new
}

function Ensure-Column([array]$rows, [string]$field, [string]$default = '0') {
    foreach ($r in $rows) {
        if (-not ($r.PSObject.Properties.Name -contains $field)) {
            $r | Add-Member -NotePropertyName $field -NotePropertyValue $default -Force
        }
    }
}

# -- Shape metadata (AltText payload "v1|<key>|<value>") -----

function Set-ShapeMetadata($shape, [string]$key, [string]$value) {
    <#
    Stamps a small metadata payload on a Shape's AlternativeText so later
    phases (Mark) can identify what the picture represents.
    Format: "v1|<key>|<value>"   e.g. "v1|GIFT_HM|JIDSC48S"
    #>
    if ($null -eq $shape) { return }
    $payload = "v1|{0}|{1}" -f $key, $value
    try { $shape.AlternativeText = $payload } catch {}
}

function Get-ShapeMetadata($shape) {
    if ($null -eq $shape) { return $null }
    $t = $null
    try { $t = [string]$shape.AlternativeText } catch { return $null }
    if ([string]::IsNullOrWhiteSpace($t)) { return $null }
    if (-not $t.StartsWith('v1|')) { return $null }
    $rest = $t.Substring(3)
    $parts = $rest -split '\|', 2
    if ($parts.Count -ne 2) { return $null }
    return @{ Key = $parts[0]; Value = $parts[1] }
}

# -- Red rectangle helper (Mark phase) -----------------------

function Add-RedRectangle($ws, [double]$left, [double]$top, [double]$width, [double]$height, [string]$name, [double]$lineWeight = 1.5) {
    <#
    Draws a hollow rectangle with red border at absolute (left, top) on $ws.
    Width/Height in points. Name is the Shape.Name (used for cleanup).
    Returns the Shape.
    #>
    # msoShapeRectangle = 1
    $shape = $ws.Shapes.AddShape(1, $left, $top, $width, $height)
    try { $shape.Fill.Visible = 0 } catch {}              # msoFalse
    try { $shape.Line.Visible = -1 } catch {}             # msoTrue
    try { $shape.Line.ForeColor.RGB = 255 } catch {}      # red (0x0000FF)
    try { $shape.Line.Weight = $lineWeight } catch {}
    try { $shape.Name = $name } catch {}
    try { $shape.ZOrder(0) | Out-Null } catch {}          # msoBringToFront
    return $shape
}

function Set-CellRangeFill($ws, [int]$row, [int]$colStart, [int]$colEnd, [long]$oleColor) {
    <#
    Fill the interior of cells ($colStart..$colEnd) in $row with $oleColor.
    oleColor is Excel OLE format: R + (G * 256) + (B * 65536).
    Yellow RGB(255,255,0) = 65535.  None (clear) = -4142.
    #>
    $range = $ws.Range($ws.Cells.Item($row, $colStart), $ws.Cells.Item($row, $colEnd))
    try { $range.Interior.Color = $oleColor } catch {}
}

function Remove-MarkShapes($ws, [string]$namePrefix) {
    <#
    Deletes every shape on $ws whose Name starts with $namePrefix.
    Returns the count deleted. Used by Mark phase for idempotent re-runs.
    #>
    if ($null -eq $ws -or [string]::IsNullOrWhiteSpace($namePrefix)) { return 0 }
    $cnt = 0
    $toDel = New-Object System.Collections.Generic.List[string]
    foreach ($s in $ws.Shapes) {
        try {
            if ([string]$s.Name -like ($namePrefix + '*')) { $toDel.Add($s.Name) }
        } catch {}
    }
    foreach ($n in $toDel) {
        try { $ws.Shapes.Item($n).Delete() | Out-Null; $cnt++ } catch {}
    }
    return $cnt
}

# -- GFIX log highlight (folded into MarkGfix) ---------------

function Invoke-GfixLogHighlight {
    <#
    Highlights the GFIX-log "Command:" row in a GFIX receive sheet. For each
    $LogAnchor cell in column B, finds the first row in that region whose B
    cell matches $CommandPattern and fills $ColStart..$ColEnd with
    $HighlightColor. Any prior matching fill in the region is cleared first,
    so re-runs are idempotent.

    Returns a hashtable: @{ Applied=<int>; Anchors=<int>; Ok=<bool>; Warnings=<string[]> }
    COM-only; the caller supplies an already-open worksheet.
    #>
    param(
        $ws,
        [string]$LogAnchor,
        [string]$CommandPattern = "Command:\s*'/appl/[A-Za-z0-9]+/shell/",
        [long]$HighlightColor = 65535,
        [int]$ColStart = 2,
        [int]$ColEnd   = 51
    )
    $warnings = New-Object System.Collections.Generic.List[string]
    $applied  = 0
    if ($null -eq $ws) { return @{ Applied = 0; Anchors = 0; Ok = $false; Warnings = @('worksheet is null') } }

    $xlUp = -4162
    $lastRow = 0
    try { $lastRow = [int]$ws.Cells.Item($ws.Rows.Count, 2).End($xlUp).Row } catch { $lastRow = 0 }
    if ($lastRow -lt 1) {
        try { $used = $ws.UsedRange; $lastRow = [int]($used.Row + $used.Rows.Count - 1) } catch { $lastRow = 200 }
    }

    $anchorRows = @()
    for ($r = 1; $r -le $lastRow; $r++) {
        $v = $null
        try { $v = [string]$ws.Cells.Item($r, 2).Value2 } catch {}
        if ($v -eq $LogAnchor) { $anchorRows += $r }
    }
    if ($anchorRows.Count -eq 0) {
        $warnings.Add(("no '{0}' anchors found in sheet" -f $LogAnchor))
        return @{ Applied = 0; Anchors = 0; Ok = $false; Warnings = $warnings.ToArray() }
    }

    $ok = $true
    for ($ai = 0; $ai -lt $anchorRows.Count; $ai++) {
        $regionStart = $anchorRows[$ai] + 1
        $regionEnd   = if ($ai + 1 -lt $anchorRows.Count) { $anchorRows[$ai + 1] - 1 } else { $lastRow }

        # Clear previous yellow fills in this region (idempotent re-run).
        for ($r = $regionStart; $r -le $regionEnd; $r++) {
            $existFill = -1
            try { $existFill = [long]$ws.Cells.Item($r, $ColStart).Interior.Color } catch {}
            if ($existFill -eq $HighlightColor) { Set-CellRangeFill $ws $r $ColStart $ColEnd -4142 }
        }

        # Find the Command: row.
        $targetRow = -1; $matchCount = 0
        for ($r = $regionStart; $r -le $regionEnd; $r++) {
            $v = $null
            try { $v = [string]$ws.Cells.Item($r, 2).Value2 } catch {}
            if (-not [string]::IsNullOrWhiteSpace($v) -and ($v -match $CommandPattern)) {
                if ($matchCount -eq 0) { $targetRow = $r }
                $matchCount++
            }
        }
        if ($targetRow -lt 0) {
            $warnings.Add(("anchor row {0}: no Command: match in region {1}..{2}" -f $anchorRows[$ai], $regionStart, $regionEnd))
            $ok = $false
            continue
        }
        if ($matchCount -gt 1) {
            $warnings.Add(("anchor row {0}: {1} Command: matches; using first (row {2})" -f $anchorRows[$ai], $matchCount, $targetRow))
        }
        Set-CellRangeFill $ws $targetRow $ColStart $ColEnd $HighlightColor
        $applied++
    }

    return @{ Applied = $applied; Anchors = $anchorRows.Count; Ok = ($ok -and $applied -gt 0); Warnings = $warnings.ToArray() }
}
