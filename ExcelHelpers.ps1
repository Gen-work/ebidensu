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

# ── Excel application lifecycle ─────────────────────────────

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

# ── Sheet operations ────────────────────────────────────────

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

# ── Cleanup helpers ─────────────────────────────────────────

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
        try { $range.Clear() | Out-Null } catch {}
    }
}

# ── Anchor row math ─────────────────────────────────────────

function Get-RowAtOrBelow($ws, [double]$targetTop, [int]$startRow = 1, [int]$maxScanRows = 2000) {
    <#
    Linear scan: returns the first row r in [$startRow, $maxScanRows]
    such that ws.Cells(r,1).Top >= $targetTop.
    Cells(r,1).Top is monotonic with r, so walking is safe.
    #>
    $r = [Math]::Max(1, $startRow)
    while ($r -le $maxScanRows) {
        $t = 0.0
        try { $t = [double]$ws.Cells.Item($r, 1).Top } catch { return $maxScanRows }
        if ($t -ge $targetTop) { return $r }
        $r++
    }
    return $maxScanRows
}

function Get-NextAnchorRow($ws, $shape, [int]$blankRows = 1, [int]$maxScanRows = 2000) {
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

# ── Inserts ─────────────────────────────────────────────────

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
    Used for "GFIX Jenkins フォルダ受信ファイルなし" and
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

# ── Bitmask helpers (mapping-side; used by callers) ─────────

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

# ── Shape metadata (AltText payload "v1|<key>|<value>") ─────

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

# ── Red rectangle helper (Mark phase) ───────────────────────

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
