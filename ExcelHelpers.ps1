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

function Insert-PictureAtPointBringToFront($ws, [double]$left, [double]$top, [string]$imgPath) {
    <#
    Insert an image at absolute sheet (left, top) points, native size,
    ZOrder = msoBringToFront (0). Same shape as Add-RedRectangle (raw point
    coordinates, e.g. an image-match hit already scaled to sheet space) but
    for a whole picture instead of a rectangle -- used by Mark.ps1's
    Mark.Boxes 'StampImage' (image-recognition-placed stamp, no cell lookup).
    Returns the Shape object.
    #>
    if (-not (Test-Path -LiteralPath $imgPath)) {
        throw ("Image not found: {0}" -f $imgPath)
    }
    $pic = $ws.Shapes.AddPicture($imgPath, 0, -1, $left, $top, -1, -1)
    try { $pic.ZOrder(0) | Out-Null } catch {}  # msoBringToFront
    return $pic
}

function Insert-PictureBringToFront($ws, [int]$row, [int]$col, [string]$imgPath) {
    <#
    Insert an image at the top-left of (row, col), native size.
    ZOrder = msoBringToFront (0) -- unlike Insert-PictureSendToBack, this is
    for annotation stamps (e.g. Mark.ps1's verifyNote stamp overlay) that must
    sit visibly on top of the base screenshot and any red-rectangle marks.
    Returns the Shape object.
    #>
    if (-not (Test-Path -LiteralPath $imgPath)) {
        throw ("Image not found: {0}" -f $imgPath)
    }
    $left = [double]$ws.Cells.Item($row, $col).Left
    $top  = [double]$ws.Cells.Item($row, $col).Top
    $pic = $ws.Shapes.AddPicture($imgPath, 0, -1, $left, $top, -1, -1)
    try { $pic.ZOrder(0) | Out-Null } catch {}  # msoBringToFront
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

function Write-LogLines($ws, [int]$startRow, [int]$col, [string[]]$lines, [string]$FontName = '', [double]$FontSize = 0) {
    <#
    Paste each line in $lines into separate rows starting at $startRow.
    When $FontName is non-blank, every pasted line's font is forced to it
    (used to paste the GFIX receive log in a fixed-width font, e.g. 'MS
    Gothic', regardless of the workbook's default font). When $FontSize is
    > 0, the font size is forced too (e.g. 11), so the pasted log renders
    identically no matter what the workbook's default style is.
    Returns the row index immediately after the last pasted line.
    #>
    if ($null -eq $lines -or $lines.Count -eq 0) { return $startRow }
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $cell = $ws.Cells.Item($startRow + $i, $col)
        $cell.Value2 = $lines[$i]
        try { $cell.Font.Bold = $false } catch {}
        if (-not [string]::IsNullOrWhiteSpace($FontName)) {
            try { $cell.Font.Name = $FontName } catch {}
        }
        if ($FontSize -gt 0) {
            try { $cell.Font.Size = $FontSize } catch {}
        }
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

function Get-TextPixelWidth {
    <#
    PURE-ish (GDI+ only, no COM): measures $Text's rendered width in pixels
    for the given font via System.Drawing.Graphics.MeasureString. Caller must
    have already loaded System.Drawing (Add-Type -AssemblyName System.Drawing)
    -- this file makes no Add-Type calls itself (dot-source safety rule: no
    param() at file scope, no side-loading).
    #>
    param(
        [string]$Text,
        [string]$FontName = 'Calibri',
        [double]$FontSize = 11,
        [bool]$Bold = $false
    )
    if ([string]::IsNullOrEmpty($Text)) { return 0.0 }
    if ([string]::IsNullOrWhiteSpace($FontName)) { $FontName = 'Calibri' }
    if ($FontSize -le 0) { $FontSize = 11 }
    $style = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $font = $null; $bmp = $null; $gfx = $null; $fmt = $null
    try {
        $font = New-Object System.Drawing.Font($FontName, [float]$FontSize, $style)
        $bmp  = New-Object System.Drawing.Bitmap(1, 1)
        # A fresh Bitmap inherits the PROCESS'S SCREEN DPI (e.g. 120/144 on a
        # 125%/150%-scaled laptop; powershell.exe is DPI-aware), which used to
        # inflate the measured pixel width by that scale factor while callers
        # converted with a fixed 96-DPI assumption (px * 0.75) -- the GFIX
        # highlight ran 25-50% too long on scaled displays. Pin the bitmap to
        # 96 DPI so the returned pixels are always 96-DPI-based and the
        # documented px->points conversion (x 0.75) is exact everywhere.
        try { $bmp.SetResolution(96, 96) } catch {}
        $gfx  = [System.Drawing.Graphics]::FromImage($bmp)
        # Plain MeasureString pads the result with layout margins (roughly an
        # em of slack), which made the auto-sized GFIX highlight run several
        # grid columns LONGER than the text on narrow-column evidence sheets.
        # GenericTypographic measures the actual glyph advance width;
        # MeasureTrailingSpaces keeps any trailing blanks in the log line.
        $fmt = New-Object System.Drawing.StringFormat([System.Drawing.StringFormat]::GenericTypographic)
        $fmt.FormatFlags = $fmt.FormatFlags -bor [System.Drawing.StringFormatFlags]::MeasureTrailingSpaces
        $size = $gfx.MeasureString($Text, $font, [int]::MaxValue, $fmt)
        return [double]$size.Width
    } finally {
        if ($null -ne $fmt)  { $fmt.Dispose() }
        if ($null -ne $gfx)  { $gfx.Dispose() }
        if ($null -ne $bmp)  { $bmp.Dispose() }
        if ($null -ne $font) { $font.Dispose() }
    }
}

function Get-TextCellUnits {
    <#
    PURE (no COM/GDI): counts the text's width in character-cell units the
    way fixed-pitch CJK fonts (MS Gothic etc.) lay glyphs out -- a half-width
    character (ASCII, halfwidth katakana) advances 0.5 em, a full-width one
    advances 1.0 em. Multiplying by the font size in points gives the ideal
    (unhinted) rendered width in points, which is used as a LOWER BOUND for
    the GFIX highlight so a failed or undershooting pixel measurement can
    never leave the highlight shorter than the text. Unit-tested in
    Tests\Test-ExcelHelpers.ps1.
    #>
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return 0.0 }
    $units = 0.0
    foreach ($ch in $Text.ToCharArray()) {
        $cp = [int]$ch
        if ($cp -le 0xFF -or ($cp -ge 0xFF61 -and $cp -le 0xFF9F)) { $units += 0.5 }
        else { $units += 1.0 }
    }
    return $units
}

function Get-TextGdiWidthSample {
    <#
    GDI measurement of $Text's rendered width via System.Windows.Forms.
    TextRenderer (NoPadding|SingleLine) -- the same renderer Excel draws cell
    text with, so hinted/bitmap advance widths (MS Gothic etc.) are included.
    Returns @{ Pixels; Dpi } (pixels are at the process's screen DPI; convert
    with 72/Dpi). Throws when System.Drawing / System.Windows.Forms are not
    usable -- deliberately kept as a SEPARATE function because PowerShell
    resolves the [System.Drawing.*]/[System.Windows.Forms.*] type literals
    when it COMPILES the containing function, so keeping them out of
    Get-TextPointWidthInfo lets that orchestrator (and its char-cell floor
    tier) still run on a host where GDI is unavailable; the call site's
    try/catch turns this function's compile/run failure into a tier skip.
    #>
    param(
        [string]$Text,
        [string]$FontName,
        [double]$FontSize,
        [bool]$Bold = $false
    )
    $style = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $dpi = 96.0
    $g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
    try { $dpi = [double]$g.DpiX } finally { $g.Dispose() }
    if ($dpi -le 0) { $dpi = 96.0 }
    $font = New-Object System.Drawing.Font($FontName, [float]$FontSize, $style)
    try {
        $flags = [System.Windows.Forms.TextFormatFlags]::NoPadding -bor [System.Windows.Forms.TextFormatFlags]::SingleLine
        $proposed = New-Object System.Drawing.Size([int]::MaxValue, [int]::MaxValue)
        $sz = [System.Windows.Forms.TextRenderer]::MeasureText($Text, $font, $proposed, $flags)
        return @{ Pixels = [double]$sz.Width; Dpi = $dpi }
    } finally { $font.Dispose() }
}

function Get-TextPointWidthInfo {
    <#
    Measures $Text's rendered width in POINTS for the given font, trying the
    renderer that matches how Excel actually draws cell text first:

      1. GDI (Get-TextGdiWidthSample: System.Windows.Forms.TextRenderer,
         NoPadding) -- Excel renders cells with GDI, whose hinted/bitmap
         advance widths for classic Japanese fonts like MS Gothic run WIDER
         than the ideal typographic advance (e.g. 8 px vs 7.33 px per
         half-width char at 11pt/96dpi). Measuring with GDI+
         GenericTypographic (the old only path) therefore undershot ~8% on a
         long Command: line and the highlight ended before the text did.
         Pixels are converted with the REAL screen DPI, not a hardcoded 96.
      2. GDI+ (Get-TextPixelWidth, 96-DPI-pinned) when System.Windows.Forms
         is not loaded / GDI measurement fails.

    Whatever tier answered, the result is floored at the ideal fixed-pitch
    character-cell width (Get-TextCellUnits x $FontSize) so the highlight
    always covers AT LEAST the text's nominal advance -- "mark enough cells";
    the caller's HighlightColEnd cap still bounds it above. When both
    measurement tiers fail (no GDI at all), the floor alone answers. Returns
    a hashtable with the chosen width plus every intermediate value so
    callers can print a diagnostic line:
      @{ Points; Source ('gdi'/'gdiplus'/'floor'/'none', '+floor' suffix when
         the floor won); Dpi; Pixels; FloorPoints; CellUnits }
    This function contains no System.Drawing/System.Windows.Forms type
    literals itself (see Get-TextGdiWidthSample's note), so it is safe to
    call -- and its floor tier keeps working -- even where GDI is missing.
    #>
    param(
        [string]$Text,
        [string]$FontName = 'Calibri',
        [double]$FontSize = 11,
        [bool]$Bold = $false
    )
    if ([string]::IsNullOrWhiteSpace($FontName)) { $FontName = 'Calibri' }
    if ($FontSize -le 0) { $FontSize = 11 }
    $units = [double](Get-TextCellUnits -Text $Text)
    $floorPoints = $units * $FontSize
    $info = @{ Points = 0.0; Source = 'none'; Dpi = 96.0; Pixels = 0.0; FloorPoints = $floorPoints; CellUnits = $units }
    if ([string]::IsNullOrEmpty($Text)) { return $info }

    # Tier 1: GDI (matches Excel's renderer). MeasureText pixels are at the
    # process's screen DPI, so convert with that same DPI.
    try {
        $s = Get-TextGdiWidthSample -Text $Text -FontName $FontName -FontSize $FontSize -Bold $Bold
        if ($null -ne $s -and $s.Pixels -gt 0) {
            $info.Pixels = [double]$s.Pixels
            $info.Dpi    = [double]$s.Dpi
            $info.Points = [double]$s.Pixels * 72.0 / [double]$s.Dpi
            $info.Source = 'gdi'
        }
    } catch {}

    # Tier 2: GDI+ (96-DPI-pinned by Get-TextPixelWidth, so x 0.75 is exact).
    if ($info.Points -le 0) {
        try {
            $px = Get-TextPixelWidth -Text $Text -FontName $FontName -FontSize $FontSize -Bold $Bold
            if ($px -gt 0) {
                $info.Pixels = [double]$px
                $info.Points = [double]$px * 0.75
                $info.Source = 'gdiplus'
            }
        } catch {}
    }

    # Floor: never narrower than the ideal fixed-pitch character-cell width.
    if ($floorPoints -gt 0 -and $floorPoints -gt $info.Points) {
        if ($info.Points -gt 0) { $info.Source = $info.Source + '+floor' } else { $info.Source = 'floor' }
        $info.Points = $floorPoints
    }
    return $info
}

function Get-ColumnsForWidth {
    <#
    PURE (no COM/GDI): given an ordered array of column widths (points, one
    per column starting at the highlight's ColStart) and a target width in
    points, returns the 0-based offset from ColStart of the last column
    needed so the cumulative width covers $NeededPoints. Split out from
    Get-AutoHighlightColEnd so the column-accumulation math is unit-testable
    without Excel/GDI+ (Tests\Test-ExcelHelpers.ps1).
    #>
    param(
        [double[]]$ColumnWidths,
        [double]$NeededPoints
    )
    if ($NeededPoints -le 0 -or $null -eq $ColumnWidths -or $ColumnWidths.Count -eq 0) { return 0 }
    $acc = 0.0
    for ($i = 0; $i -lt $ColumnWidths.Count; $i++) {
        $acc += [double]$ColumnWidths[$i]
        if ($acc -ge $NeededPoints) { return $i }
    }
    return ($ColumnWidths.Count - 1)
}

function Get-AutoHighlightColEnd {
    <#
    Determines how many columns from $ColStart the yellow highlight needs to
    cover the ACTUAL pasted text on $Row (the GFIX log 'Command:' row),
    instead of always filling out to a fixed width -- a long Command: path
    could run past a short fixed range, and a short one leaves an
    unnecessarily wide highlight. Measures the row's text with its real cell
    font (GDI+) and walks the sheet's actual column widths to find where that
    rendered width lands, then pads by $PadCols. Never narrower than
    $ColStart, never wider than $MaxColEnd (also the fixed legacy default --
    used as an upper bound so this only ever tightens the old behavior).
    Falls back to $MaxColEnd on any read/measurement problem.
    $FontName/$FontSize (optional) override the cell-font read for the
    measurement: pass the font ReplaceGfix forced onto the log paste
    (Replace.GfixLogFontName/GfixLogFontSize) so the measured width always
    matches the pasted font even when the cell read fails or the log was
    pasted before the font forcing existed.
    $DiagList (optional, List[string]): every decision/fallback appends a
    human-readable line so a wrong width on an office PC is diagnosable from
    the console instead of silently indistinguishable from AutoWidth=off.
    #>
    param(
        $ws,
        [int]$Row,
        [int]$ColStart,
        [int]$MaxColEnd,
        [int]$PadCols = 1,
        [string]$FontName = '',
        [double]$FontSize = 0,
        [System.Collections.Generic.List[string]]$DiagList = $null
    )
    if ($MaxColEnd -lt $ColStart) { return $MaxColEnd }
    $text = $null
    try { $text = [string]$ws.Cells.Item($Row, $ColStart).Value2 } catch {}
    if ([string]::IsNullOrEmpty($text)) {
        if ($null -ne $DiagList) { $DiagList.Add(("row {0}: no text in col {1} cell; using fixed ColEnd {2}" -f $Row, $ColStart, $MaxColEnd)) }
        return $MaxColEnd
    }

    # NOTE: locals deliberately NOT named $fontName/$fontSize -- PowerShell
    # variables are case-insensitive, so those would collide with the
    # $FontName/$FontSize override parameters.
    $measureFont = 'Calibri'; $measureSize = 11.0; $measureBold = $false
    try {
        $cell = $ws.Cells.Item($Row, $ColStart)
        try { $measureFont = [string]$cell.Font.Name } catch {}
        try { $measureSize = [double]$cell.Font.Size } catch {}
        try { $measureBold = [bool]$cell.Font.Bold } catch {}
    } catch {}
    if (-not [string]::IsNullOrWhiteSpace($FontName)) { $measureFont = $FontName }
    if ($FontSize -gt 0) { $measureSize = $FontSize }

    $mi = $null
    try { $mi = Get-TextPointWidthInfo -Text $text -FontName $measureFont -FontSize $measureSize -Bold $measureBold } catch {}
    if ($null -eq $mi -or $mi.Points -le 0) {
        if ($null -ne $DiagList) { $DiagList.Add(("row {0}: width measurement failed ({1} chars, font '{2}' {3}pt); using fixed ColEnd {4}" -f $Row, $text.Length, $measureFont, $measureSize, $MaxColEnd)) }
        return $MaxColEnd
    }
    $neededPoints = [double]$mi.Points

    $widths = New-Object System.Collections.Generic.List[double]
    $comMiss = 0
    for ($c = $ColStart; $c -le $MaxColEnd; $c++) {
        $w = 48.0   # ~8.43-char default column width in points, used only if the COM read below fails
        try { $w = [double]$ws.Columns.Item($c).Width } catch { $comMiss++ }
        $widths.Add($w)
    }
    $offset = Get-ColumnsForWidth -ColumnWidths $widths.ToArray() -NeededPoints $neededPoints
    $colEnd = $ColStart + $offset + [Math]::Max(0, $PadCols)
    if ($colEnd -gt $MaxColEnd) { $colEnd = $MaxColEnd }
    if ($colEnd -lt $ColStart)  { $colEnd = $ColStart }
    if ($null -ne $DiagList) {
        $miss = if ($comMiss -gt 0) { (" ({0} column-width reads failed, default 48pt used)" -f $comMiss) } else { '' }
        $DiagList.Add(("row {0}: {1} chars, font '{2}' {3}pt bold={4} -> {5:0.0}pt via {6} (dpi={7}, px={8:0.0}, cell-floor={9:0.0}pt) -> cols {10}..{11} (+{12} pad, cap {13}){14}" -f `
            $Row, $text.Length, $measureFont, $measureSize, $measureBold, $neededPoints, $mi.Source, $mi.Dpi, $mi.Pixels, $mi.FloorPoints, $ColStart, $colEnd, [Math]::Max(0, $PadCols), $MaxColEnd, $miss))
    }
    return $colEnd
}

function Invoke-GfixLogHighlight {
    <#
    Highlights the GFIX-log "Command:" row in a GFIX receive sheet. For each
    $LogAnchor cell in column B, finds the first row in that region whose B
    cell matches $CommandPattern and fills $ColStart..<end> with
    $HighlightColor. When $AutoWidth is set, <end> is computed per row from
    the row's actual text width (Get-AutoHighlightColEnd), capped at
    $ColEnd; otherwise <end> is always the fixed $ColEnd (legacy behavior).
    Any prior matching fill in the full $ColStart..$ColEnd region is cleared
    first (regardless of AutoWidth), so re-runs are idempotent even after
    the computed width changes between runs.
    $FontName/$FontSize (optional) are the font the log was PASTED in
    (Replace.GfixLogFontName/GfixLogFontSize); when set, the AutoWidth
    measurement uses them instead of trusting the cell-font read.

    Returns a hashtable: @{ Applied=<int>; Anchors=<int>; Ok=<bool>; Warnings=<string[]>; Diag=<string[]> }
    (Diag: one AutoWidth measurement/fallback line per highlighted row --
    callers print them so a wrong width is diagnosable from the console.)
    COM-only; the caller supplies an already-open worksheet.
    #>
    param(
        $ws,
        [string]$LogAnchor,
        [string]$CommandPattern = "Command:\s*'/appl/[A-Za-z0-9]+/shell/",
        [long]$HighlightColor = 65535,
        [int]$ColStart = 2,
        [int]$ColEnd   = 51,
        [bool]$AutoWidth = $true,
        [int]$PadCols = 1,
        [string]$FontName = '',
        [double]$FontSize = 0
    )
    $warnings = New-Object System.Collections.Generic.List[string]
    $diag     = New-Object System.Collections.Generic.List[string]
    $applied  = 0
    if ($null -eq $ws) { return @{ Applied = 0; Anchors = 0; Ok = $false; Warnings = @('worksheet is null'); Diag = @() } }

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
        return @{ Applied = 0; Anchors = 0; Ok = $false; Warnings = $warnings.ToArray(); Diag = @() }
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
        $rowColEnd = $ColEnd
        if ($AutoWidth) {
            try { $rowColEnd = Get-AutoHighlightColEnd -ws $ws -Row $targetRow -ColStart $ColStart -MaxColEnd $ColEnd -PadCols $PadCols -FontName $FontName -FontSize $FontSize -DiagList $diag }
            catch { $rowColEnd = $ColEnd; $diag.Add(("row {0}: Get-AutoHighlightColEnd threw ({1}); using fixed ColEnd {2}" -f $targetRow, $_.Exception.Message, $ColEnd)) }
        }
        Set-CellRangeFill $ws $targetRow $ColStart $rowColEnd $HighlightColor
        $applied++
    }

    return @{ Applied = $applied; Anchors = $anchorRows.Count; Ok = ($ok -and $applied -gt 0); Warnings = $warnings.ToArray(); Diag = $diag.ToArray() }
}
