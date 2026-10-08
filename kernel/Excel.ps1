#Requires -Version 5.1
# ============================================================
#  kernel/Excel.ps1
#
#  The Excel COM helpers the excel.* steps share (P4). Ported from
#  ExcelHelpers.ps1 (New-ExcelApp / Close-Workbook / Add-RedRectangle /
#  Set-CellRangeFill / Get-NextAnchorRow) with the house rules kept:
#    - Visible = $true BEFORE DisplayAlerts = $false
#    - every cleanup in its own try/catch (a release must survive
#      "nothing to release")
#    - one concern per try/catch, never one bare catch {} swallowing the
#      rest (the v2.15.1 lesson)
#  Dot-source only (no param(), ASCII source). Nothing here creates a COM
#  object until a step calls it, so a DryRun never touches Excel.
#
#    Get-EbiSheet            workbook + name or 1-based index -> sheet / $null
#    Get-EbiRowAtOrBelow     first row whose top is at/below a point
#    Get-EbiPictureShapes    a sheet's pictures, top-to-bottom, left-to-right
#    Add-EbiRedRect          hollow red rectangle (points)
#    Set-EbiRowFill          interior colour over row r, columns a..b
#    Set-EbiCellTextPlain    value + plain font (no bold, black, no fill)
#    Get-EbiSheetNames       names in tab order
#    Invoke-EbiComRelease    ReleaseComObject, swallowing errors
# ============================================================

function Invoke-EbiComRelease {
    param($Object)
    if ($null -eq $Object) { return }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } catch { }
}

function Get-EbiSheet {
    # A worksheet by name (exact) or 1-based index; $null when absent.
    param($Workbook, $Sheet)
    if ($null -eq $Workbook) { return $null }
    $s = [string]$Sheet
    if ($s -match '^\d+$') {
        try { return $Workbook.Worksheets.Item([int]$s) } catch { return $null }
    }
    foreach ($ws in $Workbook.Worksheets) { if ([string]$ws.Name -eq $s) { return $ws } }
    return $null
}

function Get-EbiSheetNames {
    param($Workbook)
    $names = New-Object System.Collections.ArrayList
    try { foreach ($ws in $Workbook.Worksheets) { [void]$names.Add([string]$ws.Name) } } catch { }
    return $names.ToArray()
}

function Get-EbiRowAtOrBelow {
    # The first row (>= StartRow) whose Top is at or below TopPoints - a
    # half-point tolerance. Scans at most MaxScan rows; returns the last
    # scanned row + 1 when none qualifies (never 0).
    param($Worksheet, [double]$TopPoints, [int]$StartRow = 1, [int]$MaxScan = 5000)
    $r = [Math]::Max(1, $StartRow)
    for ($i = 0; $i -lt $MaxScan; $i++) {
        $t = 0.0
        try { $t = [double]$Worksheet.Rows.Item($r).Top } catch { return $r }
        if ($t -ge ($TopPoints - 0.5)) { return $r }
        $r++
    }
    return $r
}

function Get-EbiPictureShapes {
    # Pictures (msoPicture 13, and linked 11) of a sheet as
    # @{ shape; name; top; left; width; height } sorted top, then left.
    param($Worksheet)
    $list = New-Object System.Collections.ArrayList
    try {
        foreach ($s in $Worksheet.Shapes) {
            $t = 0
            try { $t = [int]$s.Type } catch { continue }
            if ($t -ne 13 -and $t -ne 11) { continue }
            [void]$list.Add(@{ shape = $s; name = [string]$s.Name; top = [double]$s.Top; left = [double]$s.Left; width = [double]$s.Width; height = [double]$s.Height })
        }
    } catch { }
    return @($list | Sort-Object -Property @{ Expression = { $_['top'] } }, @{ Expression = { $_['left'] } })
}

function Add-EbiRedRect {
    # Hollow red rectangle at points; Name / AltText stamped when given.
    param($Worksheet, [double]$Left, [double]$Top, [double]$Width, [double]$Height, [string]$Name = '', [double]$Weight = 1.5, [string]$AltText = '')
    $shape = $Worksheet.Shapes.AddShape(1, $Left, $Top, $Width, $Height)   # msoShapeRectangle
    try { $shape.Fill.Visible = 0 } catch { }
    try { $shape.Line.Visible = -1 } catch { }
    try { $shape.Line.ForeColor.RGB = 255 } catch { }                        # red (BGR 0x0000FF)
    try { $shape.Line.Weight = $Weight } catch { }
    if ($Name -ne '') { try { $shape.Name = $Name } catch { } }
    if ($AltText -ne '') { try { $shape.AlternativeText = $AltText } catch { } }
    try { [void]$shape.ZOrder(0) } catch { }                                 # msoBringToFront
    return $shape
}

function Set-EbiRowFill {
    # Interior colour (OLE BGR long; yellow = 65535) over one row's columns.
    param($Worksheet, [int]$Row, [int]$ColStart, [int]$ColEnd, [long]$Color)
    $range = $Worksheet.Range($Worksheet.Cells.Item($Row, $ColStart), $Worksheet.Cells.Item($Row, $ColEnd))
    $range.Interior.Color = $Color
}

function Set-EbiCellTextPlain {
    # Write a value as TEXT (a leading apostrophe-free string Excel will not
    # reinterpret: NumberFormat '@' first) with a plain font.
    param($Cell, [string]$Text, [string]$FontName = '', [double]$FontSize = 0, [bool]$AsText = $true)
    if ($AsText) { try { $Cell.NumberFormat = '@' } catch { } }
    $Cell.Value2 = $Text
    try { $Cell.Font.Bold = $false } catch { }
    if ($FontName -ne '') { try { $Cell.Font.Name = $FontName } catch { } }
    if ($FontSize -gt 0) { try { $Cell.Font.Size = $FontSize } catch { } }
}
