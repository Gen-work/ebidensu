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

function Get-EbiSaveDialogs {
    <#
      PURE. The dialog titles a save answers with Esc: the step's own
      dismissDialogs when given, else the profile's (layout.saveDialogs,
      then layout.evidence.saveDialogs) -- 10/20: no save step may forget
      the classification dialog.
    #>
    param($Given, $Profile)
    $g = @(@($Given) | Where-Object { $null -ne $_ -and [string]$_ -ne '' } | ForEach-Object { [string]$_ })
    if ($g.Count -gt 0) { return $g }
    try {
        $lay = if ($Profile -is [System.Collections.IDictionary] -and $Profile.Contains('layout')) { $Profile['layout'] } else { $null }
        if ($lay -is [System.Collections.IDictionary]) {
            if ($lay.Contains('saveDialogs')) { return @(@($lay['saveDialogs']) | Where-Object { [string]$_ -ne '' } | ForEach-Object { [string]$_ }) }
            if ($lay.Contains('evidence') -and $lay['evidence'] -is [System.Collections.IDictionary] -and $lay['evidence'].Contains('saveDialogs')) { return @(@($lay['evidence']['saveDialogs']) | Where-Object { [string]$_ -ne '' } | ForEach-Object { [string]$_ }) }
        }
    } catch { }
    return @()
}

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
    # by name straight away, then a walk of the sheets; a few tries apart:
    # 10/21 QJODWDBC, right after 154 lines were written to a sheet the next
    # step was told "no sheet" -- Excel was still busy (inferred), and a
    # busy Excel can hand back an empty answer instead of an error
    for ($try = 0; $try -lt 4; $try++) {
        try { $ws = $Workbook.Worksheets.Item($s); if ($null -ne $ws -and [string]$ws.Name -eq $s) { return $ws } } catch { }
        try { foreach ($ws in $Workbook.Worksheets) { if ([string]$ws.Name -eq $s) { return $ws } } } catch { }
        if ($try -lt 3) { Start-Sleep -Milliseconds 500 }
    }
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

function Show-EbiExcelAt {
    <#
      Let the operator SEE what the run is doing to a visible workbook
      (10/15: "switch to it, so I can q any time, not wait blind"):
      activate the sheet, scroll so -Row is near the top, and -- with
      -Front -- bring Excel in front. Does nothing when Excel is hidden
      (the plan reads books with visible = false). Best effort, never
      throws: a refused call here must not fail the step that wrote.
      Looking is fine; CLICKING in Excel while the run writes makes Excel
      refuse the next call (0x800AC472, docs/ebi-dance/KNOWLEDGE.md).
      -Dwell (10/19: "the camera does not follow the inserts"): called
      again AFTER a write, at what was written, and held there for
      Get-EbiWatchDwellMs so a person can see it before the next step.
    #>
    param($Worksheet, [int]$Row = 0, [switch]$Front, [switch]$Dwell)
    $app = $null
    try { $app = $Worksheet.Application; if (-not [bool]$app.Visible) { return } } catch { return }
    try { [void]$Worksheet.Activate() } catch { }
    if ($Row -gt 0) { try { $app.ActiveWindow.ScrollRow = [Math]::Max(1, $Row - 2) } catch { } }
    if ($Front -and [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT -and $null -ne (Get-Command -Name 'Set-EbiForeground' -ErrorAction SilentlyContinue)) {
        try { [void](Set-EbiForeground -HWnd ([IntPtr]([int64]$app.Hwnd)) -SettleMs 0) } catch { }
    }
    if ($Dwell) { $ms = Get-EbiWatchDwellMs; if ($ms -gt 0) { Start-Sleep -Milliseconds $ms } }
}

function Get-EbiWatchDwellMs {
    # PURE (reads the environment). How long Show-EbiExcelAt -Dwell holds
    # the view on what was just written: EBI_WATCH_MS, default 700; 0 = off.
    $v = [string]$env:EBI_WATCH_MS
    $n = 0
    if ($v -ne '' -and [int]::TryParse($v, [ref]$n)) { return [Math]::Max(0, $n) }
    return 700
}

function Export-EbiShapePng {
    <#
      One shape -> a PNG, through a temporary chart (Shape.Copy, paste into
      a ChartObject stretched to -Scale x the shape, Chart.Export, delete
      the chart). Scale > 1 re-renders from the picture's ORIGINAL data, so
      text shown small on the sheet is big enough for OCR (EvidenceImage
      Export.ps1's lesson). Clobbers the clipboard; leaves the sheet as it
      was. Returns @{ ok; path; message }; never throws.
    #>
    param($Worksheet, $Shape, [string]$OutPath, [double]$Scale = 3.0)
    $chartObj = $null
    # 10/16: the office run reported "picture not exported to ...pic1.png:"
    # with nothing after the colon -- an exception with an empty Message.
    # Name the stage that failed and the exception's type / HRESULT so the
    # next failure says what to fix.
    $stage = 'prepare'
    try {
        $dir = Split-Path -Parent $OutPath
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        if (Test-Path -LiteralPath $OutPath) { Remove-Item -LiteralPath $OutPath -Force }
        $w0 = [double]$Shape.Width; $h0 = [double]$Shape.Height
        if ($Scale -lt 1.0) { $Scale = 1.0 }
        if ([Math]::Max($w0, $h0) * $Scale -gt 6000.0) { $Scale = 6000.0 / [Math]::Max($w0, $h0) }
        $w = $w0 * $Scale; $h = $h0 * $Scale
        $stage = 'activate sheet'
        try { [void]$Worksheet.Activate() } catch { }   # ChartObject.Activate / Paste are unreliable on a sheet that is not active
        $stage = 'copy shape'
        [void]$Shape.Copy()
        Start-Sleep -Milliseconds 150
        $stage = 'add chart'
        $chartObj = $Worksheet.ChartObjects().Add(0, 0, $w, $h)
        $stage = 'activate chart'
        [void]$chartObj.Activate()
        $stage = 'paste into chart'
        [void]$chartObj.Chart.Paste()
        Start-Sleep -Milliseconds 150
        # the pasted picture lands in Chart.Shapes or, on some Excel builds,
        # only in the legacy Chart.Pictures() (EvidenceImageExport.ps1)
        $stretched = $false
        try { $pz = $chartObj.Chart.Shapes; if ([int]$pz.Count -ge 1) { $p = $pz.Item(1); $p.LockAspectRatio = 0; $p.Left = 0; $p.Top = 0; $p.Width = $w; $p.Height = $h; $stretched = $true } } catch { }
        if (-not $stretched) { try { $pc = $chartObj.Chart.Pictures(); if ([int]$pc.Count -ge 1) { $p = $pc.Item(1); $p.Left = 0; $p.Top = 0; $p.Width = $w; $p.Height = $h } } catch { } }
        $stage = 'export'
        [void]$chartObj.Chart.Export($OutPath, 'PNG')
        if (-not (Test-Path -LiteralPath $OutPath)) { return @{ ok = $false; path = $OutPath; message = 'Chart.Export wrote no file' } }
        return @{ ok = $true; path = $OutPath; message = '' }
    } catch {
        return @{ ok = $false; path = $OutPath; message = (Format-EbiExcelError -Stage $stage -ErrorRecord $_) }
    } finally {
        if ($null -ne $chartObj) { try { [void]$chartObj.Delete() } catch { }; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($chartObj) } catch { } }
    }
}

function Format-EbiExcelError {
    <#
      PURE. One line for a caught COM / PowerShell error: "<stage> failed:
      <message> [<type> 0x<hresult>]". An empty Message (seen 10/16) still
      gives the stage and the type.
    #>
    param([string]$Stage, $ErrorRecord)
    $ex = $null
    if ($null -ne $ErrorRecord) { if ($ErrorRecord -is [System.Exception]) { $ex = $ErrorRecord } elseif ($null -ne $ErrorRecord.Exception) { $ex = $ErrorRecord.Exception } }
    # a method call on a COM object wraps the real error one level down
    if ($null -ne $ex -and $ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
    $msg = ''; $type = ''; $hr = ''
    if ($null -ne $ex) {
        $msg = ([string]$ex.Message).Trim()
        $type = $ex.GetType().Name
        try { if ([int]$ex.HResult -ne 0) { $hr = ('0x{0:X8}' -f [int]$ex.HResult) } } catch { }
    }
    if ($msg -eq '') { $msg = '(no message)' }
    $tail = (@($type, $hr) | Where-Object { $_ -ne '' }) -join ' '
    $line = $Stage + ' failed: ' + $msg
    if ($tail -ne '') { $line += ' [' + $tail + ']' }
    return $line
}

function Get-EbiGroupShapesAbove {
    # Group shapes (msoGroup 6: pictures put together with Ctrl+G) whose top
    # is at or above -Top points: they come BEFORE the first plain picture
    # in reading order, and Get-EbiPictureShapes does not see them, so "the
    # first picture" may not be the one a person sees first. Names, top
    # first. Never throws.
    param($Worksheet, [double]$Top)
    $out = New-Object System.Collections.ArrayList
    try {
        foreach ($s in $Worksheet.Shapes) {
            $t = 0
            try { $t = [int]$s.Type } catch { continue }
            if ($t -ne 6) { continue }
            if ([double]$s.Top -le ($Top + 1.0)) { [void]$out.Add([string]$s.Name) }
        }
    } catch { }
    return $out.ToArray()
}
