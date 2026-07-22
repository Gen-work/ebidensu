#Requires -Version 5.1
# ============================================================
#  Probe-SheetFormat.ps1   -- UTF-8, NO BOM, ASCII source.
#
#  Formatting probe / calibration aid (same spirit as Probe-Shapes.ps1,
#  but for CELL FORMATTING instead of shapes): dumps every color / font /
#  size / number-format / alignment / width setting of a target workbook
#  or one sheet, so a generated output workbook (e.g. ProcessTime.ps1's
#  Write-ProcessTimeWorkbook block) can be made to match a delivery
#  template exactly without opening Excel and inspecting cells by hand.
#
#  Usage (standalone; has param() -> call via &, NEVER dot-source):
#    powershell -File Probe-SheetFormat.ps1 -Path C:\...\template.xlsx
#    powershell -File Probe-SheetFormat.ps1 -Path ... -Sheet <name> -Json out.json
#
#  Output per sheet:
#    - used-range address + row/column counts
#    - column widths (ColumnWidth units) and distinct row heights
#    - distinct cell-format signatures (NumberFormat / font name+size+
#      bold / font color / interior color / horizontal+vertical alignment
#      / border line style), each with its cell count and sample
#      addresses -- colors are printed as the raw Excel Long (BGR order,
#      the exact value to feed back into Interior.Color / Font.Color).
#
#  Read-only: the workbook is opened read-only and never saved.
#  Rows/columns beyond -MaxRows/-MaxCols are skipped (a template's format
#  inventory lives in its first screenful; the caps keep huge sheets fast).
# ============================================================

param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$Sheet = '',
    [int]$MaxRows = 200,
    [int]$MaxCols = 30,
    [int]$SampleCells = 8,
    [string]$Json = '',
    [string]$ExcelHelpersScript = ''
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$helpersPath = $null
foreach ($c in @($ExcelHelpersScript, (Join-Path $PSScriptRoot 'ExcelHelpers.ps1'))) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).ProviderPath; break
    }
}
if (-not $helpersPath) {
    Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red; exit 1
}
. $helpersPath

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Host ("[ERROR] workbook not found: {0}" -f $Path) -ForegroundColor Red; exit 1
}
$Path = (Resolve-Path -LiteralPath $Path).ProviderPath

# One cell's format signature as a '|'-joined string (grouping key) plus
# the structured fields. Every read is individually guarded: a merged /
# protected / odd cell must never abort the whole probe.
function Get-CellFormatSignature {
    param($Cell)
    $f = [ordered]@{
        NumberFormat = ''; FontName = ''; FontSize = ''; FontBold = ''
        FontColor = ''; InteriorColor = ''; HAlign = ''; VAlign = ''
        BorderLineStyle = ''; MergedCells = ''
    }
    try { $f.NumberFormat  = [string]$Cell.NumberFormat } catch {}
    try { $f.FontName      = [string]$Cell.Font.Name } catch {}
    try { $f.FontSize      = [string]$Cell.Font.Size } catch {}
    try { $f.FontBold      = [string]$Cell.Font.Bold } catch {}
    try { $f.FontColor     = [string]$Cell.Font.Color } catch {}
    try { $f.InteriorColor = [string]$Cell.Interior.Color } catch {}
    try { $f.HAlign        = [string]$Cell.HorizontalAlignment } catch {}
    try { $f.VAlign        = [string]$Cell.VerticalAlignment } catch {}
    try { $f.BorderLineStyle = [string]$Cell.Borders.LineStyle } catch {}   # mixed edges read as null/-4142
    try { $f.MergedCells   = [string]$Cell.MergeCells } catch {}
    return @{
        Key    = (($f.Values | ForEach-Object { [string]$_ }) -join '|')
        Fields = $f
    }
}

$excel = $null
$wb = $null
$report = New-Object System.Collections.Generic.List[object]
try {
    $excel = New-ExcelApp
    $wb = $excel.Workbooks.Open($Path, 0, $true)   # read-only

    $sheets = @()
    foreach ($ws in $wb.Worksheets) {
        if ([string]::IsNullOrWhiteSpace($Sheet) -or [string]$ws.Name -eq $Sheet) { $sheets += ,$ws }
    }
    if ($sheets.Count -eq 0) {
        Write-Host ("[ERROR] sheet '{0}' not found in {1}" -f $Sheet, (Split-Path $Path -Leaf)) -ForegroundColor Red
        exit 1
    }

    Write-Host ''
    Write-Host ("===== Probe-SheetFormat: {0} =====" -f (Split-Path $Path -Leaf)) -ForegroundColor Green

    foreach ($ws in $sheets) {
        $used = $ws.UsedRange
        $rowCount = [int]$used.Rows.Count
        $colCount = [int]$used.Columns.Count
        $firstRow = [int]$used.Row
        $firstCol = [int]$used.Column
        $scanRows = [Math]::Min($rowCount, $MaxRows)
        $scanCols = [Math]::Min($colCount, $MaxCols)

        Write-Host ''
        Write-Host ("----- sheet '{0}' -----" -f [string]$ws.Name) -ForegroundColor Cyan
        Write-Host ("  UsedRange : {0} ({1} row(s) x {2} col(s); scanning {3} x {4})" -f `
            [string]$used.Address($false, $false), $rowCount, $colCount, $scanRows, $scanCols)

        # Column widths + row heights (grouped by distinct value).
        $colWidths = New-Object System.Collections.Generic.List[object]
        for ($c = 0; $c -lt $scanCols; $c++) {
            $colIdx = $firstCol + $c
            $w = ''
            try { $w = [string]$ws.Columns.Item($colIdx).ColumnWidth } catch {}
            $colLetter = ''
            try { $colLetter = ([string]$ws.Cells.Item(1, $colIdx).Address($true, $false)) -replace '\$?\d+$', '' } catch { $colLetter = [string]$colIdx }
            $colWidths.Add([pscustomobject]@{ Column = $colLetter; Width = $w })
        }
        Write-Host ("  Column widths: {0}" -f (($colWidths | ForEach-Object { ('{0}={1}' -f $_.Column, $_.Width) }) -join ' '))

        $heightGroups = @{}
        for ($r = 0; $r -lt $scanRows; $r++) {
            $rowIdx = $firstRow + $r
            $h = ''
            try { $h = [string]$ws.Rows.Item($rowIdx).RowHeight } catch {}
            if (-not $heightGroups.ContainsKey($h)) { $heightGroups[$h] = New-Object System.Collections.Generic.List[int] }
            $heightGroups[$h].Add($rowIdx)
        }
        foreach ($h in $heightGroups.Keys) {
            $rows = $heightGroups[$h]
            $sample = ($rows | Select-Object -First 10) -join ','
            $more = if ($rows.Count -gt 10) { (' (+{0} more)' -f ($rows.Count - 10)) } else { '' }
            Write-Host ("  RowHeight {0}: {1} row(s) -- rows {2}{3}" -f $h, $rows.Count, $sample, $more)
        }

        # Distinct cell-format signatures across the scanned grid.
        $sigs = @{}
        $sigOrder = New-Object System.Collections.Generic.List[string]
        for ($r = 0; $r -lt $scanRows; $r++) {
            for ($c = 0; $c -lt $scanCols; $c++) {
                $cell = $ws.Cells.Item($firstRow + $r, $firstCol + $c)
                $sig = Get-CellFormatSignature $cell
                if (-not $sigs.ContainsKey($sig.Key)) {
                    $sigs[$sig.Key] = @{ Fields = $sig.Fields; Cells = (New-Object System.Collections.Generic.List[string]) }
                    $sigOrder.Add($sig.Key)
                }
                $addr = ''
                try { $addr = [string]$cell.Address($false, $false) } catch {}
                $sigs[$sig.Key].Cells.Add($addr)
            }
        }

        Write-Host ("  {0} distinct cell format(s):" -f $sigOrder.Count)
        $sheetSigs = New-Object System.Collections.Generic.List[object]
        $idx = 0
        foreach ($k in $sigOrder) {
            $idx++
            $g = $sigs[$k]
            $f = $g.Fields
            $sample = ($g.Cells | Select-Object -First $SampleCells) -join ','
            $more = if ($g.Cells.Count -gt $SampleCells) { (' (+{0} more)' -f ($g.Cells.Count - $SampleCells)) } else { '' }
            Write-Host ("  [{0,2}] {1} cell(s) -- e.g. {2}{3}" -f $idx, $g.Cells.Count, $sample, $more) -ForegroundColor White
            Write-Host ("       NumberFormat='{0}'  Font='{1}' {2}pt Bold={3} Color={4}" -f `
                $f.NumberFormat, $f.FontName, $f.FontSize, $f.FontBold, $f.FontColor) -ForegroundColor DarkGray
            Write-Host ("       Interior.Color={0}  HAlign={1} VAlign={2} Borders.LineStyle={3} Merged={4}" -f `
                $f.InteriorColor, $f.HAlign, $f.VAlign, $f.BorderLineStyle, $f.MergedCells) -ForegroundColor DarkGray
            $sheetSigs.Add([pscustomobject]@{
                Count = $g.Cells.Count; SampleCells = @($g.Cells | Select-Object -First $SampleCells)
                NumberFormat = $f.NumberFormat; FontName = $f.FontName; FontSize = $f.FontSize
                FontBold = $f.FontBold; FontColor = $f.FontColor; InteriorColor = $f.InteriorColor
                HAlign = $f.HAlign; VAlign = $f.VAlign; BorderLineStyle = $f.BorderLineStyle; MergedCells = $f.MergedCells
            })
        }

        $report.Add([pscustomobject]@{
            Sheet = [string]$ws.Name
            UsedRange = [string]$used.Address($false, $false)
            Rows = $rowCount; Columns = $colCount
            ScannedRows = $scanRows; ScannedColumns = $scanCols
            ColumnWidths = $colWidths.ToArray()
            RowHeights = @($heightGroups.Keys | ForEach-Object { [pscustomobject]@{ Height = $_; Rows = $heightGroups[$_].ToArray() } })
            Formats = $sheetSigs.ToArray()
        })
    }

    if (-not [string]::IsNullOrWhiteSpace($Json)) {
        if (-not [System.IO.Path]::IsPathRooted($Json)) { $Json = Join-Path (Get-Location).Path $Json }
        ([pscustomobject]@{ Workbook = $Path; ProbedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Sheets = $report.ToArray() } |
            ConvertTo-Json -Depth 8) | Set-Content -LiteralPath $Json -Encoding UTF8
        Write-Host ''
        Write-Host ("[OK] wrote JSON report -> {0}" -f $Json) -ForegroundColor Green
    }
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $excel) { try { Close-ExcelApp $excel } catch {} }
}

Write-Host ''
Write-Host '===== Probe-SheetFormat Done =====' -ForegroundColor Green
