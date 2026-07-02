#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Only the PURE column-accumulation math (Get-ColumnsForWidth) is exercised
# here. Everything else in ExcelHelpers.ps1 (Get-TextPixelWidth needs GDI+,
# Get-AutoHighlightColEnd/Invoke-GfixLogHighlight/etc. need live Excel COM
# objects) is COM/GDI-bound and can only be confirmed on an office PC.

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'ExcelHelpers.ps1')

Reset-Tests 'ExcelHelpers (pure helpers)'

# No columns / non-positive target -> offset 0 (caller treats this as "no columns needed").
Assert-Equal 0 (Get-ColumnsForWidth -ColumnWidths @() -NeededPoints 100) 'empty widths -> 0'
Assert-Equal 0 (Get-ColumnsForWidth -ColumnWidths @(59.0, 59.0) -NeededPoints 0) 'zero target -> 0'
Assert-Equal 0 (Get-ColumnsForWidth -ColumnWidths @(59.0, 59.0) -NeededPoints -5) 'negative target -> 0'

# First column alone already covers the target -> offset 0.
Assert-Equal 0 (Get-ColumnsForWidth -ColumnWidths @(59.0, 59.0, 59.0) -NeededPoints 30) 'covered by first column -> offset 0'
Assert-Equal 0 (Get-ColumnsForWidth -ColumnWidths @(59.0, 59.0, 59.0) -NeededPoints 59) 'exact first-column width -> offset 0'

# Needs more than one column.
Assert-Equal 1 (Get-ColumnsForWidth -ColumnWidths @(59.0, 59.0, 59.0) -NeededPoints 60) 'just past first column -> offset 1'
Assert-Equal 2 (Get-ColumnsForWidth -ColumnWidths @(50.0, 50.0, 50.0, 50.0) -NeededPoints 120) 'needs 3 columns -> offset 2'

# Target wider than every column combined -> clamp to the last available column.
Assert-Equal 2 (Get-ColumnsForWidth -ColumnWidths @(20.0, 20.0, 20.0) -NeededPoints 1000) 'overflow clamps to last column'

exit (Complete-Tests)
