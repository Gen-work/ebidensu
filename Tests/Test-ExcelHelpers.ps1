#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Only the PURE helpers (Get-ColumnsForWidth column-accumulation math and
# Get-TextCellUnits character-cell counting) are exercised here. Everything
# else in ExcelHelpers.ps1 (Get-TextPixelWidth/Get-TextPointWidthInfo need
# GDI/GDI+, Get-AutoHighlightColEnd/Invoke-GfixLogHighlight/etc. need live
# Excel COM objects) is COM/GDI-bound and can only be confirmed on an
# office PC.

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

# -- Get-TextCellUnits: fixed-pitch character-cell width counting -----------
# Half-width (ASCII, halfwidth katakana) = 0.5 units; full-width = 1.0 units.
Assert-Equal 0 (Get-TextCellUnits -Text '') 'empty text -> 0 units'
Assert-Equal 0 (Get-TextCellUnits -Text $null) 'null text -> 0 units'
Assert-Equal 0.5 (Get-TextCellUnits -Text 'A') 'one ASCII char -> 0.5'
Assert-Equal 4 (Get-TextCellUnits -Text 'Command:') '8 ASCII chars -> 4.0'
Assert-Equal 0.5 (Get-TextCellUnits -Text ' ') 'ASCII space counts (trailing blanks in log lines)'

# Full-width chars: A (U+FF21), katakana RO (U+30ED), kanji (U+53D7).
$fwA    = [string][char]0xFF21
$kataRo = [string][char]0x30ED
$kanji  = [string][char]0x53D7
Assert-Equal 1 (Get-TextCellUnits -Text $fwA) 'full-width A -> 1.0'
Assert-Equal 1 (Get-TextCellUnits -Text $kataRo) 'full-width katakana -> 1.0'
Assert-Equal 1 (Get-TextCellUnits -Text $kanji) 'kanji -> 1.0'

# Halfwidth katakana (U+FF61..U+FF9F) advance half a cell like ASCII.
$hwKata = [string][char]0xFF71
Assert-Equal 0.5 (Get-TextCellUnits -Text $hwKata) 'halfwidth katakana -> 0.5'

# Mixed: 4 ASCII (2.0) + 2 full-width (2.0) = 4.0.
Assert-Equal 4 (Get-TextCellUnits -Text ('ab12' + $kataRo + $kanji)) 'mixed half/full-width'

# The floor a typical Command: line produces: pure-ASCII -> len/2 units, so
# at 11pt the floor is len * 5.5pt (ideal MS Gothic half-width advance).
$cmd = "Command: '/appl/JIDS/shell/recv_gfix.sh JIDSC02S'"
Assert-Equal ($cmd.Length * 0.5) (Get-TextCellUnits -Text $cmd) 'ASCII command line -> len/2 units'

# -- Get-TextPointWidthInfo: renderer-independent invariants ----------------
# The GDI/GDI+ tiers need Windows; on a GDI-less host the char-cell floor
# answers alone (Source='floor'). These assertions hold on EVERY platform:
# a real measurement can only make Points larger than the floor, never
# smaller (that is the whole point of the floor).
$wi = Get-TextPointWidthInfo -Text $cmd -FontName 'MS Gothic' -FontSize 11
Assert-True ($wi.Points -ge $wi.FloorPoints) 'PointWidthInfo: Points >= char-cell floor'
Assert-Equal ($cmd.Length * 0.5 * 11) $wi.FloorPoints 'PointWidthInfo: floor = units x size'
Assert-True ($wi.Source -ne 'none') 'PointWidthInfo: non-empty text always has a source'
$wiEmpty = Get-TextPointWidthInfo -Text ''
Assert-Equal 0 $wiEmpty.Points 'PointWidthInfo: empty text -> 0 points'
Assert-Equal 'none' $wiEmpty.Source 'PointWidthInfo: empty text -> source none'

exit (Complete-Tests)
