# ============================================================
#  Probe-Shapes.ps1
#
#  Calibration tool. Lists every Shape in an evidence workbook with
#  its name, type, position, size, and AlternativeText payload.
#
#  Workflow for first-time Mark calibration:
#    1. Run ReplaceGift on one Excel.
#    2. Open the resulting evidence/<name>.xlsx manually.
#    3. Draw your reference red rectangle(s) by hand around the area
#       you want Mark to circle on each picture.
#    4. Save the workbook.
#    5. Run Probe-Shapes.ps1 -File <that workbook>.
#    6. For each manual rectangle (Type=AutoShape), note its
#       (Left, Top, Width, Height).
#    7. For its parent picture (the Picture whose AltText is the
#       relevant source folder, e.g. "v1|GIFT_HM|JIDSC48S"), note
#       (Left, Top).
#    8. Offsets to put in VerifyConfig.psd1 -> Mark.Boxes:
#         OffsetX = rect.Left - picture.Left
#         OffsetY = rect.Top  - picture.Top
#         Width   = rect.Width
#         Height  = rect.Height
#
#  Usage:
#    .\Probe-Shapes.ps1 -File path\to\evidence.xlsx
#    .\Probe-Shapes.ps1 -File path\to\evidence.xlsx -Sheet 'GIFT受信結果'
# ============================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$File,

    [string]$Sheet = '',
    [string]$CommonScript = '',
    [string]$ExcelHelpersScript = ''
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

if (-not (Test-Path -LiteralPath $File)) {
    Write-Host "[ERROR] File not found: $File" -ForegroundColor Red; exit 1
}

# Dot-source ExcelHelpers
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($ExcelHelpersScript)) { $candidates += $ExcelHelpersScript }
$candidates += (Join-Path $PSScriptRoot 'ExcelHelpers.ps1')
$helpersPath = $null
foreach ($c in $candidates) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).Path; break
    }
}
if (-not $helpersPath) { Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red; exit 1 }
. $helpersPath

# msoShapeType decoder (the ones we care about)
$typeNames = @{
    1  = 'AutoShape'
    5  = 'Freeform'
    6  = 'Group'
    7  = 'EmbeddedOLE'
    11 = 'LinkedPic'
    13 = 'Picture'
    17 = 'TextBox'
    19 = 'Comment'
    23 = 'PictureLink'
    27 = 'Line'
    28 = 'Connector'
}

Write-Host ''
Write-Host '===== Probe-Shapes =====' -ForegroundColor Cyan
Write-Host ("  File  : {0}" -f $File)
if (-not [string]::IsNullOrWhiteSpace($Sheet)) { Write-Host ("  Sheet : {0}" -f $Sheet) }
Write-Host ''

$excel = New-ExcelApp
try {
    $wb = Open-Workbook $excel $File

    $sheetsToScan = @()
    foreach ($ws in $wb.Worksheets) {
        if ([string]::IsNullOrWhiteSpace($Sheet) -or $ws.Name -eq $Sheet) {
            $sheetsToScan += $ws
        }
    }
    if ($sheetsToScan.Count -eq 0) {
        Write-Host ("[WARN] No sheet matched '{0}'." -f $Sheet) -ForegroundColor Yellow
        $sheetsToScan = @($wb.Worksheets)
    }

    foreach ($ws in $sheetsToScan) {
        $shapeCount = 0
        try { $shapeCount = $ws.Shapes.Count } catch {}
        if ($shapeCount -eq 0) {
            Write-Host ("--- {0} (no shapes) ---" -f $ws.Name) -ForegroundColor DarkGray
            continue
        }

        Write-Host ("--- {0} ({1} shapes) ---" -f $ws.Name, $shapeCount) -ForegroundColor White
        $hdr = "  {0,-28} {1,-12} {2,8} {3,8} {4,8} {5,8}  {6}"
        Write-Host ($hdr -f 'Name','Type','Left','Top','Width','Height','AltText') -ForegroundColor DarkGray
        Write-Host ("  {0}" -f ('-' * 100)) -ForegroundColor DarkGray

        $shapes = @()
        foreach ($s in $ws.Shapes) { $shapes += $s }
        $shapes = $shapes | Sort-Object -Property @{ Expression = { [double]$_.Top } }, @{ Expression = { [double]$_.Left } }

        # Prints one shape row; recurses into Ctrl+G groups with indentation
        # so the children (and their possibly group-RELATIVE Top/Left) are
        # visible -- key when diagnosing OCR section-export misses.
        function Write-ShapeRow($s, [int]$Depth) {
            $name = ''; $type = ''; $alt = ''
            $L = 0.0; $T = 0.0; $W = 0.0; $H = 0.0
            $tInt = -1
            try { $name = [string]$s.Name } catch {}
            try {
                $tInt = [int]$s.Type
                if ($typeNames.ContainsKey($tInt)) { $type = $typeNames[$tInt] }
                else { $type = ("type{0}" -f $tInt) }
            } catch {}
            try { $L = [Math]::Round([double]$s.Left, 1) } catch {}
            try { $T = [Math]::Round([double]$s.Top, 1) } catch {}
            try { $W = [Math]::Round([double]$s.Width, 1) } catch {}
            try { $H = [Math]::Round([double]$s.Height, 1) } catch {}
            try { $alt = [string]$s.AlternativeText } catch {}

            $indent = ('  ' * $Depth)
            if ($Depth -gt 0) { $name = ('{0}> {1}' -f $indent, $name) }
            $color = if ($alt -like 'v1|*') { 'Green' } elseif ($type -eq 'AutoShape') { 'Yellow' } else { 'White' }
            Write-Host ($hdr -f $name, $type, $L, $T, $W, $H, $alt) -ForegroundColor $color

            if ($tInt -eq 6) {
                try {
                    foreach ($child in $s.GroupItems) { Write-ShapeRow $child ($Depth + 1) }
                } catch {
                    Write-Host ("  {0}> [WARN] cannot enumerate group children: {1}" -f $indent, $_.Exception.Message) -ForegroundColor Yellow
                }
            }
        }

        foreach ($s in $shapes) { Write-ShapeRow $s 0 }
    }

    Close-Workbook $wb $false
} finally {
    Close-ExcelApp $excel
}

Write-Host ''
Write-Host 'Legend:' -ForegroundColor DarkGray
Write-Host '  Green  = stamped Picture (has v1| metadata, this is what Mark anchors on)' -ForegroundColor Green
Write-Host '  Yellow = AutoShape (manual reference rectangle, measure these for Mark.Boxes)' -ForegroundColor Yellow
Write-Host ''
Write-Host 'Calibration recipe:' -ForegroundColor DarkGray
Write-Host '  OffsetX = AutoShape.Left - Picture.Left'
Write-Host '  OffsetY = AutoShape.Top  - Picture.Top'
Write-Host '  Width   = AutoShape.Width'
Write-Host '  Height  = AutoShape.Height'
Write-Host 'Edit VerifyConfig.psd1 -> Mark.Boxes -> <folder> accordingly.'
Write-Host ''
