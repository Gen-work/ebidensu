# ============================================================
#  MarkGfixLog.ps1
#
#  Phase: MarkGfixLog
#
#  In each evidence workbook's GFIX jushin kekka sheet, fills the GFIX-log
#  "Command:" row with yellow (see Invoke-GfixLogHighlight in ExcelHelpers.ps1).
#
#  NOTE: this highlight is normally performed as part of MarkGfix now (one
#  workbook open, tracked by the single isMarked GFIX bit). This script is
#  kept as a standalone, idempotent re-highlight utility and does NOT write a
#  mapping column (the old isGfixLogMarked flag was removed).
#
#  Idempotent: clears existing yellow fill in the scanned region before
#  re-applying, so re-runs are safe.
#
#  Usage:
#    .\MarkGfixLog.ps1 -WorkDir C:\work\myproject
#    .\MarkGfixLog.ps1 -WorkDir C:\work\myproject -TargetIds JIDSU91S
#    .\MarkGfixLog.ps1 -WorkDir C:\work\myproject -Force
# ============================================================

param(
    [string]$WorkDir,
    [string]$Owner = '',
    [string[]]$TargetIds = @(),
    [string]$ExcelPrefix = '',
    [switch]$Force,

    [string]$ExcelHelpersScript = '',

    # ▼GFIXログ anchor text (B column exact match)
    [string]$LogAnchor = '',

    # Regex that identifies the target row to highlight
    [string]$CommandPattern = "Command:\s*'/appl/[A-Za-z0-9]+/shell/",

    # Highlight color: yellow RGB(255,255,0) = OLE 65535
    [long]$HighlightColor = 65535,

    # Highlight column range (B=2, AY=51)
    [int]$HighlightColStart = 2,
    [int]$HighlightColEnd   = 51,

    # true (default) = size the highlight to the target row's actual pasted
    # text width instead of always filling to HighlightColEnd (see
    # Get-AutoHighlightColEnd in ExcelHelpers.ps1). HighlightColEnd stays the
    # upper bound either way.
    [bool]$AutoWidth = $true,
    [int]$PadCols = 1,

    # Font the GFIX log was PASTED in (Replace.GfixLogFontName/GfixLogFontSize).
    # Used by the AutoWidth measurement so the computed highlight width always
    # matches the rendered text. Blank/0 -> measure with the cell's own font.
    [string]$FontName = '',
    [double]$FontSize = 0,

    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

try {
    Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
} catch {}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1
}

$forceFlag = [bool]$Force.IsPresent
$dryFlag   = [bool]$DryRun.IsPresent

# -- Default anchor string ------------------------------------
if ([string]::IsNullOrWhiteSpace($LogAnchor)) {
    # ▼GFIXログ
    $LogAnchor = [char]0x25BC + "GFIX" + [char]0x30ED + [char]0x30B0
}

# -- Dot-source ExcelHelpers.ps1 -----------------------------
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($ExcelHelpersScript)) { $candidates += $ExcelHelpersScript }
$candidates += @(
    (Join-Path $PSScriptRoot 'ExcelHelpers.ps1')
)
$helpersPath = $null
foreach ($c in $candidates) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).ProviderPath; break
    }
}
if (-not $helpersPath) { Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red; exit 1 }
. $helpersPath
if (-not (Get-Command -Name 'Set-CellRangeFill' -ErrorAction SilentlyContinue)) {
    Write-Host '[ERROR] ExcelHelpers dot-source failed (Set-CellRangeFill not loaded).' -ForegroundColor Red; exit 1
}
# Needed by Get-TextPixelWidth (AutoWidth path of Invoke-GfixLogHighlight).
# Missing/failed load degrades gracefully (AutoWidth falls back to the fixed
# HighlightColEnd), but load it so AutoWidth actually works standalone too.
try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop } catch {
    Write-Host ("[WARN] System.Drawing unavailable ({0}); AutoWidth will fall back to the fixed HighlightColEnd." -f $_.Exception.Message) -ForegroundColor Yellow
}
# System.Windows.Forms provides TextRenderer, the GDI measurement tier of the
# AutoWidth highlight (Get-TextPointWidthInfo). GDI matches how Excel actually
# renders cell text (hinted MS Gothic advances), so it is preferred; a failed
# load only drops that tier (GDI+ + the char-cell floor remain).
try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop } catch {
    Write-Host ("[WARN] System.Windows.Forms unavailable ({0}); AutoWidth will measure via GDI+ only." -f $_.Exception.Message) -ForegroundColor Yellow
}

# -- Sheet name -----------------------------------------------
$sheetGfixRecv = "GFIX" + [char]0x53D7 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C  # GFIX受信結果

# -- Target filter --------------------------------------------
$targetSet = @{}
foreach ($rawId in @($TargetIds)) {
    if ($null -eq $rawId) { continue }
    foreach ($part in ($rawId.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $targetSet[$v] = $true }
    }
}
function Test-TargetRow($row) {
    if ($targetSet.Count -eq 0) { return $true }
    return ($targetSet.ContainsKey([string]$row.Correl_ID_S) -or
            $targetSet.ContainsKey([string]$row.Correl_ID_M) -or
            $targetSet.ContainsKey([string]$row.JOB_NAME) -or
            $targetSet.ContainsKey([string]$row.Excel_NAME))
}

# -- Header ---------------------------------------------------
$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
$evDir       = Join-Path $WorkDir 'evidence'

. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

Write-Host ''
Write-Host '===== MarkGfixLog =====' -ForegroundColor Green
Write-Host ("  WorkDir   : {0}" -f $WorkDir)
Write-Host ("  Mapping   : {0}" -f $mappingPath)
Write-Host ("  Sheet     : {0}" -f $sheetGfixRecv)
Write-Host ("  Anchor    : {0}" -f $LogAnchor)
Write-Host ("  Pattern   : {0}" -f $CommandPattern)
Write-Host ("  HighlightCol: {0}..{1}  AutoWidth: {2}  PadCols: {3}  MeasureFont: {4} {5}" -f $HighlightColStart, $HighlightColEnd, $AutoWidth, $PadCols, `
    $(if ([string]::IsNullOrWhiteSpace($FontName)) { '(cell font)' } else { $FontName }), $(if ($FontSize -le 0) { '' } else { [string]$FontSize }))
Write-Host ("  Force     : {0}" -f $forceFlag)
Write-Host ("  DryRun    : {0}" -f $dryFlag)
if ($targetSet.Count -gt 0) { Write-Host ("  TargetIds : {0}" -f (($targetSet.Keys | Sort-Object) -join ', ')) }
Write-Host ''

if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1
}
if (-not (Test-Path -LiteralPath $evDir)) {
    Write-Host "[ERROR] evidence dir missing: $evDir" -ForegroundColor Red; exit 1
}

$allRows = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)

$workRows = @($allRows | Where-Object { Test-TargetRow $_ })
$groups   = $workRows | Group-Object Excel_NAME | Sort-Object Name
if ($groups.Count -eq 0) {
    Write-Host '[INFO] No rows after filter.' -ForegroundColor Yellow
    return
}

# -- Main loop ------------------------------------------------
# Standalone re-highlight utility. The shared Invoke-GfixLogHighlight does the
# actual work (same code path MarkGfix uses). No mapping column is written.
# -Force is accepted for back-compat but is a no-op (the op is idempotent).
$excel = New-ExcelApp
$cntDone = 0
$cntSkip = 0
$cntFail = 0

try {
    foreach ($g in $groups) {
        $first     = $g.Group | Select-Object -First 1
        $excelName   = [string]$first.Excel_NAME
        if ([string]::IsNullOrWhiteSpace($excelName)) { continue }
        $excelPrefix = Resolve-ExcelPrefix -Row $first -DefaultPrefix $ExcelPrefix
        $fullStem    = Get-ExcelFullStem -Prefix $excelPrefix -Name $excelName

        $wbPath = Find-WorkbookByExcelName -Dir $evDir -ExcelName $fullStem
        if ($null -eq $wbPath) {
            Write-Host ("[SKIP] {0}: workbook missing" -f $excelName) -ForegroundColor Yellow
            $cntSkip++; continue
        }

        Write-Host ''
        Write-Host ("----- {0} -----" -f $excelName) -ForegroundColor Cyan

        if ($dryFlag) {
            Write-Host ("  [DRY]  would process {0}" -f $wbPath) -ForegroundColor DarkGray
            $cntSkip++; continue
        }

        $wb = $null
        try { $wb = Open-Workbook $excel $wbPath } catch {
            Write-Host ("  [FAIL] open: {0}" -f $_.Exception.Message) -ForegroundColor Red
            $cntFail++; continue
        }

        $hl = $null
        try {
            $ws = Get-SheetByName $wb $sheetGfixRecv
            if ($null -eq $ws) {
                Write-Host ("  [FAIL] sheet not found: {0}" -f $sheetGfixRecv) -ForegroundColor Red
            } else {
                $hl = Invoke-GfixLogHighlight -ws $ws -LogAnchor $LogAnchor `
                    -CommandPattern $CommandPattern -HighlightColor $HighlightColor `
                    -ColStart $HighlightColStart -ColEnd $HighlightColEnd `
                    -AutoWidth $AutoWidth -PadCols $PadCols `
                    -FontName $FontName -FontSize $FontSize
                foreach ($w in @($hl.Warnings)) { Write-Host ("  [WARN] {0}" -f $w) -ForegroundColor Yellow }
                foreach ($d in @($hl.Diag))     { Write-Host ("  [width] {0}" -f $d) -ForegroundColor DarkGray }
                Write-Host ("  highlights applied: {0} (anchors: {1})" -f $hl.Applied, $hl.Anchors) -ForegroundColor DarkGray
                $wb.Save()
            }
        } catch {
            Write-Host ("  [FAIL] processing: {0}" -f $_.Exception.Message) -ForegroundColor Red
        } finally {
            Close-Workbook $wb $false
        }

        if ($null -ne $hl -and $hl.Applied -gt 0) { $cntDone++ } else { $cntFail++ }
    }
} finally {
    Close-ExcelApp $excel
}

Write-Host ''
Write-Host '===== MarkGfixLog Done =====' -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
