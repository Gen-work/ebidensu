# ============================================================
#  MarkGfixLog.ps1
#
#  Phase: MarkGfixLog
#
#  In each evidence workbook's GFIX受信結果 sheet:
#    1. Scan column B for anchor cells containing the exact text ▼GFIXログ.
#    2. For each anchor region (anchor_row+1 .. next_anchor_row-1, or sheet end):
#       a. Find the first row whose B cell matches $CommandPattern.
#       b. Fill cells B:AY of that row with yellow (RGB 255,255,0).
#    3. Updates isGfixLogMarked = 1 on all rows in the group when all OK.
#
#  Idempotent: clears existing yellow fill in column B:AY of every
#  scanned row before re-applying, so re-runs are safe.
#
#  Usage:
#    .\MarkGfixLog.ps1 -WorkDir C:\work\myproject
#    .\MarkGfixLog.ps1 -WorkDir C:\work\myproject -TargetIds JIDSU91S
#    .\MarkGfixLog.ps1 -WorkDir C:\work\myproject -Force
# ============================================================

param(
    [string]$WorkDir,
    [string]$Owner = ([char]0x53B3),
    [string[]]$TargetIds = @(),
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

# ── Default anchor string ────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($LogAnchor)) {
    # ▼GFIXログ
    $LogAnchor = [char]0x25BC + "GFIX" + [char]0x30ED + [char]0x30B0
}

# ── Dot-source ExcelHelpers.ps1 ─────────────────────────────
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($ExcelHelpersScript)) { $candidates += $ExcelHelpersScript }
$candidates += @(
    (Join-Path $PSScriptRoot 'ExcelHelpers.ps1')
)
$helpersPath = $null
foreach ($c in $candidates) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).Path; break
    }
}
if (-not $helpersPath) { Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red; exit 1 }
. $helpersPath
if (-not (Get-Command -Name 'Set-CellRangeFill' -ErrorAction SilentlyContinue)) {
    Write-Host '[ERROR] ExcelHelpers dot-source failed (Set-CellRangeFill not loaded).' -ForegroundColor Red; exit 1
}

# ── Sheet name ───────────────────────────────────────────────
$sheetGfixRecv = "GFIX" + [char]0x53D7 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C  # GFIX受信結果

# ── Target filter ────────────────────────────────────────────
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

# ── Header ───────────────────────────────────────────────────
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
Write-Host ("  HighlightCol: {0}..{1}" -f $HighlightColStart, $HighlightColEnd)
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
Ensure-Column $allRows 'isGfixLogMarked' '0'

$workRows = @($allRows | Where-Object { Test-TargetRow $_ })
$groups   = $workRows | Group-Object Excel_NAME | Sort-Object Name
if ($groups.Count -eq 0) {
    Write-Host '[INFO] No rows after filter.' -ForegroundColor Yellow
    return
}

# ── Main loop ────────────────────────────────────────────────
$excel = New-ExcelApp
$cntDone = 0
$cntSkip = 0
$cntFail = 0

try {
    foreach ($g in $groups) {
        $first     = $g.Group | Select-Object -First 1
        $excelName = [string]$first.Excel_NAME
        if ([string]::IsNullOrWhiteSpace($excelName)) { continue }

        $wbPath = Find-WorkbookByExcelName -Dir $evDir -ExcelName $excelName
        if ($null -eq $wbPath) {
            Write-Host ("[SKIP] {0}: workbook missing" -f $excelName) -ForegroundColor Yellow
            $cntSkip++; continue
        }

        $curMark = 0
        try { $curMark = [int]$first.isGfixLogMarked } catch { $curMark = 0 }
        if (-not $forceFlag -and $curMark -eq 1) {
            Write-Host ("[SKIP] {0}: isGfixLogMarked already 1" -f $excelName) -ForegroundColor DarkGray
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

        $allOk        = $true
        $marksApplied = 0
        try {
            $ws = Get-SheetByName $wb $sheetGfixRecv
            if ($null -eq $ws) {
                Write-Host ("  [FAIL] sheet not found: {0}" -f $sheetGfixRecv) -ForegroundColor Red
                $allOk = $false
            } else {
                # 1) Collect anchor rows (B column exact-match to LogAnchor)
                $xlUp    = -4162
                $lastRow = 0
                try {
                    $lastRow = [int]$ws.Cells.Item($ws.Rows.Count, 2).End($xlUp).Row
                } catch { $lastRow = 0 }
                if ($lastRow -lt 1) {
                    try {
                        $used    = $ws.UsedRange
                        $lastRow = [int]($used.Row + $used.Rows.Count - 1)
                    } catch { $lastRow = 200 }
                }

                $anchorRows = @()
                for ($r = 1; $r -le $lastRow; $r++) {
                    $v = $null
                    try { $v = [string]$ws.Cells.Item($r, 2).Value2 } catch {}
                    if ($v -eq $LogAnchor) { $anchorRows += $r }
                }

                if ($anchorRows.Count -eq 0) {
                    Write-Host ("  [WARN] no '{0}' anchors found in sheet" -f $LogAnchor) -ForegroundColor Yellow
                    $allOk = $false
                } else {
                    Write-Host ("  [INFO] {0} anchor(s) found: rows {1}" -f $anchorRows.Count, ($anchorRows -join ',')) -ForegroundColor DarkGray

                    for ($ai = 0; $ai -lt $anchorRows.Count; $ai++) {
                        $regionStart = $anchorRows[$ai] + 1
                        $regionEnd   = if ($ai + 1 -lt $anchorRows.Count) { $anchorRows[$ai + 1] - 1 } else { $lastRow }

                        # 2) Clear previous yellow fills in this region
                        for ($r = $regionStart; $r -le $regionEnd; $r++) {
                            $existFill = -1
                            try { $existFill = [long]$ws.Cells.Item($r, $HighlightColStart).Interior.Color } catch {}
                            if ($existFill -eq $HighlightColor) {
                                Set-CellRangeFill $ws $r $HighlightColStart $HighlightColEnd -4142
                            }
                        }

                        # 3) Find Command: row
                        $targetRow = -1
                        $matchCount = 0
                        for ($r = $regionStart; $r -le $regionEnd; $r++) {
                            $v = $null
                            try { $v = [string]$ws.Cells.Item($r, 2).Value2 } catch {}
                            if (-not [string]::IsNullOrWhiteSpace($v) -and ($v -match $CommandPattern)) {
                                if ($matchCount -eq 0) { $targetRow = $r }
                                $matchCount++
                            }
                        }

                        if ($targetRow -lt 0) {
                            Write-Host ("  [WARN] anchor row {0}: no Command: match in region {1}..{2}" -f $anchorRows[$ai], $regionStart, $regionEnd) -ForegroundColor Yellow
                            $allOk = $false
                            continue
                        }

                        if ($matchCount -gt 1) {
                            Write-Host ("  [WARN] anchor row {0}: {1} Command: matches; using first (row {2})" -f $anchorRows[$ai], $matchCount, $targetRow) -ForegroundColor Yellow
                        }

                        Set-CellRangeFill $ws $targetRow $HighlightColStart $HighlightColEnd $HighlightColor
                        $marksApplied++
                        Write-Host ("  [MARK] anchor R{0} -> highlight R{1} (B:{2})" -f $anchorRows[$ai], $targetRow, [string]($ws.Cells.Item($targetRow, 2).Value2).Substring(0, [Math]::Min(40, ([string]($ws.Cells.Item($targetRow, 2).Value2)).Length))) -ForegroundColor Green
                    }
                }

                $wb.Save()
            }
        } catch {
            Write-Host ("  [FAIL] processing: {0}" -f $_.Exception.Message) -ForegroundColor Red
            $allOk = $false
        } finally {
            Close-Workbook $wb $false
        }

        Write-Host ("  highlights applied: {0}" -f $marksApplied) -ForegroundColor DarkGray

        if ($allOk -and $marksApplied -gt 0) {
            $groupNames = @($g.Group | ForEach-Object { [string]$_.Correl_ID_M })
            foreach ($r in $allRows) {
                if ($groupNames -contains [string]$r.Correl_ID_M) {
                    $r.isGfixLogMarked = '1'
                }
            }
            Write-Host ("  isGfixLogMarked = 1 for {0} row(s)" -f $g.Count) -ForegroundColor Green
            $cntDone++
        } elseif ($marksApplied -eq 0) {
            Write-Host '  [WARN] no highlights applied' -ForegroundColor Yellow
            $cntFail++
        } else {
            Write-Host '  isGfixLogMarked NOT updated (allOk=false)' -ForegroundColor Yellow
            $cntFail++
        }
    }

    if ($cntDone -gt 0) {
        $allRows | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force
        Write-Host ''
        Write-Host ("Mapping saved: {0}" -f $mappingPath) -ForegroundColor DarkGreen
    }
} finally {
    Close-ExcelApp $excel
}

Write-Host ''
Write-Host '===== MarkGfixLog Done =====' -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
