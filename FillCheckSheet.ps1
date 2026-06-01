# ============================================================
# FillCheckSheet.ps1 - Phase CheckSheet
#
# Appends one row per evidence Excel (grouped by Excel_NAME) to the shared
# review check sheet workbook, sheet "Check Sheet_J4".
#
# The check sheet is a PUBLIC document, so the write is double-checked:
#   1. Snapshot the original file's timestamp + size.
#   2. Copy it to a TEMP file, apply the planned rows there, open it visible
#      so Misaki can eyeball the result.
#   3. She presses Enter to commit, or q to abort.
#   4. On commit the original is re-stat'd; if it changed at all during the
#      preview the write is HELD (nothing is overwritten) so she can re-check.
#      Otherwise the identical edits are applied to the original and saved.
#
# Columns written (1-indexed, configurable; Japanese headers in parens):
#   A No.            : continued from the last numeric No. (only if blank)
#   B (kinyuubi)     : today's date, number format copied from the row above
#   C (COBOL/JAVA)   : Language   (JAVA)
#   D (resource id)  : left blank
#   E (kakunin phase): Phase      (J4 internal review label)
#   F (review target): full evidence filename  <Excel_Prefix>_<Excel_NAME>.xlsx
#   G (tantou)       : Owner
#   H (kakuninsha)   : Viewer     (reviewer short name)
#   I (kanryou kibou-bi) / J~ : left blank
#
# All Japanese (sheet label, phase label, reviewer name) arrives as
# parameters from VerifyConfig.psd1 (BOM) so this source stays pure ASCII.
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner = ([char]0x53B3),
    [string[]]$TargetIds = @(),

    [string]$CheckSheetPath = '',
    [string]$SheetName = 'Check Sheet_J4',
    [string]$Language = 'JAVA',
    [string]$Phase = '',
    [string]$Viewer = '',

    [int]$ColNo = 1,
    [int]$ColDate = 2,
    [int]$ColLang = 3,
    [int]$ColResourceId = 4,
    [int]$ColPhase = 5,
    [int]$ColTarget = 6,
    [int]$ColOwner = 7,
    [int]$ColViewer = 8,

    [string]$DateFormat = 'yyyy/m/d',

    [switch]$Force,
    [switch]$DryRun,
    [string]$ExcelHelpersScript = ''
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$forceFlag  = [bool]$Force.IsPresent
$dryRunFlag = [bool]$DryRun.IsPresent

# -- Dot-source ExcelHelpers.ps1 + shared libs ---------------
$helpersPath = $null
foreach ($c in @($ExcelHelpersScript, (Join-Path $PSScriptRoot 'ExcelHelpers.ps1'))) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).Path; break
    }
}
if (-not $helpersPath) {
    Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red; exit 1
}
. $helpersPath
. (Join-Path $PSScriptRoot 'MappingStore.ps1')
. (Join-Path $PSScriptRoot 'ProgressLog.ps1')
. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

function Get-CellText($ws, [int]$row, [int]$col) {
    $v = $null
    try { $v = $ws.Cells.Item($row, $col).Value2 } catch { return '' }
    if ($null -eq $v) { return '' }
    return ([string]$v).Trim()
}

# Computes and (optionally) writes the new rows on $ws. Returns an array of
# summary objects { Row; No; Target; Action }. Re-running computes identical
# results on the original as on the temp copy (same content), so the preview
# and the commit stay in lock-step.
function Apply-CheckSheetRows($ws, [object[]]$Candidates, [bool]$WriteCells) {
    $summary = New-Object System.Collections.Generic.List[object]

    $xlUp = -4162
    $lastB = 0
    try { $lastB = [int]$ws.Cells.Item($ws.Rows.Count, $ColDate).End($xlUp).Row } catch { $lastB = 0 }
    # End(xlUp) lands on row 1 even when empty; treat a blank B1 as "no data".
    if ($lastB -eq 1 -and [string]::IsNullOrWhiteSpace((Get-CellText $ws 1 $ColDate))) { $lastB = 0 }
    $startRow = $lastB + 1

    # Existing review targets already listed (column F), for de-dup.
    $existing = @{}
    for ($r = 1; $r -le $lastB; $r++) {
        $t = Get-CellText $ws $r $ColTarget
        if (-not [string]::IsNullOrWhiteSpace($t)) { $existing[$t.ToLower()] = $true }
    }

    # Last numeric No. above the insertion point (skip the "rei" sample row).
    $lastNo = 0
    for ($r = $lastB; $r -ge 1; $r--) {
        $a = Get-CellText $ws $r $ColNo
        $n = 0
        if ([int]::TryParse($a, [ref]$n)) { $lastNo = $n; break }
    }

    # Date format to mirror from the last filled B cell.
    $dateFmt = $DateFormat
    if ($lastB -ge 1) {
        try {
            $f = [string]$ws.Cells.Item($lastB, $ColDate).NumberFormat
            if (-not [string]::IsNullOrWhiteSpace($f)) { $dateFmt = $f }
        } catch {}
    }

    $today    = (Get-Date).Date
    $todaySer = [double]$today.ToOADate()
    $nextNo   = $lastNo + 1
    $row      = $startRow

    foreach ($cand in $Candidates) {
        $target = [string]$cand.TargetFile
        if ((-not $forceFlag) -and $existing.ContainsKey($target.ToLower())) {
            $summary.Add([pscustomobject]@{ Row = 0; No = ''; Target = $target; Action = 'skip (already listed)' })
            continue
        }

        # No. -- only fill when blank; always keep numbering monotonic.
        $noText = Get-CellText $ws $row $ColNo
        $thisNo = $noText
        if ([string]::IsNullOrWhiteSpace($noText)) {
            $thisNo = [string]$nextNo
            if ($WriteCells) { try { $ws.Cells.Item($row, $ColNo).Value2 = $nextNo } catch {} }
            $nextNo++
        } else {
            $parsed = 0
            if ([int]::TryParse($noText, [ref]$parsed) -and $parsed -ge $nextNo) { $nextNo = $parsed + 1 }
        }

        if ($WriteCells) {
            try {
                $bCell = $ws.Cells.Item($row, $ColDate)
                $bCell.Value2 = $todaySer
                try { $bCell.NumberFormat = $dateFmt } catch {}
            } catch {}
            try { $ws.Cells.Item($row, $ColLang).Value2   = $Language } catch {}
            try { $ws.Cells.Item($row, $ColPhase).Value2  = $Phase } catch {}
            try { $ws.Cells.Item($row, $ColTarget).Value2 = $target } catch {}
            try { $ws.Cells.Item($row, $ColOwner).Value2  = $Owner } catch {}
            try { $ws.Cells.Item($row, $ColViewer).Value2 = $Viewer } catch {}
        }

        $summary.Add([pscustomobject]@{ Row = $row; No = $thisNo; Target = $target; Action = 'add' })
        $row++
    }

    return $summary.ToArray()
}

function Show-Plan([object[]]$Plan) {
    Write-Host ''
    Write-Host '  Planned check-sheet rows:' -ForegroundColor Cyan
    foreach ($p in $Plan) {
        if ($p.Action -eq 'add') {
            Write-Host ("    row {0,-4} No.{1,-4} {2}" -f $p.Row, $p.No, $p.Target) -ForegroundColor White
        } else {
            Write-Host ("    {0,-14} {1}" -f $p.Action, $p.Target) -ForegroundColor DarkGray
        }
    }
}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1
}

$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1
}

# Resolve the check sheet path; prompt if the configured one does not exist.
if ([string]::IsNullOrWhiteSpace($CheckSheetPath) -or -not (Test-Path -LiteralPath $CheckSheetPath)) {
    if (-not [string]::IsNullOrWhiteSpace($CheckSheetPath)) {
        Write-Host ("[INFO] configured check sheet not found: {0}" -f $CheckSheetPath) -ForegroundColor Yellow
    }
    $CheckSheetPath = Read-Host 'Review check sheet (.xlsx) full path'
}
if ([string]::IsNullOrWhiteSpace($CheckSheetPath) -or -not (Test-Path -LiteralPath $CheckSheetPath)) {
    Write-Host "[ERROR] check sheet not found: $CheckSheetPath" -ForegroundColor Red; exit 1
}
$CheckSheetPath = (Resolve-Path -LiteralPath $CheckSheetPath).Path

$targets = @(ConvertTo-TargetIdList $TargetIds)

$allRows = @(Import-Mapping $mappingPath)
if ($allRows.Count -eq 0) {
    Write-Host "[ERROR] mapping has no rows: $mappingPath" -ForegroundColor Red; exit 1
}

# One candidate per Excel_NAME (mapping order), honoring the target filter.
$candidates = New-Object System.Collections.Generic.List[object]
$seen = @{}
foreach ($r in $allRows) {
    if (-not (Test-TargetRow $r $targets)) { continue }
    $name = [string]$r.Excel_NAME
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    if ($seen.ContainsKey($name)) { continue }
    $seen[$name] = $true
    $fullStem = Get-ExcelFullStem -Prefix (Get-RowProp $r 'Excel_Prefix') -Name $name
    $candidates.Add([pscustomobject]@{ ExcelName = $name; TargetFile = (Get-ExcelDestLeaf $fullStem) })
}
if ($candidates.Count -eq 0) {
    Write-Host '[INFO] no target rows.' -ForegroundColor Yellow
    return
}

Write-Host ''
Write-Host '===== FillCheckSheet =====' -ForegroundColor Green
Write-Host ("  WorkDir    : {0}" -f $WorkDir)
Write-Host ("  CheckSheet : {0}" -f $CheckSheetPath)
Write-Host ("  Sheet      : {0}" -f $SheetName)
Write-Host ("  Candidates : {0}" -f $candidates.Count)
Write-Host ("  Owner      : {0}" -f $Owner)
Write-Host ("  Force      : {0}" -f $forceFlag)
if ($targets.Count -gt 0) { Write-Host ("  TargetIds  : {0}" -f ($targets -join ', ')) }

$excel = $null
$wbTmp = $null
$wbOrig = $null
$tmpPath = $null

try {
    $excel = New-ExcelApp
    $excel.Visible = $true
    try { $excel.DisplayAlerts = $false } catch {}
    # New-ExcelApp turns ScreenUpdating OFF for headless speed. This phase
    # shows the workbook to the operator for a visual check, so turn it back
    # ON -- otherwise the visible window never repaints and the preview looks
    # blank (no name, no rows/columns) even though the content is correct.
    try { $excel.ScreenUpdating = $true } catch {}

    # -- DryRun: compute on a read-only open of the original, no writes --
    if ($dryRunFlag) {
        $wbRo = $excel.Workbooks.Open($CheckSheetPath, 0, $true)   # ReadOnly
        try {
            $wsRo = Get-SheetByName $wbRo $SheetName
            if ($null -eq $wsRo) { Write-Host ("[ERROR] sheet not found: {0}" -f $SheetName) -ForegroundColor Red; return }
            $plan = Apply-CheckSheetRows $wsRo $candidates.ToArray() $false
            Show-Plan $plan
            Write-Host '  [DRY RUN] nothing written.' -ForegroundColor Yellow
        } finally {
            try { $wbRo.Close($false) } catch {}
        }
        return
    }

    # -- 1) snapshot original --------------------------------
    $origItem = Get-Item -LiteralPath $CheckSheetPath
    $snapTime = $origItem.LastWriteTimeUtc
    $snapLen  = [long]$origItem.Length

    # -- 2) preview on a temp copy ---------------------------
    $tmpPath = Join-Path $env:TEMP ("checksheet_preview_{0}_{1}.xlsx" -f $PID, (Get-Date -Format 'yyyyMMddHHmmss'))
    Copy-Item -LiteralPath $CheckSheetPath -Destination $tmpPath -Force

    $wbTmp = $excel.Workbooks.Open($tmpPath)
    $wsTmp = Get-SheetByName $wbTmp $SheetName
    if ($null -eq $wsTmp) {
        Write-Host ("[ERROR] sheet not found in workbook: {0}" -f $SheetName) -ForegroundColor Red
        return
    }
    try { $wsTmp.Visible = -1 } catch {}
    $wsTmp.Activate() | Out-Null

    $plan = Apply-CheckSheetRows $wsTmp $candidates.ToArray() $true
    $added = @($plan | Where-Object { $_.Action -eq 'add' })
    try { $wbTmp.Save() } catch {}

    Show-Plan $plan
    if ($added.Count -gt 0) {
        try { $wsTmp.Range($wsTmp.Cells.Item($added[0].Row, 1), $wsTmp.Cells.Item($added[-1].Row, $ColViewer)).Select() | Out-Null } catch {}
    }

    if ($added.Count -eq 0) {
        Write-Host '  Nothing to add (all candidates already listed). Use -Force to add anyway.' -ForegroundColor Yellow
        return
    }

    Write-Host ''
    Write-Host '  PREVIEW open (temp copy). Check the rows above in Excel.' -ForegroundColor DarkGray
    $ans = Read-Host '  Enter=commit to the real check sheet, q=abort (nothing written)'
    if ($ans.Trim().ToLower() -eq 'q') {
        Write-Host '  [ABORT] no changes written to the check sheet.' -ForegroundColor Yellow
        return
    }

    # -- 3) verify the original did not change during preview --
    $origItem2 = Get-Item -LiteralPath $CheckSheetPath
    if ($origItem2.LastWriteTimeUtc -ne $snapTime -or [long]$origItem2.Length -ne $snapLen) {
        Write-Host ''
        Write-Host '  [HOLD] the check sheet changed on disk during the preview.' -ForegroundColor Red
        Write-Host '         Nothing was written. Re-run CheckSheet to rebuild against the new content.' -ForegroundColor Red
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'CheckSheet' -Action 'commit' -Status 'skip' -Message 'original changed during preview'
        return
    }

    # -- 4) publish the previewed temp copy over the original ----
    # The temp copy already IS the original plus the new rows, and we just
    # verified above that the original did not change during the preview.
    # So copy it back instead of reopening the original in Excel. This writes
    # exactly what Misaki saw, and -- importantly -- it sidesteps Excel's
    # ~218-char path limit: Workbooks.Open fails ("見つかりません") on the long
    # UNC check-sheet path even though the file exists, whereas .NET Copy-Item
    # handles the long path fine.
    try { $wbTmp.Save() } catch {}
    try { $wbTmp.Close($true) } catch {}
    $wbTmp = $null

    try {
        Copy-Item -LiteralPath $tmpPath -Destination $CheckSheetPath -Force
    } catch {
        Write-Host ('  [HOLD] could not write the check sheet (open elsewhere or no access). Nothing written.') -ForegroundColor Red
        Write-Host ('         {0}' -f $_.Exception.Message) -ForegroundColor Red
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'CheckSheet' -Action 'commit' -Status 'fail' -Message $_.Exception.Message
        return
    }

    $okAdded = $added.Count
    Write-Host ("  [OK] wrote {0} row(s) to {1}" -f $okAdded, $CheckSheetPath) -ForegroundColor Green
    foreach ($p in $added) {
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'CheckSheet' -JobName $p.Target -Action 'commit' -Status 'ok' -Message ("No.{0}" -f $p.No)
    }
} catch {
    Write-Host ("[ERROR] {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-ProgressEvent -WorkDir $WorkDir -Phase 'CheckSheet' -Action 'commit' -Status 'fail' -Message $_.Exception.Message
} finally {
    if ($wbOrig) { try { $wbOrig.Close($false) } catch {} }
    if ($wbTmp)  { try { $wbTmp.Close($false) } catch {} }
    if ($excel)  { Close-ExcelApp $excel }
    if ($tmpPath -and (Test-Path -LiteralPath $tmpPath)) {
        try { Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue } catch {}
    }
}

Write-Host ''
Write-Host '===== FillCheckSheet Done =====' -ForegroundColor Green
