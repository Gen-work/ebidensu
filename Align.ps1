#Requires -Version 5.1
# ============================================================
#  Align.ps1   (Phase: Align / Precheck)   UTF-8, NO BOM, ASCII source.
#
#  Compares each work evidence workbook (evidence\<Excel_NAME>.xlsx, or
#  a workbook whose filename ends with _<Excel_NAME>.xlsx) against the matching
#  J4 baseline workbook, and (with -Apply) syncs the sheets
#  that differ (spec 6).
#
#  Migration-type branching:
#    Host->Open : the 3 receive sheets only.
#    Open->Open / Open->Host : GIFT/GFIX send-result sheets + 3 receive.
#    (See AlignCompare.ps1 Get-AlignSheetsForMigration.)
#  Host vs Open is decided from FROM_sys/TO_sys via -HostSystemTypes. Until
#  those literal values are supplied, the type is 'Unknown' and Align safely
#  falls back to the Host->Open (3 receive) scope with a warning.
#
#  Default is a READ-ONLY DryRun (reports which sheets would sync). -Apply
#  performs a VALUES sync (clear work sheet contents, copy J4 UsedRange
#  values). Cell formatting is NOT synced yet (spec 6: format diffs = TODO).
#  -Apply is EXPERIMENTAL and untested on real Excel -- run it on a copy.
#
#  Usage:
#    .\Align.ps1 -WorkDir C:\work\proj -J4BaseDir \\fs\...\40.J4\07.GPCS
#    .\Align.ps1 -WorkDir ... -J4BaseDir ... -HostSystemTypes HOST,MF
#    .\Align.ps1 -WorkDir ... -J4BaseDir ... -Apply
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner = ([char]0x53B3),
    [string[]]$TargetIds = @(),
    [string]$J4BaseDir = '',
    [string[]]$HostSystemTypes = @(),
    [string]$MigrationTypeOverride = '',
    [switch]$Apply,
    [string]$ExcelHelpersScript = ''
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) { Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1 }

$applyFlag = [bool]$Apply.IsPresent

# dot-source ExcelHelpers + shared libs
$helpersPath = $null
foreach ($c in @($ExcelHelpersScript, (Join-Path $PSScriptRoot 'ExcelHelpers.ps1'))) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) { $helpersPath = (Resolve-Path -LiteralPath $c).Path; break }
}
if (-not $helpersPath) { Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red; exit 1 }
. $helpersPath
. (Join-Path $PSScriptRoot 'MappingStore.ps1')
. (Join-Path $PSScriptRoot 'ProgressLog.ps1')
. (Join-Path $PSScriptRoot 'ProjectLabels.ps1')
. (Join-Path $PSScriptRoot 'AlignCompare.ps1')

$labels     = Get-ProjectLabels
$sendSheets = Get-AlignSendSheets
$recvSheets = Get-AlignRecvSheets

$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
$evDir       = Join-Path $WorkDir 'evidence'

if (-not (Test-Path -LiteralPath $mappingPath)) { Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1 }
if (-not (Test-Path -LiteralPath $evDir))       { Write-Host "[ERROR] evidence dir missing: $evDir" -ForegroundColor Red; exit 1 }
if ([string]::IsNullOrWhiteSpace($J4BaseDir)) { $J4BaseDir = Read-Host 'J4BaseDir path' }
if ([string]::IsNullOrWhiteSpace($J4BaseDir) -or -not (Test-Path -LiteralPath $J4BaseDir)) {
    Write-Host "[ERROR] -J4BaseDir is required and must exist (baseline workbooks)." -ForegroundColor Red; exit 1
}

$allRows  = Import-Mapping $mappingPath
Ensure-MappingColumns -Rows $allRows | Out-Null
$targets  = ConvertTo-TargetIdList $TargetIds
$workRows = @($allRows | Where-Object { Test-TargetRow $_ $targets })
if ($workRows.Count -eq 0) { Write-Host '[INFO] No rows after filter.' -ForegroundColor Yellow; return }
$groups = $workRows | Group-Object Excel_NAME | Sort-Object Name

Write-Host ''
Write-Host '===== Align / Precheck =====' -ForegroundColor Green
Write-Host ("  Mode      : {0}" -f $(if ($applyFlag) { 'APPLY (values sync)' } else { 'DryRun (report only)' }))
Write-Host ("  J4BaseDir : {0}" -f $J4BaseDir)
Write-Host ("  HostTypes : {0}" -f $(if (@($HostSystemTypes).Count) { ($HostSystemTypes -join ',') } else { '(none configured)' }))
Write-Host ("  Groups    : {0}" -f $groups.Count)
Write-Host ''
Write-ProgressEvent -WorkDir $WorkDir -Phase 'Align' -Action 'start' -Status 'info' -Message ("apply={0} groups={1}" -f $applyFlag, $groups.Count)

. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

function Find-J4Workbook([string]$name) {
    return (Find-WorkbookByExcelName -Dir $J4BaseDir -ExcelName $name -Recurse)
}

function Read-SheetGrid($ws) {
    $used = $null
    try { $used = $ws.UsedRange } catch { return @{ Rows = 0; Cols = 0; Flat = @() } }
    $rows = [int]$used.Rows.Count
    $cols = [int]$used.Columns.Count
    $flat = [System.Collections.Generic.List[string]]::new()
    $vals = $used.Value2
    if ($rows -le 1 -and $cols -le 1) {
        $s = ''; if ($null -ne $vals) { $s = ([string]$vals).Trim() }
        $flat.Add($s)
    } else {
        for ($r = 1; $r -le $rows; $r++) {
            for ($c = 1; $c -le $cols; $c++) {
                $v = $vals[$r, $c]
                $s = ''; if ($null -ne $v) { $s = ([string]$v).Trim() }
                $flat.Add($s)
            }
        }
    }
    return @{ Rows = $rows; Cols = $cols; Flat = $flat.ToArray() }
}

function Sync-SheetValues($workWs, $j4Ws) {
    $used = $j4Ws.UsedRange
    $r0 = [int]$used.Row; $c0 = [int]$used.Column
    $rows = [int]$used.Rows.Count; $cols = [int]$used.Columns.Count
    try { $workWs.UsedRange.ClearContents() | Out-Null } catch {}
    $dst = $workWs.Range($workWs.Cells.Item($r0, $c0), $workWs.Cells.Item($r0 + $rows - 1, $c0 + $cols - 1))
    $dst.Value2 = $used.Value2
}

$excel = New-ExcelApp
$cntDiff = 0; $cntSame = 0; $cntSynced = 0; $cntSkip = 0

try {
    foreach ($g in $groups) {
        $first = $g.Group | Select-Object -First 1
        $excelName = [string]$first.Excel_NAME
        if ([string]::IsNullOrWhiteSpace($excelName)) { continue }
        $fromSys = ''; $toSys = ''
        if ($first.PSObject.Properties.Name -contains 'FROM_sys') { $fromSys = [string]$first.FROM_sys }
        if ($first.PSObject.Properties.Name -contains 'TO_sys')   { $toSys   = [string]$first.TO_sys }

        Write-Host ''
        Write-Host ("----- {0} -----" -f $excelName) -ForegroundColor Cyan

        $workPath = Find-WorkbookByExcelName -Dir $evDir -ExcelName $excelName
        if ($null -eq $workPath) { Write-Host ("  [SKIP] work workbook missing (*{0}.xlsx)" -f $excelName) -ForegroundColor Yellow; $cntSkip++; continue }
        $j4Path = Find-J4Workbook $excelName
        if ($null -eq $j4Path) { Write-Host ("  [SKIP] J4 baseline not found (*{0}.xlsx)" -f $excelName) -ForegroundColor Yellow; $cntSkip++; continue }
        Write-Host ("  work: {0}" -f (Split-Path $workPath -Leaf)) -ForegroundColor DarkGray
        Write-Host ("  J4  : {0}" -f (Split-Path $j4Path -Leaf)) -ForegroundColor DarkGray

        $migType = $MigrationTypeOverride
        if ([string]::IsNullOrWhiteSpace($migType)) { $migType = Get-MigrationType -FromSys $fromSys -ToSys $toSys -HostTypes $HostSystemTypes }
        if ($migType -eq 'Unknown') {
            Write-Host '  [WARN] migration type unknown (set -HostSystemTypes); using Host->Open scope (3 receive sheets)' -ForegroundColor Yellow
        }
        $sheets = Get-AlignSheetsForMigration -MigrationType $migType -SendSheets $sendSheets -RecvSheets $recvSheets
        Write-Host ("  migration: {0}   sheets to check: {1}" -f $migType, $sheets.Count) -ForegroundColor DarkGray

        $workWb = $null; $j4Wb = $null
        $changed = $false
        try {
            $workWb = Open-Workbook $excel $workPath
            $j4Wb   = Open-Workbook $excel $j4Path
            Unhide-AllSheets $workWb
            Unhide-AllSheets $j4Wb

            foreach ($sheetName in $sheets) {
                $wsW = Get-SheetByName $workWb $sheetName
                $wsJ = Get-SheetByName $j4Wb $sheetName
                if ($null -eq $wsW -or $null -eq $wsJ) {
                    Write-Host ("    [WARN] sheet missing in {0}: {1}" -f $(if ($null -eq $wsW) { 'work' } else { 'J4' }), $sheetName) -ForegroundColor Yellow
                    continue
                }
                $gw = Read-SheetGrid $wsW
                $gj = Read-SheetGrid $wsJ
                $cmp = Compare-SheetGrid -RowsA $gw.Rows -ColsA $gw.Cols -FlatA $gw.Flat -RowsB $gj.Rows -ColsB $gj.Cols -FlatB $gj.Flat
                if ($cmp.Same) {
                    Write-Host ("    [same] {0}" -f $sheetName) -ForegroundColor DarkGray
                    $cntSame++
                } else {
                    Write-Host ("    [DIFF] {0}  ({1})" -f $sheetName, $cmp.Reason) -ForegroundColor Yellow
                    $cntDiff++
                    if ($applyFlag) {
                        Sync-SheetValues $wsW $wsJ
                        Write-Host ("    [SYNC] {0} <- J4 (values; formats TODO)" -f $sheetName) -ForegroundColor Green
                        $changed = $true; $cntSynced++
                    }
                }
            }
            if ($applyFlag -and $changed) { $workWb.Save() }
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'Align' -JobName ([string]$first.JOB_NAME) -Action 'compare' `
                -Status $(if ($changed) { 'ok' } else { 'info' }) -Message ("{0} mig={1} synced={2}" -f $excelName, $migType, $changed)
        } catch {
            Write-Host ("  [FAIL] {0}" -f $_.Exception.Message) -ForegroundColor Red
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'Align' -Action 'compare' -Status 'fail' -Message ("{0}: {1}" -f $excelName, $_.Exception.Message)
        } finally {
            Close-Workbook $j4Wb $false
            Close-Workbook $workWb $false
        }
    }
} finally {
    Close-ExcelApp $excel
}

Write-Host ''
Write-Host '===== Align Done =====' -ForegroundColor Green
Write-Host ("  Sheets same   : {0}" -f $cntSame)
Write-Host ("  Sheets diff   : {0}" -f $cntDiff)
Write-Host ("  Sheets synced : {0}" -f $cntSynced) -ForegroundColor $(if ($applyFlag) { 'Green' } else { 'DarkGray' })
Write-Host ("  Groups skipped: {0}" -f $cntSkip)
if (-not $applyFlag -and $cntDiff -gt 0) {
    Write-Host '  (DryRun: re-run with -Apply to sync the DIFF sheets)' -ForegroundColor Magenta
}
