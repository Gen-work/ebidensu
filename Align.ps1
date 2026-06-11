#Requires -Version 5.1
# ============================================================
#  Align.ps1   (Phase: Align / Precheck)   UTF-8, NO BOM, ASCII source.
#
#  Compares each work evidence workbook (evidence\<Excel_NAME>.xlsx, or
#  a workbook whose filename ends with _<Excel_NAME>.xlsx) against the matching
#  J4 baseline workbook, and (with -Apply) syncs the sheets
#  that differ (spec 6).
#
#  Migration-type branching (recv sheets are NEVER synced -- operator's own evidence):
#    Host->Open : send[0] + GIFT/GFIX send-result sheets (host team manages these in J4).
#    Open->Open : GIFT/GFIX send-result sheets (coworker alignment, future use).
#    Open->Host : GIFT/GFIX send-result sheets.
#    Host->Host : all send sheets.
#    (See AlignCompare.ps1 Get-AlignSheetsForMigration.)
#  Host vs Open is decided from FROM_sys/TO_sys via -HostSystemTypes. Until
#  those literal values are supplied, the type is 'Unknown' and Align safely
#  falls back to the legacy receive-sheet scope with a warning.
#
#  Default is force-replace (Apply always on); -DiffMode switches to report-only.
#  -Apply is kept for backward compatibility (now a no-op; Apply is always active
#  unless -DiffMode is set).
#
#  Usage:
#    .\Align.ps1 -WorkDir C:\work\proj -J4BaseDir \\fs\...\40.J4\07.GPCS
#    .\Align.ps1 -WorkDir ... -J4BaseDir ... -HostSystemTypes HOST,MF
#    .\Align.ps1 -WorkDir ... -J4BaseDir ... -DiffMode    # report only, no sync
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner = '',
    [string[]]$TargetIds = @(),
    [string]$J4BaseDir = '',
    [string]$CloneSourceDir = '',
    [string]$ExcelPrefix = '',
    [string[]]$HostSystemTypes = @(),
    [string]$MigrationTypeOverride = '',
    [switch]$Apply,
    [switch]$DiffMode,
    [string]$ExcelHelpersScript = ''
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) { Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1 }

$diffFlag  = [bool]$DiffMode.IsPresent
$applyFlag = -not $diffFlag

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
Write-Host ("  Mode      : {0}" -f $(if ($diffFlag) { 'DiffMode (report only)' } else { 'Apply (force replace: work <- J4)' }))
Write-Host ("  J4BaseDir : {0}" -f $J4BaseDir)
if (-not [string]::IsNullOrWhiteSpace($CloneSourceDir)) { Write-Host ("  CloneSrc  : {0}" -f $CloneSourceDir) }
Write-Host ("  HostTypes : {0}" -f $(if (@($HostSystemTypes).Count) { ($HostSystemTypes -join ',') } else { '(none configured)' }))
Write-Host ("  Groups    : {0}" -f $groups.Count)
Write-Host ''
Write-ProgressEvent -WorkDir $WorkDir -Phase 'Align' -Action 'start' -Status 'info' -Message ("apply={0} groups={1}" -f $applyFlag, $groups.Count)

. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

function Read-AlignYesNo([string]$Prompt, [bool]$DefaultYes = $true) {
    $suffix = if ($DefaultYes) { 'Y/n' } else { 'y/N' }
    $ans = Read-Host ("{0} [{1}]" -f $Prompt, $suffix)
    if ([string]::IsNullOrWhiteSpace($ans)) { return $DefaultYes }
    return ($ans.Trim().ToLowerInvariant() -in @('y', 'yes'))
}

function Invoke-CloneForAlignMissingWorkbook([string]$ExcelName) {
    $cloneScript = Join-Path $PSScriptRoot 'Clone.ps1'
    if (-not (Test-Path -LiteralPath $cloneScript)) {
        Write-Host ("  [WARN] Clone.ps1 not found: {0}" -f $cloneScript) -ForegroundColor Yellow
        return
    }

    $args = @{ WorkDir = $WorkDir; Owner = $Owner; TargetIds = @($ExcelName) }
    if (-not [string]::IsNullOrWhiteSpace($ExcelPrefix)) { $args['ExcelPrefix'] = $ExcelPrefix }
    if (-not [string]::IsNullOrWhiteSpace($CloneSourceDir)) { $args['SourceDir'] = $CloneSourceDir }

    Write-Host ("  [RUN] Clone missing work workbook: {0}" -f $ExcelName) -ForegroundColor Green
    & $cloneScript @args
}

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

function Sync-Sheet($workWb, $workWs, $j4Ws) {
    $sheetName = $workWs.Name
    $idx       = $workWs.Index
    $workWs.Delete()
    $insertAfter = $workWb.Sheets.Item([Math]::Min($idx, $workWb.Sheets.Count))
    $j4Ws.Copy([System.Reflection.Missing]::Value, $insertAfter)
    $newWs = $workWb.ActiveSheet
    if ($newWs.Name -ne $sheetName) { $newWs.Name = $sheetName }
}

$excel = New-ExcelApp
$cntDiff = 0; $cntSame = 0; $cntSynced = 0; $cntSkip = 0

try {
    foreach ($g in $groups) {
        $first = $g.Group | Select-Object -First 1
        $excelName   = [string]$first.Excel_NAME
        if ([string]::IsNullOrWhiteSpace($excelName)) { continue }
        $excelPrefix = Resolve-ExcelPrefix -Row $first -DefaultPrefix $ExcelPrefix
        $fullStem    = Get-ExcelFullStem -Prefix $excelPrefix -Name $excelName
        $fromSys = ''; $toSys = ''
        if ($first.PSObject.Properties.Name -contains 'FROM_sys') { $fromSys = [string]$first.FROM_sys }
        if ($first.PSObject.Properties.Name -contains 'TO_sys')   { $toSys   = [string]$first.TO_sys }

        Write-Host ''
        Write-Host ("----- {0} -----" -f $excelName) -ForegroundColor Cyan

        $workPath = Find-WorkbookByExcelName -Dir $evDir -ExcelName $fullStem
        if ($null -eq $workPath) {
            Write-Host ("  [SKIP] work workbook missing ({0}.xlsx)" -f $fullStem) -ForegroundColor Yellow
            Write-Host ("         evidence dir: {0}" -f $evDir) -ForegroundColor DarkGray
            Write-Host ("         This usually means the Clone phase has not created this workbook yet.") -ForegroundColor DarkGray
            if (Read-AlignYesNo ("  Clone this workbook now?") $true) {
                Invoke-CloneForAlignMissingWorkbook $excelName
                $workPath = Find-WorkbookByExcelName -Dir $evDir -ExcelName $fullStem
                if ($null -eq $workPath) {
                    Write-Host ("  [SKIP] work workbook still missing after Clone ({0}.xlsx)" -f $fullStem) -ForegroundColor Yellow
                    $cntSkip++
                    continue
                }
                Write-Host ("  [OK] cloned work workbook: {0}" -f (Split-Path $workPath -Leaf)) -ForegroundColor Green
            } else {
                Write-Host '  [SKIP] operator chose not to clone now' -ForegroundColor Yellow
                $cntSkip++
                continue
            }
        }
        $j4Path = Find-J4Workbook $fullStem
        if ($null -eq $j4Path) { Write-Host ("  [SKIP] J4 baseline not found (*{0}.xlsx)" -f $excelName) -ForegroundColor Yellow; $cntSkip++; continue }
        Write-Host ("  work: {0}" -f (Split-Path $workPath -Leaf)) -ForegroundColor DarkGray
        Write-Host ("  J4  : {0}" -f (Split-Path $j4Path -Leaf)) -ForegroundColor DarkGray

        $migType = $MigrationTypeOverride
        if ([string]::IsNullOrWhiteSpace($migType)) { $migType = Get-MigrationType -FromSys $fromSys -ToSys $toSys -HostTypes $HostSystemTypes }
        if ($migType -eq 'Unknown') {
            Write-Host '  [WARN] migration type unknown (set -HostSystemTypes); using legacy receive-sheet scope' -ForegroundColor Yellow
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
                if ($null -eq $wsJ) {
                    Write-Host ("    [WARN] sheet missing in J4: {0}" -f $sheetName) -ForegroundColor Yellow
                    continue
                }
                if ($null -eq $wsW) {
                    if ($applyFlag) {
                        $insertAfter2 = $workWb.Sheets.Item($workWb.Sheets.Count)
                        $wsJ.Copy([System.Reflection.Missing]::Value, $insertAfter2)
                        $newWs2 = $workWb.ActiveSheet
                        if ($newWs2.Name -ne $sheetName) { try { $newWs2.Name = $sheetName } catch {} }
                        Write-Host ("    [ADD]  {0} (missing in work, copied from J4)" -f $sheetName) -ForegroundColor Green
                        $cntSynced++; $changed = $true
                    } else {
                        Write-Host ("    [MISS] {0} (in J4, not in work -- run without -DiffMode to copy)" -f $sheetName) -ForegroundColor Yellow
                    }
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
                        Sync-Sheet $workWb $wsW $wsJ
                        Write-Host ("    [SYNC] {0} <- J4 (full copy: values + formats + shapes + pictures)" -f $sheetName) -ForegroundColor Green
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
if ($diffFlag -and $cntDiff -gt 0) {
    Write-Host '  (DiffMode: re-run without -DiffMode to sync the DIFF sheets)' -ForegroundColor Magenta
}
