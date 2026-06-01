# ============================================================
#  ReplaceEvidence.ps1   (Phase: ReplaceGift / ReplaceGfix / ReplaceDf)
#  UTF-8, NO BOM, ASCII source. Japanese sheet/label names come from
#  ProjectLabels.ps1 ([char]-built), so this file is codepage-agnostic.
#
#  Rewritten to be PLAN-DRIVEN (spec section 6 priority):
#    1. Build a correl-major insert plan with the pure, unit-tested
#       EvidencePlan.ps1 (Build-Gift/Gfix/Df EvidencePlan).
#    2. Execute it with EvidenceExecutor.ps1 against the target sheet.
#  This replaces the old folder-major loop (all HM, then all MQ, ...),
#  which conflicted with the review standard.
#
#  Target sheet per -Mode:
#    Gift -> GIFT jushin kekka       (excel, then HM/MQ per correl,
#            then a Jenkins section, then a NoGfix section)
#    Gfix -> GFIX jushin kekka       (excel, then HM + bold log-header +
#            whole matched log per correl, then a Jenkins section)
#    Df   -> GIFT-vs-GFIX            (per correl: id text + DF snap;
#            order taken from the 'Soushin data' sheet column A)
#
#  Completion: isReplaced |= bit (Gift=1, Gfix=2, Df=4) on every row in
#  the Excel_NAME group, but ONLY when all REQUIRED pieces were inserted.
#  NoGfix snaps are optional; missing ones fail the group by default
#  unless -AllowMissingOptionalNoGfix is given (spec 8.11 / 11).
#  Per-correl misses are written to status\progress.jsonl.
#
#  Usage:
#    .\ReplaceEvidence.ps1 -Mode Gift -WorkDir C:\work\proj
#    .\ReplaceEvidence.ps1 -Mode Gfix -TargetIds JIGPL48S
#    .\ReplaceEvidence.ps1 -Mode Df   -Force
# ============================================================

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Gift','Gfix','Df')]
    [string]$Mode,

    [string]$WorkDir,
    [string]$Owner = ([char]0x53B3),
    [string[]]$TargetIds = @(),
    [switch]$Force,
    [switch]$AllowMissingOptionalNoGfix,

    [string]$CommonScript = '',
    [string]$ExcelHelpersScript = '',

    # Accepted for VerifyTool back-compat (it splats Config.Replace.BlankRowsBetween).
    # Superseded by the per-spec plan spacing in EvidencePlan.ps1 (sections 7/8/9),
    # so it is intentionally not used.
    [int]$BlankRowsBetween = 1,

    # NoGfix section header override. Empty -> ProjectLabels default.
    [string]$GiftNoGfixLabel = '',
    # Kept for VerifyTool back-compat; the bold log header is now the
    # standard ProjectLabels 'GfixLogLabel' so these are unused.
    [string]$GfixLogLabel = '',
    [string]$GfixLogTodoText = ''
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
if (-not (Test-Path -LiteralPath $WorkDir)) { Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1 }

# Capture switches BEFORE dot-source (dot-source rule).
$forceFlag        = [bool]$Force.IsPresent
$allowNoGfixFlag  = [bool]$AllowMissingOptionalNoGfix.IsPresent

# ── dot-source ExcelHelpers (robust path resolve) + shared libs ──
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($ExcelHelpersScript)) { $candidates += $ExcelHelpersScript }
$candidates += @(
    (Join-Path $PSScriptRoot 'ExcelHelpers.ps1'),
    (Join-Path (Split-Path $PSScriptRoot -Parent) 'VerifyTool\ExcelHelpers.ps1')
)
$helpersPath = $null
foreach ($c in $candidates) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).Path; break
    }
}
if (-not $helpersPath) { Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red; exit 1 }
. $helpersPath
. (Join-Path $PSScriptRoot 'MappingStore.ps1')
. (Join-Path $PSScriptRoot 'ProgressLog.ps1')
. (Join-Path $PSScriptRoot 'EvidencePlan.ps1')
. (Join-Path $PSScriptRoot 'ProjectLabels.ps1')
. (Join-Path $PSScriptRoot 'GfixLog.ps1')
. (Join-Path $PSScriptRoot 'EvidenceExecutor.ps1')
if (-not (Get-Command -Name 'New-ExcelApp' -ErrorAction SilentlyContinue)) {
    Write-Host '[ERROR] ExcelHelpers dot-source failed.' -ForegroundColor Red; exit 1
}

$labels = Get-ProjectLabels
$modeCfg = switch ($Mode) {
    'Gift' { @{ Sheet = $labels['SheetGiftRecv']; Bit = 1 } }
    'Gfix' { @{ Sheet = $labels['SheetGfixRecv']; Bit = 2 } }
    'Df'   { @{ Sheet = $labels['SheetDfCompare']; Bit = 4 } }
}

$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
$evDir       = Join-Path $WorkDir 'evidence'

. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

$snapRoot    = Join-Path $WorkDir 'snap'
$logDir      = Join-Path $WorkDir 'log'

Write-Host ''
Write-Host ("===== ReplaceEvidence ({0}) =====" -f $Mode) -ForegroundColor Green
Write-Host ("  WorkDir  : {0}" -f $WorkDir)
Write-Host ("  Sheet    : {0}" -f $modeCfg.Sheet)
Write-Host ("  Bit      : {0}   Force: {1}   AllowMissingNoGfix: {2}" -f $modeCfg.Bit, $forceFlag, $allowNoGfixFlag)
Write-Host ''

if (-not (Test-Path -LiteralPath $mappingPath)) { Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1 }
if (-not (Test-Path -LiteralPath $evDir))       { Write-Host "[ERROR] evidence dir missing: $evDir (run Clone first)" -ForegroundColor Red; exit 1 }

$allRows = Import-Mapping $mappingPath
Ensure-MappingColumns -Rows $allRows | Out-Null
$targets  = ConvertTo-TargetIdList $TargetIds
$workRows = @($allRows | Where-Object { Test-TargetRow $_ $targets })
if ($workRows.Count -eq 0) { Write-Host '[INFO] No rows after filter.' -ForegroundColor Yellow; return }

$groups = $workRows | Group-Object Excel_NAME | Sort-Object Name
Write-Host ("Groups (Excel_NAME): {0}" -f $groups.Count) -ForegroundColor Cyan
Write-ProgressEvent -WorkDir $WorkDir -Phase ("Replace:{0}" -f $Mode) -Action 'start' -Status 'info' `
    -Message ("groups={0} force={1}" -f $groups.Count, $forceFlag)

# Read correl order from the 'Soushin data' sheet col A (spec 7.2). Returns
# $null when the sheet is absent so the caller can fall back to mapping order.
function Read-SoshinDataOrder($wb) {
    $ws = Get-SheetByName $wb $labels['SheetSoshinData']
    if ($null -eq $ws) { return $null }
    $last = 1
    try { $last = [int]$ws.UsedRange.Row + [int]$ws.UsedRange.Rows.Count - 1 } catch { $last = 1 }
    if ($last -gt 5000) { $last = 5000 }
    $vals = [System.Collections.Generic.List[string]]::new()
    for ($r = 1; $r -le $last; $r++) {
        $v = $ws.Cells.Item($r, 1).Value2
        if ($null -ne $v) { $vals.Add([string]$v) }
    }
    return (Select-ValidCorrelIds $vals.ToArray())
}

$excel = New-ExcelApp
$cntDone = 0; $cntSkip = 0; $cntFail = 0

try {
    foreach ($g in $groups) {
        $first = $g.Group | Select-Object -First 1
        $excelName = [string]$first.Excel_NAME
        if ([string]::IsNullOrWhiteSpace($excelName)) { continue }
        $jobName = [string]$first.JOB_NAME
        $toCode  = [string]$first.TO_code

        Write-Host ''
        Write-Host ("=" * 72) -ForegroundColor White
        Write-Host ("  {0}   ({1} correl id)" -f $excelName, $g.Count) -ForegroundColor White
        Write-Host ("=" * 72) -ForegroundColor White

        $wbPath = Find-WorkbookByExcelName -Dir $evDir -ExcelName $excelName
        if ($null -eq $wbPath) {
            Write-Host ("  [SKIP] workbook missing: {0}" -f $wbPath) -ForegroundColor Yellow
            $cntSkip++; continue
        }

        $sampleBits = Get-RowProp $first 'isReplaced'
        if (-not $forceFlag -and (Test-BitDone $sampleBits $modeCfg.Bit)) {
            Write-Host ("  [SKIP] bit {0} already set (isReplaced={1})" -f $modeCfg.Bit, $sampleBits) -ForegroundColor DarkGray
            $cntSkip++; continue
        }

        $correlIds = @($g.Group | ForEach-Object { [string]$_.Correl_ID_S } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $correlToCode = @{}
        foreach ($row in @($g.Group)) {
            $cid = [string]$row.Correl_ID_S
            $rowToCode = [string]$row.TO_code
            if (-not [string]::IsNullOrWhiteSpace($cid) -and -not [string]::IsNullOrWhiteSpace($rowToCode)) {
                $correlToCode[$cid] = $rowToCode
            }
        }

        $wb = $null
        $ok = $false
        $exec = $null
        try {
            $wb = Open-Workbook $excel $wbPath
            Unhide-AllSheets $wb
            $ws = Get-SheetByName $wb $modeCfg.Sheet
            if ($null -eq $ws) {
                Write-Host ("  [FAIL] sheet not found: {0}" -f $modeCfg.Sheet) -ForegroundColor Red
            } else {
                Reset-SheetBelowRow $ws 3 20

                $plan = $null
                switch ($Mode) {
                    'Gift' { $plan = Build-GiftEvidencePlan -SnapRoot $snapRoot -JobName $jobName -CorrelOrder $correlIds }
                    'Gfix' { $plan = Build-GfixEvidencePlan -SnapRoot $snapRoot -JobName $jobName -CorrelOrder $correlIds -ToCode $toCode -CorrelToCode $correlToCode }
                    'Df'   {
                        $dfOrder = Read-SoshinDataOrder $wb
                        if ($null -eq $dfOrder -or @($dfOrder).Count -eq 0) {
                            Write-Host "  [WARN] could not read Soushin-data col A order; using mapping order" -ForegroundColor Yellow
                            $dfOrder = $correlIds
                        }
                        $plan = Build-DfEvidencePlan -SnapRoot $snapRoot -CorrelOrder $dfOrder
                    }
                }

                $exec = Invoke-EvidencePlan -Worksheet $ws -Plan $plan -Labels $labels `
                            -LogDir $logDir -StartRow 3 -Col 2 -GiftNoGfixLabelOverride $GiftNoGfixLabel

                foreach ($w in @($exec.Warnings)) { Write-Host ("  [WARN] {0}" -f $w) -ForegroundColor Yellow }

                $ok = ($exec.MissingRequired.Count -eq 0)
                if ($Mode -eq 'Gift' -and $ok -and (-not $allowNoGfixFlag) -and $exec.MissingOptional.Count -gt 0) {
                    $ok = $false
                    Write-Host ("  [FAIL-STRICT] {0} NoGfix snap(s) missing; pass -AllowMissingOptionalNoGfix to complete anyway." -f $exec.MissingOptional.Count) -ForegroundColor Yellow
                }
            }
            $wb.Save()
        } catch {
            Write-Host ("  [FAIL] processing: {0}" -f $_.Exception.Message) -ForegroundColor Red
            $ok = $false
        } finally {
            Close-Workbook $wb $false
        }

        if ($ok) {
            $groupMs = @($g.Group | ForEach-Object { [string]$_.Correl_ID_M })
            foreach ($r in $allRows) {
                if ($groupMs -contains [string]$r.Correl_ID_M) { Set-MappingBit -Row $r -Field 'isReplaced' -Bit $modeCfg.Bit }
            }
            Write-Host ("  isReplaced |= {0} for {1} row(s)" -f $modeCfg.Bit, $g.Count) -ForegroundColor Green
            Write-ProgressEvent -WorkDir $WorkDir -Phase ("Replace:{0}" -f $Mode) -JobName $jobName `
                -Action 'group' -Status 'ok' -Message ("{0}: inserted={1} log={2}" -f $excelName, $exec.Inserted, $exec.LogMatched)
            $cntDone++
        } else {
            Write-Host '  isReplaced NOT updated (required pieces missing).' -ForegroundColor Yellow
            if ($null -ne $exec) {
                foreach ($m in @($exec.MissingRequired)) {
                    Write-ProgressEvent -WorkDir $WorkDir -Phase ("Replace:{0}" -f $Mode) -CorrelIdS ([string]$m.CorrelIdS) `
                        -JobName $jobName -Action 'missing' -Status 'fail' -Message ("{0}: {1}" -f $m.Folder, $m.Path)
                }
            }
            Write-ProgressEvent -WorkDir $WorkDir -Phase ("Replace:{0}" -f $Mode) -JobName $jobName `
                -Action 'group' -Status 'fail' -Message ("{0}: missing required pieces" -f $excelName)
            $cntFail++
        }
    }

    if ($cntDone -gt 0) {
        Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
        Write-Host ''
        Write-Host ("Mapping saved: {0}" -f $mappingPath) -ForegroundColor DarkGreen
    }
} finally {
    Close-ExcelApp $excel
}

Write-Host ''
Write-Host ("===== ReplaceEvidence ({0}) Done =====" -f $Mode) -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
