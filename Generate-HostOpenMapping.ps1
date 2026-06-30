# ============================================================
#  Generate-HostOpenMapping.ps1  (Phase 1, v2)
#
#  Generates $WorkDir\mapping_<Owner>.csv for Host-Open VerifyTool.
#
#  Sources (auto-detected in $WorkDir):
#    - *WBS*.xlsx     sheet "WBS"            col A=JOB_NAME  col P=Owner
#    - *GFIX*.xlsx    sheet "GFIX送受信一覧" (columns resolved via headers)
#
#  Filtering (AND when both given):
#    Owner          : WBS col P matches -Owner (no personal default)
#                     counted when: Owner, Owner←*, or *→Owner
#                     example: Owner / Owner←Other / Other→Owner = counted; Other←Owner = not counted
#    FromBizCode    : GFIX from 業務コード == -FromBizCode  (optional)
#    Row range      : WBS rows in [-WbsStartRow, -WbsEndRow]  (optional;
#                     unset -> scan full WBS UsedRange)
#
#  Usage:
#    .\Generate-HostOpenMapping.ps1
#    .\Generate-HostOpenMapping.ps1 -FromBizCode JRV
#    .\Generate-HostOpenMapping.ps1 -FromBizCode JRV -Owner <Owner>
#    .\Generate-HostOpenMapping.ps1 -WbsStartRow 1275 -WbsEndRow 2250
#    .\Generate-HostOpenMapping.ps1 -Force
#
#  Incremental add (grow an existing mapping, keep all done progress):
#    .\Generate-HostOpenMapping.ps1 -Add -JobNames CJODJDEU,CJODJDB5
#    .\Generate-HostOpenMapping.ps1 -Add -CorrelIdsM JIDSC09M
#    .\Generate-HostOpenMapping.ps1 -Add -ExcelNames LJRVWD64
#    .\Generate-HostOpenMapping.ps1 -Add -WbsStartRow 2300 -WbsEndRow 2400
#    -Add appends only Correl_ID_M values not already in mapping_<Owner>.csv;
#    existing rows (and their snap/replace/mark/review state) are untouched.
#    Owner filter composes: explicit JOB_NAME / Correl_ID_M / Excel_NAME
#    selectors are now looked up in the WBS and dropped when their owner cell
#    (col P) belongs to another operator. A JOB_NAME absent from the WBS is
#    kept (temp / not-yet-listed) and reported. (Pure helper: OwnerFilter.ps1.)
#
#  File encoding: UTF-8, NO BOM. All Japanese used at runtime (owner
#  arrows, sheet/label names) is built from [char] code points so PS 5.1
#  on a JP-locale host cannot mis-decode it. See Check-Encoding.ps1.
# ============================================================

param(
    [string]$WorkDir,
    [int]$WbsStartRow = 0,
    [int]$WbsEndRow   = 0,
    [string]$Owner       = '',
    [string]$FromBizCode = "",
    [string[]]$CorrelIdsM = @(),
    [string[]]$JobNames = @(),
    [string[]]$ExcelNames = @(),
    [switch]$AllowTempMapping,
    [switch]$Add,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path $MyInvocation.MyCommand.Path
$forceFlag = [bool]$Force.IsPresent    # capture before dot-source
. (Join-Path $scriptDir 'MappingStore.ps1')
. (Join-Path $scriptDir 'OwnerFilter.ps1')   # pure Test-OwnerMatch / Select-JobsByOwner

# -- Force console to UTF-8 --
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

# -- Interactive fallback --
if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host "WorkDir path" }

function Convert-ToCleanList([object[]]$Values) {
    $out = @()
    foreach ($raw in @($Values)) {
        if ($null -eq $raw) { continue }
        foreach ($part in ([string]$raw -split ',')) {
            $v = $part.Trim()
            if (-not [string]::IsNullOrWhiteSpace($v)) { $out += $v }
        }
    }
    return @($out | Select-Object -Unique)
}

function Read-YesNo([string]$Prompt, [bool]$DefaultYes = $true) {
    $suffix = if ($DefaultYes) { 'Y/n' } else { 'y/N' }
    while ($true) {
        $v = Read-Host ("{0} [{1}]" -f $Prompt, $suffix)
        if ([string]::IsNullOrWhiteSpace($v)) { return $DefaultYes }
        switch ($v.Trim().ToLower()) {
            { $_ -in @('y','yes','1','true') } { return $true }
            { $_ -in @('n','no','0','false') } { return $false }
            default { Write-Host '  please enter y or n.' -ForegroundColor Yellow }
        }
    }
}

$requestedCorrelIdsM = @(Convert-ToCleanList $CorrelIdsM)

# Accept Excel_NAME format (W at index 4, 8 chars) in either -JobNames or -ExcelNames,
# and normalise case so lowercase input works too. Both JJODJDEI and JJODWDEI therefore
# resolve to the same JOB_NAME and hit the same GFIX rows.
$requestedExcelNames = @(Convert-ToCleanList $ExcelNames)
$requestedJobNames = @(
    (@(Convert-ToCleanList $JobNames) + $requestedExcelNames) | ForEach-Object {
        $j = ([string]$_).ToUpper()
        if ($j.Length -eq 8 -and $j[4] -eq 'W') { $j.Substring(0, 4) + 'J' + $j.Substring(5) } else { $j }
    } | Select-Object -Unique
)

$addFlag = [bool]$Add.IsPresent

$useRowRange   = ($WbsStartRow -gt 0 -and $WbsEndRow -gt 0)
$useFromFilter = -not [string]::IsNullOrWhiteSpace($FromBizCode)
$useExplicitTempSelectors = ($requestedCorrelIdsM.Count -gt 0 -or $requestedJobNames.Count -gt 0)

Write-Host ""
Write-Host "===== Generate-HostOpenMapping (Phase 1 v2) =====" -ForegroundColor Green
Write-Host ("  WorkDir     : {0}" -f $WorkDir)
Write-Host ("  Owner       : {0}" -f $Owner)
Write-Host ("  FromBizCode : {0}" -f $(if ($useFromFilter) { $FromBizCode } else { "(none)" }))
Write-Host ("  Row range   : {0}" -f $(if ($useRowRange) { "$WbsStartRow - $WbsEndRow" } else { "(full WBS scan)" }))
Write-Host ("  CorrelIdsM  : {0}" -f $(if ($requestedCorrelIdsM.Count -gt 0) { $requestedCorrelIdsM -join ", " } else { "(none)" }))
Write-Host ("  JobNames    : {0}" -f $(if ($requestedJobNames.Count -gt 0) { $requestedJobNames -join ", " } else { "(none)" }))
Write-Host ("  ExcelNames  : {0}" -f $(if ($requestedExcelNames.Count -gt 0) { $requestedExcelNames -join ", " } else { "(none)" }))
Write-Host ("  TempMapping : {0}" -f $AllowTempMapping.IsPresent)
Write-Host ("  Add (merge) : {0}" -f $addFlag)
Write-Host ("  Force       : {0}" -f $Force.IsPresent)
Write-Host ""

# -- Validate --
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found." -ForegroundColor Red; exit 1
}
if ($useRowRange -and $WbsEndRow -lt $WbsStartRow) {
    Write-Host "[ERROR] WbsEndRow < WbsStartRow." -ForegroundColor Red; exit 1
}

$mappingFileName = "mapping_$Owner.csv"
$mappingPath     = Join-Path $WorkDir $mappingFileName
Write-Host ("[INFO] Output : {0}" -f $mappingFileName)

# -- Find WBS & GFIX files --
function Find-SingleFile([string]$dir, [string]$pattern, [string]$desc) {
    $files = @(Get-ChildItem -LiteralPath $dir -Filter $pattern -File -ErrorAction SilentlyContinue |
                 Where-Object { -not $_.Name.StartsWith("~$") })
    if ($files.Count -eq 0) {
        Write-Host ("[ERROR] No {0} file (pattern: {1})." -f $desc, $pattern) -ForegroundColor Red; exit 1
    }
    if ($files.Count -gt 1) {
        Write-Host ("[ERROR] Multiple {0} files:" -f $desc) -ForegroundColor Red
        $files | ForEach-Object { Write-Host ("        - {0}" -f $_.Name) -ForegroundColor Red }
        exit 1
    }
    return $files[0].FullName
}

$wbsPath  = Find-SingleFile $WorkDir "*WBS*.xlsx"  "WBS"
$gfixPath = Find-SingleFile $WorkDir "*GFIX*.xlsx" "GFIX list"
Write-Host ("[INFO] WBS    : {0}" -f (Split-Path -Leaf $wbsPath))
Write-Host ("[INFO] GFIX   : {0}" -f (Split-Path -Leaf $gfixPath))
Write-Host ""

# -- Japanese label constants --
$LBL_WBS_SHEET  = "WBS"
$LBL_GFIX_SHEET = "GFIX" + [char]0x9001 + [char]0x53D7 + [char]0x4FE1 + [char]0x4E00 + [char]0x89A7  # GFIX送受信一覧
$LBL_SYS_TYPE   = [char]0x30B7 + [char]0x30B9 + [char]0x30C6 + [char]0x30E0 + [char]0x7A2E + [char]0x5225
$LBL_BIZ_CODE   = [char]0x696D + [char]0x52D9 + [char]0x30B3 + [char]0x30FC + [char]0x30C9
$LBL_CORREL_ID  = [char]0x76F8 + [char]0x95A2 + "ID"
$LBL_IF_NUMBER  = "IF" + [char]0x756A + [char]0x53F7
$LBL_EDA        = [char]0x679D + [char]0x756A
$LBL_ZIP_FILE   = [char]0x5727 + [char]0x7E2E + [char]0x30D5 + [char]0x30A1 + [char]0x30A4 + [char]0x30EB
$LBL_SEND       = [char]0x9001 + [char]0x4FE1
$LBL_JOB        = [char]0x30B8 + [char]0x30E7 + [char]0x30D6
$LBL_FROM = "from"
$LBL_TO   = "to"

# -- Helpers --
function Get-ColLetter([int]$c) {
    if ($c -le 0)  { return "?" }
    if ($c -le 26) { return [char]([byte][char]'A' + $c - 1) }
    $a = [char]([byte][char]'A' + [Math]::Floor(($c - 1) / 26) - 1)
    $b = [char]([byte][char]'A' + (($c - 1) % 26))
    return "$a$b"
}
function Read-CellStr($ws, [int]$r, [int]$c) {
    $v = $ws.Cells.Item($r, $c).Value2
    if ($null -eq $v) { return "" }
    return ([string]$v).Trim()
}
function Read-OwnerCell($ws, [int]$r) {
    # Col P (16) with merge-cell awareness
    $cell = $ws.Cells.Item($r, 16)
    if ($cell.MergeCells) {
        $v = $cell.MergeArea.Cells.Item(1, 1).Value2
    } else {
        $v = $cell.Value2
    }
    if ($null -eq $v) { return "" }
    return ([string]$v).Trim()
}

function Add-UniqueJobName($List, $Seen, [string]$JobName) {
    $job = ([string]$JobName).Trim()
    if ([string]::IsNullOrWhiteSpace($job)) { return }
    if ($Seen.Add($job)) { $List.Add($job) }
}

# Test-OwnerMatch moved to OwnerFilter.ps1 (pure, dot-sourced above) so the
# explicit -Add selector path can reuse it and it can be unit-tested.

function Build-WbsJobOwnerMap([string]$wbsFilePath) {
    # Scan the WBS (col A = JOB_NAME, col P = owner) into a JOB_NAME -> owner
    # cell map for owner filtering of explicit -Add selectors. Mirrors Step C
    # semantics: if a JOB_NAME appears on multiple WBS rows, an owner-matching
    # occurrence wins over a non-matching one. Self-contained: opens and closes
    # its own read-only workbook.
    # Bulk-reads both columns at once to avoid per-cell COM calls on large WBS files.
    # Col P may use merged cells; merged interiors read as null -> forward-fill.
    $map = @{}
    $wb = $excel.Workbooks.Open($wbsFilePath, $false, $true)
    try {
        $ws = $null
        foreach ($w in $wb.Worksheets) { if ($w.Name -eq $LBL_WBS_SHEET) { $ws = $w; break } }
        if (-not $ws) { throw ("WBS sheet '{0}' not found." -f $LBL_WBS_SHEET) }
        $ur    = $ws.UsedRange
        $first = $ur.Row
        $last  = $ur.Row + $ur.Rows.Count - 1
        $count = $last - $first + 1
        $colAArr = $ws.Range($ws.Cells.Item($first, 1),  $ws.Cells.Item($last, 1)).Value2
        $colPArr = $ws.Range($ws.Cells.Item($first, 16), $ws.Cells.Item($last, 16)).Value2
        $lastOwner = ""
        for ($i = 1; $i -le $count; $i++) {
            $ownerV = if ($count -eq 1) { $colPArr } else { $colPArr[$i, 1] }
            $ownerCell = if ($null -ne $ownerV -and -not [string]::IsNullOrWhiteSpace([string]$ownerV)) {
                $lastOwner = ([string]$ownerV).Trim(); $lastOwner
            } else { $lastOwner }

            $jobV = if ($count -eq 1) { $colAArr } else { $colAArr[$i, 1] }
            if ($null -eq $jobV) { continue }
            $job = ([string]$jobV).Trim()
            if ([string]::IsNullOrWhiteSpace($job)) { continue }

            if (-not $map.ContainsKey($job)) {
                $map[$job] = $ownerCell
            } elseif (-not (Test-OwnerMatch $map[$job] $Owner) -and (Test-OwnerMatch $ownerCell $Owner)) {
                $map[$job] = $ownerCell
            }
        }
    } finally {
        if ($wb) {
            try { $wb.Close($false) } catch {}
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
        }
    }
    return $map
}

# -- Excel COM --
$excel  = $null
$wbWbs  = $null
$wbGfix = $null

try {
    Write-Host "[*] Starting Excel COM..." -ForegroundColor DarkGray
$excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # -- Suppress add-in / startup workbooks --
    $excel.AutomationSecurity = 3              # msoAutomationSecurityForceDisable
    $excel.AskToUpdateLinks   = $false
    $excel.EnableEvents       = $false

    # ============================================================
    # Step A: Open GFIX and resolve header columns
    # ============================================================
    Write-Host "[Step A] Parsing GFIX header (rows 4 & 5)..." -ForegroundColor Cyan
    $wbGfix = $excel.Workbooks.Open($gfixPath, $false, $true)
    $wsGfix = $null
    foreach ($ws in $wbGfix.Worksheets) {
        if ($ws.Name -eq $LBL_GFIX_SHEET) { $wsGfix = $ws; break }
    }
    if (-not $wsGfix) { throw ("GFIX sheet '{0}' not found." -f $LBL_GFIX_SHEET) }

    $MAX_COL = 30
    $row4 = @{}; $row5 = @{}
    for ($c = 1; $c -le $MAX_COL; $c++) {
        $v4 = $wsGfix.Cells.Item(4, $c).Value2
        $v5 = $wsGfix.Cells.Item(5, $c).Value2
        if ($null -ne $v4) { $row4[$c] = ([string]$v4).Trim() }
        if ($null -ne $v5) { $row5[$c] = ([string]$v5).Trim() }
    }
    $groupMap = @{}; $current = ""
    for ($c = 1; $c -le $MAX_COL; $c++) {
        if ($row4.ContainsKey($c) -and -not [string]::IsNullOrWhiteSpace($row4[$c])) {
            $current = $row4[$c]
        }
        $groupMap[$c] = $current
    }
    $gsTable = @{}
    for ($c = 1; $c -le $MAX_COL; $c++) {
        $g = $groupMap[$c]
        $s = if ($row5.ContainsKey($c)) { $row5[$c] } else { "" }
        $gsTable[$c] = @($g, $s)
    }
    function Find-GS([hashtable]$gs, [string]$g, [string]$s) {
        foreach ($c in ($gs.Keys | Sort-Object)) {
            if ($gs[$c][0] -eq $g -and $gs[$c][1] -eq $s) { return $c }
        }
        return 0
    }

    $col_from_sys  = Find-GS $gsTable $LBL_FROM       $LBL_SYS_TYPE
    $col_from_code = Find-GS $gsTable $LBL_FROM       $LBL_BIZ_CODE
    $col_to_sys    = Find-GS $gsTable $LBL_TO         $LBL_SYS_TYPE
    $col_to_code   = Find-GS $gsTable $LBL_TO         $LBL_BIZ_CODE
    $col_correlid  = Find-GS $gsTable $LBL_CORREL_ID  ""
    $col_ifno      = Find-GS $gsTable $LBL_IF_NUMBER  ""
    $col_eda       = Find-GS $gsTable $LBL_IF_NUMBER  $LBL_EDA
    $col_zip       = Find-GS $gsTable $LBL_ZIP_FILE   ""
    $col_job       = Find-GS $gsTable $LBL_SEND       $LBL_JOB

    Write-Host "  Resolved columns:"
    Write-Host ("    from_sys  : {0,2} ({1})" -f (Get-ColLetter $col_from_sys),  $col_from_sys)
    Write-Host ("    from_code : {0,2} ({1})" -f (Get-ColLetter $col_from_code), $col_from_code)
    Write-Host ("    to_sys    : {0,2} ({1})" -f (Get-ColLetter $col_to_sys),    $col_to_sys)
    Write-Host ("    to_code   : {0,2} ({1})" -f (Get-ColLetter $col_to_code),   $col_to_code)
    Write-Host ("    correl_id : {0,2} ({1})" -f (Get-ColLetter $col_correlid),  $col_correlid)
    Write-Host ("    if_no     : {0,2} ({1})" -f (Get-ColLetter $col_ifno),      $col_ifno)
    Write-Host ("    eda       : {0,2} ({1})" -f (Get-ColLetter $col_eda),       $col_eda)
    Write-Host ("    zip       : {0,2} ({1})" -f (Get-ColLetter $col_zip),       $col_zip)
    Write-Host ("    job       : {0,2} ({1})" -f (Get-ColLetter $col_job),       $col_job)

    $missing = @()
    foreach ($p in @(
        @{n="from_sys";v=$col_from_sys},   @{n="from_code";v=$col_from_code},
        @{n="to_sys";v=$col_to_sys},       @{n="to_code";v=$col_to_code},
        @{n="correl_id";v=$col_correlid},  @{n="if_no";v=$col_ifno},
        @{n="eda";v=$col_eda},             @{n="zip";v=$col_zip},
        @{n="job";v=$col_job}
    )) {
        if ($p.v -eq 0) { $missing += $p.n }
    }
    if ($missing.Count -gt 0) {
        Write-Host ("[ERROR] Missing columns: {0}" -f ($missing -join ", ")) -ForegroundColor Red
        return
    }

    # ============================================================
    # Step B (optional): GFIX from-code filter -> gfixSet of JOB_NAMEs
    # ============================================================
    $gfixSet = New-Object 'System.Collections.Generic.HashSet[System.String]'
    $lastRow = $wsGfix.UsedRange.Rows.Count

    if ($useFromFilter) {
    Write-Host ("[Step B] GFIX filter: from biz_code == '{0}'" -f $FromBizCode) -ForegroundColor Cyan
    $startRow = 6

        $fromRange = $wsGfix.Range(
            $wsGfix.Cells.Item($startRow, $col_from_code),
            $wsGfix.Cells.Item($lastRow,  $col_from_code))
        $jobRange = $wsGfix.Range(
            $wsGfix.Cells.Item($startRow, $col_job),
            $wsGfix.Cells.Item($lastRow,  $col_job))
        $fromArr = $fromRange.Value2
        $jobArr  = $jobRange.Value2
        $rc = $lastRow - $startRow + 1

        for ($i = 1; $i -le $rc; $i++) {
            $fv = if ($rc -eq 1) { $fromArr } else { $fromArr[$i, 1] }
            if ($null -eq $fv) { continue }
            if (([string]$fv).Trim() -ne $FromBizCode) { continue }

            $jv = if ($rc -eq 1) { $jobArr } else { $jobArr[$i, 1] }
            if ($null -eq $jv) { continue }
            $jn = ([string]$jv).Trim()
            if (-not [string]::IsNullOrWhiteSpace($jn)) { [void]$gfixSet.Add($jn) }
        }
        Write-Host ("  GFIX-side JOB_NAMEs for from='{0}' : {1}" -f $FromBizCode, $gfixSet.Count) -ForegroundColor Green
        if ($gfixSet.Count -eq 0) {
            Write-Host "[ABORT] No GFIX rows match FromBizCode." -ForegroundColor Yellow; return
        }
    }

    # ============================================================
    # Step C: Read WBS -> distinct JOB_NAMEs (Owner [∩ gfixSet])
    # ============================================================
    $jobNameList = (New-Object 'System.Collections.Generic.List[System.String]')
    $seen        = New-Object 'System.Collections.Generic.HashSet[System.String]'
    $warnings    = (New-Object 'System.Collections.Generic.List[System.String]')

    if ($useExplicitTempSelectors) {
        Write-Host "[Step C] Explicit Correl_ID_M / JOB_NAME supplied; skip WBS scan." -ForegroundColor Cyan
    } else {
        Write-Host "[Step C] Reading WBS..." -ForegroundColor Cyan
        $wbWbs = $excel.Workbooks.Open($wbsPath, $false, $true)
        $wsWbs = $null
        foreach ($ws in $wbWbs.Worksheets) {
            if ($ws.Name -eq $LBL_WBS_SHEET) { $wsWbs = $ws; break }
        }
        if (-not $wsWbs) { throw ("WBS sheet '{0}' not found." -f $LBL_WBS_SHEET) }

        if (-not $useRowRange) {
            $ur = $wsWbs.UsedRange
            $WbsStartRow = $ur.Row
            $WbsEndRow   = $ur.Row + $ur.Rows.Count - 1
            Write-Host ("  Full UsedRange scan: row {0} - {1}" -f $WbsStartRow, $WbsEndRow)
        } else {
            Write-Host ("  Given row range: row {0} - {1}" -f $WbsStartRow, $WbsEndRow)
        }
        $rowCount = $WbsEndRow - $WbsStartRow + 1

        $colA_range = $wsWbs.Range($wsWbs.Cells.Item($WbsStartRow, 1), $wsWbs.Cells.Item($WbsEndRow, 1))
        $colA = $colA_range.Value2

        $excludedByFromFilter = 0
        $ownerMatched = 0

        for ($r = $WbsStartRow; $r -le $WbsEndRow; $r++) {
            $ownerStr = Read-OwnerCell $wsWbs $r
            if (-not (Test-OwnerMatch $ownerStr $Owner)) { continue }
            $ownerMatched++

            $i = $r - $WbsStartRow + 1
            $job_v = if ($rowCount -eq 1) { $colA } else { $colA[$i, 1] }
            if ($null -eq $job_v) { continue }
            $job = ([string]$job_v).Trim()
            if ([string]::IsNullOrWhiteSpace($job)) { continue }

            if ($useFromFilter -and -not $gfixSet.Contains($job)) {
                $excludedByFromFilter++; continue
            }
            if ($seen.Add($job)) { $jobNameList.Add($job) }
        }

        Write-Host ("  WBS owner-matched rows                 : {0}" -f $ownerMatched) -ForegroundColor Green
        Write-Host ("  Distinct JOB_NAMEs (after all filters) : {0}" -f $jobNameList.Count) -ForegroundColor Green
        if ($useFromFilter) {
            Write-Host ("    Excluded by from-filter            : {0}" -f $excludedByFromFilter) -ForegroundColor DarkGray
        }
        if ($jobNameList.Count -le 30) {
            foreach ($j in $jobNameList) { Write-Host ("    - {0}" -f $j) -ForegroundColor DarkGray }
        } else {
            Write-Host ("    (first 10) " + (($jobNameList | Select-Object -First 10) -join ", ")) -ForegroundColor DarkGray
        }
    }

    if ($jobNameList.Count -eq 0 -or $useExplicitTempSelectors) {
        if ($useExplicitTempSelectors) {
            Write-Host "[INFO] Explicit Correl_ID_M / JOB_NAME list supplied; using GFIX-selected temp JOB_NAMEs instead of WBS list." -ForegroundColor Yellow
        } elseif ($jobNameList.Count -eq 0) {
            Write-Host "[WARN] No matching JOB_NAMEs from WBS." -ForegroundColor Yellow
        }

        $useTemp = $false
        if ($useExplicitTempSelectors) {
            $useTemp = $true
            $jobNameList = (New-Object 'System.Collections.Generic.List[System.String]')
            $seen        = New-Object 'System.Collections.Generic.HashSet[System.String]'
        } elseif ($AllowTempMapping.IsPresent) {
            Write-Host "  WBS may be incomplete. You can create a temporary mapping from GFIX Correl_ID_M or JOB_NAME." -ForegroundColor Yellow
            $useTemp = Read-YesNo "Create temp mapping_$Owner from GFIX directly?" $true
        }

        if (-not $useTemp) {
            Write-Host "[ABORT] No matching JOB_NAMEs." -ForegroundColor Yellow; return
        }

        if ($requestedCorrelIdsM.Count -eq 0 -and $requestedJobNames.Count -eq 0) {
            $rawIds = Read-Host "Correl_ID_M list, comma-separated (example: JIDSC02M,JIDSC03M). Empty = skip"
            $requestedCorrelIdsM = @(Convert-ToCleanList @($rawIds))
            $rawJobs = Read-Host "JOB_NAME list, comma-separated (example: CJODJDEI,CJODJDB7). Empty = skip"
            $requestedJobNames = @(Convert-ToCleanList @($rawJobs))
        }

        if ($requestedCorrelIdsM.Count -eq 0 -and $requestedJobNames.Count -eq 0) {
            Write-Host "[ABORT] No Correl_ID_M or JOB_NAME supplied for temp mapping." -ForegroundColor Yellow; return
        }

        Write-Host "[Step C2] Building temp JOB_NAME list from GFIX..." -ForegroundColor Cyan
        $wantedIds = New-Object 'System.Collections.Generic.HashSet[System.String]'
        foreach ($id in $requestedCorrelIdsM) { [void]$wantedIds.Add($id) }
        foreach ($jn in $requestedJobNames) { Add-UniqueJobName $jobNameList $seen $jn }

        # Only scan GFIX rows when Correl_ID_M selectors were given.
        # Bulk-read both columns at once to avoid per-cell COM calls.
        if ($wantedIds.Count -gt 0) {
            $startRow = 6
            $gfixRowCount = $lastRow - $startRow + 1
            $cidArr = $wsGfix.Range($wsGfix.Cells.Item($startRow, $col_correlid), $wsGfix.Cells.Item($lastRow, $col_correlid)).Value2
            $jobArr2 = $wsGfix.Range($wsGfix.Cells.Item($startRow, $col_job),      $wsGfix.Cells.Item($lastRow, $col_job)).Value2
            $foundIds = New-Object 'System.Collections.Generic.HashSet[System.String]'
            for ($i = 1; $i -le $gfixRowCount; $i++) {
                $cid = if ($gfixRowCount -eq 1) { $cidArr } else { $cidArr[$i, 1] }
                if ($null -eq $cid) { continue }
                $cid = ([string]$cid).Trim()
                if ([string]::IsNullOrWhiteSpace($cid) -or -not $wantedIds.Contains($cid)) { continue }
                [void]$foundIds.Add($cid)
                $jv = if ($gfixRowCount -eq 1) { $jobArr2 } else { $jobArr2[$i, 1] }
                if ($null -eq $jv -or [string]::IsNullOrWhiteSpace(([string]$jv).Trim())) {
                    $warnings.Add(("Correl_ID_M has empty JOB at GFIX row {0}: {1}" -f ($startRow + $i - 1), $cid)); continue
                }
                Add-UniqueJobName $jobNameList $seen ([string]$jv).Trim()
            }
            foreach ($id in $requestedCorrelIdsM) {
                if (-not $foundIds.Contains($id)) { $warnings.Add(("Correl_ID_M not in GFIX list: {0}" -f $id)) }
            }
        }

        Write-Host ("  Temp Owner / mapping name : {0} / mapping_{0}.csv" -f $Owner) -ForegroundColor Green
        Write-Host ("  Temp distinct JOB_NAMEs   : {0}" -f $jobNameList.Count) -ForegroundColor Green
        foreach ($j in $jobNameList) { Write-Host ("    - {0}" -f $j) -ForegroundColor DarkGray }

        # ----------------------------------------------------------------
        # Step C3: compose explicit -Add selectors with the owner filter.
        # The WBS-range path already owner-filters in Step C; explicit
        # JOB_NAME / Correl_ID_M / Excel_NAME selectors used to bypass it.
        # Look each candidate up in the WBS (col A) and keep only those whose
        # owner cell (col P) matches -Owner. A JOB_NAME absent from the WBS is
        # kept (a temp / not-yet-listed job the WBS cannot judge) but reported.
        # ----------------------------------------------------------------
        if (-not [string]::IsNullOrWhiteSpace($Owner) -and $jobNameList.Count -gt 0) {
            Write-Host "[Step C3] Applying owner filter to explicit JOB_NAMEs..." -ForegroundColor Cyan
            $wbsOwnerMap = $null
            try {
                $wbsOwnerMap = Build-WbsJobOwnerMap $wbsPath
            } catch {
                Write-Host ("  [WARN] could not read WBS for owner filter (keeping all): {0}" -f $_.Exception.Message) -ForegroundColor Yellow
            }
            if ($null -ne $wbsOwnerMap) {
                $sel = Select-JobsByOwner -Jobs $jobNameList -JobOwnerMap $wbsOwnerMap -OwnerInput $Owner
                $jobNameList = (New-Object 'System.Collections.Generic.List[System.String]')
                foreach ($j in $sel.Kept) { [void]$jobNameList.Add($j) }
                Write-Host ("  Owner-matched / kept JOB_NAMEs : {0}" -f $sel.Kept.Count) -ForegroundColor Green
                foreach ($ex in $sel.Excluded) {
                    $warnings.Add(("JOB_NAME excluded by owner filter (WBS owner '{0}'): {1}" -f $ex.OwnerCell, $ex.Job))
                    Write-Host ("    - excluded (WBS owner '{0}'): {1}" -f $ex.OwnerCell, $ex.Job) -ForegroundColor DarkGray
                }
                foreach ($uk in $sel.Unknown) {
                    $warnings.Add(("JOB_NAME not in WBS, kept without owner check: {0}" -f $uk))
                }
            }
        }
    }

    if ($jobNameList.Count -eq 0) {
        Write-Host "[ABORT] No matching JOB_NAMEs." -ForegroundColor Yellow; return
    }

    # ============================================================
    # Step D: Build job -> GFIX rows index
    # ============================================================
    Write-Host "[Step D] Locating GFIX rows for matched JOB_NAMEs..." -ForegroundColor Cyan
    $startRow = 6
    $jobRange = $wsGfix.Range($wsGfix.Cells.Item($startRow, $col_job), $wsGfix.Cells.Item($lastRow, $col_job))
    $jobArr   = $jobRange.Value2
    $jobRowCount = $lastRow - $startRow + 1

    $jobToRows = @{}
    foreach ($jn in $jobNameList) { $jobToRows[$jn] = (New-Object 'System.Collections.Generic.List[System.Int32]') }

    for ($i = 1; $i -le $jobRowCount; $i++) {
        $v = if ($jobRowCount -eq 1) { $jobArr } else { $jobArr[$i, 1] }
        if ($null -eq $v) { continue }
        $jc = ([string]$v).Trim()
        if ($jobToRows.ContainsKey($jc)) {
            $jobToRows[$jc].Add($startRow + $i - 1)
        }
    }
    foreach ($jn in $jobNameList) {
        Write-Host ("    {0} -> {1} row(s)" -f $jn, $jobToRows[$jn].Count) -ForegroundColor DarkGray
    }
    $gfixMatchedRows = 0
    foreach ($jn in $jobNameList) { $gfixMatchedRows += $jobToRows[$jn].Count }
    Write-Host ("  GFIX matched rows (total)              : {0}" -f $gfixMatchedRows) -ForegroundColor Green

    # ============================================================
    # Step E: Build mapping records
    # ============================================================
    Write-Host "[Step E] Building mapping records..." -ForegroundColor Cyan

    $records  = (New-Object 'System.Collections.Generic.List[System.Management.Automation.PSObject]')

    foreach ($jn in $jobNameList) {
        $rows = $jobToRows[$jn]
        if ($rows.Count -eq 0) {
            $warnings.Add(("JOB_NAME not in GFIX list: {0}" -f $jn)); continue
        }

        $excelName = $jn
        if ($jn.Length -eq 8 -and $jn[4] -eq 'J') {
            $excelName = $jn.Substring(0, 4) + 'W' + $jn.Substring(5)
        } elseif ($jn.Length -ne 8) {
            $warnings.Add(("JOB_NAME not 8 chars, EXCEL_NAME = JOB_NAME: {0}" -f $jn))
        } else {
            $warnings.Add(("JOB_NAME[5] != 'J', EXCEL_NAME = JOB_NAME: {0}" -f $jn))
        }

        $toCodeSet = New-Object 'System.Collections.Generic.HashSet[System.String]'
        foreach ($r in $rows) {
            $tc = Read-CellStr $wsGfix $r $col_to_code
            if (-not [string]::IsNullOrWhiteSpace($tc)) { [void]$toCodeSet.Add($tc) }
        }
        $isMultiAppl = $toCodeSet.Count
        $amount      = $rows.Count

        foreach ($r in $rows) {
            $correlidM = Read-CellStr $wsGfix $r $col_correlid
            if ([string]::IsNullOrWhiteSpace($correlidM)) {
                $warnings.Add(("Empty Correl_ID at GFIX row {0}" -f $r)); continue
            }
            $correlidS = $correlidM
            if ($correlidM[$correlidM.Length - 1] -eq 'M') {
                $correlidS = $correlidM.Substring(0, $correlidM.Length - 1) + 'S'
            } else {
                $warnings.Add(("Correl_ID doesn't end in 'M': {0}" -f $correlidM))
            }

            $from_sys  = Read-CellStr $wsGfix $r $col_from_sys
            $from_code = Read-CellStr $wsGfix $r $col_from_code
            $to_sys    = Read-CellStr $wsGfix $r $col_to_sys
            $to_code   = Read-CellStr $wsGfix $r $col_to_code
            $ifno      = Read-CellStr $wsGfix $r $col_ifno
            $eda       = Read-CellStr $wsGfix $r $col_eda
            $zipStr    = Read-CellStr $wsGfix $r $col_zip

            $ifFull = if ([string]::IsNullOrWhiteSpace($eda)) { $ifno } else { "{0}-{1}" -f $ifno, $eda }
            $isZip  = if ($zipStr -eq "Yes") { 1 } else { 0 }

            $rec = [pscustomobject][ordered]@{
                Correl_ID_M           = $correlidM
                Correl_ID_S           = $correlidS
                JOB_NAME              = $jn
                Excel_NAME            = $excelName
                FROM_sys              = $from_sys
                FROM_code             = $from_code
                TO_sys                = $to_sys
                TO_code               = $to_code
                IF                    = $ifFull
                Amount                = $amount
                isMultiAppl           = $isMultiAppl
                isZip                 = $isZip
                Excel_snap            = ""
                GIFT_HM_snap          = ""
                GIFT_MQ_snap          = ""
                GIFT_Jenkins_snap     = ""
                GIFT_noGfixfile_snap  = ""
                GFIX_HM_snap          = ""
                GFIX_Jenkins_snap     = ""
                GFIX_log              = ""
                DF_snap               = ""
                isReplaced            = 0
                isMarked              = 0
                isReviewed            = 0
                ReviewComment         = ""
            }
            $records.Add($rec)
        }
    }

    Write-Host ("  Total mapping rows : {0}" -f $records.Count) -ForegroundColor Green
    if ($warnings.Count -gt 0) {
        Write-Host ("  Warnings ({0}):" -f $warnings.Count) -ForegroundColor Yellow
        foreach ($w in $warnings) { Write-Host ("    [WARN] {0}" -f $w) -ForegroundColor Yellow }
    }
    if ($records.Count -eq 0) {
        Write-Host "[ABORT] No records to write." -ForegroundColor Yellow; return
    }

    # ============================================================
    # Step F: Diff with existing mapping_<Owner>.csv
    # ============================================================
    $isOverwrite = Test-Path -LiteralPath $mappingPath

    # ------------------------------------------------------------
    # Step F0 (incremental add): -Add merges the freshly-built records
    # INTO the existing mapping. Existing rows -- and all their progress
    # (snaps / isReplaced / isMarked / isReviewed / delivery flags / comments / legacy columns) --
    # are kept verbatim; only Correl_ID_M values not already present are
    # appended. Use this to grow the mapping day by day (new JOB_NAMEs /
    # Correl_IDs / Excel_NAMEs / WBS range) without losing finished work.
    # ------------------------------------------------------------
    if ($addFlag -and $isOverwrite) {
        Write-Host ""
        Write-Host "===== Add (incremental merge) =====" -ForegroundColor Cyan
        $existing = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)
        $existingM = New-Object 'System.Collections.Generic.HashSet[System.String]'
        foreach ($x in $existing) {
            $k = [string]$x.Correl_ID_M
            if (-not [string]::IsNullOrWhiteSpace($k)) { [void]$existingM.Add($k) }
        }

        $newRecs  = (New-Object 'System.Collections.Generic.List[System.Management.Automation.PSObject]')
        $dupCount = 0
        foreach ($rec in $records) {
            if ($existingM.Contains([string]$rec.Correl_ID_M)) { $dupCount++; continue }
            $newRecs.Add($rec)
        }

        Write-Host ("  Existing rows        : {0}" -f $existing.Count)
        Write-Host ("  Built from selectors : {0}" -f $records.Count)
        Write-Host ("  Already present      : {0}" -f $dupCount) -ForegroundColor DarkGray
        Write-Host ("  New rows to add      : {0}" -f $newRecs.Count) -ForegroundColor $(if ($newRecs.Count) { 'Green' } else { 'DarkGray' })
        foreach ($r in $newRecs) {
            Write-Host ("    + {0}  (JOB: {1}, Excel: {2})" -f $r.Correl_ID_M, $r.JOB_NAME, $r.Excel_NAME) -ForegroundColor Green
        }
        if ($warnings.Count -gt 0) {
            Write-Host ("  Warnings ({0}):" -f $warnings.Count) -ForegroundColor Yellow
            foreach ($w in $warnings) { Write-Host ("    [WARN] {0}" -f $w) -ForegroundColor Yellow }
        }

        if ($newRecs.Count -eq 0) {
            Write-Host "[INFO] Nothing new to add; mapping left unchanged." -ForegroundColor Yellow
            return
        }

        # Union = existing (untouched) + new. Ensure-MappingColumns aligns the
        # schema (adds any status columns the existing file already carries,
        # e.g. isDelivered, to the fresh rows) so the CSV header is complete.
        $combined = (New-Object 'System.Collections.Generic.List[System.Management.Automation.PSObject]')
        foreach ($x in $existing) { $combined.Add($x) }
        foreach ($r in $newRecs)  { $combined.Add($r) }
        Ensure-MappingColumns -Rows $combined.ToArray() | Out-Null

        Export-MappingAtomic -Rows $combined.ToArray() -Path $mappingPath | Out-Null
        Write-Host ("  Saved : {0} ({1} -> {2} rows)" -f $mappingPath, $existing.Count, $combined.Count) -ForegroundColor Green
        return
    }
    if ($addFlag -and -not $isOverwrite) {
        Write-Host "[INFO] -Add given but no existing mapping; creating a fresh one." -ForegroundColor Yellow
    }

    if ($isOverwrite) {
        Write-Host ""
        Write-Host "===== Diff Report =====" -ForegroundColor Cyan
        $old = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)

        $oldIds = New-Object 'System.Collections.Generic.HashSet[System.String]'
        $oldIdToJob = @{}
        foreach ($x in $old) {
            [void]$oldIds.Add($x.Correl_ID_M)
            $oldIdToJob[$x.Correl_ID_M] = $x.JOB_NAME
        }
        $newIds = New-Object 'System.Collections.Generic.HashSet[System.String]'
        $newIdToJob = @{}
        foreach ($x in $records) {
            [void]$newIds.Add($x.Correl_ID_M)
            $newIdToJob[$x.Correl_ID_M] = $x.JOB_NAME
        }

        $added   = @($newIds | Where-Object { -not $oldIds.Contains($_) }) | Sort-Object
        $removed = @($oldIds | Where-Object { -not $newIds.Contains($_) }) | Sort-Object
        $jobChanged = @()
        foreach ($id in $oldIds) {
            if ($newIds.Contains($id) -and $oldIdToJob[$id] -ne $newIdToJob[$id]) {
                $jobChanged += [pscustomobject]@{
                    Correl_ID = $id
                    Old_JOB   = $oldIdToJob[$id]
                    New_JOB   = $newIdToJob[$id]
                }
            }
        }

        Write-Host ("  Old rows : {0}" -f $old.Count)
        Write-Host ("  New rows : {0}" -f $records.Count)
        Write-Host ("  Added (+{0}):" -f $added.Count) -ForegroundColor $(if ($added.Count) { 'Green' } else { 'DarkGray' })
        if ($added.Count -eq 0) { Write-Host "    (none)" -ForegroundColor DarkGray }
        foreach ($id in $added) { Write-Host ("    + {0}  (JOB: {1})" -f $id, $newIdToJob[$id]) -ForegroundColor Green }

        Write-Host ("  Removed (-{0}):" -f $removed.Count) -ForegroundColor $(if ($removed.Count) { 'Red' } else { 'DarkGray' })
        if ($removed.Count -eq 0) { Write-Host "    (none)" -ForegroundColor DarkGray }
        foreach ($id in $removed) { Write-Host ("    - {0}  (JOB: {1})" -f $id, $oldIdToJob[$id]) -ForegroundColor Red }

        Write-Host ("  JOB_NAME changed ({0}):" -f $jobChanged.Count) -ForegroundColor $(if ($jobChanged.Count) { 'Yellow' } else { 'DarkGray' })
        if ($jobChanged.Count -eq 0) { Write-Host "    (none)" -ForegroundColor DarkGray }
        foreach ($x in $jobChanged) { Write-Host ("    * {0}: {1} -> {2}" -f $x.Correl_ID, $x.Old_JOB, $x.New_JOB) -ForegroundColor Yellow }
        Write-Host ""
    }

    # ============================================================
    # Step G: Write CSV (or skip if exists and no -Force)
    # ============================================================
    if ($isOverwrite -and -not $Force) {
        Write-Host "[INFO] mapping file already exists; -Force not set." -ForegroundColor Yellow
        Write-Host ("       Existing file kept: {0}" -f $mappingPath) -ForegroundColor Yellow
        Write-Host "       Re-run with -Force to overwrite." -ForegroundColor Yellow
        return
    }

    # Carry over completion status from the existing mapping so a regenerate
    # (e.g. after WBS/GFIX edits) does NOT wipe finished work. Matched by
    # Correl_ID_M; only non-blank old values overwrite the fresh defaults.
    if ($isOverwrite) {
        $statusCols = @('Excel_snap','GIFT_HM_snap','GIFT_MQ_snap','GIFT_Jenkins_snap',
                        'GIFT_noGfixfile_snap','GFIX_HM_snap','GFIX_Jenkins_snap','GFIX_log',
                        'DF_snap','isReplaced','isMarked','isReviewed','ReviewComment')
        $oldByM = @{}
        foreach ($x in $old) {
            $key = [string]$x.Correl_ID_M
            if (-not [string]::IsNullOrWhiteSpace($key)) { $oldByM[$key] = $x }
        }
        $carried = 0
        foreach ($rec in $records) {
            $key = [string]$rec.Correl_ID_M
            if ($oldByM.ContainsKey($key)) {
                $o = $oldByM[$key]
                foreach ($col in $statusCols) {
                    if ($o.PSObject.Properties.Name -contains $col) {
                        $ov = [string]$o.$col
                        if (-not [string]::IsNullOrWhiteSpace($ov)) { $rec.$col = $ov }
                    }
                }
                $carried++
            }
        }
        Write-Host ("  Carried over status from {0} existing row(s)." -f $carried) -ForegroundColor DarkGray
    }

    Write-Host "[Step G] Writing mapping file..." -ForegroundColor Cyan
    Export-MappingAtomic -Rows $records -Path $mappingPath | Out-Null
    Write-Host ("  Saved : {0}" -f $mappingPath) -ForegroundColor Green

    Write-Host ""
    Write-Host "===== Preview (first 5 rows) =====" -ForegroundColor Cyan
    $records | Select-Object -First 5 Correl_ID_M, Correl_ID_S, JOB_NAME, Excel_NAME, TO_code, IF, Amount, isMultiAppl, isZip |
        Format-Table -AutoSize | Out-String | Write-Host
}
finally {
    if ($wbWbs)  { try { $wbWbs.Close($false)  } catch {} }
    if ($wbGfix) { try { $wbGfix.Close($false) } catch {} }
    if ($excel) {
        # Close any stray workbooks Excel auto-opened
        try {
            foreach ($wb in @($excel.Workbooks)) {
                try { $wb.Close($false) } catch {}
            }
        } catch {}
        try { $excel.Quit() } catch {}
    }

    if ($wbWbs)  { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbWbs)  }
    if ($wbGfix) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbGfix) }
    if ($excel)  { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)  }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "===== Phase 1 v2 Done =====" -ForegroundColor Green