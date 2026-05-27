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
#    Owner          : WBS col P matches -Owner (default: 厳)
#                     counted when: Owner, Owner←*, or *→Owner
#                     example: 厳 / 厳←ニンヌ / 小野→厳 = counted; ニンヌ←厳 = not counted
#    FromBizCode    : GFIX from 業務コード == -FromBizCode  (optional)
#    Row range      : WBS rows in [-WbsStartRow, -WbsEndRow]  (optional;
#                     unset -> scan full WBS UsedRange)
#
#  Usage:
#    .\Generate-HostOpenMapping.ps1
#    .\Generate-HostOpenMapping.ps1 -FromBizCode JRV
#    .\Generate-HostOpenMapping.ps1 -FromBizCode JRV -Owner 厳
#    .\Generate-HostOpenMapping.ps1 -WbsStartRow 1275 -WbsEndRow 2250
#    .\Generate-HostOpenMapping.ps1 -Force
#
#  File encoding: save THIS file as UTF-8 with BOM, CRLF line endings.
# ============================================================

param(
    [string]$WorkDir,
    [int]$WbsStartRow = 0,
    [int]$WbsEndRow   = 0,
    [string]$Owner       = ([char]0x53B3),  # 厳
    [string]$FromBizCode = "",
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# ── Force console to UTF-8 ──
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

# ── Interactive fallback ──
if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host "WorkDir path" }

$useRowRange   = ($WbsStartRow -gt 0 -and $WbsEndRow -gt 0)
$useFromFilter = -not [string]::IsNullOrWhiteSpace($FromBizCode)

Write-Host ""
Write-Host "===== Generate-HostOpenMapping (Phase 1 v2) =====" -ForegroundColor Green
Write-Host ("  WorkDir     : {0}" -f $WorkDir)
Write-Host ("  Owner       : {0}" -f $Owner)
Write-Host ("  FromBizCode : {0}" -f $(if ($useFromFilter) { $FromBizCode } else { "(none)" }))
Write-Host ("  Row range   : {0}" -f $(if ($useRowRange) { "$WbsStartRow - $WbsEndRow" } else { "(full WBS scan)" }))
Write-Host ("  Force       : {0}" -f $Force.IsPresent)
Write-Host ""

# ── Validate ──
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found." -ForegroundColor Red; exit 1
}
if ($useRowRange -and $WbsEndRow -lt $WbsStartRow) {
    Write-Host "[ERROR] WbsEndRow < WbsStartRow." -ForegroundColor Red; exit 1
}

$mappingFileName = "mapping_$Owner.csv"
$mappingPath     = Join-Path $WorkDir $mappingFileName
Write-Host ("[INFO] Output : {0}" -f $mappingFileName)

# ── Find WBS & GFIX files ──
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
$gfixPath = Find-SingleFile $WorkDir "*GFIX*.xlsx" "GFIX一覧"
Write-Host ("[INFO] WBS    : {0}" -f (Split-Path -Leaf $wbsPath))
Write-Host ("[INFO] GFIX   : {0}" -f (Split-Path -Leaf $gfixPath))
Write-Host ""

# ── Japanese label constants ──
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

# ── Helpers ──
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

function Test-OwnerMatch([string]$ownerCell, [string]$ownerInput) {
    # Match the current owner represented by -Owner.
    # Count only these patterns:
    #   1) exact owner              : 厳
    #   2) owner followed by ←...   : 厳←ニンヌ / 厳←小野
    #   3) ... followed by →owner   : 小野→厳
    # Do NOT count reverse direction such as ニンヌ←厳.
    if ([string]::IsNullOrWhiteSpace($ownerCell) -or [string]::IsNullOrWhiteSpace($ownerInput)) {
        return $false
    }

    $cell  = $ownerCell.Trim()
    $input = $ownerInput.Trim()

    if ($cell -eq $input) { return $true }
    if ($cell.StartsWith($input + "←")) { return $true }
    if ($cell.EndsWith("→" + $input)) { return $true }

    return $false
}

# ── Excel COM ──
$excel  = $null
$wbWbs  = $null
$wbGfix = $null

try {
    Write-Host "[*] Starting Excel COM..." -ForegroundColor DarkGray
$excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # ── Suppress add-in / startup workbooks ──
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
    $gfixSet = New-Object 'System.Collections.Generic.HashSet[string]'
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

    $jobNames = [System.Collections.Generic.List[string]]::new()
    $seen     = New-Object 'System.Collections.Generic.HashSet[string]'
    $excludedByFromFilter = 0

    for ($r = $WbsStartRow; $r -le $WbsEndRow; $r++) {
        $ownerStr = Read-OwnerCell $wsWbs $r
        if (-not (Test-OwnerMatch $ownerStr $Owner)) { continue }

        $i = $r - $WbsStartRow + 1
        $job_v = if ($rowCount -eq 1) { $colA } else { $colA[$i, 1] }
        if ($null -eq $job_v) { continue }
        $job = ([string]$job_v).Trim()
        if ([string]::IsNullOrWhiteSpace($job)) { continue }

        if ($useFromFilter -and -not $gfixSet.Contains($job)) {
            $excludedByFromFilter++; continue
        }
        if ($seen.Add($job)) { $jobNames.Add($job) }
    }

    Write-Host ("  Distinct JOB_NAMEs (after all filters) : {0}" -f $jobNames.Count) -ForegroundColor Green
    if ($useFromFilter) {
        Write-Host ("    Excluded by from-filter            : {0}" -f $excludedByFromFilter) -ForegroundColor DarkGray
    }
    if ($jobNames.Count -le 30) {
        foreach ($j in $jobNames) { Write-Host ("    - {0}" -f $j) -ForegroundColor DarkGray }
    } else {
        Write-Host ("    (first 10) " + (($jobNames | Select-Object -First 10) -join ", ")) -ForegroundColor DarkGray
    }
    if ($jobNames.Count -eq 0) {
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
    foreach ($jn in $jobNames) { $jobToRows[$jn] = [System.Collections.Generic.List[int]]::new() }

    for ($i = 1; $i -le $jobRowCount; $i++) {
        $v = if ($jobRowCount -eq 1) { $jobArr } else { $jobArr[$i, 1] }
        if ($null -eq $v) { continue }
        $jc = ([string]$v).Trim()
        if ($jobToRows.ContainsKey($jc)) {
            $jobToRows[$jc].Add($startRow + $i - 1)
        }
    }
    foreach ($jn in $jobNames) {
        Write-Host ("    {0} -> {1} row(s)" -f $jn, $jobToRows[$jn].Count) -ForegroundColor DarkGray
    }

    # ============================================================
    # Step E: Build mapping records
    # ============================================================
    Write-Host "[Step E] Building mapping records..." -ForegroundColor Cyan

    $records  = [System.Collections.Generic.List[psobject]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    foreach ($jn in $jobNames) {
        $rows = $jobToRows[$jn]
        if ($rows.Count -eq 0) {
            $warnings.Add(("JOB_NAME not in GFIX一覧: {0}" -f $jn)); continue
        }

        $excelName = $jn
        if ($jn.Length -eq 8 -and $jn[4] -eq 'J') {
            $excelName = $jn.Substring(0, 4) + 'W' + $jn.Substring(5)
        } elseif ($jn.Length -ne 8) {
            $warnings.Add(("JOB_NAME not 8 chars, EXCEL_NAME = JOB_NAME: {0}" -f $jn))
        } else {
            $warnings.Add(("JOB_NAME[5] != 'J', EXCEL_NAME = JOB_NAME: {0}" -f $jn))
        }

        $toCodeSet = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($r in $rows) {
            $tc = Read-CellStr $wsGfix $r $col_to_code
            if (-not [string]::IsNullOrWhiteSpace($tc)) { [void]$toCodeSet.Add($tc) }
        }
        $isMultiAppl = $toCodeSet.Count
        $amount      = $rows.Count

        foreach ($r in $rows) {
            $correlidM = Read-CellStr $wsGfix $r $col_correlid
            if ([string]::IsNullOrWhiteSpace($correlidM)) {
                $warnings.Add(("Empty 相関ID at GFIX row {0}" -f $r)); continue
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
                isReviewed            = 0
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

    if ($isOverwrite) {
        Write-Host ""
        Write-Host "===== Diff Report =====" -ForegroundColor Cyan
        $old = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)

        $oldIds = New-Object 'System.Collections.Generic.HashSet[string]'
        $oldIdToJob = @{}
        foreach ($x in $old) {
            [void]$oldIds.Add($x.Correl_ID_M)
            $oldIdToJob[$x.Correl_ID_M] = $x.JOB_NAME
        }
        $newIds = New-Object 'System.Collections.Generic.HashSet[string]'
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

    Write-Host "[Step G] Writing mapping file..." -ForegroundColor Cyan
    $records | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force
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
`
``n
--- File: GfixLogDownload.ps1 ---