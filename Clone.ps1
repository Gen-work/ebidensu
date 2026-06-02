# ============================================================
#  Clone.ps1
#
#  Phase: Clone (a.k.a. mkexcel / renameexcel)
#
#  For each unique Excel_NAME in mapping_<Owner>.csv:
#    1. Try existing file:  <SourceDir>\<BizCode>\<Excel_NAME>.xlsx or *<Excel_NAME>.xlsx
#       (BizCode candidates: TO_code then FROM_code, deduped.
#        Or override with -BizCodes.)
#    2. Fallback to template: <WorkDir>\template_<BizCode>.xlsx
#    3. Universal fallback : <WorkDir>\template.xlsx
#    Dest filename:
#      From SourceDir   -> source filename preserved (e.g. J4..._LJRVWD64.xlsx)
#      Full-stem name   -> <full stem>.xlsx  (J4..._LJRVWD64.xlsx) -- preferred for J4 upload
#      Suffix form (_X) -> X.xlsx  (leading _ stripped)
#      Short stem       -> <stem>.xlsx  (legacy)
#
#  Skip if dest exists, unless -Force.
#
#  Usage:
#    .\Clone.ps1 -WorkDir <work> -SourceDir <external>
#    .\Clone.ps1 -WorkDir <work> -SourceDir <ext> -BizCodes IGP2,ILP2
#    .\Clone.ps1 -WorkDir <work> -TargetIds JIGPL48S -Force
# ============================================================

param(
    [string]$WorkDir,
    [string]$Owner = ([char]0x53B3),
    [string]$SourceDir = '',
    [string[]]$BizCodes = @(),
    [string[]]$TargetIds = @(),
    [switch]$Force,
    [string]$CommonScript = ''
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

# Unblock UNC files
try {
    Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
} catch {}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) { Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1 }

# Resolve switch BEFORE any dot-source
$forceFlag = [bool]$Force.IsPresent

# Dot-source MappingStore EARLY (it has no param() so it is safe) so the
# local single-arg Test-TargetRow defined below overrides MappingStore's
# two-arg version. We use Export-MappingAtomic / Update-MappingRows /
# Ensure-MappingColumns to write the captured Excel_Prefix back.
. (Join-Path $PSScriptRoot 'MappingStore.ps1')

# Optional target narrowing
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

# Header
Write-Host ''
Write-Host '===== Clone (mkexcel) =====' -ForegroundColor Green
Write-Host ("  WorkDir   : {0}" -f $WorkDir)
Write-Host ("  Owner     : {0}" -f $Owner)
Write-Host ("  Force     : {0}" -f $forceFlag)
if ($targetSet.Count -gt 0) { Write-Host ("  TargetIds : {0}" -f (($targetSet.Keys | Sort-Object) -join ', ')) }
if ($BizCodes.Count -gt 0)  { Write-Host ("  BizCodes  : {0} (override)" -f ($BizCodes -join ', ')) }
Write-Host ''

# Prompt for SourceDir if empty
if ([string]::IsNullOrWhiteSpace($SourceDir)) {
    Write-Host 'External SourceDir for existing evidence files.'
    Write-Host '  Expected layout: <SourceDir>\<BizCode>\<Excel_NAME>.xlsx'
    Write-Host '  Enter to skip (use templates only): ' -ForegroundColor Magenta -NoNewline
    $SourceDir = Read-Host
}
$SourceDir = $SourceDir.Trim()
if (-not [string]::IsNullOrWhiteSpace($SourceDir) -and -not (Test-Path -LiteralPath $SourceDir)) {
    Write-Host "[WARN] SourceDir not reachable, will fall back to templates: $SourceDir" -ForegroundColor Yellow
    $SourceDir = ''
}

# Mapping
$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red
    exit 1
}
# $allRows = full set (so prefix write-back never drops non-target rows).
# $rows    = target-filtered references INTO $allRows (used for processing).
$allRows = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)
$rows = @($allRows | Where-Object { Test-TargetRow $_ })
if ($rows.Count -eq 0) {
    Write-Host '[INFO] No rows after filter.' -ForegroundColor Yellow
    return
}

# Evidence dir
$evDir = Join-Path $WorkDir 'evidence'

. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

if (-not (Test-Path -LiteralPath $evDir)) {
    New-Item -ItemType Directory -Path $evDir -Force | Out-Null
}

# Group unique Excel_NAME
$groups = $rows | Group-Object Excel_NAME | Sort-Object Name
Write-Host ("Unique Excel_NAME(s): {0}" -f $groups.Count) -ForegroundColor Cyan
Write-Host ''

$cntDone = 0
$cntSkip = 0
$cntMiss = 0
$cntFromSource = 0
$cntFromTemplate = 0
$cntPrefix = 0
$mappingDirty = $false

foreach ($g in $groups) {
    $first = $g.Group | Select-Object -First 1
    $excelName = [string]$first.Excel_NAME
    if ([string]::IsNullOrWhiteSpace($excelName)) {
        Write-Host '[SKIP] empty Excel_NAME row group' -ForegroundColor DarkGray
        continue
    }
    $excelPrefix = if ($first.PSObject.Properties.Name -contains 'Excel_Prefix') { [string]$first.Excel_Prefix } else { '' }
    $fullStem    = Get-ExcelFullStem -Prefix $excelPrefix -Name $excelName

    $existingPath = Find-WorkbookByExcelName -Dir $evDir -ExcelName $fullStem
    if ($existingPath -and -not $forceFlag) {
        Write-Host ("[SKIP] {0}  (already in evidence: {1})" -f $excelName, (Split-Path $existingPath -Leaf)) -ForegroundColor DarkGray
        $cntSkip++
        continue
    }

    # BizCode candidates
    if ($BizCodes.Count -gt 0) {
        $codes = @($BizCodes)
    } else {
        $codes = @()
        foreach ($k in @('TO_code', 'FROM_code')) {
            if ($first.PSObject.Properties.Name -contains $k) {
                $v = [string]$first.$k
                if (-not [string]::IsNullOrWhiteSpace($v)) { $codes += $v.Trim() }
            }
        }
        $codes = @($codes | Select-Object -Unique)
    }

    # 1) existing under SourceDir
    $foundPath = $null
    $foundFrom = ''
    $ambigCount = 0
    if (-not [string]::IsNullOrWhiteSpace($SourceDir)) {
        $searchDirs = @()
        foreach ($code in $codes) {
            if ([string]::IsNullOrWhiteSpace($code)) { continue }
            $sd = Join-Path $SourceDir $code
            if (Test-Path -LiteralPath $sd) {
                $searchDirs += ,@{ Dir = $sd; Tag = ("source\{0}" -f $code) }
            }
        }
        # Also try flat SourceDir as last resort
        $searchDirs += ,@{ Dir = $SourceDir; Tag = 'source\(flat)' }

        foreach ($sd in $searchDirs) {
            $dir = $sd.Dir
            $tag = $sd.Tag

            # a) exact name: try full stem first (J4..._LJRVWD64.xlsx), then short name
            $exact = Join-Path $dir ("{0}.xlsx" -f $fullStem)
            if (Test-Path -LiteralPath $exact) {
                $foundPath = $exact
                $foundFrom = $tag
                break
            }
            if ($fullStem -ne $excelName) {
                $exact2 = Join-Path $dir ("{0}.xlsx" -f $excelName)
                if (Test-Path -LiteralPath $exact2) {
                    $foundPath = $exact2
                    $foundFrom = $tag
                    break
                }
            }

            # b) suffix wildcard *<short name>.xlsx  (broad discovery)
            $hits = @(Get-ChildItem -LiteralPath $dir -Filter ("*{0}.xlsx" -f $excelName) -File -ErrorAction SilentlyContinue)
            if ($hits.Count -ge 1) {
                $best = $hits | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                $foundPath = $best.FullName
                $foundFrom = ("{0} ~ {1}" -f $tag, $best.Name)
                if ($hits.Count -gt 1) { $ambigCount = $hits.Count }
                break
            }
        }
    }

    # 2) template_<bizcode>.xlsx
    if (-not $foundPath) {
        foreach ($code in $codes) {
            if ([string]::IsNullOrWhiteSpace($code)) { continue }
            $tpl = Join-Path $WorkDir ("template_{0}.xlsx" -f $code)
            if (Test-Path -LiteralPath $tpl) {
                $foundPath = $tpl
                $foundFrom = ("template_{0}" -f $code)
                break
            }
        }
    }

    # 3) universal template.xlsx
    if (-not $foundPath) {
        $tpl = Join-Path $WorkDir 'template.xlsx'
        if (Test-Path -LiteralPath $tpl) {
            $foundPath = $tpl
            $foundFrom = 'template'
        }
    }

    if (-not $foundPath) {
        Write-Host ("[MISS] {0}  (codes tried: {1})" -f $excelName, ($codes -join ',')) -ForegroundColor Red
        $cntMiss++
        continue
    }

    try {
        $destLeaf = if ($foundFrom -like 'source*') { Split-Path $foundPath -Leaf } else { Get-ExcelDestLeaf $fullStem }
        $destPath = if ($existingPath) { $existingPath } else { Join-Path $evDir $destLeaf }
        Copy-Item -LiteralPath $foundPath -Destination $destPath -Force
        if ($foundFrom -like 'source*') {
            $cntFromSource++
            Write-Host ("[COPY] {0}  <- {1}  => {2}" -f $excelName, $foundFrom, (Split-Path $destPath -Leaf)) -ForegroundColor Green
            if ($ambigCount -gt 1) {
                Write-Host ("       NOTE: {0} candidates matched, picked newest" -f $ambigCount) -ForegroundColor DarkYellow
            }

            # Capture the J4 filename prefix from the real source filename so
            # downstream phases (CheckSheet / DeliverMail / DeliverFiles) get
            # the exact <prefix>_<Excel_NAME>.xlsx name. This is the operator's
            # "input point" for Excel_Prefix -- no manual entry needed.
            $capPrefix = Get-PrefixFromFilename -FileName (Split-Path $foundPath -Leaf) -Name $excelName
            if (-not [string]::IsNullOrWhiteSpace($capPrefix) -and $capPrefix -ne $excelPrefix) {
                $n = Update-MappingRows -Rows $allRows -KeyField 'Excel_NAME' -KeyValue $excelName -Updates @{ Excel_Prefix = $capPrefix }
                if ($n -gt 0) {
                    $mappingDirty = $true
                    $cntPrefix++
                    Write-Host ("       Excel_Prefix captured: {0}" -f $capPrefix) -ForegroundColor DarkCyan
                }
            }
        } else {
            $cntFromTemplate++
            Write-Host ("[TPL ] {0}  <- {1}  => {2}" -f $excelName, $foundFrom, (Split-Path $destPath -Leaf)) -ForegroundColor DarkYellow
        }
        $cntDone++
    } catch {
        Write-Host ("[FAIL] {0}  : {1}" -f $excelName, $_.Exception.Message) -ForegroundColor Red
    }
}

# Persist any captured Excel_Prefix values (atomic write; retries if the CSV
# is open in Excel). Existing rows/progress are untouched.
if ($mappingDirty) {
    try {
        Ensure-MappingColumns -Rows $allRows | Out-Null
        Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
        Write-Host ''
        Write-Host ("[INFO] Excel_Prefix captured for {0} Excel_NAME(s) into mapping." -f $cntPrefix) -ForegroundColor Cyan
    } catch {
        Write-Host ("[WARN] could not write Excel_Prefix back to mapping: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

Write-Host ''
Write-Host '===== Clone Done =====' -ForegroundColor Green
Write-Host ("  Done           : {0}" -f $cntDone)
Write-Host ("    from source  : {0}" -f $cntFromSource)
Write-Host ("    from template: {0}" -f $cntFromTemplate)
Write-Host ("  Prefix captured: {0}" -f $cntPrefix)
Write-Host ("  Skipped exists : {0}" -f $cntSkip)
Write-Host ("  Missing source : {0}" -f $cntMiss) -ForegroundColor $(if ($cntMiss -gt 0) { 'Yellow' } else { 'White' })

# Next-step hint
if ($cntMiss -gt 0) {
    Write-Host ''
    Write-Host 'Missing Excel(s) need either:' -ForegroundColor Yellow
    Write-Host '  - <SourceDir>\<bizcode>\*<Excel_NAME>.xlsx  (full J4 name or exact stem), or' -ForegroundColor DarkGray
    Write-Host '  - <WorkDir>\template_<bizcode>.xlsx, or' -ForegroundColor DarkGray
    Write-Host '  - <WorkDir>\template.xlsx (universal fallback)' -ForegroundColor DarkGray
}
if ($cntDone -gt 0 -or $cntSkip -gt 0) {
    Write-Host ''
    Write-Host 'Next:' -ForegroundColor Cyan
    Write-Host '  .\VerifyTool.ps1 -Phase Validate     # see what is ready for Replace' -ForegroundColor DarkGray
    Write-Host '  .\VerifyTool.ps1 -Phase ReplaceGift  # start replacing GIFT side' -ForegroundColor DarkGray
}
