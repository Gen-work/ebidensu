#Requires -Version 5.1
# ============================================================
# DeliverFiles.ps1 - Phase DeliverFiles
#
# Copies evidence Excel workbooks from <EvidenceDir> to J4EvidenceDir,
# and copies (or moves) DATA\GFIX and DATA\GIFT files to the
# corresponding J4 data directories.
#
# Completion is tracked per Excel_NAME via the isFilesDelivered column
# in the mapping CSV. Uses net use drive mapping to handle long UNC paths.
#
# Usage:
#   .\DeliverFiles.ps1 -WorkDir C:\work\proj -J4EvidenceDir \\srv\...\JRV
#   .\DeliverFiles.ps1 -WorkDir ... -MoveData      # Move DATA files (delete source)
#   .\DeliverFiles.ps1 -WorkDir ... -Force          # redo already-delivered files
# ============================================================

param(
    [string]$WorkDir       = '',
    [string]$Owner         = '',
    [string[]]$TargetIds   = @(),
    [string]$J4EvidenceDir = '',
    [string]$J4GfixDataDir = '',
    [string]$J4GiftDataDir = '',
    [string]$EvidenceDir   = '',
    [string]$ExcelPrefix   = '',
    [switch]$MoveData,
    [switch]$Force,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$forceFlag    = [bool]$Force.IsPresent
$dryRunFlag   = [bool]$DryRun.IsPresent
$moveDataFlag = [bool]$MoveData.IsPresent

. (Join-Path $PSScriptRoot 'MappingStore.ps1')
. (Join-Path $PSScriptRoot 'ProgressLog.ps1')
. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1
}

if ([string]::IsNullOrWhiteSpace($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir 'evidence' }
elseif (-not [System.IO.Path]::IsPathRooted($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir $EvidenceDir }

$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1
}

if ([string]::IsNullOrWhiteSpace($J4EvidenceDir)) {
    Write-Host '[ERROR] -J4EvidenceDir is required (destination for evidence Excel files).' -ForegroundColor Red; exit 1
}

$targets = @(ConvertTo-TargetIdList $TargetIds)
$allRows  = @(Import-Mapping $mappingPath)
if ($allRows.Count -eq 0) {
    Write-Host "[ERROR] mapping empty: $mappingPath" -ForegroundColor Red; exit 1
}
Ensure-MappingColumns $allRows | Out-Null

# Ensure isFilesDelivered column exists
foreach ($r in $allRows) {
    if (-not ($r.PSObject.Properties.Name -contains 'isFilesDelivered')) {
        $r | Add-Member -NotePropertyName 'isFilesDelivered' -NotePropertyValue '0' -Force
    }
}

# Helper: net use drive mapping for long UNC paths
function Get-FreeDriveLetter2 {
    $used = @(Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Name) + @('A','B','C')
    foreach ($l in 'Z','Y','X','W','V','U','T','S','R','Q','P') {
        if ($used -notcontains $l) { return $l }
    }
    return $null
}

function Copy-LongPath([string]$Src, [string]$DstDir) {
    $mappedLetter2 = $null
    $dstDir2 = $DstDir
    if ($DstDir.Length -gt 200 -and $DstDir -match '^\\\\') {
        $letter2 = Get-FreeDriveLetter2
        if ($letter2) {
            $out = & net use "${letter2}:" $DstDir /persistent:no 2>&1
            if ($LASTEXITCODE -eq 0) {
                $mappedLetter2 = $letter2
                $dstDir2 = "${letter2}:\"
            }
        }
    }
    try {
        $dst = Join-Path $dstDir2 (Split-Path $Src -Leaf)
        Copy-Item -LiteralPath $Src -Destination $dst -Force
        return $dst
    } finally {
        if ($mappedLetter2) { & net use "${mappedLetter2}:" /delete /y 2>&1 | Out-Null }
    }
}

# Resolve data dirs
$localGfixDir = Join-Path $WorkDir "DATA\GFIX"
$localGiftDir = Join-Path $WorkDir "DATA\GIFT"
if ([string]::IsNullOrWhiteSpace($J4GfixDataDir)) { $J4GfixDataDir = Join-Path $J4EvidenceDir "DATA\GFIX" }
if ([string]::IsNullOrWhiteSpace($J4GiftDataDir)) { $J4GiftDataDir = Join-Path $J4EvidenceDir "DATA\GIFT" }

# Build unique Excel_NAME list (mapping order) with target filter
$names = New-Object System.Collections.Generic.List[string]
$prefixByName = @{}
$rowsByName   = @{}
foreach ($r in $allRows) {
    if (-not (Test-TargetRow $r $targets)) { continue }
    $name = [string]$r.Excel_NAME
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    if (-not $names.Contains($name)) {
        $names.Add($name)
        $prefixByName[$name] = Resolve-ExcelPrefix -Row $r -DefaultPrefix $ExcelPrefix
        $rowsByName[$name]   = [System.Collections.Generic.List[object]]::new()
    }
    $rowsByName[$name].Add($r)
}

if ($names.Count -eq 0) {
    Write-Host '[INFO] no target rows.' -ForegroundColor Yellow; return
}

# Collect unique Correl_ID_S for DATA files
$correlIds = @($allRows | Where-Object { Test-TargetRow $_ $targets } |
    Select-Object -ExpandProperty Correl_ID_S | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Select-Object -Unique)

Write-Host ''
Write-Host '===== DeliverFiles =====' -ForegroundColor Green
Write-Host ("  WorkDir       : {0}" -f $WorkDir)
Write-Host ("  EvidenceDir   : {0}" -f $EvidenceDir)
Write-Host ("  J4EvidenceDir : {0}" -f $J4EvidenceDir)
Write-Host ("  J4GfixDataDir : {0}" -f $J4GfixDataDir)
Write-Host ("  J4GiftDataDir : {0}" -f $J4GiftDataDir)
Write-Host ("  Excels        : {0}" -f $names.Count)
Write-Host ("  Correlations  : {0}" -f $correlIds.Count)
Write-Host ("  MoveData      : {0}" -f $moveDataFlag)
Write-Host ("  Force         : {0}" -f $forceFlag)
if ($dryRunFlag) { Write-Host '  [DRY RUN] no files will be copied/moved.' -ForegroundColor Yellow }

$cntOk = 0; $cntSkip = 0; $cntFail = 0

# ---- Phase A: Evidence Excel files ----
Write-Host ''
Write-Host '-- Evidence Excel --' -ForegroundColor Cyan
foreach ($name in $names) {
    $rows = $rowsByName[$name]
    $alreadyDone = ($rows | Where-Object { [string]$_.isFilesDelivered -ne '1' }).Count -eq 0
    if ($alreadyDone -and -not $forceFlag) {
        Write-Host ("  [SKIP] already delivered: {0}" -f $name) -ForegroundColor DarkGray
        $cntSkip++; continue
    }
    $prefix   = $prefixByName[$name]
    $fullStem = Get-ExcelFullStem -Prefix $prefix -Name $name
    $srcFile  = Find-WorkbookByExcelName -Dir $EvidenceDir -ExcelName $fullStem
    if ($null -eq $srcFile) {
        Write-Host ("  [MISS] evidence not found: {0}" -f $fullStem) -ForegroundColor Yellow
        $cntFail++; continue
    }
    $dstLeaf = Split-Path $srcFile -Leaf
    Write-Host ("  -> {0}" -f $dstLeaf) -ForegroundColor White
    if (-not $dryRunFlag) {
        try {
            Copy-LongPath $srcFile $J4EvidenceDir | Out-Null
            foreach ($r in $rows) { $r.isFilesDelivered = '1' }
            Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'DeliverFiles' -JobName $name -Action 'copy' -Status 'ok' -Message $dstLeaf
            Write-Host '     copied to J4EvidenceDir' -ForegroundColor Green
            $cntOk++
        } catch {
            Write-Host ("  [FAIL] {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'DeliverFiles' -JobName $name -Action 'copy' -Status 'fail' -Message $_.Exception.Message
            $cntFail++
        }
    } else { $cntOk++ }
}

# ---- Phase B: DATA\GFIX files ----
Write-Host ''
Write-Host '-- DATA GFIX --' -ForegroundColor Cyan
if (Test-Path -LiteralPath $localGfixDir) {
    foreach ($correl in $correlIds) {
        $files = @(Get-ChildItem -LiteralPath $localGfixDir -Filter ("${correl}*") -File -ErrorAction SilentlyContinue)
        if ($files.Count -eq 0) { continue }
        foreach ($f in $files) {
            Write-Host ("  -> {0}" -f $f.Name) -ForegroundColor White
            if (-not $dryRunFlag) {
                try {
                    Copy-LongPath $f.FullName $J4GfixDataDir | Out-Null
                    if ($moveDataFlag) { Remove-Item -LiteralPath $f.FullName -Force }
                    Write-Host ("     {0}" -f $(if ($moveDataFlag) { 'moved' } else { 'copied' })) -ForegroundColor Green
                    $cntOk++
                } catch {
                    Write-Host ("  [FAIL] {0}: {1}" -f $f.Name, $_.Exception.Message) -ForegroundColor Red
                    $cntFail++
                }
            } else { $cntOk++ }
        }
    }
} else {
    Write-Host ("  (no local GFIX dir: {0})" -f $localGfixDir) -ForegroundColor DarkGray
}

# ---- Phase C: DATA\GIFT files ----
Write-Host ''
Write-Host '-- DATA GIFT --' -ForegroundColor Cyan
if (Test-Path -LiteralPath $localGiftDir) {
    foreach ($correl in $correlIds) {
        $files = @(Get-ChildItem -LiteralPath $localGiftDir -Filter ("${correl}*") -File -ErrorAction SilentlyContinue)
        if ($files.Count -eq 0) { continue }
        foreach ($f in $files) {
            Write-Host ("  -> {0}" -f $f.Name) -ForegroundColor White
            if (-not $dryRunFlag) {
                try {
                    Copy-LongPath $f.FullName $J4GiftDataDir | Out-Null
                    if ($moveDataFlag) { Remove-Item -LiteralPath $f.FullName -Force }
                    Write-Host ("     {0}" -f $(if ($moveDataFlag) { 'moved' } else { 'copied' })) -ForegroundColor Green
                    $cntOk++
                } catch {
                    Write-Host ("  [FAIL] {0}: {1}" -f $f.Name, $_.Exception.Message) -ForegroundColor Red
                    $cntFail++
                }
            } else { $cntOk++ }
        }
    }
} else {
    Write-Host ("  (no local GIFT dir: {0})" -f $localGiftDir) -ForegroundColor DarkGray
}

Write-Host ''
Write-Host '===== DeliverFiles Done =====' -ForegroundColor Green
Write-Host ("  OK   : {0}" -f $cntOk)
Write-Host ("  Skip : {0}" -f $cntSkip)
Write-Host ("  Fail : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
