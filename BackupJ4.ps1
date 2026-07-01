#Requires -Version 5.1
# ============================================================
# BackupJ4.ps1 - Phase BackupJ4 ("bk")
#
# Standalone local backup: for each targeted Excel_NAME, finds the
# CORRESPONDING J4 evidence workbook (in J4EvidenceDir) and copies it as-is
# into a LOCAL backup folder under WorkDir (default <WorkDir>\bk), timestamped
# so repeated runs never overwrite an earlier snapshot. J4 itself is never
# modified or deleted by this phase -- read-only against J4, write-only
# against the local bk folder.
#
# Typical use: run this before DeliverFiles (which now replaces sheets in
# the existing J4 workbook in place) to keep a local rollback copy of J4
# files as they stood before delivery.
#
# Usage:
#   .\BackupJ4.ps1 -WorkDir C:\work\proj -J4EvidenceDir \\srv\...\JRV
#   .\BackupJ4.ps1 -WorkDir ... -TargetIds SJRVWD64
#   .\BackupJ4.ps1 -WorkDir ... -LocalDir C:\work\proj\bk_2026-07-01
# ============================================================

param(
    [string]$WorkDir       = '',
    [string]$Owner         = '',
    [string[]]$TargetIds   = @(),
    [string]$J4EvidenceDir = '',
    [string]$LocalDir      = '',
    [string]$ExcelPrefix   = '',
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$dryRunFlag = [bool]$DryRun.IsPresent

. (Join-Path $PSScriptRoot 'MappingStore.ps1')
. (Join-Path $PSScriptRoot 'ProgressLog.ps1')
. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1
}

$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1
}

if ([string]::IsNullOrWhiteSpace($J4EvidenceDir)) {
    Write-Host '[ERROR] J4 source not configured.' -ForegroundColor Red
    Write-Host '  Set DeliverFiles.J4EvidenceDir (or Mail.EvidenceFolder) in this work' -ForegroundColor Yellow
    Write-Host '  folder''s verify_config.json -- run: .\VerifyTool.ps1 -Phase InitConfig -Interactive' -ForegroundColor Yellow
    Write-Host '  (group "path" or "mail"), or pass -J4EvidenceDir on the command line.' -ForegroundColor Yellow
    exit 1
}

if ([string]::IsNullOrWhiteSpace($LocalDir)) { $LocalDir = Join-Path $WorkDir 'bk' }
elseif (-not [System.IO.Path]::IsPathRooted($LocalDir)) { $LocalDir = Join-Path $WorkDir $LocalDir }

$targets = @(ConvertTo-TargetIdList $TargetIds)
$allRows = @(Import-Mapping $mappingPath)
if ($allRows.Count -eq 0) {
    Write-Host "[ERROR] mapping empty: $mappingPath" -ForegroundColor Red; exit 1
}
Ensure-MappingColumns $allRows | Out-Null

# Helper: net use drive mapping so a long UNC J4 source path can still be read.
function Get-FreeDriveLetterBk {
    $used = @(Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Name) + @('A','B','C')
    foreach ($l in 'Z','Y','X','W','V','U','T','S','R','Q','P') {
        if ($used -notcontains $l) { return $l }
    }
    return $null
}

function Copy-J4FileToLocal([string]$SrcPath, [string]$DestDir) {
    $srcDir  = Split-Path $SrcPath -Parent
    $srcLeaf = Split-Path $SrcPath -Leaf
    $ts      = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $destName = "{0}.{1}.bak" -f $srcLeaf, $ts
    $mappedLetter = $null
    $srcPathEff = $SrcPath
    if ($SrcPath.Length -gt 200 -and $SrcPath -match '^\\\\') {
        $letter = Get-FreeDriveLetterBk
        if ($letter) {
            $out = & net use "${letter}:" $srcDir /persistent:no 2>&1
            if ($LASTEXITCODE -eq 0) {
                $mappedLetter = $letter
                $srcPathEff = Join-Path "${letter}:\" $srcLeaf
            }
        }
    }
    try {
        if (-not (Test-Path -LiteralPath $DestDir)) { New-Item -ItemType Directory -Path $DestDir -Force | Out-Null }
        $dest = Join-Path $DestDir $destName
        Copy-Item -LiteralPath $srcPathEff -Destination $dest -Force
        return $dest
    } finally {
        if ($mappedLetter) { & net use "${mappedLetter}:" /delete /y 2>&1 | Out-Null }
    }
}

# Build unique Excel_NAME list (mapping order) with target filter
$names = New-Object System.Collections.Generic.List[string]
$prefixByName = @{}
foreach ($r in $allRows) {
    if (-not (Test-TargetRow $r $targets)) { continue }
    $name = [string]$r.Excel_NAME
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    if (-not $names.Contains($name)) {
        $names.Add($name)
        $prefixByName[$name] = Resolve-ExcelPrefix -Row $r -DefaultPrefix $ExcelPrefix
    }
}

if ($names.Count -eq 0) {
    Write-Host '[INFO] no target rows.' -ForegroundColor Yellow; return
}

Write-Host ''
Write-Host '===== BackupJ4 =====' -ForegroundColor Green
Write-Host ("  WorkDir       : {0}" -f $WorkDir)
Write-Host ("  J4EvidenceDir : {0}" -f $J4EvidenceDir)
Write-Host ("  LocalDir      : {0}" -f $LocalDir)
Write-Host ("  Excels        : {0}" -f $names.Count)
if ($dryRunFlag) { Write-Host '  [DRY RUN] no files will be copied.' -ForegroundColor Yellow }

$cntOk = 0; $cntMiss = 0; $cntFail = 0

foreach ($name in $names) {
    $prefix   = $prefixByName[$name]
    $fullStem = Get-ExcelFullStem -Prefix $prefix -Name $name
    $j4File   = Find-WorkbookByExcelName -Dir $J4EvidenceDir -ExcelName $fullStem -FullWidthFallback Reject
    if ($null -eq $j4File) {
        Write-Host ("  [MISS] no J4 workbook found: {0}" -f $fullStem) -ForegroundColor Yellow
        $cntMiss++; continue
    }
    Write-Host ("  -> {0}" -f (Split-Path $j4File -Leaf)) -ForegroundColor White
    if ($dryRunFlag) { $cntOk++; continue }
    try {
        $dest = Copy-J4FileToLocal -SrcPath $j4File -DestDir $LocalDir
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'BackupJ4' -JobName $name -Action 'backup' -Status 'ok' -Message (Split-Path $dest -Leaf)
        Write-Host ("     backed up -> {0}" -f (Split-Path $dest -Leaf)) -ForegroundColor Green
        $cntOk++
    } catch {
        Write-Host ("  [FAIL] {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'BackupJ4' -JobName $name -Action 'backup' -Status 'fail' -Message $_.Exception.Message
        $cntFail++
    }
}

Write-Host ''
Write-Host '===== BackupJ4 Done =====' -ForegroundColor Green
Write-Host ("  OK   : {0}" -f $cntOk)
Write-Host ("  Miss : {0}" -f $cntMiss)
Write-Host ("  Fail : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
