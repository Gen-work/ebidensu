#Requires -Version 5.1
# ============================================================
# DeliverFiles.ps1 - Phase DeliverFiles
#
# Copies evidence Excel workbooks from <EvidenceDir> to J4EvidenceDir,
# and copies DATA\GFIX and DATA\GIFT files to the corresponding J4 data
# directories. Source files are NEVER deleted -- this phase only copies.
#
# Completion is tracked per Excel_NAME via the isFilesDelivered column
# in the mapping CSV. Uses net use drive mapping to handle long UNC paths.
#
# Usage:
#   .\DeliverFiles.ps1 -WorkDir C:\work\proj -J4EvidenceDir \\srv\...\JRV
#   .\DeliverFiles.ps1 -WorkDir ... -SkipData        # evidence Excel only
#   .\DeliverFiles.ps1 -WorkDir ... -SkipExcel       # DATA files only
#   .\DeliverFiles.ps1 -WorkDir ... -Backup          # back up J4 files before overwrite
#   .\DeliverFiles.ps1 -WorkDir ... -Force            # redo already-delivered files
#
# J4 filename ambiguity: J4 sometimes carries a workbook name typed with
# full-width ASCII characters (e.g. a full-width digit instead of a normal
# half-width "0") while the work folder copy is half-width. When a
# same-name-but-width variant is found in
# J4EvidenceDir, the operator is asked whether to remove the J4 variant so
# only the half-width (work folder) name survives -- never both. -Backup
# saves a copy of anything removed/overwritten under J4EvidenceDir\_bak.
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
    [switch]$SkipExcel,
    [switch]$SkipData,
    [switch]$Backup,
    [switch]$Force,
    [switch]$NonInteractive,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$forceFlag         = [bool]$Force.IsPresent
$dryRunFlag        = [bool]$DryRun.IsPresent
$skipExcelFlag     = [bool]$SkipExcel.IsPresent
$skipDataFlag      = [bool]$SkipData.IsPresent
$backupFlag        = [bool]$Backup.IsPresent
$nonInteractiveFlag = [bool]$NonInteractive.IsPresent

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
    Write-Host '[ERROR] J4 destination not configured.' -ForegroundColor Red
    Write-Host '  Set DeliverFiles.J4EvidenceDir (or Mail.EvidenceFolder) in this work' -ForegroundColor Yellow
    Write-Host '  folder''s verify_config.json -- run: .\VerifyTool.ps1 -Phase InitConfig -Interactive' -ForegroundColor Yellow
    Write-Host '  (group "path" or "mail"), or pass -J4EvidenceDir on the command line.' -ForegroundColor Yellow
    exit 1
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

function Copy-LongPath([string]$Src, [string]$DstDir, [string]$DestName = '') {
    if ([string]::IsNullOrWhiteSpace($DestName)) { $DestName = Split-Path $Src -Leaf }
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
        $dst = Join-Path $dstDir2 $DestName
        Copy-Item -LiteralPath $Src -Destination $dst -Force
        return $dst
    } finally {
        if ($mappedLetter2) { & net use "${mappedLetter2}:" /delete /y 2>&1 | Out-Null }
    }
}

# Copies the existing file at $Path into a <dir>\_bak\<name>.<timestamp>.bak
# sidecar before it gets overwritten/removed. Never touches the original.
function Backup-ExistingFile([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return }
    $ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $bakDir = Join-Path (Split-Path $Path -Parent) '_bak'
    if (-not (Test-Path -LiteralPath $bakDir)) { New-Item -ItemType Directory -Path $bakDir -Force | Out-Null }
    $bakName = "{0}.{1}.bak" -f (Split-Path $Path -Leaf), $ts
    Copy-Item -LiteralPath $Path -Destination (Join-Path $bakDir $bakName) -Force
    return (Join-Path $bakDir $bakName)
}

# J4 sometimes already holds the same workbook saved under a full-width-ASCII
# variant of the name (mojibake-adjacent typing, not an OCR/encoding bug).
# Finds such variants next to $DestLeaf in $Dir and, after confirmation,
# removes them so only the half-width (work folder) name remains in J4.
function Remove-FullWidthDuplicates([string]$Dir, [string]$DestLeaf, [string]$DestPath, [string]$ExcelName) {
    $candidates = @(Get-FullWidthWorkbookCandidates -Dir $Dir -ExcelName $DestLeaf)
    foreach ($old in $candidates) {
        if ($old.FullName -ieq $DestPath) { continue }
        Write-Host ("  [WARN] possible full-width duplicate in J4 for {0}: {1}" -f $ExcelName, $old.Name) -ForegroundColor Yellow
        if ($dryRunFlag) {
            Write-Host '  [DRY] would ask to remove this and keep the work-folder (half-width) copy.' -ForegroundColor Yellow
            continue
        }
        if ($nonInteractiveFlag) {
            Write-Host '  [SKIP] non-interactive: leaving both files -- clean up manually.' -ForegroundColor Yellow
            continue
        }
        $resp = Read-Host ("  Replace '{0}' with the work-folder copy? [y/N]" -f $old.Name)
        if ($resp -notmatch '^(?i:y|yes)$') { continue }
        if ($backupFlag) { Backup-ExistingFile $old.FullName | Out-Null }
        Remove-Item -LiteralPath $old.FullName -Force
        Write-Host ("  [OK] removed full-width duplicate: {0}" -f $old.Name) -ForegroundColor Green
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'DeliverFiles' -JobName $ExcelName -Action 'dedup' -Status 'ok' -Message ("removed {0}" -f $old.Name)
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
Write-Host ("  SkipExcel     : {0}" -f $skipExcelFlag)
Write-Host ("  SkipData      : {0}" -f $skipDataFlag)
Write-Host ("  Backup        : {0}" -f $backupFlag)
Write-Host ("  Force         : {0}" -f $forceFlag)
if ($dryRunFlag) { Write-Host '  [DRY RUN] no files will be copied.' -ForegroundColor Yellow }

$cntOk = 0; $cntSkip = 0; $cntFail = 0

# ---- Phase A: Evidence Excel files ----
if (-not $skipExcelFlag) {
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
    if ($null -eq $srcFile -and -not [string]::IsNullOrWhiteSpace($prefix)) {
        # The local evidence copy may predate the configured prefix (or was
        # cloned before Workbook.ExcelPrefix was set). Fall back to the bare
        # Excel_NAME and add the prefix on the J4 copy below.
        $bareFile = Find-WorkbookByExcelName -Dir $EvidenceDir -ExcelName $name
        if ($null -ne $bareFile) {
            Write-Host ("  [INFO] local evidence has no prefix; adding it on the J4 copy: {0}" -f (Split-Path $bareFile -Leaf)) -ForegroundColor DarkGray
            $srcFile = $bareFile
        }
    }
    if ($null -eq $srcFile) {
        Write-Host ("  [MISS] evidence not found: {0}" -f $fullStem) -ForegroundColor Yellow
        $cntFail++; continue
    }
    # Destination name always carries the resolved prefix, regardless of what
    # the source file happens to be named on disk.
    $destLeaf = Get-ExcelDestLeaf $fullStem
    $destPath = Join-Path $J4EvidenceDir $destLeaf
    Write-Host ("  -> {0}" -f $destLeaf) -ForegroundColor White
    if (-not $dryRunFlag) {
        try {
            Remove-FullWidthDuplicates -Dir $J4EvidenceDir -DestLeaf $destLeaf -DestPath $destPath -ExcelName $name
            if ($backupFlag -and (Test-Path -LiteralPath $destPath)) { Backup-ExistingFile $destPath | Out-Null }
            Copy-LongPath $srcFile $J4EvidenceDir $destLeaf | Out-Null
            foreach ($r in $rows) { $r.isFilesDelivered = '1' }
            Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'DeliverFiles' -JobName $name -Action 'copy' -Status 'ok' -Message $destLeaf
            Write-Host '     copied to J4EvidenceDir' -ForegroundColor Green
            $cntOk++
        } catch {
            Write-Host ("  [FAIL] {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'DeliverFiles' -JobName $name -Action 'copy' -Status 'fail' -Message $_.Exception.Message
            $cntFail++
        }
    } else {
        Remove-FullWidthDuplicates -Dir $J4EvidenceDir -DestLeaf $destLeaf -DestPath $destPath -ExcelName $name
        $cntOk++
    }
}
} else {
    Write-Host ''
    Write-Host '-- Evidence Excel -- (skipped)' -ForegroundColor DarkGray
}

# ---- Phase B/C: DATA files ----
if (-not $skipDataFlag) {
foreach ($spec in @(
    @{ Label = 'GFIX'; LocalDir = $localGfixDir; J4Dir = $J4GfixDataDir },
    @{ Label = 'GIFT'; LocalDir = $localGiftDir; J4Dir = $J4GiftDataDir }
)) {
    Write-Host ''
    Write-Host ("-- DATA {0} --" -f $spec.Label) -ForegroundColor Cyan
    if (Test-Path -LiteralPath $spec.LocalDir) {
        foreach ($correl in $correlIds) {
            $files = @(Get-ChildItem -LiteralPath $spec.LocalDir -Filter ("${correl}*") -File -ErrorAction SilentlyContinue)
            if ($files.Count -eq 0) { continue }
            foreach ($f in $files) {
                Write-Host ("  -> {0}" -f $f.Name) -ForegroundColor White
                if (-not $dryRunFlag) {
                    try {
                        $destPath = Join-Path $spec.J4Dir $f.Name
                        if ($backupFlag -and (Test-Path -LiteralPath $destPath)) { Backup-ExistingFile $destPath | Out-Null }
                        Copy-LongPath $f.FullName $spec.J4Dir | Out-Null
                        Write-Host '     copied' -ForegroundColor Green
                        $cntOk++
                    } catch {
                        Write-Host ("  [FAIL] {0}: {1}" -f $f.Name, $_.Exception.Message) -ForegroundColor Red
                        $cntFail++
                    }
                } else { $cntOk++ }
            }
        }
    } else {
        Write-Host ("  (no local {0} dir: {1})" -f $spec.Label, $spec.LocalDir) -ForegroundColor DarkGray
    }
}
} else {
    Write-Host ''
    Write-Host '-- DATA GFIX / GIFT -- (skipped)' -ForegroundColor DarkGray
}

Write-Host ''
Write-Host '===== DeliverFiles Done =====' -ForegroundColor Green
Write-Host ("  OK   : {0}" -f $cntOk)
Write-Host ("  Skip : {0}" -f $cntSkip)
Write-Host ("  Fail : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
