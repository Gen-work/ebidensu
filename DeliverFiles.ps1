#Requires -Version 5.1
# ============================================================
# DeliverFiles.ps1 - Phase DeliverFiles
#
# Delivers evidence to J4: for each targeted Excel_NAME, replaces the three
# delivery-scope sheets (GIFT recv result / GFIX recv result / GIFT-vs-GFIX
# data compare -- the operator's own captured evidence, same set as
# Align.ps1's Get-AlignRecvSheets) in the CORRESPONDING J4 workbook with the
# matching sheets from the work evidence workbook, in place -- like Align,
# but in the opposite direction (work -> J4) and only for these three
# sheets. Every other J4 sheet (host-managed send sheets, etc.) is left
# untouched. If J4 does not have a workbook for this Excel_NAME yet (first
# delivery), the whole work evidence file is copied in as a bootstrap
# (same as this phase's old, pre-rework behavior).
#
# Also copies DATA\GFIX and DATA\GIFT files (including each side's unzip
# subfolder -- DfSnap isZip extractions -- into the matching J4
# DATA\...\unzip folder) to the corresponding J4 data directories. Source
# files are NEVER deleted -- this phase only copies / in-place
# sheet-replaces.
#
# Completion is tracked per Excel_NAME via the isFilesDelivered column
# in the mapping CSV. Uses net use drive mapping to handle long UNC paths.
#
# Usage:
#   .\DeliverFiles.ps1 -WorkDir C:\work\proj -J4EvidenceDir \\srv\...\JRV
#   .\DeliverFiles.ps1 -WorkDir ... -SkipData        # evidence Excel only
#   .\DeliverFiles.ps1 -WorkDir ... -SkipExcel       # DATA files only
#   .\DeliverFiles.ps1 -WorkDir ... -Backup          # back up the whole J4 file before its sheets are replaced
#   .\DeliverFiles.ps1 -WorkDir ... -Force            # redo already-delivered files
#
# NOTE: -Backup only covers the run's own J4EvidenceDir\_bak safety copy.
# To keep a LOCAL rollback copy of J4 files before running this phase, use
# the standalone BackupJ4 phase (-Phase BackupJ4) first.
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
. (Join-Path $PSScriptRoot 'ExcelHelpers.ps1')
. (Join-Path $PSScriptRoot 'ProjectLabels.ps1')

# The three delivery-scope sheets: same set as Align.ps1's Get-AlignRecvSheets
# (the operator's own captured evidence, never the host-managed send sheets).
$deliverSheetNames = @(Get-AlignRecvSheets)

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
    Write-Host '  Set J4EvidenceDir in this work folder''s verify_config.json -- run:' -ForegroundColor Yellow
    Write-Host '  .\VerifyTool.ps1 -Phase InitConfig -Interactive  (group "path" or "mail"),' -ForegroundColor Yellow
    Write-Host '  or pass -J4EvidenceDir on the command line.' -ForegroundColor Yellow
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

# Sheet-name match tolerant of stray whitespace / full-width ASCII (mirrors
# Align.ps1's Get-AlignSheetMatch -- kept local since Align.ps1 has a param()
# block and cannot be dot-sourced).
function Get-DeliverSheetMatch($wb, [string]$name) {
    $ws = Get-SheetByName $wb $name
    if ($null -ne $ws) { return $ws }
    $target = (Convert-FullWidthAsciiToHalfWidth $name).Trim()
    foreach ($s in $wb.Worksheets) {
        $cand = (Convert-FullWidthAsciiToHalfWidth ([string]$s.Name)).Trim()
        if ($cand -eq $target) { return $s }
    }
    return $null
}

# Opens $workPath for reading. Work and J4 copies of an Excel_NAME are named
# identically by design, and Excel cannot hold two same-named workbooks open
# at once. Since J4 (opened by the caller) must be saved in place, a temp
# copy of the WORK file is opened instead when the names collide.
function Open-WorkForDeliverRead($excel, [string]$workPath, [string]$j4Path) {
    $workLeaf = [System.IO.Path]::GetFileName($workPath)
    $j4Leaf   = [System.IO.Path]::GetFileName($j4Path)
    $samePath = $false
    try { $samePath = ([System.IO.Path]::GetFullPath($workPath) -eq [System.IO.Path]::GetFullPath($j4Path)) } catch {}
    if (-not ($samePath -or ($workLeaf -eq $j4Leaf))) {
        return @{ Wb = (Open-Workbook $excel $workPath); TempPath = $null }
    }
    $ext      = [System.IO.Path]::GetExtension($workPath)
    $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("verify_deliver_{0}{1}" -f ([guid]::NewGuid().ToString('N')), $ext)
    Copy-Item -LiteralPath $workPath -Destination $tempPath -Force
    return @{ Wb = (Open-Workbook $excel $tempPath); TempPath = $tempPath }
}

# Copies $srcWs into $destWb, replacing the existing same-named sheet in
# place (kept at its original position), or appending it at the end when J4
# does not have it yet. Mirrors Align.ps1's Sync-Sheet, direction reversed
# (work -> J4).
function Copy-SheetIntoJ4($destWb, $destWsOld, $srcWs) {
    $sheetName = $srcWs.Name
    if ($null -ne $destWsOld) {
        $srcWs.Copy($destWsOld)              # Before := the existing J4 sheet
        $newWs = $destWb.ActiveSheet
        $destWsOld.Delete()
    } else {
        $srcWs.Copy([System.Reflection.Missing]::Value, $destWb.Sheets.Item($destWb.Sheets.Count))
        $newWs = $destWb.ActiveSheet
    }
    if ($newWs.Name -ne $sheetName) { try { $newWs.Name = $sheetName } catch {} }
    return $newWs
}

# Replaces the delivery-scope sheets ($SheetNames) in the J4 workbook at
# $J4Path with the matching sheets from the work evidence workbook at
# $WorkPath, saves J4, and reports per-sheet lines. Other J4 sheets are
# never touched.
function Sync-DeliverSheets($excel, [string]$WorkPath, [string]$J4Path, [string[]]$SheetNames) {
    $workWb = $null; $j4Wb = $null; $workTemp = $null
    $lines = [System.Collections.Generic.List[string]]::new()
    $changed = $false
    try {
        $workOpen = Open-WorkForDeliverRead $excel $WorkPath $J4Path
        $workWb   = $workOpen.Wb
        $workTemp = $workOpen.TempPath
        $j4Wb     = Open-Workbook $excel $J4Path
        Unhide-AllSheets $workWb
        Unhide-AllSheets $j4Wb
        foreach ($sheetName in $SheetNames) {
            $wsWork = Get-DeliverSheetMatch $workWb $sheetName
            if ($null -eq $wsWork) {
                $lines.Add(("     [MISS] {0} not found in work evidence -- skipped" -f $sheetName))
                continue
            }
            $wsJ4Old = Get-DeliverSheetMatch $j4Wb $sheetName
            Copy-SheetIntoJ4 $j4Wb $wsJ4Old $wsWork | Out-Null
            $changed = $true
            $lines.Add(("     [OK]   {0} {1}" -f $sheetName, $(if ($wsJ4Old) { 'replaced' } else { 'added' })))
        }
        if ($changed) { $j4Wb.Save() }
    } finally {
        Close-Workbook $j4Wb $false
        Close-Workbook $workWb $false
        if (-not [string]::IsNullOrWhiteSpace($workTemp)) {
            try { Remove-Item -LiteralPath $workTemp -Force -ErrorAction SilentlyContinue } catch {}
        }
    }
    return @{ Changed = $changed; Lines = $lines.ToArray() }
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
Write-Host ("  (delivery sheets: {0})" -f ($deliverSheetNames -join ' / ')) -ForegroundColor DarkGray
$excel = $null
if (-not $dryRunFlag) { $excel = New-ExcelApp }
try {
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
            $existingJ4 = $null
            if (Test-Path -LiteralPath $destPath) { $existingJ4 = $destPath }
            else { $existingJ4 = Find-WorkbookByExcelName -Dir $J4EvidenceDir -ExcelName $fullStem -FullWidthFallback Reject }
            if ($null -ne $existingJ4) {
                if ($backupFlag) { Backup-ExistingFile $existingJ4 | Out-Null }
                $result = Sync-DeliverSheets -excel $excel -WorkPath $srcFile -J4Path $existingJ4 -SheetNames $deliverSheetNames
                foreach ($ln in $result.Lines) {
                    $color = if ($ln -match '\[MISS\]') { 'Yellow' } else { 'Green' }
                    Write-Host $ln -ForegroundColor $color
                }
                if (-not $result.Changed) {
                    Write-Host '     [WARN] no delivery sheets matched in work evidence -- J4 file left unchanged' -ForegroundColor Yellow
                }
            } else {
                Write-Host '     [INFO] no existing J4 workbook found -- first delivery: copying whole file' -ForegroundColor DarkGray
                Copy-LongPath $srcFile $J4EvidenceDir $destLeaf | Out-Null
            }
            foreach ($r in $rows) { $r.isFilesDelivered = '1' }
            Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'DeliverFiles' -JobName $name -Action 'copy' -Status 'ok' -Message $destLeaf
            Write-Host '     delivered to J4EvidenceDir' -ForegroundColor Green
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
} finally {
    if ($null -ne $excel) { Close-ExcelApp $excel }
}
} else {
    Write-Host ''
    Write-Host '-- Evidence Excel -- (skipped)' -ForegroundColor DarkGray
}

# ---- Phase B/C: DATA files ----
# Each side also delivers its unzip subfolder (DfSnap isZip extractions,
# named after the correl id) into the matching J4 DATA\...\unzip folder.
# The unzip specs are Optional: absent locally on most rows, so no dir is
# no message; their J4 subfolder is created on first delivery.
if (-not $skipDataFlag) {
foreach ($spec in @(
    @{ Label = 'GFIX';       LocalDir = $localGfixDir;                   J4Dir = $J4GfixDataDir;                   Optional = $false },
    @{ Label = 'GFIX unzip'; LocalDir = (Join-Path $localGfixDir 'unzip'); J4Dir = (Join-Path $J4GfixDataDir 'unzip'); Optional = $true },
    @{ Label = 'GIFT';       LocalDir = $localGiftDir;                   J4Dir = $J4GiftDataDir;                   Optional = $false },
    @{ Label = 'GIFT unzip'; LocalDir = (Join-Path $localGiftDir 'unzip'); J4Dir = (Join-Path $J4GiftDataDir 'unzip'); Optional = $true }
)) {
    if ($spec.Optional -and -not (Test-Path -LiteralPath $spec.LocalDir)) { continue }
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
                        if ($spec.Optional -and -not (Test-Path -LiteralPath $spec.J4Dir)) {
                            New-Item -ItemType Directory -Path $spec.J4Dir -Force | Out-Null
                        }
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
