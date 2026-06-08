# ============================================================
#  Validate.ps1
#
#  Phase: Validate (read-only diagnostic)
#
#  Scans the WorkDir without touching Excel and reports:
#    - mapping_<owner>.csv presence and required columns
#    - isReplaced bitmask distribution
#    - Directory structure (evidence, snap/*, DATA, log, templates)
#    - Per Excel_NAME readiness matrix for Clone / ReplaceGift /
#      ReplaceGfix / ReplaceDf
#    - Aggregate readiness totals
#    - Next suggested action
#
#  Pure read-only -- does not modify mapping, evidence, or session.
#
#  Usage:
#    .\Validate.ps1
#    .\Validate.ps1 -TargetIds JIGPL48S
#    .\Validate.ps1 -Compact            # hide the per-Excel matrix
# ============================================================

param(
    [string]$WorkDir,
    [string]$Owner = '',
    [string[]]$TargetIds = @(),
    [switch]$Compact,
    [string]$CommonScript = ''
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1
}

# Target filter
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

Write-Host ''
Write-Host '===== Validate =====' -ForegroundColor Cyan
Write-Host ("  WorkDir : {0}" -f $WorkDir)
Write-Host ("  Owner   : {0}" -f $Owner)
if ($targetSet.Count -gt 0) { Write-Host ("  Targets : {0}" -f (($targetSet.Keys | Sort-Object) -join ', ')) }

# ── Mapping ─────────────────────────────────────────────────
$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
Write-Host ''
Write-Host '-- Mapping --' -ForegroundColor White
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host ("  File          : MISSING ({0})" -f $mappingPath) -ForegroundColor Red
    Write-Host '  Cannot validate further. Run -Phase Mapping first.' -ForegroundColor Yellow
    return
}
$rows = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)
Write-Host ("  File          : OK ({0} rows)" -f $rows.Count) -ForegroundColor Green

$first = $rows | Select-Object -First 1
if ($null -eq $first) {
    Write-Host '  Mapping has no rows.' -ForegroundColor Yellow
    return
}

$requiredCols = @('Correl_ID_M','Correl_ID_S','JOB_NAME','Excel_NAME','TO_code','FROM_code')
$missingCols = @()
foreach ($c in $requiredCols) {
    if (-not ($first.PSObject.Properties.Name -contains $c)) { $missingCols += $c }
}
if ($missingCols.Count -eq 0) {
    Write-Host '  Required cols : OK' -ForegroundColor Green
} else {
    Write-Host ("  Required cols : MISSING {0}" -f ($missingCols -join ', ')) -ForegroundColor Red
}

# Status bitmask distribution helper
function Show-BitStatus([string]$FieldName, [object[]]$Rows) {
    $bitDist = @{}
    foreach ($r in $Rows) {
        $v = 0
        try { $v = [int]$r.$FieldName } catch { $v = 0 }
        if ($v -lt 0 -or $v -gt 7) { $v = -1 }
        if (-not $bitDist.ContainsKey($v)) { $bitDist[$v] = 0 }
        $bitDist[$v]++
    }
    $labels = @{
        0 = 'none'
        1 = 'Gift'
        2 = 'Gfix'
        3 = 'Gift+Gfix'
        4 = 'Df'
        5 = 'Gift+Df'
        6 = 'Gfix+Df'
        7 = 'all'
       -1 = 'invalid'
    }
    $parts = @()
    foreach ($k in (-1, 0, 1, 2, 3, 4, 5, 6, 7)) {
        if ($bitDist.ContainsKey($k) -and $bitDist[$k] -gt 0) {
            $parts += ("{0}={1}" -f $labels[$k], $bitDist[$k])
        }
    }
    return ($parts -join ', ')
}

# isReplaced bitmask distribution
$hasIsReplaced = $first.PSObject.Properties.Name -contains 'isReplaced'
if ($hasIsReplaced) {
    Write-Host ("  isReplaced    : {0}" -f (Show-BitStatus 'isReplaced' $rows))
} else {
    Write-Host '  isReplaced    : column not yet present (added on first Replace run)' -ForegroundColor DarkGray
}

# isMarked bitmask distribution
$hasIsMarked = $first.PSObject.Properties.Name -contains 'isMarked'
if ($hasIsMarked) {
    Write-Host ("  isMarked      : {0}" -f (Show-BitStatus 'isMarked' $rows))
} else {
    Write-Host '  isMarked      : column not yet present (added on first Mark run)' -ForegroundColor DarkGray
}

# isReviewed bitmask distribution
$hasIsReviewed = $first.PSObject.Properties.Name -contains 'isReviewed'
if ($hasIsReviewed) {
    Write-Host ("  isReviewed    : {0}" -f (Show-BitStatus 'isReviewed' $rows))
} else {
    Write-Host '  isReviewed    : column not yet present (added on first Review run)' -ForegroundColor DarkGray
}

# isDelivered flag distribution (0/1, not a bitmask)
$hasIsDelivered = $first.PSObject.Properties.Name -contains 'isDelivered'
if ($hasIsDelivered) {
    $delivDist = @{}
    foreach ($r in $rows) {
        $v = [string]$r.isDelivered
        if ([string]::IsNullOrWhiteSpace($v)) { $v = '0' }
        if (-not $delivDist.ContainsKey($v)) { $delivDist[$v] = 0 }
        $delivDist[$v]++
    }
    $parts = @()
    foreach ($k in ('0','1','')) {
        if ($delivDist.ContainsKey($k) -and $delivDist[$k] -gt 0) {
            $label = if ($k -eq '1') { 'sent' } elseif ($k -eq '0') { 'pending' } else { 'blank' }
            $parts += ("{0}={1}" -f $label, $delivDist[$k])
        }
    }
    Write-Host ("  isDelivered   : {0}" -f ($parts -join ', '))
} else {
    Write-Host '  isDelivered   : column not yet present (added on first DeliverMail run)' -ForegroundColor DarkGray
}

# isFilesDelivered flag distribution (0/1, per Excel_NAME)
$hasIsFilesDelivered = $first.PSObject.Properties.Name -contains 'isFilesDelivered'
if ($hasIsFilesDelivered) {
    $fileDist = @{}
    foreach ($r in $rows) {
        $v = [string]$r.isFilesDelivered
        if ([string]::IsNullOrWhiteSpace($v)) { $v = '0' }
        if (-not $fileDist.ContainsKey($v)) { $fileDist[$v] = 0 }
        $fileDist[$v]++
    }
    $parts = @()
    foreach ($k in ('0','1','')) {
        if ($fileDist.ContainsKey($k) -and $fileDist[$k] -gt 0) {
            $label = if ($k -eq '1') { 'copied' } elseif ($k -eq '0') { 'pending' } else { 'blank' }
            $parts += ("{0}={1}" -f $label, $fileDist[$k])
        }
    }
    Write-Host ("  isFilesDeliv. : {0}" -f ($parts -join ', '))
} else {
    Write-Host '  isFilesDeliv. : column not yet present (added on first DeliverFiles run)' -ForegroundColor DarkGray
}

# Target-filtered workset
$workRows = @($rows | Where-Object { Test-TargetRow $_ })
$groups = $workRows | Group-Object Excel_NAME |
          Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) } |
          Sort-Object Name
Write-Host ("  Excel_NAMEs   : {0} unique{1}" -f $groups.Count, $(if ($targetSet.Count -gt 0) { ' (after filter)' } else { '' }))

# ── Filesystem ──────────────────────────────────────────────
Write-Host ''
Write-Host '-- Filesystem --' -ForegroundColor White
$dirsToCheck = @(
    @{ Path='evidence';             Required=$true;  Note='destination of Clone' }
    @{ Path='snap';                 Required=$true;  Note='' }
    @{ Path='snap\excel';           Required=$false; Note='used by Gift/Gfix Replace' }
    @{ Path='snap\GIFT_HM';         Required=$false; Note='used by ReplaceGift' }
    @{ Path='snap\GIFT_MQ';         Required=$false; Note='used by ReplaceGift' }
    @{ Path='snap\GIFT_Jenkins';    Required=$false; Note='used by ReplaceGift' }
    @{ Path='snap\GIFT_noGfixfile'; Required=$false; Note='used by ReplaceGift tail (may stay empty)' }
    @{ Path='snap\GFIX_HM';         Required=$false; Note='used by ReplaceGfix' }
    @{ Path='snap\GFIX_Jenkins';    Required=$false; Note='used by ReplaceGfix' }
    @{ Path='snap\DF';              Required=$false; Note='used by ReplaceDf' }
    @{ Path='DATA';                 Required=$false; Note='Jenkins download root' }
    @{ Path='log';                  Required=$false; Note='planned for GfixLodDownload' }
)
foreach ($d in $dirsToCheck) {
    $full = Join-Path $WorkDir $d.Path
    if (Test-Path -LiteralPath $full) {
        $pngCount  = @(Get-ChildItem -LiteralPath $full -Filter '*.png'  -File -ErrorAction SilentlyContinue).Count
        $xlsxCount = @(Get-ChildItem -LiteralPath $full -Filter '*.xlsx' -File -ErrorAction SilentlyContinue).Count
        $extra = @()
        if ($pngCount  -gt 0) { $extra += ("{0} PNG"  -f $pngCount) }
        if ($xlsxCount -gt 0) { $extra += ("{0} XLSX" -f $xlsxCount) }
        $tail = if ($extra.Count -gt 0) { '  ' + ($extra -join ', ') } else { '' }
        Write-Host ("  {0,-24}: OK{1}" -f $d.Path, $tail) -ForegroundColor Green
    } else {
        if ($d.Required) {
            Write-Host ("  {0,-24}: MISSING ({1})" -f $d.Path, $d.Note) -ForegroundColor Red
        } else {
            $note = if (-not [string]::IsNullOrWhiteSpace($d.Note)) { ("  ({0})" -f $d.Note) } else { '' }
            Write-Host ("  {0,-24}: absent{1}" -f $d.Path, $note) -ForegroundColor DarkGray
        }
    }
}

# Templates
$templates = @(Get-ChildItem -LiteralPath $WorkDir -Filter 'template*.xlsx' -File -ErrorAction SilentlyContinue)
Write-Host ''
Write-Host '-- Templates --' -ForegroundColor White
if ($templates.Count -gt 0) {
    foreach ($t in $templates) { Write-Host ("  {0}" -f $t.Name) -ForegroundColor Green }
} else {
    Write-Host '  none found in WorkDir' -ForegroundColor DarkGray
    Write-Host '  Clone will rely entirely on -CloneSourceDir if no template_<bizcode>.xlsx exists.' -ForegroundColor DarkGray
}

# ── Per Excel_NAME readiness ────────────────────────────────
$snapBase = Join-Path $WorkDir 'snap'
$evDir    = Join-Path $WorkDir 'evidence'

. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

function Has-Image([string]$folder, [string]$key) {
    if ([string]::IsNullOrWhiteSpace($folder) -or [string]::IsNullOrWhiteSpace($key)) { return $false }
    $f = Join-Path (Join-Path $snapBase $folder) ("{0}.png" -f $key)
    return [bool](Test-Path -LiteralPath $f)
}

$totals = @{ Clone=0; ReplaceGift=0; ReplaceGfix=0; ReplaceDf=0 }
$counts = @{ Clone=0; ReplaceGift=0; ReplaceGfix=0; ReplaceDf=0 }
$detailRows = @()

foreach ($g in $groups) {
    $f       = $g.Group | Select-Object -First 1
    $exName      = [string]$f.Excel_NAME
    $exPrefix    = if ($f.PSObject.Properties.Name -contains 'Excel_Prefix') { [string]$f.Excel_Prefix } else { '' }
    $fullStem    = Get-ExcelFullStem -Prefix $exPrefix -Name $exName
    $jobName = [string]$f.JOB_NAME
    $cidList = @($g.Group | ForEach-Object { [string]$_.Correl_ID_S } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    $cloneReady = $null -ne (Find-WorkbookByExcelName -Dir $evDir -ExcelName $fullStem)

    $excelOk        = Has-Image 'excel' $jobName
    $giftHmOk       = $true
    $giftMqOk       = $true
    $giftJenkinsOk  = $true
    $gfixHmOk       = $true
    $gfixJenkinsOk  = $true
    $dfOk           = $true
    foreach ($cid in $cidList) {
        if (-not (Has-Image 'GIFT_HM'      $cid)) { $giftHmOk      = $false }
        if (-not (Has-Image 'GIFT_MQ'      $cid)) { $giftMqOk      = $false }
        if (-not (Has-Image 'GIFT_Jenkins' $cid)) { $giftJenkinsOk = $false }
        if (-not (Has-Image 'GFIX_HM'      $cid)) { $gfixHmOk      = $false }
        if (-not (Has-Image 'GFIX_Jenkins' $cid)) { $gfixJenkinsOk = $false }
        if (-not (Has-Image 'DF'           $cid)) { $dfOk          = $false }
    }

    $giftReady = $cloneReady -and $excelOk -and $giftHmOk -and $giftMqOk -and $giftJenkinsOk
    $gfixReady = $cloneReady -and $excelOk -and $gfixHmOk -and $gfixJenkinsOk
    $dfReady   = $cloneReady -and $dfOk

    $counts.Clone++;       if ($cloneReady) { $totals.Clone++ }
    $counts.ReplaceGift++; if ($giftReady)  { $totals.ReplaceGift++ }
    $counts.ReplaceGfix++; if ($gfixReady)  { $totals.ReplaceGfix++ }
    $counts.ReplaceDf++;   if ($dfReady)    { $totals.ReplaceDf++ }

    # Per-row missing summary
    $miss = @()
    if (-not $cloneReady)    { $miss += 'evidence' }
    if (-not $excelOk)       { $miss += 'excel' }
    if (-not $giftHmOk)      { $miss += 'gHM' }
    if (-not $giftMqOk)      { $miss += 'gMQ' }
    if (-not $giftJenkinsOk) { $miss += 'gJk' }
    if (-not $gfixHmOk)      { $miss += 'GHM' }
    if (-not $gfixJenkinsOk) { $miss += 'GJk' }
    if (-not $dfOk)          { $miss += 'DF' }

    $detailRows += [PSCustomObject]@{
        ExcelName = $exName
        Count     = $g.Count
        Clone     = if ($cloneReady) { 'Y' } else { 'N' }
        Gift      = if ($giftReady)  { 'Y' } else { '-' }
        Gfix      = if ($gfixReady)  { 'Y' } else { '-' }
        Df        = if ($dfReady)    { 'Y' } else { '-' }
        Missing   = ($miss -join ',')
    }
}

# ── Detail matrix ───────────────────────────────────────────
if (-not $Compact.IsPresent -and $detailRows.Count -gt 0) {
    Write-Host ''
    Write-Host '-- Per Excel_NAME readiness --' -ForegroundColor White
    Write-Host ('  legend: gHM=GIFT_HM gMQ=GIFT_MQ gJk=GIFT_Jenkins  GHM=GFIX_HM GJk=GFIX_Jenkins') -ForegroundColor DarkGray
    Write-Host ('')
    $hdrFmt = "  {0,-34} {1,3}  {2,5} {3,4} {4,4} {5,3}   {6}"
    Write-Host ($hdrFmt -f 'Excel_NAME', 'cnt', 'Clone', 'Gift', 'Gfix', 'Df', 'missing') -ForegroundColor DarkGray
    Write-Host ("  {0}" -f ('-' * 74)) -ForegroundColor DarkGray
    foreach ($d in $detailRows) {
        $name = $d.ExcelName
        if ($name.Length -gt 34) { $name = $name.Substring(0, 31) + '...' }
        $color = if ($d.Clone -eq 'Y' -and $d.Gift -eq 'Y' -and $d.Gfix -eq 'Y' -and $d.Df -eq 'Y') { 'Green' }
                 elseif ($d.Clone -eq 'Y' -and ($d.Gift -eq 'Y' -or $d.Gfix -eq 'Y' -or $d.Df -eq 'Y')) { 'Yellow' }
                 elseif ($d.Clone -eq 'Y') { 'DarkYellow' }
                 else { 'DarkGray' }
        Write-Host ($hdrFmt -f $name, $d.Count, $d.Clone, $d.Gift, $d.Gfix, $d.Df, $d.Missing) -ForegroundColor $color
    }
}

# ── Totals ──────────────────────────────────────────────────
Write-Host ''
Write-Host '-- Phase totals --' -ForegroundColor White
foreach ($k in 'Clone','ReplaceGift','ReplaceGfix','ReplaceDf') {
    $t = $totals[$k]; $c = $counts[$k]
    $color = if ($c -eq 0) { 'DarkGray' }
             elseif ($t -eq $c) { 'Green' }
             elseif ($t -eq 0) { 'DarkGray' }
             else { 'Yellow' }
    Write-Host ("  {0,-13}: {1}/{2} ready" -f $k, $t, $c) -ForegroundColor $color
}

# ── Next suggested action ───────────────────────────────────
Write-Host ''
Write-Host '-- Next suggested action --' -ForegroundColor White
if ($totals.Clone -lt $counts.Clone) {
    $miss = $counts.Clone - $totals.Clone
    Write-Host ("  -> Run Clone (still {0} Excel(s) missing in work\evidence\)" -f $miss) -ForegroundColor Cyan
    Write-Host '       .\VerifyTool.ps1 -Phase Clone -CloneSourceDir <path>' -ForegroundColor DarkGray
} elseif ($totals.ReplaceGift -lt $counts.ReplaceGift) {
    Write-Host '  -> Take missing GIFT-side snaps, then ReplaceGift' -ForegroundColor Cyan
    Write-Host '       .\VerifyTool.ps1 -Phase GiftHmSnap' -ForegroundColor DarkGray
    Write-Host '       .\VerifyTool.ps1 -Phase GiftMqSnap' -ForegroundColor DarkGray
    Write-Host '       .\VerifyTool.ps1 -Phase GiftJenkins' -ForegroundColor DarkGray
    Write-Host '       .\VerifyTool.ps1 -Phase ReplaceGift' -ForegroundColor DarkGray
} elseif ($totals.ReplaceGfix -lt $counts.ReplaceGfix) {
    Write-Host '  -> Take GFIX-side snaps, then ReplaceGfix' -ForegroundColor Cyan
    Write-Host '       .\VerifyTool.ps1 -Phase GfixHmSnap' -ForegroundColor DarkGray
    Write-Host '       .\VerifyTool.ps1 -Phase GfixJenkins' -ForegroundColor DarkGray
    Write-Host '       .\VerifyTool.ps1 -Phase ReplaceGfix' -ForegroundColor DarkGray
} elseif ($totals.ReplaceDf -lt $counts.ReplaceDf) {
    Write-Host '  -> DfSnap (planned), then ReplaceDf' -ForegroundColor Cyan
} else {
    Write-Host '  All replace phases ready or complete. -> ReviewEvidence' -ForegroundColor Green
    Write-Host '       .\VerifyTool.ps1 -Phase ReviewEvidence' -ForegroundColor DarkGray
}
Write-Host ''
