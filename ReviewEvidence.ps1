# ============================================================
# ReviewEvidence.ps1 - manual visual review driver
#
# Opens evidence workbooks one by one in editable mode. Misaki reviews/adjusts by hand.
#
# Important behavior:
#   - Opening a workbook does NOT move sheets/cursor, save, close, or mark reviewed.
#   - Only after the user returns to this shell and presses Enter, the script:
#       1. places cursor on each sheet from last to first
#          - A1 if sheet name contains the "=>" Excel sheet arrow
#          - otherwise CursorCell, default A3
#       2. sends Ctrl+S, waits, sends Esc to dismiss GenBa comment prompt
#       3. closes the workbook
#       4. updates mapping isReviewed bitmask
#
# Read-only handling:
#   - The script requests read-write open explicitly.
#   - "Read-only recommended" is ignored.
#   - If the file has the Windows ReadOnly attribute, the attribute is cleared before open.
#   - If Excel still opens it as read-only, the item fails instead of silently saving nothing.
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner = ([char]0x53B3),
    [string]$EvidenceDir = '',
    [string]$CursorCell = 'A3',
    [string]$ReviewField = 'isReviewed',
    [int]$ReviewBit = 7,
    [string[]]$TargetIds = @(),
    [int]$SaveWaitMs = 1000,
    [switch]$Force,
    [switch]$DryRun,
    [switch]$Maximize,
    [string]$ExcelHelpersScript = ''
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

$forceFlag = [bool]$Force.IsPresent
$dryRunFlag = [bool]$DryRun.IsPresent
$maximizeFlag = [bool]$Maximize.IsPresent

# ── Dot-source ExcelHelpers.ps1 ─────────────────────────────
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($ExcelHelpersScript)) { $candidates += $ExcelHelpersScript }
$candidates += @(
    (Join-Path $PSScriptRoot 'ExcelHelpers.ps1')
)

$helpersPath = $null
foreach ($c in $candidates) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).Path
        break
    }
}
if (-not $helpersPath) {
    Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red
    exit 1
}
. $helpersPath

function Test-TargetRow($Row, [hashtable]$TargetSet) {
    if ($TargetSet.Count -eq 0) { return $true }
    return ($TargetSet.ContainsKey([string]$Row.Correl_ID_S) -or
            $TargetSet.ContainsKey([string]$Row.Correl_ID_M) -or
            $TargetSet.ContainsKey([string]$Row.JOB_NAME) -or
            $TargetSet.ContainsKey([string]$Row.Excel_NAME))
}

function Test-BitDone([array]$Rows, [string]$Field, [int]$Bit) {
    if ($Rows.Count -eq 0) { return $false }
    foreach ($r in $Rows) {
        $v = Get-BitValue $r $Field
        if (($v -band $Bit) -ne $Bit) { return $false }
    }
    return $true
}

function Get-EvidencePath([string]$Dir, [string]$ExcelName) {
    if ($ExcelName -match '\.xlsx$') { return (Join-Path $Dir $ExcelName) }
    return (Join-Path $Dir ($ExcelName + '.xlsx'))
}

function Get-ExcelProcessId($Excel) {
    try {
        $hwnd = [IntPtr]([int]$Excel.Hwnd)
        $p = Get-Process | Where-Object { $_.MainWindowHandle -eq $hwnd } | Select-Object -First 1
        if ($p) { return [int]$p.Id }
    } catch {}
    return 0
}

function Activate-ExcelWindow($Shell, $Excel, $Workbook) {
    $excelPid = Get-ExcelProcessId $Excel
    if ($excelPid -gt 0) {
        try {
            if ($Shell.AppActivate($excelPid)) { return $true }
        } catch {}
    }
    try {
        if ($Workbook -and $Shell.AppActivate([string]$Workbook.Name)) { return $true }
    } catch {}
    try {
        if ($Shell.AppActivate('Excel')) { return $true }
    } catch {}
    return $false
}


function Clear-FileReadOnlyAttribute([string]$File) {
    try {
        $item = Get-Item -LiteralPath $File -ErrorAction Stop
        if (($item.Attributes -band [System.IO.FileAttributes]::ReadOnly) -ne 0) {
            $item.Attributes = ($item.Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly))
            Write-Host '  [INFO] Windows ReadOnly attribute cleared before open.' -ForegroundColor DarkGray
        }
    } catch {
        Write-Host ("  [WARN] failed to clear Windows ReadOnly attribute: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

function Open-WorkbookEditable($Excel, [string]$File) {
    Clear-FileReadOnlyAttribute $File

    $missing = [Type]::Missing

    # Workbooks.Open parameters used here:
    #   Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword,
    #   IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify,
    #   Converter, AddToMru, Local
    #
    # Points:
    #   - ReadOnly=$false requests editable open.
    #   - IgnoreReadOnlyRecommended=$true avoids Excel's "read-only recommended" prompt/default.
    #   - Notify=$false prevents Excel from opening a locked workbook as read-only while waiting.
    #   - AddToMru=$false avoids polluting Recent files.
    return $Excel.Workbooks.Open(
        $File,
        0,
        $false,
        $missing,
        $missing,
        $missing,
        $true,
        $missing,
        $missing,
        $true,
        $false,
        $missing,
        $false,
        $true
    )
}

function Read-ReviewChoice([datetime]$OpenedAt) {
    $choice = Read-Host '  Press Enter=save+close+mark, s=skip(no mark), q=quit'

    # Safety guard:
    # If blank Enter is received immediately after open, it is often an accidental/stale key input.
    # Ask once more so the workbook does not close right after opening.
    if ([string]::IsNullOrWhiteSpace($choice)) {
        $elapsed = ((Get-Date) - $OpenedAt).TotalSeconds
        if ($elapsed -lt 2) {
            Write-Host '  [WARN] Enter was received immediately after open. Confirm once more to close this workbook.' -ForegroundColor Yellow
            $choice = Read-Host '  Press Enter again=save+close+mark, s=skip(no mark), q=quit'
        }
    }

    return $choice
}

function Set-ReviewCursorAllSheets($Workbook, [string]$DefaultCell) {
    $arrow = [char]0x21D2
    $count = 0
    try { $count = [int]$Workbook.Worksheets.Count } catch { return }

    for ($i = $count; $i -ge 1; $i--) {
        $ws = $Workbook.Worksheets.Item($i)
        $cell = $DefaultCell
        try {
            if ([string]$ws.Name -like ('*' + $arrow + '*')) { $cell = 'A1' }
            $ws.Activate() | Out-Null
            $ws.Range($cell).Select() | Out-Null
        } catch {
            Write-Host ("  [WARN] failed to set cursor: sheet={0}, cell={1}" -f $ws.Name, $cell) -ForegroundColor Yellow
        }
    }
}

function Save-BySendKeys($Shell, $Excel, $Workbook, [int]$WaitMs) {
    if ($Workbook.ReadOnly) {
        throw "Workbook is read-only. Save is blocked: $($Workbook.FullName)"
    }

    [void](Activate-ExcelWindow $Shell $Excel $Workbook)
    Start-Sleep -Milliseconds 200

    # GenBa macro prompt appears immediately after Ctrl+S.
    # Dismiss it first, then wait for network-share save to settle.
    $Shell.SendKeys('^s')
    Start-Sleep -Milliseconds 300
    $Shell.SendKeys('{ESC}')
    Start-Sleep -Milliseconds ([Math]::Max(5000, $WaitMs))
}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir 'evidence' }
elseif (-not [System.IO.Path]::IsPathRooted($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir $EvidenceDir }

if (-not (Test-Path -LiteralPath $EvidenceDir)) {
    Write-Host "[ERROR] EvidenceDir not found: $EvidenceDir" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($ReviewField)) { $ReviewField = 'isReviewed' }
if ($ReviewBit -le 0) { $ReviewBit = 7 }

$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red
    exit 1
}

$targetSet = @{}
foreach ($rawId in @($TargetIds)) {
    if ($null -eq $rawId) { continue }
    foreach ($part in ($rawId.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $targetSet[$v] = $true }
    }
}

$allRows = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)
if ($allRows.Count -eq 0) {
    Write-Host "[ERROR] mapping has no rows: $mappingPath" -ForegroundColor Red
    exit 1
}
Ensure-Column $allRows $ReviewField '0'

$selectedRows = @($allRows | Where-Object { Test-TargetRow $_ $targetSet })
if ($selectedRows.Count -eq 0) {
    Write-Host '[INFO] no target rows.' -ForegroundColor Yellow
    return
}

$excelNames = New-Object System.Collections.Generic.List[string]
foreach ($r in $selectedRows) {
    $name = [string]$r.Excel_NAME
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    if (-not $excelNames.Contains($name)) { $excelNames.Add($name) }
}

Write-Host ''
Write-Host '===== ReviewEvidence =====' -ForegroundColor Green
Write-Host ("  WorkDir     : {0}" -f $WorkDir)
Write-Host ("  Mapping     : {0}" -f $mappingPath)
Write-Host ("  EvidenceDir : {0}" -f $EvidenceDir)
Write-Host ("  Field       : {0}" -f $ReviewField)
Write-Host ("  ReviewBit   : {0}" -f $ReviewBit)
Write-Host ("  Workbooks   : {0}" -f $excelNames.Count)
Write-Host ("  Force       : {0}" -f $forceFlag)
if ($targetSet.Count -gt 0) { Write-Host ("  TargetIds   : {0}" -f (($targetSet.Keys | Sort-Object) -join ', ')) }

if ($dryRunFlag) {
    foreach ($name in $excelNames) {
        $groupRows = @($allRows | Where-Object { [string]$_.Excel_NAME -eq $name })
        $done = Test-BitDone $groupRows $ReviewField $ReviewBit
        Write-Host ("  [DRY] {0} rows={1} done={2}" -f $name, $groupRows.Count, $done)
    }
    return
}

$excel = $null
$shell = $null
$cntDone = 0
$cntSkip = 0
$cntMissing = 0
$cntFail = 0

try {
    $excel = New-ExcelApp
    $excel.Visible = $true
    try { $excel.DisplayAlerts = $false } catch {}
    try { $excel.ScreenUpdating = $true } catch {}
    if ($maximizeFlag) {
        try { $excel.WindowState = -4137 } catch {} # xlMaximized
    }

    $shell = New-Object -ComObject WScript.Shell

    for ($idx = 0; $idx -lt $excelNames.Count; $idx++) {
        $name = $excelNames[$idx]
        $groupRows = @($allRows | Where-Object { [string]$_.Excel_NAME -eq $name })

        if ((Test-BitDone $groupRows $ReviewField $ReviewBit) -and -not $forceFlag) {
            Write-Host ("[{0}/{1}] SKIP reviewed: {2}" -f ($idx + 1), $excelNames.Count, $name) -ForegroundColor DarkGray
            $cntSkip++
            continue
        }

        $file = Get-EvidencePath $EvidenceDir $name
        if (-not (Test-Path -LiteralPath $file)) {
            Write-Host ("[{0}/{1}] MISSING: {2}" -f ($idx + 1), $excelNames.Count, $file) -ForegroundColor Yellow
            $cntMissing++
            continue
        }

        $wb = $null
        try {
            Write-Host ''
            Write-Host ("[{0}/{1}] OPEN: {2}" -f ($idx + 1), $excelNames.Count, $file) -ForegroundColor Cyan
            $openedAt = Get-Date
            $wb = Open-WorkbookEditable $excel $file

            if ($wb.ReadOnly) {
                try { $wb.Close($false) } catch {}
                $wb = $null
                throw "Workbook opened as read-only. Possible causes: another Excel window/user is locking it, write permission is missing, or the file/share is read-only: $file"
            }

            [void](Activate-ExcelWindow $shell $excel $wb)

            Write-Host '  Workbook opened in editable mode. Review/edit in Excel, then return here manually.' -ForegroundColor DarkGray
            $choice = Read-ReviewChoice $openedAt
            if ($choice -match '^\s*q\s*$') {
                try { $wb.Close($false) } catch {}
                break
            }
            if ($choice -match '^\s*s\s*$') {
                try { $wb.Close($false) } catch {}
                $cntSkip++
                continue
            }

            Set-ReviewCursorAllSheets $wb $CursorCell
            Save-BySendKeys $shell $excel $wb $SaveWaitMs

            if (-not $wb.Saved) {
                Write-Host '  [WARN] Excel still reports unsaved after Ctrl+S/Esc/wait. Closing without second save to avoid macro loop.' -ForegroundColor Yellow
            }

            try { $wb.Close($false) } catch {}

            foreach ($r in $groupRows) {
                Set-BitValue $r $ReviewField $ReviewBit
            }
            $allRows | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force

            Write-Host ("  [OK] reviewed bit set: {0} -> {1}" -f $name, $ReviewBit) -ForegroundColor Green
            $cntDone++
        } catch {
            Write-Host ("  [ERROR] {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
            try { if ($wb) { $wb.Close($false) } } catch {}
            $cntFail++
        }
    }
} finally {
    if ($shell) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch {}
    }
    if ($excel) {
        try { $excel.DisplayAlerts = $true } catch {}
        Close-ExcelApp $excel
    }
}

Write-Host ''
Write-Host '===== ReviewEvidence Done =====' -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Missing : {0}" -f $cntMissing)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
Write-Host ("  Mapping : {0}" -f $mappingPath)
