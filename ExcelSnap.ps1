# ============================================================
#  ExcelSnap.ps1  (Phase 2)
#
#  For each distinct JOB_NAME in $WorkDir\mapping_<Owner>.csv:
#    1. AutoFilter GFIX一覧's 送信ジョブ col to JOB_NAME
#    2. CopyPicture of B4:O<lastRow>  (dynamic left/right cols)
#    3. Export PNG to $WorkDir\snap\excel\<JOB_NAME>.png
#    4. Update mapping: Excel_snap = 1 for all rows with this JOB
#
#  Default mode: Excel hidden, xlPrinter (works headless).
#  -Visible    : Excel shows on screen, xlScreen (truer to what you see).
#  -Force      : re-snap even if Excel_snap already = 1.
#
#  GFIX一覧 is COPIED TO TEMP first, so you can keep yours open.
#
#  File encoding: save as UTF-8 with BOM, CRLF.
# ============================================================

param(
    [string]$WorkDir,
    [string]$Owner = '',
    [switch]$Visible,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# -- Force console to UTF-8 --
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding           = [System.Text.UTF8Encoding]::new()
} catch {}

# -- Interactive fallback --
if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host "WorkDir path" }

Write-Host ""
Write-Host "===== ExcelSnap (Phase 2) =====" -ForegroundColor Green
Write-Host ("  WorkDir : {0}" -f $WorkDir)
Write-Host ("  Owner   : {0}" -f $Owner)
Write-Host ("  Mode    : {0}" -f $(if ($Visible.IsPresent) { "Visible + xlScreen" } else { "Hidden + xlScreen + xlBitmap" }))
Write-Host ("  Force   : {0}" -f $Force.IsPresent)
Write-Host ""

# -- Validate --
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found." -ForegroundColor Red; exit 1
}
$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host ("[ERROR] mapping file not found: {0}" -f $mappingPath) -ForegroundColor Red; exit 1
}

# -- Find GFIX file --
$gfixFiles = @(Get-ChildItem -LiteralPath $WorkDir -Filter "*GFIX*.xlsx" -File -ErrorAction SilentlyContinue |
                Where-Object { -not $_.Name.StartsWith("~$") })
if ($gfixFiles.Count -eq 0) { Write-Host "[ERROR] No GFIX*.xlsx" -ForegroundColor Red; exit 1 }
if ($gfixFiles.Count -gt 1) {
    Write-Host "[ERROR] Multiple GFIX*.xlsx:" -ForegroundColor Red
    $gfixFiles | ForEach-Object { Write-Host ("    - {0}" -f $_.Name) -ForegroundColor Red }; exit 1
}
$gfixPath = $gfixFiles[0].FullName
Write-Host ("[INFO] GFIX     : {0}" -f (Split-Path -Leaf $gfixPath))
Write-Host ("[INFO] Mapping  : {0}" -f (Split-Path -Leaf $mappingPath))

# -- Output dir --
$snapDir = Join-Path $WorkDir "snap\excel"
if (-not (Test-Path -LiteralPath $snapDir)) {
    New-Item -ItemType Directory -Path $snapDir -Force | Out-Null
}
Write-Host ("[INFO] OutDir   : {0}" -f $snapDir)

# -- Warn about clipboard --
Write-Host ""
Write-Host "[NOTE] Script uses clipboard for image transfer. Avoid Ctrl+C/V during run." -ForegroundColor Yellow

# -- Japanese label constants --
$LBL_GFIX_SHEET = "GFIX" + [char]0x9001 + [char]0x53D7 + [char]0x4FE1 + [char]0x4E00 + [char]0x89A7  # GFIX送受信一覧
$LBL_SYS_TYPE   = [char]0x30B7 + [char]0x30B9 + [char]0x30C6 + [char]0x30E0 + [char]0x7A2E + [char]0x5225  # システム種別
$LBL_SEND       = [char]0x9001 + [char]0x4FE1                                                            # 送信
$LBL_JOB        = [char]0x30B8 + [char]0x30E7 + [char]0x30D6                                             # ジョブ
$LBL_SRC_FILE   = [char]0x5143 + [char]0x30D5 + [char]0x30A1 + [char]0x30A4 + [char]0x30EB              # 元ファイル
$LBL_FROM       = "from"

# -- Helpers --
function Get-ColLetter([int]$c) {
    if ($c -le 0)  { return "?" }
    if ($c -le 26) { return [char]([byte][char]'A' + $c - 1) }
    $a = [char]([byte][char]'A' + [Math]::Floor(($c - 1) / 26) - 1)
    $b = [char]([byte][char]'A' + (($c - 1) % 26))
    return "$a$b"
}

# Robust Range.CopyPicture: the bare call fails with
#   "Range クラスの CopyPicture プロパティを取得できません" (0x800A03EC)
# when the Excel window is parked off-screen / not foreground, because
# xlScreen ("as shown on screen") has nothing rendered on a physical
# monitor to copy, and the clipboard can be transiently busy. We re-
# activate the app window + sheet + selection before each try, retry a
# few times, then fall back to xlPrinter appearance (rendered via the
# print path, no monitor needed). Appearance: 1=xlScreen 2=xlPrinter.
# Format: -4147=xlPicture(EMF) -4144=xlBitmap.
function Invoke-RangeCopyPicture {
    param($Excel, $Worksheet, $Range)
    # (Appearance, Format) attempts. xlBitmap leads because xlPicture/EMF
    # was found to fail in the default off-screen window mode; xlPrinter
    # appearance needs no monitor, so it is the off-screen safety net.
    $combos = @(
        @(1, -4144),   # xlScreen  + xlBitmap  (proven reliable on-screen)
        @(2, -4144),   # xlPrinter + xlBitmap  (no monitor needed; off-screen safe)
        @(1, -4147),   # xlScreen  + xlPicture (vector; higher fidelity fallback)
        @(2, -4147)    # xlPrinter + xlPicture (last resort)
    )
    $lastErr = $null
    foreach ($combo in $combos) {
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                $Excel.Visible = $true
                try { $Worksheet.Activate() | Out-Null } catch {}
                try { $Range.Select()       | Out-Null } catch {}
                Start-Sleep -Milliseconds (80 * $attempt)
                $Range.CopyPicture($combo[0], $combo[1]) | Out-Null
                Start-Sleep -Milliseconds 120
                return @($combo[0], $combo[1])   # success -> return the combo used
            } catch {
                $lastErr = $_
                Start-Sleep -Milliseconds (150 * $attempt)
            }
        }
    }
    throw ("CopyPicture failed after all appearance/format fallbacks. Last error: {0}" -f $lastErr)
}

# ============================================================
# Load mapping & figure out pending JOB_NAMEs
# ============================================================
Write-Host ""
Write-Host "[Step 1] Loading mapping..." -ForegroundColor Cyan
$mapping = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)
Write-Host ("  Rows : {0}" -f $mapping.Count)

# Distinct JOB_NAME, with "is this JOB fully done?" flag
$jobOrder = [System.Collections.Generic.List[string]]::new()
$jobDone  = @{}
foreach ($r in $mapping) {
    $j = $r.JOB_NAME
    if (-not $jobDone.ContainsKey($j)) {
        [void]$jobOrder.Add($j)
        $jobDone[$j] = $true
    }
    if ($r.Excel_snap -ne "1") { $jobDone[$j] = $false }
}

$pending = @()
$skipped = @()
foreach ($j in $jobOrder) {
    if ($jobDone[$j] -and -not $Force) { $skipped += $j } else { $pending += $j }
}
Write-Host ("  Distinct JOBs: {0}, pending: {1}, skipped: {2}" -f $jobOrder.Count, $pending.Count, $skipped.Count)
if ($pending.Count -eq 0) {
    Write-Host "[INFO] Nothing to do (all done). Use -Force to redo." -ForegroundColor Yellow
    return
}

# ============================================================
# Copy GFIX to TEMP (so user can keep theirs open)
# ============================================================
Write-Host ""
Write-Host "[Step 2] Preparing temp copy of GFIX..." -ForegroundColor Cyan
$tempGfix = Join-Path $env:TEMP ("gfix_snap_{0}_{1}.xlsx" -f $PID, (Get-Random))
Copy-Item -LiteralPath $gfixPath -Destination $tempGfix -Force
Write-Host ("  Temp : {0}" -f (Split-Path -Leaf $tempGfix))

# ============================================================
# Excel COM
# ============================================================
$excel  = $null
$wbGfix = $null

try {
    Write-Host ""
    Write-Host "[Step 3] Starting Excel COM..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    
    # CopyPicture needs an unminimized window that renders the area.
    # Default: visible+off-screen so user doesn't see it but Excel still renders.
    $excel.Visible            = $true
    $excel.DisplayAlerts      = $false
    $excel.AutomationSecurity = 3
    $excel.AskToUpdateLinks   = $false
    $excel.EnableEvents       = $false
    try { $excel.WindowState = -4143 } catch {}   # xlNormal
    if (-not $Visible.IsPresent) {
        # Park window off-screen (large positive Top/Left)
        try {
            $excel.Top  = 30000
            $excel.Left = 30000
            $excel.Width  = 1200
            $excel.Height = 800
        } catch {}
    }

    $wbGfix = $excel.Workbooks.Open($tempGfix, $false, $false)  # ReadOnly=false (need filter)
    $wsGfix = $null
    foreach ($ws in $wbGfix.Worksheets) {
        if ($ws.Name -eq $LBL_GFIX_SHEET) { $wsGfix = $ws; break }
    }
    if (-not $wsGfix) { throw ("GFIX sheet '{0}' not found." -f $LBL_GFIX_SHEET) }
    $wsGfix.Activate() | Out-Null

    # -- Resolve columns (B = from システム種別 ... O = 送信 元ファイル ... filter on 送信 ジョブ) --
    Write-Host "[Step 4] Resolving header columns..." -ForegroundColor Cyan
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

    $col_left  = Find-GS $gsTable $LBL_FROM $LBL_SYS_TYPE   # left bound (from system-type col)
    $col_job   = Find-GS $gsTable $LBL_SEND $LBL_JOB        # filter col (send job)
    $col_right = $col_job                                    # right bound = filter col (snap ends here)

    if ($col_left -eq 0 -or $col_right -eq 0 -or $col_job -eq 0) {
        Write-Host ""
        Write-Host "[ERROR] Column resolution failed:" -ForegroundColor Red
        Write-Host ("  col_left  = {0} ({1})" -f $col_left,  (Get-ColLetter $col_left))
        Write-Host ("  col_right = {0} ({1})" -f $col_right, (Get-ColLetter $col_right))
        Write-Host ("  col_job   = {0} ({1})" -f $col_job,   (Get-ColLetter $col_job))
        Write-Host ""
        Write-Host "  Detected headers (col / group hex / sub hex):" -ForegroundColor Yellow
        Write-Host "  --------------------------------------------------" -ForegroundColor Yellow
        for ($c = 1; $c -le $MAX_COL; $c++) {
            $g = $gsTable[$c][0]; $s = $gsTable[$c][1]
            if (-not [string]::IsNullOrWhiteSpace($g) -or -not [string]::IsNullOrWhiteSpace($s)) {
                $gHex = if ($g) { ($g.ToCharArray() | ForEach-Object { "{0:X4}" -f [int]$_ }) -join " " } else { "" }
                $sHex = if ($s) { ($s.ToCharArray() | ForEach-Object { "{0:X4}" -f [int]$_ }) -join " " } else { "" }
                Write-Host ("    {0,2} ({1}) : group=[{2}]  sub=[{3}]" -f $c, (Get-ColLetter $c), $gHex, $sHex) -ForegroundColor DarkYellow
            }
        }
        Write-Host ""
        Write-Host "  Searching for:" -ForegroundColor Yellow
        $sendHex = ($LBL_SEND.ToCharArray()     | ForEach-Object { "{0:X4}" -f [int]$_ }) -join " "
        $srcHex  = ($LBL_SRC_FILE.ToCharArray() | ForEach-Object { "{0:X4}" -f [int]$_ }) -join " "
        $jobHex  = ($LBL_JOB.ToCharArray()      | ForEach-Object { "{0:X4}" -f [int]$_ }) -join " "
        Write-Host ("    LBL_SEND     = [{0}]  (send)" -f $sendHex) -ForegroundColor Yellow
        Write-Host ("    LBL_SRC_FILE = [{0}]  (source file)" -f $srcHex) -ForegroundColor Yellow
        Write-Host ("    LBL_JOB      = [{0}]  (job)" -f $jobHex) -ForegroundColor Yellow
        throw "Column resolution failed."
    }
    Write-Host ("  Snap range cols : {0} ({1}) - {2} ({3})" -f `
        (Get-ColLetter $col_left),  $col_left,
        (Get-ColLetter $col_right), $col_right)
    Write-Host ("  Filter col      : {0} ({1})" -f (Get-ColLetter $col_job), $col_job)

    $lastRow = $wsGfix.UsedRange.Rows.Count
    Write-Host ("  Data last row   : {0}" -f $lastRow)

    # -- Bottom-line clear of any pre-existing filter --
    if ($wsGfix.AutoFilterMode) { $wsGfix.AutoFilterMode = $false }

    # CopyPicture: Invoke-RangeCopyPicture tries xlScreen first (truest to
    # the on-screen render), then falls back to xlPrinter when the off-screen
    # window has nothing on a monitor to capture. -Visible only controls where
    # the window sits (on-screen vs off-screen); the fallback makes the
    # off-screen (hidden) mode reliable regardless.

    # ============================================================
    # Loop: each JOB_NAME -> filter -> copy picture -> chart export
    # ============================================================
    Write-Host ""
    Write-Host ("[Step 5] Snapping {0} JOB(s)..." -f $pending.Count) -ForegroundColor Cyan

    $idx = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($jobName in $pending) {
        $idx++
        $tStart = $sw.Elapsed.TotalSeconds
        Write-Host ("  [{0}/{1}] {2}" -f $idx, $pending.Count, $jobName) -ForegroundColor White

        $outPath = Join-Path $snapDir ("{0}.png" -f $jobName)
        if ((Test-Path -LiteralPath $outPath) -and -not $Force) {
            Remove-Item -LiteralPath $outPath -Force  # we're snapping anyway, ensure fresh
        }

        # AutoFilter on row 5 (sub-header row) so labels are filter headers
        $filterRange = $wsGfix.Range(
            $wsGfix.Cells.Item(5, $col_left),
            $wsGfix.Cells.Item($lastRow, $col_right))
        $filterField = $col_job - $col_left + 1
        $filterRange.AutoFilter($filterField, $jobName) | Out-Null

        # Re-activate sheet after AutoFilter
        $wsGfix.Activate() | Out-Null
        Start-Sleep -Milliseconds 50

        $snapRange = $wsGfix.Range(
            $wsGfix.Cells.Item(4, $col_left),
            $wsGfix.Cells.Item($lastRow, $col_right))

        # -- Compute visible bounding box (filter hides middle rows) --
        # Sum visible row heights between row 4 and lastRow.
        $visibleHeight = 0
        for ($rr = 4; $rr -le $lastRow; $rr++) {
            if (-not $wsGfix.Rows.Item($rr).Hidden) {
                $visibleHeight += $wsGfix.Rows.Item($rr).Height
            }
        }
        # Width: cols aren't hidden, use range.Width directly
        $rangeWidth = $snapRange.Width

        # -- Copy as picture (retry + appearance/format fallbacks) --
        # xlScreen can fail outright when the window is parked off-screen;
        # Invoke-RangeCopyPicture re-activates + retries + falls back to
        # xlPrinter so the capture works even when nothing is on a monitor.
        # Leads with xlBitmap (xlPicture/EMF was found to fail in off-screen
        # mode); xlPicture is kept only as a lower-priority fallback.
        [void](Invoke-RangeCopyPicture -Excel $excel -Worksheet $wsGfix -Range $snapRange)

        # -- Insert chart sized to actual visible content, paste, export --
        $chartObj = $wsGfix.ChartObjects().Add(0, 0, $rangeWidth + 2, $visibleHeight + 2)
        try {
            # Remove chart's default fill so PNG has no gray border
            $chartObj.Chart.ChartArea.Format.Fill.Visible = -1                          # msoTrue
            $chartObj.Chart.ChartArea.Format.Fill.ForeColor.RGB = 16777215              # white (0xFFFFFF)
            $chartObj.Chart.ChartArea.Format.Line.Visible = 0
            $chartObj.Chart.Paste()
            # After paste, snap pasted shape to top-left and refit chart if needed
            if ($chartObj.Chart.Shapes.Count -gt 0) {
                $shape = $chartObj.Chart.Shapes.Item(1)
                $shape.Left = 0
                $shape.Top  = 0
                # If paste came in with different size, prefer the shape's reported size
                if ($shape.Width  -gt 0) { $chartObj.Width  = $shape.Width  + 2 }
                if ($shape.Height -gt 0) { $chartObj.Height = $shape.Height + 2 }
            }
            $chartObj.Chart.Export($outPath, "PNG") | Out-Null
        } finally {
            $chartObj.Delete() | Out-Null
        }

        # Update mapping in-memory + persist
        foreach ($r in $mapping) {
            if ($r.JOB_NAME -eq $jobName) { $r.Excel_snap = "1" }
        }
        $mapping | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force

        $tEnd = $sw.Elapsed.TotalSeconds
        Write-Host ("       -> {0}  ({1:N1}s)" -f (Split-Path -Leaf $outPath), ($tEnd - $tStart)) -ForegroundColor Green
    }

    # Clean filter
    if ($wsGfix.AutoFilterMode) { $wsGfix.AutoFilterMode = $false }

    $sw.Stop()
    Write-Host ""
    Write-Host ("[Done] {0} snap(s) generated in {1:N1}s." -f $pending.Count, $sw.Elapsed.TotalSeconds) -ForegroundColor Green
    if ($skipped.Count -gt 0) {
        Write-Host ("       {0} skipped (already done). Use -Force to redo." -f $skipped.Count) -ForegroundColor DarkGray
    }
}
finally {
    if ($wbGfix) { try { $wbGfix.Close($false) } catch {} }
    if ($excel) {
        try {
            foreach ($wb in @($excel.Workbooks)) { try { $wb.Close($false) } catch {} }
        } catch {}
        try { $excel.Quit() } catch {}
    }
    if ($wbGfix) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbGfix) }
    if ($excel)  { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)  }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Cleanup temp file
    if (Test-Path -LiteralPath $tempGfix) {
        Remove-Item -LiteralPath $tempGfix -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host "===== Phase 2 Done =====" -ForegroundColor Green
