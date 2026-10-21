# modules/excel/excel.snap_list.ps1
# Snap one job's rows of the overview list (the GFIX list workbook) the
# way ExcelSnap.ps1 did: AutoFilter the send-job column to the job,
# Range.CopyPicture from the "from" column to the job column, paste into a
# temporary chart, export a PNG. 10/16 (the reviewer's 10/13 notes on
# JJDSWM10, check sheet rows 326/327): when two or more of the job's rows
# are the same from "from" to the job column, the snap runs on to the
# source-file (SNDFILE) column so the rows can be told apart (the fix
# shown in JJDSWM51). Columns are found by the two header rows (group row
# + sub row), never by letter; the labels come from the profile.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.snap_list'
  group      = 'excel'
  summary    = 'Filter the overview list to one job and snap its rows to a PNG (on to SNDFILE when the rows repeat)'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook    = @{ type='session'; sessionKind='workbook'; required=$true; desc='the overview workbook (excel.open readOnly)' }
    sheet       = @{ type='string'; required=$true }
    value       = @{ type='string'; required=$true; desc='the job to filter on (the J name)' }
    groupRow    = @{ type='int';    default=4; desc='header row with the group labels (from / to / send)' }
    subRow      = @{ type='int';    default=5; desc='header row with the column labels; data starts below it' }
    left        = @{ type='map';    required=$true; desc='@{ group; sub }: the first column of the snap' }
    job         = @{ type='map';    required=$true; desc='@{ group; sub }: the filtered column, the normal right edge' }
    extendTo    = @{ type='map';    default=@{}; desc='@{ group; sub }: the right edge when rows repeat from left to job; empty = never extend' }
    maxColumns  = @{ type='int';    default=40; desc='header columns scanned' }
    savePng     = @{ type='path';   required=$true; desc='the PNG (relative: under the work dir)' }
  }
  outputs    = @{
    png      = @{ type='path' }
    rows     = @{ type='int';    desc='rows the filter left' }
    extended = @{ type='bool';   desc='true when the snap ran on to extendTo (rows repeated from left to job)' }
    range    = @{ type='string'; desc='the snapped address, e.g. B4:O7' }
    lines    = @{ type='list';   desc='each row''s cells from left to the right edge, tab-joined (for the gate)' }
    jobFrom  = @{ type='any';    desc='left edge of the job column across the picture, 0..1 of its width (screen.compare_strip)' }
    jobTo    = @{ type='any';    desc='right edge of the job column, 0..1' }
    bodyFrom = @{ type='any';    desc='top of the data rows (under the two header rows), 0..1 of the picture''s height' }
  }
  failures   = @(
    @{ id = 'sheet_not_found';  transient = $false }
    @{ id = 'column_not_found'; transient = $false }
    @{ id = 'not_found';        transient = $false }
    @{ id = 'copy_failed';      transient = $true  }
  )
  example    = @{ use = 'excel.snap_list'; with = @{ workbook = 'list'; sheet = 'overview'; value = '{{item.JOB}}'; left = @{ group = 'from'; sub = 'system' }; job = @{ group = 'send'; sub = 'job' }; extendTo = @{ group = 'send'; sub = 'source file' }; savePng = 'capture/{{item.keySafe}}/pic1.png' } }
  notes      = 'Changes only the AutoFilter of the overview sheet and clears it again; open the overview read-only (excel.open readOnly) so nothing is ever saved. Clobbers the clipboard. A header label matches when the cell, spaces and line breaks removed, STARTS with it (a cell "<label>(SNDFILE)" matches "<label>").'
}

function ExcelSnapList-NormalizeLabel {
    # PURE. A header cell -> the text compared: line breaks and spaces out.
    param($Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value -replace '[\s\u3000]', '')
}

function ExcelSnapList-ResolveColumns {
    <#
      PURE. Group-row and sub-row texts (index 1..n; [0] unused) -> the
      column of each wanted @{ group; sub }. A group label covers every
      column to its right until the next group label (merged cells read
      as the first cell only). A sub label matches when the normalized
      cell STARTS with the normalized wanted label ("<label>(SNDFILE)"
      for "<label>"). Returns @{ ok; columns = @{ name -> col }; message }.
    #>
    param($GroupTexts, $SubTexts, [hashtable]$Wanted)
    $g = @($GroupTexts); $s = @($SubTexts)
    $n = [Math]::Max($g.Count, $s.Count)
    $groupOf = @{}
    $cur = ''
    for ($c = 1; $c -lt $n; $c++) {
        $gt = if ($c -lt $g.Count) { ExcelSnapList-NormalizeLabel $g[$c] } else { '' }
        if ($gt -ne '') { $cur = $gt }
        $groupOf[$c] = $cur
    }
    $cols = @{}
    $missing = New-Object System.Collections.ArrayList
    foreach ($name in @($Wanted.Keys | Sort-Object)) {
        $w = $Wanted[$name]
        if ($null -eq $w -or -not ($w -is [System.Collections.IDictionary]) -or -not $w.Contains('sub')) { continue }
        $wg = ExcelSnapList-NormalizeLabel $w['group']; $ws = ExcelSnapList-NormalizeLabel $w['sub']
        $found = 0
        for ($c = 1; $c -lt $n; $c++) {
            $st = if ($c -lt $s.Count) { ExcelSnapList-NormalizeLabel $s[$c] } else { '' }
            if ($st -eq '' -or -not $st.StartsWith($ws)) { continue }
            if ($wg -ne '' -and [string]$groupOf[$c] -ne $wg) { continue }
            $found = $c; break
        }
        if ($found -eq 0) { [void]$missing.Add(('{0} ({1} / {2})' -f $name, $w['group'], $w['sub'])) } else { $cols[$name] = $found }
    }
    if ($missing.Count -gt 0) { return @{ ok = $false; columns = $cols; message = ('header not found: ' + ($missing -join ', ')) } }
    return @{ ok = $true; columns = $cols; message = '' }
}

function ExcelSnapList-HasRepeats {
    # PURE. Rows (each a list of cell texts from left to job) -> true when
    # two of them are the same in every cell (the reviewer, 10/13: rows
    # that repeat from "from" to the send job). A different IF number is
    # not a repeat.
    param($Rows)
    $seen = @{}
    foreach ($r in @($Rows)) {
        $k = (@($r) | ForEach-Object { ([string]$_).Trim() }) -join "`t"
        if ($seen.Contains($k)) { return $true }
        $seen[$k] = $true
    }
    return $false
}

function ExcelSnapList-ColumnLetter {
    # PURE. 1 -> A, 27 -> AA.
    param([int]$Col)
    $s = ''
    $c = $Col
    while ($c -gt 0) { $m = ($c - 1) % 26; $s = [string][char](65 + $m) + $s; $c = [int][Math]::Floor(($c - 1) / 26) }
    return $s
}

function ExcelSnapList-CopyPicture {
    # Range.CopyPicture with the ExcelSnap.ps1 fallbacks: xlScreen/xlBitmap,
    # xlPrinter/xlBitmap, then the xlPicture pair; each tried 3 times.
    param($Worksheet, $Range)
    $combos = @(@(1, -4144), @(2, -4144), @(1, -4147), @(2, -4147))
    $last = $null
    foreach ($combo in $combos) {
        for ($a = 1; $a -le 3; $a++) {
            try {
                try { $Worksheet.Application.Visible = $true } catch { }
                try { [void]$Worksheet.Activate() } catch { }
                try { [void]$Range.Select() } catch { }
                Start-Sleep -Milliseconds (80 * $a)
                [void]$Range.CopyPicture($combo[0], $combo[1])
                Start-Sleep -Milliseconds 120
                return @{ ok = $true; message = '' }
            } catch { $last = $_; Start-Sleep -Milliseconds (150 * $a) }
        }
    }
    return @{ ok = $false; message = (Format-EbiExcelError -Stage 'CopyPicture' -ErrorRecord $last) }
}

function Invoke-Step {
    param($In, $Ctx)
    $none = @{ png = ''; rows = 0; extended = $false; range = ''; lines = @(); jobFrom = 0.0; jobTo = 0.0; bodyFrom = 0.0 }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would filter {0} to {1} and snap it' -f $In['sheet'], $In['value'])); $none['ok'] = $true; return $none }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'sheet_not_found'; $r['message'] = ('sheet "' + $In['sheet'] + '" not found'); return $r }
    $gRow = [int]$In['groupRow']; $sRow = [int]$In['subRow']; $maxC = [int]$In['maxColumns']
    $gT = @(''); $sT = @('')
    for ($c = 1; $c -le $maxC; $c++) { $gT += [string]$ws.Cells.Item($gRow, $c).Value2; $sT += [string]$ws.Cells.Item($sRow, $c).Value2 }
    $wanted = @{ left = $In['left']; job = $In['job'] }
    $ext = $In['extendTo']
    if ($ext -is [System.Collections.IDictionary] -and $ext.Contains('sub') -and [string]$ext['sub'] -ne '') { $wanted['extendTo'] = $ext }
    $rc = ExcelSnapList-ResolveColumns -GroupTexts $gT -SubTexts $sT -Wanted $wanted
    if (-not $rc['ok']) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'column_not_found'; $r['message'] = ($In['sheet'] + ' rows ' + $gRow + '/' + $sRow + ': ' + $rc['message']); return $r }
    $cL = [int]$rc['columns']['left']; $cJ = [int]$rc['columns']['job']
    $cX = if ($rc['columns'].Contains('extendTo')) { [int]$rc['columns']['extendTo'] } else { 0 }
    $last = 0
    try { $last = [int]$ws.Cells.Item([int]$ws.Rows.Count, $cJ).End(-4162).Row } catch { $last = 0 }   # xlUp
    if ($last -le $sRow) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'not_found'; $r['message'] = ($In['sheet'] + ': no data below row ' + $sRow); return $r }
    $right = [Math]::Max($cJ, $cX)
    $png = Resolve-EbiWorkPath -PathValue ([string]$In['savePng']) -WorkDir ([string]$Ctx['WorkDir'])
    $chartObj = $null
    try {
        try { if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false } } catch { }
        $filterRange = $ws.Range($ws.Cells.Item($sRow, $cL), $ws.Cells.Item($last, $right))
        [void]$filterRange.AutoFilter(($cJ - $cL + 1), [string]$In['value'])
        # the rows the filter left: their cells from left to job decide the edge
        $vis = New-Object System.Collections.ArrayList
        for ($rr = $sRow + 1; $rr -le $last; $rr++) {
            if ([bool]$ws.Rows.Item($rr).Hidden) { continue }
            if (([string]$ws.Cells.Item($rr, $cJ).Value2).Trim() -ne ([string]$In['value']).Trim()) { continue }
            [void]$vis.Add($rr)
        }
        if ($vis.Count -eq 0) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'not_found'; $r['message'] = ('no row of ' + $In['sheet'] + ' has ' + $In['value'] + ' in the job column'); return $r }
        $cells = New-Object System.Collections.ArrayList
        foreach ($rr in $vis) { $one = @(); for ($c = $cL; $c -le $cJ; $c++) { $one += [string]$ws.Cells.Item($rr, $c).Text }; [void]$cells.Add($one) }
        $extended = ($cX -gt $cJ) -and (ExcelSnapList-HasRepeats -Rows $cells.ToArray())
        $edge = if ($extended) { $cX } else { $cJ }
        $lastVis = [int]$vis[$vis.Count - 1]
        $snap = $ws.Range($ws.Cells.Item($gRow, $cL), $ws.Cells.Item($lastVis, $edge))
        $lines = New-Object System.Collections.ArrayList
        foreach ($rr in $vis) { $one = @(); for ($c = $cL; $c -le $edge; $c++) { $one += [string]$ws.Cells.Item($rr, $c).Text }; [void]$lines.Add(($one -join "`t")) }
        $h = 0.0
        for ($rr = $gRow; $rr -le $lastVis; $rr++) { if (-not [bool]$ws.Rows.Item($rr).Hidden) { $h += [double]$ws.Rows.Item($rr).Height } }
        $w = [double]$snap.Width
        $bodyFrom = 0.0
        try { if ($h -gt 0) { $hh = 0.0; for ($rr = $gRow; $rr -le $sRow; $rr++) { if (-not [bool]$ws.Rows.Item($rr).Hidden) { $hh += [double]$ws.Rows.Item($rr).Height } }; $bodyFrom = $hh / $h } } catch { }
        # where the job column sits across the picture (0..1 of its width), for
        # screen.compare_strip: the GIFT side's picture 1 must show the same
        # job there (10/21, the operator: compare that part, not OCR)
        $jobFrom = 0.0; $jobTo = 0.0
        try { if ($w -gt 0) { $jl = [double]$ws.Cells.Item($gRow, $cJ).Left - [double]$snap.Left; $jobFrom = $jl / $w; $jobTo = ($jl + [double]$ws.Columns.Item($cJ).Width) / $w } } catch { }
        $cp = ExcelSnapList-CopyPicture -Worksheet $ws -Range $snap
        if (-not $cp['ok']) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'copy_failed'; $r['message'] = $cp['message']; return $r }
        $dir = Split-Path -Parent $png
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        if (Test-Path -LiteralPath $png) { Remove-Item -LiteralPath $png -Force }
        $stage = 'add chart'
        try {
            $chartObj = $ws.ChartObjects().Add(0, 0, $w + 2, $h + 2)
            $stage = 'chart fill'
            try { $chartObj.Chart.ChartArea.Format.Fill.Visible = -1; $chartObj.Chart.ChartArea.Format.Fill.ForeColor.RGB = 16777215; $chartObj.Chart.ChartArea.Format.Line.Visible = 0 } catch { }
            $stage = 'paste into chart'
            [void]$chartObj.Chart.Paste()
            try { if ([int]$chartObj.Chart.Shapes.Count -gt 0) { $p = $chartObj.Chart.Shapes.Item(1); $p.Left = 0; $p.Top = 0; if ([double]$p.Width -gt 0) { $chartObj.Width = [double]$p.Width + 2 }; if ([double]$p.Height -gt 0) { $chartObj.Height = [double]$p.Height + 2 } } } catch { }
            $stage = 'export'
            [void]$chartObj.Chart.Export($png, 'PNG')
        } catch { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'copy_failed'; $r['message'] = (Format-EbiExcelError -Stage $stage -ErrorRecord $_); return $r }
        if (-not (Test-Path -LiteralPath $png)) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'copy_failed'; $r['message'] = ('Chart.Export wrote no file: ' + $png); return $r }
        $addr = (ExcelSnapList-ColumnLetter $cL) + $gRow + ':' + (ExcelSnapList-ColumnLetter $edge) + $lastVis
        if ($extended) { $Ctx.Log.Info(('{0}: {1} rows repeat from {2} to the job column -> snapped on to {3} ({4})' -f $In['value'], $vis.Count, (ExcelSnapList-ColumnLetter $cL), (ExcelSnapList-ColumnLetter $cX), $addr)) }
        else { $Ctx.Log.Info(('{0}: {1} row(s), snapped {2}' -f $In['value'], $vis.Count, $addr)) }
        return @{ ok = $true; png = $png; rows = $vis.Count; extended = [bool]$extended; range = $addr; lines = $lines.ToArray(); jobFrom = [double]$jobFrom; jobTo = [double]$jobTo; bodyFrom = [double]$bodyFrom }
    } catch {
        $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'copy_failed'; $r['message'] = (Format-EbiExcelError -Stage 'filter' -ErrorRecord $_); return $r
    } finally {
        if ($null -ne $chartObj) { try { [void]$chartObj.Delete() } catch { }; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($chartObj) } catch { } }
        try { if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false } } catch { }
        try { $ws.Application.CutCopyMode = $false } catch { }
    }
}
