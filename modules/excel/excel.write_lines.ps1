# modules/excel/excel.write_lines.ps1
# Paste text lines down one column, one line per row, optionally under a
# label row: a log excerpt in an evidence sheet. Lines come from text
# files (concatenated in order, no blank line between -- several job logs
# of one transfer read as one block) and/or a list. Fixed-pitch font
# forced so the evidence reads the same whatever the workbook default is
# (ExcelHelpers.ps1 Write-LogLines). The text goes in as TEXT: a log line
# that looks like a date stays the line.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')

$Manifest = @{
  id         = 'excel.write_lines'
  group      = 'excel'
  summary    = 'Paste text lines (from files or a list) down a column, under an optional label'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet    = @{ type='string'; required=$true }
    row      = @{ type='int';    required=$true; desc='first row to write (the label, when there is one)' }
    gapRows  = @{ type='int';    default=0; desc='blank rows left before row (row + gapRows is where writing starts)' }
    column   = @{ type='string'; default='B' }
    label    = @{ type='string'; default=''; desc='written at row; the lines start one row below' }
    paths    = @{ type='list';   default=@(); desc='text files, concatenated in order' }
    lines    = @{ type='list';   default=@(); desc='more lines, after the files' }
    encoding = @{ type='string'; default='mixed'; enum=@('mixed', 'utf8', 'cp932') }
    fontName = @{ type='string'; default='' }
    fontSize = @{ type='int';    default=0 }
  }
  outputs    = @{
    labelRow = @{ type='int'; desc='row of the label; 0 when none' }
    firstRow = @{ type='int'; desc='row of the first line' }
    lastRow  = @{ type='int'; desc='row of the last line (labelRow when there are no lines)' }
    nextRow  = @{ type='int'; desc='the row after the last one written' }
    written  = @{ type='int'; desc='lines written (label not counted)' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'file_not_found';  transient = $true  }
    @{ id = 'write_failed';    transient = $true  }
  )
  example    = @{ use = 'excel.write_lines'; with = @{ workbook = 'wb'; sheet = 'result'; row = 8; label = 'receive log'; paths = @('log/jobs/1.log'); fontName = 'MS Gothic'; fontSize = 10 } }
  notes      = 'Rows are written one COM call each; a few hundred lines take a few seconds.'
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $all = New-Object System.Collections.ArrayList
    foreach ($p in @($In['paths'])) {
        if ([string]::IsNullOrWhiteSpace([string]$p)) { continue }
        $fp = Resolve-EbiWorkPath -PathValue ([string]$p) -WorkDir $work
        if (-not (Test-Path -LiteralPath $fp -PathType Leaf)) {
            if ($Ctx['DryRun']) { continue }
            return @{ ok = $false; failure = 'file_not_found'; message = $fp; labelRow = 0; firstRow = 0; lastRow = 0; nextRow = 0; written = 0 }
        }
        $bytes = [System.IO.File]::ReadAllBytes($fp)
        $text = switch ([string]$In['encoding']) { 'cp932' { (Get-EbiCodePage -CodePage 932).GetString($bytes) } 'utf8' { (New-Object System.Text.UTF8Encoding($false)).GetString($bytes) } default { (ConvertFrom-EbiMixedBytes -Bytes $bytes)['text'] } }
        foreach ($l in @(Get-EbiTextLines -Text $text)) { [void]$all.Add([string]$l) }
    }
    foreach ($l in @($In['lines'])) { [void]$all.Add([string]$l) }
    $row = [Math]::Max(1, [int]$In['row'] + [Math]::Max(0, [int]$In['gapRows']))
    $label = [string]$In['label']
    $labelRow = if ($label -ne '') { $row } else { 0 }
    $firstRow = if ($label -ne '') { $row + 1 } else { $row }
    $lastRow = if ($all.Count -gt 0) { $firstRow + $all.Count - 1 } elseif ($label -ne '') { $row } else { $row - 1 }
    $result = @{ ok = $true; labelRow = $labelRow; firstRow = $firstRow; lastRow = $lastRow; nextRow = ($lastRow + 1); written = $all.Count }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would write {0}{1} line(s) to {2}!{3}{4}..{5}' -f $(if ($label) { 'label + ' } else { '' }), $all.Count, $In['sheet'], $In['column'], $row, $lastRow)); return $result }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $In['sheet'] + '"'); labelRow = 0; firstRow = 0; lastRow = 0; nextRow = 0; written = 0 } }
    $col = ConvertTo-EbiColumnNumber ([string]$In['column'])
    try {
        if ($label -ne '') { Set-EbiCellTextPlain -Cell $ws.Cells.Item($row, $col) -Text $label -FontName ([string]$In['fontName']) -FontSize ([double]$In['fontSize']) }
        for ($i = 0; $i -lt $all.Count; $i++) {
            Set-EbiCellTextPlain -Cell $ws.Cells.Item($firstRow + $i, $col) -Text ([string]$all[$i]) -FontName ([string]$In['fontName']) -FontSize ([double]$In['fontSize'])
        }
    } catch { return @{ ok = $false; failure = 'write_failed'; message = $_.Exception.Message; labelRow = $labelRow; firstRow = $firstRow; lastRow = $lastRow; nextRow = ($lastRow + 1); written = 0 } }
    $Ctx.Log.Info(('wrote {0} line(s) to {1}!rows {2}..{3}' -f $all.Count, $In['sheet'], $firstRow, $lastRow))
    return $result
}
