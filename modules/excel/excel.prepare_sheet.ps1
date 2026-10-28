# modules/excel/excel.prepare_sheet.ps1
# Make a sheet ready to be filled: create it when the book does not have it
# yet (a copy of another sheet's look -- column widths, fonts, the A1
# sheet-name formula), and clear what an earlier fill left on it: every
# shape (pictures, red boxes, Ctrl+G groups), every value, every fill
# colour -- except the cells named in keepCells.
#
# 10/28 POM books: GFIX recv result (after) starts as a COPY of the first
# run's sheet (Convert-PomBook.ps1), so the rerun's evidence would land on
# top of the first run's log lines, yellow rows and boxes; and
# GFIX data HOST vs OPEN exists only when the before / after data differ.

. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.prepare_sheet'
  group      = 'excel'
  summary    = 'Create a sheet from another one''s look when missing, and clear its shapes / values / fills except keepCells'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook  = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet     = @{ type='string'; required=$true; desc='the sheet to prepare' }
    template  = @{ type='string'; default=''; desc='when sheet is missing: copy this sheet (its look) and name the copy sheet; empty = a missing sheet is a failure' }
    after     = @{ type='string'; default=''; desc='a new sheet goes right after this one; empty = right after template' }
    keepCells = @{ type='list';   default=@('A1'); desc='cells whose value survives the clear (the A1 sheet-name formula)' }
    clearExisting = @{ type='bool'; default=$true; desc='false = only create a missing sheet, leave an existing one as it is' }
  }
  outputs    = @{
    sheet         = @{ type='string'; desc='the sheet name' }
    created       = @{ type='bool';   desc='the sheet was made by this call' }
    shapesRemoved = @{ type='int';    desc='shapes deleted by the clear (a group counts once)' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'write_failed';    transient = $true  }
  )
  example    = @{ use = 'excel.prepare_sheet'; with = @{ workbook = 'wb'; sheet = 'hostVsOpen'; template = 'compare'; keepCells = @('A1', 'B3') } }
  notes      = 'The clear keeps the cells'' formats (fonts, borders) but removes fill colours (the yellow highlight rows). Column widths and row heights are not touched.'
}

function Invoke-Step {
    param($In, $Ctx)
    $name = [string]$In['sheet']
    $keep = @(@($In['keepCells']) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would prepare sheet "{0}" (template "{1}", keep {2}, clear an existing one: {3})' -f $name, $In['template'], ($keep -join ','), [bool]$In['clearExisting']))
        return @{ ok = $true; sheet = $name; created = $false; shapesRemoved = 0 }
    }
    $wb = $In['workbook']
    $ws = Get-EbiSheet -Workbook $wb -Sheet $name
    $created = $false
    if ($null -eq $ws) {
        $tplName = [string]$In['template']
        if ($tplName -eq '') { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $name + '" and no template to make it from'); sheet = $name; created = $false; shapesRemoved = 0 } }
        $tpl = Get-EbiSheet -Workbook $wb -Sheet $tplName
        if ($null -eq $tpl) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no template sheet "' + $tplName + '"'); sheet = $name; created = $false; shapesRemoved = 0 } }
        $afterWs = $tpl
        if ([string]$In['after'] -ne '') {
            $afterWs = Get-EbiSheet -Workbook $wb -Sheet ([string]$In['after'])
            if ($null -eq $afterWs) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $In['after'] + '" to put the new sheet after'); sheet = $name; created = $false; shapesRemoved = 0 } }
        }
        try {
            $tpl.Copy([System.Reflection.Missing]::Value, $afterWs)
            $ws = $wb.Worksheets.Item([int]$afterWs.Index + 1)
            $ws.Name = $name
        } catch { return @{ ok = $false; failure = 'write_failed'; message = ('could not copy "' + $tplName + '" to "' + $name + '": ' + $_.Exception.Message); sheet = $name; created = $false; shapesRemoved = 0 } }
        $created = $true
        $Ctx.Log.Info(('made sheet "{0}" from "{1}"' -f $name, $tplName))
    }
    $removed = 0
    if ([bool]$In['clearExisting'] -or $created) {
        try {
            # Shapes lists top-level items only: a Ctrl+G group goes as one
            for ($i = [int]$ws.Shapes.Count; $i -ge 1; $i--) { $ws.Shapes.Item($i).Delete(); $removed++ }
            $saved = @{}
            foreach ($c in $keep) { $saved[$c] = [string]$ws.Range($c).Formula }
            [void]$ws.Cells.ClearContents()
            $ws.Cells.Interior.Pattern = -4142   # xlNone: the yellow highlight rows
            foreach ($c in $keep) { if ($saved[$c] -ne '') { $ws.Range($c).Formula = $saved[$c] } }
        } catch { return @{ ok = $false; failure = 'write_failed'; message = ('could not clear "' + $name + '": ' + $_.Exception.Message); sheet = $name; created = $created; shapesRemoved = $removed } }
        $Ctx.Log.Info(('cleared sheet "{0}": {1} shape(s), values and fills (kept {2})' -f $name, $removed, ($keep -join ',')))
    }
    return @{ ok = $true; sheet = $name; created = $created; shapesRemoved = $removed }
}
