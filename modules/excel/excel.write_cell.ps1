# modules/excel/excel.write_cell.ps1
# Write one cell: a given value, or the value of another cell (a label the
# deliverable repeats from its first sheet). Written as text with a plain
# font so Excel does not turn a code into a date or a number.

. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.write_cell'
  group      = 'excel'
  summary    = 'Write a value (or another cell''s value) into one cell as text'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook  = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet     = @{ type='string'; required=$true }
    cell      = @{ type='string'; required=$true; desc='address, e.g. B3' }
    value     = @{ type='string'; default=''; desc='the text to write' }
    fromSheet = @{ type='string'; default=''; desc='instead of value: copy the value of fromCell on this sheet' }
    fromCell  = @{ type='string'; default='' }
    fontName  = @{ type='string'; default='' }
    fontSize  = @{ type='int';    default=0 }
  }
  outputs    = @{
    value = @{ type='string'; desc='what the cell holds now' }
    row   = @{ type='int' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'write_failed';    transient = $true  }
  )
  example    = @{ use = 'excel.write_cell'; with = @{ workbook = 'wb'; sheet = 'compare'; cell = 'B3'; fromSheet = 'data'; fromCell = 'A3' } }
}

function Invoke-Step {
    param($In, $Ctx)
    $wb = $In['workbook']
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would write {0}!{1} = {2}' -f $In['sheet'], $In['cell'], $(if ([string]$In['fromSheet'] -ne '') { $In['fromSheet'] + '!' + $In['fromCell'] } else { '"' + $In['value'] + '"' }))); return @{ ok = $true; value = [string]$In['value']; row = 0 } }
    $ws = Get-EbiSheet -Workbook $wb -Sheet $In['sheet']
    if ($null -eq $ws) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $In['sheet'] + '"'); value = ''; row = 0 } }
    $val = [string]$In['value']
    if ([string]$In['fromSheet'] -ne '') {
        $src = Get-EbiSheet -Workbook $wb -Sheet $In['fromSheet']
        if ($null -eq $src) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $In['fromSheet'] + '"'); value = ''; row = 0 } }
        try { $val = [string]$src.Range([string]$In['fromCell']).Text } catch { return @{ ok = $false; failure = 'write_failed'; message = $_.Exception.Message; value = ''; row = 0 } }
    }
    try {
        $c = $ws.Range([string]$In['cell'])
        Set-EbiCellTextPlain -Cell $c -Text $val -FontName ([string]$In['fontName']) -FontSize ([double]$In['fontSize'])
        $row = [int]$c.Row
    } catch { return @{ ok = $false; failure = 'write_failed'; message = $_.Exception.Message; value = ''; row = 0 } }
    return @{ ok = $true; value = $val; row = $row }
}
