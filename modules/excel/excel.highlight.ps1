# modules/excel/excel.highlight.ps1
# Fill the rows whose text matches a pattern (yellow by default), from
# the text column across as many columns as the text covers: the
# operator's own highlights in the sample evidence end exactly where the
# text ends -- ceil(width / 3) of the 2.625-wide columns for MS Gothic
# 10pt (kernel/LogText.ps1 Get-EbiHighlightEndColumn). Counting characters
# instead of measuring pixels gives the same answer on every PC.

. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')

$Manifest = @{
  id         = 'excel.highlight'
  group      = 'excel'
  summary    = 'Fill the rows whose text matches a pattern, as wide as the text'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook       = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet          = @{ type='string'; required=$true }
    fromRow        = @{ type='int';    required=$true }
    toRow          = @{ type='int';    required=$true }
    column         = @{ type='string'; default='B'; desc='the column holding the text' }
    patterns       = @{ type='list';   required=$true; desc='regexes; a row matching any of them is filled' }
    color          = @{ type='int';    default=65535; desc='OLE colour (BGR); 65535 = yellow' }
    unitsPerColumn = @{ type='int';    default=3; desc='half-width characters one column holds' }
    padColumns     = @{ type='int';    default=0 }
    maxColumn      = @{ type='int';    default=0; desc='never fill past this column number; 0 = no cap' }
    expect         = @{ type='int';    default=-1; desc='how many rows should match; -1 = any. A different count is a warning' }
  }
  outputs    = @{
    rows   = @{ type='list'; desc='@{ row; endColumn (letter); pattern; text } per filled row' }
    filled = @{ type='int' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'not_found';       transient = $false }
    @{ id = 'write_failed';    transient = $true  }
  )
  example    = @{ use = 'excel.highlight'; with = @{ workbook = 'wb'; sheet = 'result'; fromRow = '{{steps.log.out.firstRow}}'; toRow = '{{steps.log.out.lastRow}}'; patterns = @('Command: ') } }
  notes      = 'not_found when no row in the range matches any pattern: the evidence would be missing its mark, which is worth stopping for.'
}

function Invoke-Step {
    param($In, $Ctx)
    $from = [int]$In['fromRow']; $to = [int]$In['toRow']
    $pats = @(@($In['patterns']) | ForEach-Object { [string]$_ })
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would highlight rows {0}..{1} of {2} matching {3}' -f $from, $to, $In['sheet'], ($pats -join ' | '))); return @{ ok = $true; rows = @(); filled = 0 } }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $In['sheet'] + '"'); rows = @(); filled = 0 } }
    $col = ConvertTo-EbiColumnNumber ([string]$In['column'])
    $rx = @(foreach ($p in $pats) { New-Object System.Text.RegularExpressions.Regex($p) })
    $done = New-Object System.Collections.ArrayList
    try {
        for ($r = $from; $r -le $to; $r++) {
            $t = [string]$ws.Cells.Item($r, $col).Value2
            if ([string]::IsNullOrEmpty($t)) { continue }
            for ($k = 0; $k -lt $rx.Count; $k++) {
                if (-not $rx[$k].IsMatch($t)) { continue }
                $end = Get-EbiHighlightEndColumn -Text $t -StartColumn $col -UnitsPerColumn ([double]$In['unitsPerColumn']) -PadColumns ([int]$In['padColumns']) -MaxColumn ([int]$In['maxColumn'])
                Set-EbiRowFill -Worksheet $ws -Row $r -ColStart $col -ColEnd $end -Color ([long]$In['color'])
                [void]$done.Add(@{ row = $r; endColumn = (ConvertTo-EbiColumnLetter $end); pattern = $pats[$k]; text = $(if ($t.Length -gt 120) { $t.Substring(0, 120) } else { $t }) })
                break
            }
        }
    } catch { return @{ ok = $false; failure = 'write_failed'; message = $_.Exception.Message; rows = $done.ToArray(); filled = $done.Count } }
    if ($done.Count -eq 0) { return @{ ok = $false; failure = 'not_found'; message = ('no row in {0}..{1} matches {2}' -f $from, $to, ($pats -join ' | ')); rows = @(); filled = 0 } }
    $out = @{ ok = $true; rows = $done.ToArray(); filled = $done.Count }
    if ([int]$In['expect'] -ge 0 -and $done.Count -ne [int]$In['expect']) { $out['warnings'] = @(@{ code = 'count_mismatch'; message = ('expected ' + $In['expect'] + ' highlighted row(s), filled ' + $done.Count) }) }
    $Ctx.Log.Info(('highlighted {0} row(s): {1}' -f $done.Count, (@($done | ForEach-Object { [string]$_['row'] + '..' + $_['endColumn'] }) -join ', ')))
    return $out
}
