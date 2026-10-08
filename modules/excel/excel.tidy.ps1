# modules/excel/excel.tidy.ps1
# Leave a workbook the way a reviewer expects to open it: every sheet
# scrolled to the top-left with A1 selected (Ctrl+Home on each), copy mode
# off (Esc), the first sheet active -- then save. Optionally report cells
# in a range whose font differs from the house font, as warnings (it does
# not change them: an odd font may be deliberate).

. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.tidy'
  group      = 'excel'
  summary    = 'Select A1 on every sheet, activate the first, report odd fonts, save'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook   = @{ type='session'; sessionKind='workbook'; required=$true }
    activate   = @{ type='string'; default='1'; desc='sheet to leave active (name or 1-based index)' }
    save       = @{ type='bool';   default=$true }
    fontSheets = @{ type='list';   default=@(); desc='sheets whose column-B text is checked against fontName / fontSize' }
    fontName   = @{ type='string'; default='' }
    fontSize   = @{ type='int';    default=0 }
  }
  outputs    = @{
    sheets   = @{ type='int';  desc='sheets visited' }
    oddFonts = @{ type='int';  desc='cells reported' }
    saved    = @{ type='bool' }
  }
  failures   = @(
    @{ id = 'save_failed'; transient = $true }
  )
  example    = @{ use = 'excel.tidy'; with = @{ workbook = 'wb'; fontSheets = @('result'); fontName = 'MS Gothic'; fontSize = 10 } }
}

function Invoke-Step {
    param($In, $Ctx)
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would select A1 on every sheet, activate {0}{1}' -f $In['activate'], $(if ([bool]$In['save']) { ', save' } else { '' }))); return @{ ok = $true; sheets = 0; oddFonts = 0; saved = $false } }
    $wb = $In['workbook']
    $warn = New-Object System.Collections.ArrayList
    $odd = 0
    $fn = [string]$In['fontName']; $fs = [int]$In['fontSize']
    foreach ($name in @($In['fontSheets'])) {
        if ($fn -eq '' -and $fs -le 0) { break }
        $ws = Get-EbiSheet -Workbook $wb -Sheet $name
        if ($null -eq $ws) { continue }
        try {
            $ur = $ws.UsedRange
            $last = [int]$ur.Row + [int]$ur.Rows.Count - 1
            for ($r = 1; $r -le $last; $r++) {
                $c = $ws.Cells.Item($r, 2)
                if ([string]::IsNullOrEmpty([string]$c.Value2)) { continue }
                $bad = ($fn -ne '' -and [string]$c.Font.Name -ne $fn) -or ($fs -gt 0 -and [double]$c.Font.Size -ne $fs)
                if ($bad) { $odd++; if ($odd -le 10) { [void]$warn.Add(@{ code = 'odd_font'; message = ('{0}!B{1}: {2} {3}pt' -f $name, $r, $c.Font.Name, $c.Font.Size) }) } }
            }
        } catch { [void]$warn.Add(@{ code = 'font_check_failed'; message = ($name + ': ' + $_.Exception.Message) }) }
    }
    if ($odd -gt 10) { [void]$warn.Add(@{ code = 'odd_fonts'; message = ('' + $odd + ' cell(s) in all (10 listed)') }) }
    $n = 0
    try { $wb.Application.CutCopyMode = $false } catch { }
    foreach ($ws in @($wb.Worksheets)) {
        try {
            if ([int]$ws.Visible -ne -1) { continue }
            [void]$ws.Activate()
            $win = $wb.Application.ActiveWindow
            $win.ScrollRow = 1
            $win.ScrollColumn = 1
            [void]$ws.Range('A1').Select()
            $n++
        } catch { [void]$warn.Add(@{ code = 'sheet_not_reset'; message = ([string]$ws.Name + ': ' + $_.Exception.Message) }) }
    }
    $first = Get-EbiSheet -Workbook $wb -Sheet $In['activate']
    if ($null -ne $first) { try { [void]$first.Activate() } catch { } }
    $saved = $false
    if ([bool]$In['save']) {
        try { $wb.Save(); $saved = $true } catch { return @{ ok = $false; failure = 'save_failed'; message = $_.Exception.Message; sheets = $n; oddFonts = $odd; saved = $false; warnings = $warn.ToArray() } }
    }
    return @{ ok = $true; sheets = $n; oddFonts = $odd; saved = $saved; warnings = $warn.ToArray() }
}
