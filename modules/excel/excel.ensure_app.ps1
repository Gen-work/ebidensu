# modules/excel/excel.ensure_app.ps1
# Start an Excel instance of our own and register it as an 'excelApp'
# resource (P4-01, one of the four lifecycle steps: ensure_app / open /
# close / quit_app -- a step registers at most one resource, and the
# Application outlives every workbook, STEP-CONTRACT 3.4 point 6).
# Ported from ExcelHelpers.ps1 New-ExcelApp; Visible is set BEFORE
# DisplayAlerts (the project's Excel COM rule).

. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.ensure_app'
  group      = 'excel'
  summary    = 'Start a dedicated Excel instance and register it'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('excel')
  provides   = @('excelApp')
  releases   = @()
  idempotent = $true
  inputs     = @{
    visible = @{ type='bool'; default=$true; desc='show the window (the operator sees what is written)' }
  }
  outputs    = @{
    version = @{ type='string'; desc='Excel version, e.g. 16.0' }
  }
  failures   = @(
    @{ id = 'excel_unavailable'; transient = $false }
  )
  example    = @{ use = 'excel.ensure_app'; with = @{ as = 'xl' } }
  notes      = 'Always a NEW instance: workbooks the operator has open in their own Excel are not touched (excel.open reports read_only if one of them is the target). Release with excel.quit_app in teardown.'
}

function Invoke-Step {
    param($In, $Ctx)
    if ($Ctx['DryRun']) { $Ctx.Log.Info('would start a new Excel instance'); return @{ ok = $true; resource = $null; version = '' } }
    $xl = $null
    try { $xl = New-Object -ComObject Excel.Application }
    catch { return @{ ok = $false; failure = 'excel_unavailable'; message = $_.Exception.Message; version = '' } }
    try { $xl.Visible = [bool]$In['visible'] } catch { }
    try { $xl.DisplayAlerts = $false } catch { }
    $v = ''
    try { $v = [string]$xl.Version } catch { }
    return @{ ok = $true; resource = $xl; version = $v }
}
