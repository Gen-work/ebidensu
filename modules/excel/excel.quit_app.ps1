# modules/excel/excel.quit_app.ps1
# Quit the Excel instance excel.ensure_app started and release it (P4-01).
# Safe when the instance is already gone: teardown runs on every resume,
# including after a Ctrl+C that never registered anything (STEP-CONTRACT
# 6.2). Ported from ExcelHelpers.ps1 Close-ExcelApp.

. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.quit_app'
  group      = 'excel'
  summary    = 'Quit the registered Excel instance and release it'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @('excelApp')
  idempotent = $true
  inputs     = @{
    app = @{ type='session'; sessionKind='excelApp'; required=$true }
  }
  outputs    = @{
    quit = @{ type='bool'; desc='false when there was nothing to quit' }
  }
  failures   = @(
    @{ id = 'internal_error'; transient = $false }
  )
  example    = @{ use = 'excel.quit_app'; with = @{ app = 'xl' } }
  notes      = 'Workbooks still open in the instance are closed WITHOUT saving -- close them with excel.close (save = true) first.'
}

function Invoke-Step {
    param($In, $Ctx)
    $xl = $In['app']
    if ($Ctx['DryRun']) { $Ctx.Log.Info('would quit the Excel instance'); return @{ ok = $true; quit = $false } }
    if ($null -eq $xl) { return @{ ok = $true; quit = $false } }
    try { $xl.DisplayAlerts = $false } catch { }
    try { foreach ($wb in @($xl.Workbooks)) { try { $wb.Close($false) } catch { } } } catch { }
    try { $xl.Quit() } catch { }
    Invoke-EbiComRelease $xl
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    return @{ ok = $true; quit = $true }
}
