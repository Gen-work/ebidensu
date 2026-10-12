# modules/excel/excel.close.ps1
# Close a registered workbook, saving it or not, and release it (P4-01).
# Safe when the workbook is already gone (ExcelHelpers.ps1 Close-Workbook:
# a $null is not an error).

. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.close'
  group      = 'excel'
  summary    = 'Close a registered workbook (save or discard) and release it'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @('workbook')
  idempotent = $true
  inputs     = @{
    workbook = @{ type='session'; sessionKind='workbook'; required=$true }
    save     = @{ type='bool'; default=$false; desc='save before closing' }
  }
  outputs    = @{
    saved  = @{ type='bool' }
    closed = @{ type='bool'; desc='false when there was nothing to close' }
  }
  failures   = @(
    @{ id = 'save_failed'; transient = $true }
  )
  example    = @{ use = 'excel.close'; with = @{ workbook = 'wb'; save = $true } }
  notes      = 'save_failed keeps the workbook open and registered (the runner only releases on ok), so a retry after fixing the cause (a full disk, a locked file) can still save it.'
}

function Invoke-Step {
    param($In, $Ctx)
    $wb = $In['workbook']
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would close the workbook{0}' -f $(if ([bool]$In['save']) { ' after saving' } else { ' without saving' }))); return @{ ok = $true; saved = $false; closed = $false } }
    if ($null -eq $wb) { return @{ ok = $true; saved = $false; closed = $false } }
    $saved = $false
    if ([bool]$In['save']) {
        try { $wb.Save(); $saved = $true }
        catch { return @{ ok = $false; failure = 'save_failed'; message = $_.Exception.Message; saved = $false; closed = $false } }
    }
    try { $wb.Close($false) } catch { }
    Invoke-EbiComRelease $wb
    return @{ ok = $true; saved = $saved; closed = $true }
}
