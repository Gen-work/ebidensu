# modules/excel/excel.close.ps1
# Close a registered workbook, saving it or not, and release it (P4-01).
# Safe when the workbook is already gone (ExcelHelpers.ps1 Close-Workbook:
# a $null is not an error).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Start/Stop-EbiDialogWatch around Save
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
    dismissDialogs = @{ type='list'; default=@(); desc='title parts of dialogs the save opens (the classification add-in) answered with Esc, as excel.tidy / excel.append_row do; empty = the profile layout.saveDialogs / layout.evidence.saveDialogs' }
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
        $err = ''
        $watch = $null
        $dlg = @(Get-EbiSaveDialogs -Given $In['dismissDialogs'] -Profile $Ctx['Profile'])
        if ($dlg.Count -gt 0) { $watch = Start-EbiDialogWatch -TitleParts $dlg }
        # 10/20: the WBS save "succeeded" but the file kept no change (a save
        # dialog -- the classification add-in -- cancelled it without an
        # error). A workbook with changes must leave a newer file behind.
        $changed = $false; $file = ''; $before = [datetime]::MinValue
        try { $changed = -not [bool]$wb.Saved; $file = [string]$wb.FullName; if (Test-Path -LiteralPath $file) { $before = (Get-Item -LiteralPath $file).LastWriteTimeUtc } } catch { }
        try { [void]$wb.Save(); $saved = $true } catch { $err = $_.Exception.Message }
        if ($null -ne $watch) { $seen = Stop-EbiDialogWatch $watch; foreach ($t in @($seen['answered'])) { $Ctx.Log.Info(('answered "{0}" with Esc while saving' -f $t)) } }
        if ($err -eq '' -and $changed -and $file -ne '' -and $before -ne [datetime]::MinValue) {
            $after = $before
            try { $after = (Get-Item -LiteralPath $file).LastWriteTimeUtc } catch { }
            $stillDirty = $true
            try { $stillDirty = -not [bool]$wb.Saved } catch { }
            if ($after -le $before -and $stillDirty) { $err = ('the save did not reach the file (' + $file + ' unchanged; a dialog cancelled it?) -- the workbook stays open: save it by hand, then r') }
        }
        if ($err -ne '') { return @{ ok = $false; failure = 'save_failed'; message = $err; saved = $false; closed = $false } }
    }
    try { $wb.Close($false) } catch { }
    Invoke-EbiComRelease $wb
    return @{ ok = $true; saved = $saved; closed = $true }
}
