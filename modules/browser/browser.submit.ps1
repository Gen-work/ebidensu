# modules/browser/browser.submit.ps1
# Press Enter in a registered window. Ported from Common.ps1 Send-Enter
# (P1-12); verifyChange as in browser.fill (P0-R12).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'browser.submit'
  group      = 'browser'
  summary    = 'Press Enter in a registered window'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('foreground')
  provides   = @()
  releases   = @()
  idempotent = $false
  inputs     = @{
    window       = @{ type='session'; sessionKind='window'; required=$true; desc='the window that receives the keys' }
    waitMs       = @{ type='int'; default=800; desc='wait after Enter (page load time)' }
    verifyChange = @{ type='bool'; default=$false; desc='read the page text before and after; unchanged -> no_effect' }
  }
  outputs    = @{ changed = @{ type='bool'; desc='page text changed (only meaningful with verifyChange)' } }
  failures   = @(
    @{ id = 'foreground_lost'; transient = $true }
    @{ id = 'no_effect';       transient = $true }
  )
  example    = @{ use = 'browser.submit'; with = @{ window = 'mainWindow' } }
  notes      = 'Not idempotent: Enter twice may submit twice. A form whose result takes longer than waitMs to render needs browser.wait_for next.'
}

function Invoke-Step {
    param($In, $Ctx)
    $verify = [bool]$In['verifyChange']
    if ($Ctx['DryRun']) { $Ctx.Log.Info('would press Enter'); return @{ ok = $true; changed = $false } }
    $fg = Set-EbiForeground -HWnd (ConvertTo-EbiHandle $In['window'])
    if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message'] } }
    $before = ''
    if ($verify) { $before = Read-EbiPageText }
    Send-EbiKeys -Keys '{ENTER}' -WaitMs ([int]$In['waitMs'])
    $changed = $false
    if ($verify) {
        $changed = ((Read-EbiPageText) -ne $before)
        if (-not $changed) { return @{ ok = $false; failure = 'no_effect'; message = 'the page text did not change after Enter'; changed = $false } }
    }
    return @{ ok = $true; changed = $changed }
}
