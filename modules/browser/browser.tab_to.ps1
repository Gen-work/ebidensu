# modules/browser/browser.tab_to.ps1
# Tab (or Shift+Tab) N times in a registered window. Ported from Common.ps1
# Send-Tab / Send-ShiftTab (P1-12) -- the HM key sequence in
# PROFILE-SCHEMA 3.0 is "Tab n -> paste -> Shift+Tab m -> Enter", so both
# directions are one step with a switch.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'browser.tab_to'
  group      = 'browser'
  summary    = 'Press Tab (or Shift+Tab) N times in a registered window'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('foreground')
  provides   = @()
  releases   = @()
  idempotent = $false
  inputs     = @{
    window = @{ type='session'; sessionKind='window'; required=$true; desc='the window that receives the keys' }
    times  = @{ type='int'; required=$true; desc='how many presses' }
    shift  = @{ type='bool'; default=$false; desc='Shift+Tab (backwards) instead of Tab' }
    waitMs = @{ type='int'; default=150; desc='wait after each press' }
  }
  outputs    = @{ pressed = @{ type='int'; desc='presses sent' } }
  failures   = @(
    @{ id = 'foreground_lost'; transient = $true }
  )
  example    = @{ use = 'browser.tab_to'; with = @{ window = 'mainWindow'; times = 4 } }
}

function BrowserTabTo-Sequence {
    # PURE. The SendKeys text for one press.
    param([bool]$Shift)
    if ($Shift) { return '+{TAB}' }
    return '{TAB}'
}

function Invoke-Step {
    param($In, $Ctx)
    $n = [int]$In['times']; $shift = [bool]$In['shift']
    if ($n -lt 0) { $n = 0 }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would press {0} x{1}' -f (BrowserTabTo-Sequence -Shift $shift), $n)); return @{ ok = $true; pressed = $n } }
    $fg = Set-EbiForeground -HWnd (ConvertTo-EbiHandle $In['window'])
    if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message'] } }
    $seq = BrowserTabTo-Sequence -Shift $shift
    for ($i = 0; $i -lt $n; $i++) { Send-EbiKeys -Keys $seq -WaitMs ([int]$In['waitMs']) }
    return @{ ok = $true; pressed = $n }
}
