# modules/browser/browser.send_keys.ps1
# Send a SendKeys sequence to a registered window. Ported from Common.ps1
# (the input is named 'sequence', not 'keys': a hashtable entry called keys
# shadows .Keys for every reader of the manifest.)
# Send-Key (P1-12). The wait is an input, not $Global:Timing; the window is
# brought to front and verified first (P0-R12).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'browser.send_keys'
  group      = 'browser'
  summary    = 'Send a SendKeys sequence to a registered window'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('foreground')
  provides   = @()
  releases   = @()
  idempotent = $false
  inputs     = @{
    window = @{ type='session'; sessionKind='window'; required=$true; desc='the window that receives the keys' }
    sequence = @{ type='string'; required=$true; desc='SendKeys syntax, e.g. ^{f} or {ESC}' }
    waitMs = @{ type='int'; default=300; desc='wait after sending' }
  }
  outputs    = @{ sent = @{ type='string'; desc='the sequence that was sent' } }
  failures   = @(
    @{ id = 'foreground_lost'; transient = $true }
  )
  example    = @{ use = 'browser.send_keys'; with = @{ window = 'mainWindow'; sequence = '{ESC}' } }
  notes      = 'Not idempotent: keys sent twice are two actions. Prefer tab_to / fill / submit / find, which say what they mean; use this for the rest.'
}

function Invoke-Step {
    param($In, $Ctx)
    $keys = [string]$In['sequence']
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would send keys {0}' -f $keys)); return @{ ok = $true; sent = $keys } }
    $fg = Set-EbiForeground -HWnd (ConvertTo-EbiHandle $In['window'])
    if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message'] } }
    Send-EbiKeys -Keys $keys -WaitMs ([int]$In['waitMs'])
    return @{ ok = $true; sent = $keys }
}
