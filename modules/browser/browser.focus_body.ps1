# modules/browser/browser.focus_body.ps1
# Click inside the page body of a registered window so the keyboard focus is
# in the page (not in the address bar, not in the console). Ported from
# Common.ps1 Click-PageBody (P1-11), with the P0-R12 rule: the window is
# named, brought to the front and verified before the click -- the old
# function clicked whatever was in front.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'browser.focus_body'
  group      = 'browser'
  summary    = 'Bring a registered window to front and click inside its body'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('foreground')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    window   = @{ type='session'; sessionKind='window'; required=$true; desc='the window to focus' }
    offsetX  = @{ type='int'; default=150; desc='click x, from the window left edge' }
    offsetY  = @{ type='int'; default=150; desc='click y, from the window top edge' }
    settleMs = @{ type='int'; default=400; desc='wait after the click' }
  }
  outputs    = @{
    x = @{ type='int'; desc='screen x clicked' }
    y = @{ type='int'; desc='screen y clicked' }
  }
  failures   = @(
    @{ id = 'foreground_lost'; transient = $true }
    @{ id = 'window_gone';     transient = $true }
  )
  example    = @{ use = 'browser.focus_body'; with = @{ window = 'mainWindow' } }
  notes      = 'The click lands offsetX/offsetY inside the window, which on a browser is the page body, never the toolbar; keep the offsets above the tab strip height.'
}

function BrowserFocusBody-Point {
    # PURE. Where to click for a window rect and offsets.
    param([hashtable]$Rect, [int]$OffsetX, [int]$OffsetY)
    return @{ x = ([int]$Rect['X'] + $OffsetX); y = ([int]$Rect['Y'] + $OffsetY) }
}

function Invoke-Step {
    param($In, $Ctx)
    $ox = [int]$In['offsetX']; $oy = [int]$In['offsetY']; $settle = [int]$In['settleMs']
    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would bring the window to front and click at +{0},+{1}' -f $ox, $oy))
        return @{ ok = $true; x = 0; y = 0 }
    }
    $hWnd = ConvertTo-EbiHandle $In['window']
    $fg = Set-EbiForeground -HWnd $hWnd
    if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message'] } }
    $rect = Get-EbiWindowRect -HWnd $hWnd
    if (-not $rect['ok']) { return @{ ok = $false; failure = 'window_gone'; message = 'the window has no rectangle (closed?)' } }
    $pt = BrowserFocusBody-Point -Rect $rect -OffsetX $ox -OffsetY $oy
    Invoke-EbiClick -X $pt['x'] -Y $pt['y'] -SettleMs $settle
    return @{ ok = $true; x = $pt['x']; y = $pt['y'] }
}
