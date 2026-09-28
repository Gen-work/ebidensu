# modules/screen/screen.fit_window.ps1
# Move and size a registered window (MoveWindow). Ported from MqSnap.ps1
# Move-EdgeAwayFromBorder (P1-19): a window flush against the screen edge
# loses its border to the DWM shadow in a capture, so the old tool parked it
# at 40,40 with the configured size before every snap.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'screen.fit_window'
  group      = 'screen'
  summary    = 'Move and size a registered window to x,y width x height'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    window   = @{ type='session'; sessionKind='window'; required=$true; desc='the window to move' }
    x        = @{ type='int'; default=40; desc='left edge, screen pixels' }
    y        = @{ type='int'; default=40; desc='top edge, screen pixels' }
    width    = @{ type='int'; required=$true }
    height   = @{ type='int'; required=$true }
    settleMs = @{ type='int'; default=300; desc='wait after the move' }
  }
  outputs    = @{
    x      = @{ type='int'; desc='where the window actually is afterwards' }
    y      = @{ type='int' }
    width  = @{ type='int' }
    height = @{ type='int' }
  }
  failures   = @(
    @{ id = 'window_gone'; transient = $true }
    @{ id = 'move_failed'; transient = $true }
  )
  example    = @{ use = 'screen.fit_window'; with = @{ window = 'mainWindow'; width = '{{profile.window.width}}'; height = '{{profile.window.height}}' } }
  notes      = 'The outputs are read back with GetWindowRect: a window with a minimum size larger than asked reports what it settled on, and a workflow that cares compares them (verify.assert).'
}

function Invoke-Step {
    param($In, $Ctx)
    $x = [int]$In['x']; $y = [int]$In['y']; $w = [int]$In['width']; $h = [int]$In['height']
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would move the window to {0},{1} {2}x{3}' -f $x, $y, $w, $h)); return @{ ok = $true; x = $x; y = $y; width = $w; height = $h } }
    $hWnd = ConvertTo-EbiHandle $In['window']
    [void](Get-EbiNative)
    if ($hWnd -eq [IntPtr]::Zero -or -not [EbiNative]::IsWindow($hWnd)) { return @{ ok = $false; failure = 'window_gone'; message = 'the window handle is not a window (closed?)' } }
    [void][EbiNative]::ShowWindowAsync($hWnd, 9)   # SW_RESTORE: a maximized window ignores MoveWindow
    Start-Sleep -Milliseconds 100
    if (-not [EbiNative]::MoveWindow($hWnd, $x, $y, $w, $h, $true)) { return @{ ok = $false; failure = 'move_failed'; message = 'MoveWindow returned false' } }
    Start-Sleep -Milliseconds ([Math]::Max(0, [int]$In['settleMs']))
    $rect = Get-EbiWindowRect -HWnd $hWnd
    if (-not $rect['ok']) { return @{ ok = $false; failure = 'window_gone'; message = 'the window has no rectangle after the move' } }
    return @{ ok = $true; x = $rect['X']; y = $rect['Y']; width = $rect['W']; height = $rect['H'] }
}
