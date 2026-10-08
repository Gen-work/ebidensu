# modules/screen/screen.capture_window.ps1
# Screenshot one window, given its handle from $Ctx.Session, to a PNG.
# Ported from Common.ps1 Take-WindowScreenshot (P0-08); Win32 and GDI+ moved
# to kernel/Native.ps1 + kernel/Image.ps1 (P1-18).
#
# The window arrives as a 'window' session input: the workflow names it, the
# runner replaces the name with the handle (STEP-CONTRACT.md 3.4 point 7),
# and this step never looks up the foreground window itself -- that is what
# browser.ensure is for. Cropping is screen.crop (P1-20), not an option here.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Image.ps1')

$Manifest = @{
  id         = 'screen.capture_window'
  group      = 'screen'
  summary    = 'Save a PNG screenshot of the window held in a session resource'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    window = @{ type='session'; sessionKind='window'; required=$true; desc='name of a registered window resource' }
    saveAs = @{ type='path';    required=$true;                        desc='PNG path; relative paths resolve under the work dir; parent dirs are created' }
    bounds = @{ type='string';  default='window'; enum=@('window', 'visible'); desc='window = GetWindowRect (Windows 10 pads it with invisible borders, which the per-side crop then takes off); visible = the DWM frame bounds, no padding' }
  }
  outputs    = @{
    path   = @{ type='path'; desc='the file that was written (absolute)' }
    width  = @{ type='int' }
    height = @{ type='int' }
  }
  failures   = @(
    @{ id = 'window_gone';  transient = $true }
    @{ id = 'save_failed';  transient = $true }
  )
  example    = @{
    use  = 'screen.capture_window'
    with = @{ window = 'mainWindow'; saveAs = 'capture/before_list/{{item.keySafe}}.png' }
  }
  notes      = 'Captures the window rectangle from the screen (CopyFromScreen), so the window must be in front and unobscured -- run browser.ensure first.'
}

function Invoke-Step {
    param($In, $Ctx)

    $dest = Resolve-EbiWorkPath -PathValue ([string]$In['saveAs']) -WorkDir ([string]$Ctx['WorkDir'])

    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would capture the window to {0}' -f $dest))
        return @{ ok = $true; path = $dest; width = 0; height = 0 }
    }

    $hWnd = ConvertTo-EbiHandle $In['window']
    if ($hWnd -eq [IntPtr]::Zero) {
        return @{ ok = $false; failure = 'window_gone'; message = 'the window resource holds no handle' }
    }
    $rect = if ([string]$In['bounds'] -eq 'visible') { Get-EbiWindowVisibleRect -HWnd $hWnd } else { Get-EbiWindowRect -HWnd $hWnd }
    if (-not $rect['ok']) {
        return @{ ok = $false; failure = 'window_gone'; message = ('the window has no rectangle ({0}x{1}); closed?' -f $rect['W'], $rect['H']); path = $dest; width = 0; height = 0 }
    }
    $grab = Save-EbiScreenRegionPng -X $rect['X'] -Y $rect['Y'] -W $rect['W'] -H $rect['H'] -Dest $dest
    if (-not $grab['ok']) {
        return @{ ok = $false; failure = 'save_failed'; message = $grab['message']; path = $dest; width = $rect['W']; height = $rect['H'] }
    }
    return @{ ok = $true; path = $dest; width = $rect['W']; height = $rect['H'] }
}
