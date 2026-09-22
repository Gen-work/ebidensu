# modules/screen/screen.capture_window.ps1
# Screenshot one window, given its handle from $Ctx.Session, to a PNG.
# Ported from Common.ps1 Take-WindowScreenshot (P0-08).
#
# The window arrives as a 'window' session input: the workflow names it, the
# runner replaces the name with the handle (STEP-CONTRACT.md 3.4 point 7),
# and this step never looks up the foreground window itself -- that is what
# browser.ensure is for.
#
# P1-18 adds screen.capture_region (clamped rect) next to this; cropping is
# screen.crop (P1-20), not an option here.

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

function ScreenCaptureWindow-EnsureNative {
    if (-not ('EbiScreenCaptureNative' -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class EbiScreenCaptureNative {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
}
"@
    }
    Add-Type -AssemblyName System.Drawing
}

function ScreenCaptureWindow-ResolvePath {
    # PURE. Relative -> under the work dir; the runner's path normalization
    # (STEP-CONTRACT.md 2.2) is P1-02, so the step does it for now.
    param([string]$SaveAs, [string]$WorkDir)
    if ([string]::IsNullOrWhiteSpace($SaveAs)) { return '' }
    if ([System.IO.Path]::IsPathRooted($SaveAs)) { return $SaveAs }
    if ([string]::IsNullOrWhiteSpace($WorkDir)) { return $SaveAs }
    # GetFullPath also normalizes separators: a workflow's 'capture/spike/x.png'
    # under 'C:\work' becomes 'C:\work\capture\spike\x.png', not the mixed
    # form the first office-PC run reported.
    return [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($WorkDir, $SaveAs))
}

function ScreenCaptureWindow-ToHandle {
    # Whatever browser.ensure registered: IntPtr, or an int from a fixture.
    param($Value)
    if ($null -eq $Value) { return [IntPtr]::Zero }
    if ($Value -is [IntPtr]) { return $Value }
    try { return [IntPtr]([int64]$Value) } catch { return [IntPtr]::Zero }
}

function ScreenCaptureWindow-Grab {
    # The only function that names System.Drawing types. Kept apart from
    # Invoke-Step so a DryRun (and the Linux CI that dry-runs the shipped
    # workflow) never compiles a reference to GDI+, which is Windows-only.
    param([IntPtr]$HWnd, [string]$Dest)

    ScreenCaptureWindow-EnsureNative
    $rect = New-Object EbiScreenCaptureNative+RECT
    if (-not [EbiScreenCaptureNative]::GetWindowRect($HWnd, [ref]$rect)) {
        return @{ ok = $false; failure = 'window_gone'; message = 'GetWindowRect failed; the window may have been closed'; width = 0; height = 0 }
    }
    $width  = $rect.Right  - $rect.Left
    $height = $rect.Bottom - $rect.Top
    if ($width -le 0 -or $height -le 0) {
        return @{ ok = $false; failure = 'window_gone'; message = ('window rectangle is empty ({0}x{1})' -f $width, $height); width = $width; height = $height }
    }

    $bmp = $null; $gfx = $null
    try {
        $dir = Split-Path -Path $Dest -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $bmp = New-Object System.Drawing.Bitmap($width, $height)
        $gfx = [System.Drawing.Graphics]::FromImage($bmp)
        $gfx.CopyFromScreen($rect.Left, $rect.Top, 0, 0, (New-Object System.Drawing.Size($width, $height)))
        $bmp.Save($Dest, [System.Drawing.Imaging.ImageFormat]::Png)
    } catch {
        return @{ ok = $false; failure = 'save_failed'; message = $_.Exception.Message; width = $width; height = $height }
    } finally {
        if ($null -ne $gfx) { $gfx.Dispose() }
        if ($null -ne $bmp) { $bmp.Dispose() }
    }
    return @{ ok = $true; failure = ''; message = ''; width = $width; height = $height }
}

function Invoke-Step {
    param($In, $Ctx)

    $dest = ScreenCaptureWindow-ResolvePath -SaveAs ([string]$In['saveAs']) -WorkDir ([string]$Ctx['WorkDir'])

    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would capture the window to {0}' -f $dest))
        return @{ ok = $true; path = $dest; width = 0; height = 0 }
    }

    $hWnd = ScreenCaptureWindow-ToHandle $In['window']
    if ($hWnd -eq [IntPtr]::Zero) {
        return @{ ok = $false; failure = 'window_gone'; message = 'the window resource holds no handle' }
    }

    $grab = ScreenCaptureWindow-Grab -HWnd $hWnd -Dest $dest
    if (-not $grab['ok']) {
        return @{ ok = $false; failure = $grab['failure']; message = $grab['message']; path = $dest; width = $grab['width']; height = $grab['height'] }
    }
    return @{ ok = $true; path = $dest; width = $grab['width']; height = $grab['height'] }
}
