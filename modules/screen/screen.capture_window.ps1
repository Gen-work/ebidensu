# modules/screen/screen.capture_window.ps1
# Screenshot one registered window (by its Session name) to a PNG.
# Source: Common.ps1 Take-WindowScreenshot. The window comes from
# $Ctx.Session, never from "whatever is foreground" (P0-R2).

$kernelDir = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'kernel'
. (Join-Path $kernelDir 'Win32.ps1')

$Manifest = @{
  id         = 'screen.capture_window'
  group      = 'screen'
  summary    = 'Screenshot a registered window to a PNG file'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    window = @{ type='session'; sessionKind='window'; required=$true; desc='Session name of the window to capture (registered by browser.ensure with as:)' }
    saveAs = @{ type='path';    required=$true; desc='PNG path; relative paths are rooted under WorkDir; parent folders are created' }
  }
  outputs    = @{
    path   = @{ type='path'; desc='the file that was written' }
    width  = @{ type='int' }
    height = @{ type='int' }
  }
  failures   = @(
    @{ id = 'window_rect_failed'; transient = $true }
    @{ id = 'capture_failed';     transient = $true }
    @{ id = 'save_failed';        transient = $false }
  )
  example    = @{ use='screen.capture_window'; with=@{ window='mainWindow'; saveAs='capture/{{vars.side}}_{{page.id}}/{{item.keySafe}}.png' } }
  notes      = 'Returns the reserved session_invalid when the registered handle no longer names a window (STEP-CONTRACT 3.1). Does not crop: screen.crop (P1-20) does that.'
}

function ScreenCaptureWindow-EnsureDir {
    param([string]$FilePath)
    $dir = Split-Path -Path $FilePath -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

function ScreenCaptureWindow-SavePng {
    # All System.Drawing use lives here, in a function Invoke-Step only
    # calls on the real path. PowerShell binds [Type] literals when it
    # compiles a function body, so a GDI+ reference inside Invoke-Step
    # itself would run System.Drawing's type initializer even under DryRun
    # -- which is how the CI dry run of this step failed on Linux.
    # Returns @{ ok=$true } or @{ ok=$false; failure=; message= }.
    param($Rect, [string]$Path)
    $bmp = $null; $gfx = $null
    try {
        try {
            Add-Type -AssemblyName System.Drawing
            $bmp = New-Object System.Drawing.Bitmap($Rect.W, $Rect.H)
            $gfx = [System.Drawing.Graphics]::FromImage($bmp)
            $gfx.CopyFromScreen($Rect.X, $Rect.Y, 0, 0, (New-Object System.Drawing.Size($Rect.W, $Rect.H)))
        } catch {
            return @{ ok = $false; failure = 'capture_failed'; message = $_.Exception.Message }
        }
        try {
            ScreenCaptureWindow-EnsureDir -FilePath $Path
            $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
        } catch {
            return @{ ok = $false; failure = 'save_failed'; message = $_.Exception.Message }
        }
    } finally {
        if ($null -ne $gfx) { $gfx.Dispose() }
        if ($null -ne $bmp) { $bmp.Dispose() }
    }
    return @{ ok = $true }
}

function Invoke-Step {
    param($In, $Ctx)

    $handle = $Ctx.Session[[string]$In.window]

    if ($Ctx.DryRun) {
        $Ctx.Log.Info(('would capture window "' + [string]$In.window + '" to ' + [string]$In.saveAs))
        return @{ ok = $true; path = [string]$In.saveAs; width = 0; height = 0 }
    }

    if (-not (Test-EbiWindowHandle -Handle $handle)) {
        return @{ ok = $false; failure = 'session_invalid'; message = ('Session resource "' + [string]$In.window + '" is not a live window any more') }
    }

    $rect = Get-EbiWindowRect -Handle $handle
    if ($null -eq $rect -or $rect.W -le 0 -or $rect.H -le 0) {
        return @{ ok = $false; failure = 'window_rect_failed'; message = ('GetWindowRect gave no usable size for "' + [string]$In.window + '"') }
    }

    $saved = ScreenCaptureWindow-SavePng -Rect $rect -Path ([string]$In.saveAs)
    if (-not $saved.ok) { return $saved }

    return @{ ok = $true; path = [string]$In.saveAs; width = [int]$rect.W; height = [int]$rect.H }
}
