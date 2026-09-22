# modules/browser/browser.ensure.ps1
# Find a window by process name, bring it to the foreground and (when the
# call says as:) register its handle in $Ctx.Session as a 'window'.
# Source: Common.ps1 Get-EdgeMainWindowHandle / Activate-EdgeWindow, with
# the process name made a parameter (P0-R12: the 'window' kind is not tied
# to a browser; df.exe and Excel windows reuse this step).

$kernelDir = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'kernel'
. (Join-Path $kernelDir 'Win32.ps1')

$Manifest = @{
  id         = 'browser.ensure'
  group      = 'browser'
  summary    = 'Find a window by process name, bring it to the foreground, register it'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @('window')
  releases   = @()
  idempotent = $true
  inputs     = @{
    process  = @{ type='string'; default='msedge'; desc='process name without .exe whose main window to find' }
    title    = @{ type='string'; default='';       desc='window-title substring for the AppActivate fallback used only when the process lookup finds no window' }
    settleMs = @{ type='int';    default=400;      desc='wait after activation before checking which window is foreground' }
  }
  outputs    = @{
    processId = @{ type='int';    desc='owning process id' }
    title     = @{ type='string'; desc='main window title as reported by Get-Process' }
    activated = @{ type='bool';   desc='GetForegroundWindow equalled the found window after activation' }
  }
  failures   = @(
    @{ id = 'process_not_found'; transient = $true }
    @{ id = 'activate_failed';   transient = $true }
  )
  example    = @{ use='browser.ensure'; with=@{ as='mainWindow' } }
  notes      = 'Process lookup first, title match only as a fallback, and a real failure when both miss -- the old AppActivate-only path silently "activated" whatever window was already foreground. The handle never appears in outputs; it travels through the resource return key into $Ctx.Session (STEP-CONTRACT 3.4 point 7).'
}

function BrowserEnsure-FindWindow {
    # First process of that name that owns a main window, or $null.
    param([string]$Process)
    $procs = Get-Process -Name $Process -ErrorAction SilentlyContinue
    foreach ($p in $procs) {
        if ($p.MainWindowHandle -ne [IntPtr]::Zero) {
            return @{ Handle = $p.MainWindowHandle; Id = [int]$p.Id; Title = [string]$p.MainWindowTitle }
        }
    }
    return $null
}

function Invoke-Step {
    param($In, $Ctx)

    $asName = $null
    if ($In.ContainsKey('as')) { $asName = $In['as'] }

    if ($Ctx.DryRun) {
        $Ctx.Log.Info(('would activate the main window of process ' + [string]$In.process))
        $r = @{ ok = $true; processId = 0; title = ''; activated = $false }
        if (-not [string]::IsNullOrEmpty([string]$asName)) { $r['resource'] = [IntPtr]::Zero }
        return $r
    }

    $found = BrowserEnsure-FindWindow -Process $In.process
    if ($null -eq $found -and -not [string]::IsNullOrWhiteSpace([string]$In.title)) {
        try {
            $shell = New-Object -ComObject WScript.Shell
            [void]$shell.AppActivate([string]$In.title)
        } catch { }
        Start-Sleep -Milliseconds 700
        $found = BrowserEnsure-FindWindow -Process $In.process
    }
    if ($null -eq $found) {
        return @{ ok = $false; failure = 'process_not_found'; message = ('no window owned by a process named ' + [string]$In.process) }
    }

    $activated = Set-EbiForegroundWindow -Handle $found.Handle -SettleMs ([int]$In.settleMs)
    if (-not $activated) {
        return @{ ok = $false; failure = 'activate_failed'; message = ('found window of ' + [string]$In.process + ' (pid ' + $found.Id + ') but it did not become foreground');
                  processId = $found.Id; title = $found.Title; activated = $false }
    }

    $r = @{ ok = $true; processId = $found.Id; title = $found.Title; activated = $true }
    if (-not [string]::IsNullOrEmpty([string]$asName)) { $r['resource'] = $found.Handle }
    return $r
}
