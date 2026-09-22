# modules/browser/browser.ensure.ps1
# Find the browser's main window, bring it to the foreground, and hand its
# handle to the runner as a 'window' resource (STEP-CONTRACT.md 3.4 point 7).
# Ported from Common.ps1 Activate-EdgeWindow (P0-08).
#
# Process handle first, title match as a fallback only: the old
# AppActivate-by-title path silently "activated" whatever window was already
# in front when the title text did not match (see the Common.ps1 comment
# above Get-EdgeMainWindowHandle). Both paths failing is a real failure here,
# not a warning.
#
# Nothing in this file knows what a page is. Which browser to look for is an
# input with a default, so a profile can say 'chrome' without touching code.
#
# The Win32 declarations are compiled lazily inside a prefixed helper so the
# file dot-sources cleanly anywhere (the contract checker loads it on Linux).
# screen.capture_window carries its own three-line copy of the same
# declarations; P1-11/P1-18 decide whether a shared native binding is worth
# a file of its own.

$Manifest = @{
  id         = 'browser.ensure'
  group      = 'browser'
  summary    = 'Find the browser main window, bring it to front, register it'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @('window')
  releases   = @()
  idempotent = $true
  inputs     = @{
    process  = @{ type='string'; default='msedge';         desc='process name (without .exe) whose main window to use' }
    title    = @{ type='string'; default='Microsoft Edge'; desc='window-title substring for the AppActivate fallback' }
    settleMs = @{ type='int';    default=400;              desc='wait after bringing the window to front' }
  }
  outputs    = @{
    processId = @{ type='int';    desc='PID that owns the window' }
    title     = @{ type='string'; desc='window title at the time it was found' }
  }
  failures   = @(
    @{ id = 'no_browser_window'; transient = $true }
  )
  example    = @{
    use  = 'browser.ensure'
    with = @{ as = 'mainWindow' }
  }
  notes      = 'Registers the window handle under with.as; the handle itself never appears in outputs. Several windows of the process: the first with a main window wins, as before.'
}

function BrowserEnsure-EnsureNative {
    if (-not ('EbiBrowserEnsureNative' -as [type])) {
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class EbiBrowserEnsureNative {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}
"@
    }
}

function BrowserEnsure-FindProcess {
    # The first process of that name owning a main window, or $null.
    param([string]$ProcessName)
    $procs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero })
    if ($procs.Count -eq 0) { return $null }
    return $procs[0]
}

function BrowserEnsure-TryTitleFallback {
    # WScript.Shell.AppActivate by title substring. Only reached when the
    # process lookup found nothing; a local COM object, never a global.
    param([string]$Title)
    try {
        $shell = New-Object -ComObject WScript.Shell
        [void]$shell.AppActivate($Title)
    } catch { }
    Start-Sleep -Milliseconds 700
}

function Invoke-Step {
    param($In, $Ctx)

    $processName = if ($In.Contains('process')  -and -not [string]::IsNullOrWhiteSpace([string]$In['process'])) { [string]$In['process'] } else { 'msedge' }
    $title       = if ($In.Contains('title'))    { [string]$In['title'] }  else { 'Microsoft Edge' }
    $settleMs    = if ($In.Contains('settleMs')) { [int]$In['settleMs'] }  else { 400 }

    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would find the main window of {0} and bring it to front' -f $processName))
        return @{ ok = $true; resource = $null; processId = 0; title = '' }
    }

    $proc = BrowserEnsure-FindProcess -ProcessName $processName
    if ($null -eq $proc) {
        BrowserEnsure-TryTitleFallback -Title $title
        $proc = BrowserEnsure-FindProcess -ProcessName $processName
    }
    if ($null -eq $proc) {
        return @{ ok = $false; failure = 'no_browser_window';
                  message = ('no {0} process with a main window (process lookup and "{1}" title match both failed)' -f $processName, $title) }
    }

    BrowserEnsure-EnsureNative
    $hWnd = $proc.MainWindowHandle
    [void][EbiBrowserEnsureNative]::ShowWindowAsync($hWnd, 9)   # SW_RESTORE
    [void][EbiBrowserEnsureNative]::SetForegroundWindow($hWnd)
    if ($settleMs -gt 0) { Start-Sleep -Milliseconds $settleMs }

    return @{ ok = $true; resource = $hWnd; processId = [int]$proc.Id; title = [string]$proc.MainWindowTitle }
}
