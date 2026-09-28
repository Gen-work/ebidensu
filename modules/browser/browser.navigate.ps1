# modules/browser/browser.navigate.ps1
# Ctrl+L, paste a URL, Enter in a registered window (P1-16). With no URL the
# step does nothing and says so with a warning: opening the page is then a
# human.prepare gate's job, and a workflow that expects this step to have
# navigated can check {{steps.<id>.out.navigated}}.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'browser.navigate'
  group      = 'browser'
  summary    = 'Ctrl+L, paste a URL, Enter; no URL means leave it to a person'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('foreground')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    window       = @{ type='session'; sessionKind='window'; required=$true; desc='the browser window' }
    url          = @{ type='string'; default=''; desc='where to go; empty = do nothing (warning no_url)' }
    waitMs       = @{ type='int'; default=1500; desc='wait after Enter' }
    verifyChange = @{ type='bool'; default=$false; desc='read the page text before and after; unchanged -> no_effect' }
  }
  outputs    = @{
    navigated = @{ type='bool'; desc='false when no URL was given' }
    url       = @{ type='string' }
  }
  failures   = @(
    @{ id = 'foreground_lost'; transient = $true }
    @{ id = 'no_effect';       transient = $true }
    @{ id = 'clipboard_error'; transient = $true }
  )
  example    = @{ use = 'browser.navigate'; with = @{ window = 'mainWindow'; url = '{{page.url}}' } }
}

function Invoke-Step {
    param($In, $Ctx)
    $url = [string]$In['url']
    if ([string]::IsNullOrWhiteSpace($url)) {
        $Ctx.Log.Warn('no URL: nothing navigated; a human.prepare gate should have the page opened')
        return @{ ok = $true; navigated = $false; url = ''; warnings = @( @{ code = 'no_url'; message = 'no URL given; the page must be opened by hand'; data = @{} } ) }
    }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would navigate to {0}' -f $url)); return @{ ok = $true; navigated = $true; url = $url } }
    $fg = Set-EbiForeground -HWnd (ConvertTo-EbiHandle $In['window'])
    if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message'] } }
    $before = ''
    if ([bool]$In['verifyChange']) { $before = Read-EbiPageText }
    try { Set-EbiClipboardText -Text $url } catch { return @{ ok = $false; failure = 'clipboard_error'; message = $_.Exception.Message } }
    Send-EbiKeys -Keys '^{l}' -WaitMs 200
    Send-EbiKeys -Keys '^v' -WaitMs 200
    Send-EbiKeys -Keys '{ENTER}' -WaitMs ([int]$In['waitMs'])
    if ([bool]$In['verifyChange'] -and ((Read-EbiPageText) -eq $before)) { return @{ ok = $false; failure = 'no_effect'; message = 'the page text did not change after navigating'; navigated = $true; url = $url } }
    return @{ ok = $true; navigated = $true; url = $url }
}
