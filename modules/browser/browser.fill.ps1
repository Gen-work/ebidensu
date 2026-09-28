# modules/browser/browser.fill.ps1
# Replace the focused input's content with a text: clipboard + Ctrl+A +
# Ctrl+V. Ported from Common.ps1 Paste-Replace (P1-12). With verifyChange
# the page text is read before and after and a page that did not change is
# no_effect (P0-R12: the verify_action wrapper became this input).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'browser.fill'
  group      = 'browser'
  summary    = 'Paste a text over the focused input (Ctrl+A, Ctrl+V)'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('foreground')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    window       = @{ type='session'; sessionKind='window'; required=$true; desc='the window that receives the keys' }
    text         = @{ type='string'; required=$true; desc='what to put in the field' }
    waitMs       = @{ type='int'; default=300; desc='wait after pasting' }
    verifyChange = @{ type='bool'; default=$false; desc='read the page text before and after; unchanged -> no_effect' }
  }
  outputs    = @{
    length  = @{ type='int';  desc='characters pasted' }
    changed = @{ type='bool'; desc='page text changed (only meaningful with verifyChange)' }
  }
  failures   = @(
    @{ id = 'foreground_lost'; transient = $true }
    @{ id = 'no_effect';       transient = $true }
    @{ id = 'clipboard_error'; transient = $true }
  )
  example    = @{ use = 'browser.fill'; with = @{ window = 'mainWindow'; text = '{{item.Correl_ID_S}}'; verifyChange = $true } }
  notes      = 'Idempotent: pasting the same text twice leaves the same field content. Reading the page text for verifyChange sends Ctrl+A/Ctrl+C/Esc, which moves the selection; the field keeps its value.'
}

function Invoke-Step {
    param($In, $Ctx)
    $text = [string]$In['text']; $verify = [bool]$In['verifyChange']
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would paste {0} character(s) over the focused field' -f $text.Length)); return @{ ok = $true; length = $text.Length; changed = $false } }
    $hWnd = ConvertTo-EbiHandle $In['window']
    $fg = Set-EbiForeground -HWnd $hWnd
    if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message'] } }
    $before = ''
    if ($verify) { $before = Read-EbiPageText }
    try { Set-EbiClipboardText -Text $text } catch { return @{ ok = $false; failure = 'clipboard_error'; message = $_.Exception.Message } }
    Send-EbiKeys -Keys '^{a}' -WaitMs 150
    Send-EbiKeys -Keys '^v' -WaitMs ([int]$In['waitMs'])
    $changed = $false
    if ($verify) {
        $after = Read-EbiPageText
        $changed = ($after -ne $before)
        if (-not $changed) { return @{ ok = $false; failure = 'no_effect'; message = 'the page text did not change after pasting'; length = $text.Length; changed = $false } }
    }
    return @{ ok = $true; length = $text.Length; changed = $changed }
}
