# modules/browser/browser.find.ps1
# Ctrl+F for an EXACT string in a registered window and report whether the
# page has it (P1-17). Ctrl+F is a substring search that stops on the FIRST
# listed occurrence, so the caller passes the full identifier of the row it
# means (the one verify.match_record chose), never a bare key -- the old tool
# highlighted the wrong rerun that way. The hit itself is decided from the
# page text (read before the search), not from the find bar, which SendKeys
# cannot read.
#
# The box is EMPTIED before pasting (Invoke-EbiFindText's rule): Edge opens
# Ctrl+F with the last search still in it, and pasting over a selection that
# does not start a new search left the OLD term's matches highlighted --
# 10/26 FJODWAF1, the Jenkins boxes went on the old search's rows (K-46).
# Everything is one step, so nothing re-activates the window between Ctrl+F
# and the paste.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'browser.find'
  group      = 'browser'
  summary    = 'Ctrl+F search for an exact string; report whether it hit'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('foreground')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    window     = @{ type='session'; sessionKind='window'; required=$true; desc='the browser window' }
    term       = @{ type='string'; required=$true; desc='the exact string to search for' }
    closeAfter = @{ type='bool'; default=$false; desc='press Esc afterwards (the highlight goes away too)' }
    readPage   = @{ type='bool'; default=$true; desc='read the page text first (Ctrl+A / Ctrl+C) for hit; false = no page selection is made, hit stays false' }
    waitMs     = @{ type='int'; default=400; desc='wait after the search' }
  }
  outputs    = @{
    hit  = @{ type='bool'; desc='the page text contains the term (false when readPage is false)' }
    rect = @{ type='rect'; desc='pixel rect of the active match; null (P3: browser.find_active_row)' }
  }
  failures   = @(
    @{ id = 'foreground_lost'; transient = $true }
    @{ id = 'clipboard_error'; transient = $true }
  )
  example    = @{ use = 'browser.find'; with = @{ window = 'mainWindow'; term = '{{steps.row.out.name}}' } }
  notes      = 'A miss is not a failure: hit=false is an answer the workflow decides about (verify.assert / human.gate). Leave closeAfter false when a screenshot of the highlighted row follows.'
}

function Invoke-Step {
    param($In, $Ctx)
    $term = [string]$In['term']
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would Ctrl+F for "{0}"' -f $term)); return @{ ok = $true; hit = $false; rect = $null } }
    $fg = Set-EbiForeground -HWnd (ConvertTo-EbiHandle $In['window'])
    if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message'] } }
    $hit = $false
    if ([bool]$In['readPage']) {
        $text = Read-EbiPageText
        $hit = (-not [string]::IsNullOrEmpty($term)) -and ($text.IndexOf($term, [System.StringComparison]::Ordinal) -ge 0)
    }
    try { Set-EbiClipboardText -Text $term } catch { return @{ ok = $false; failure = 'clipboard_error'; message = $_.Exception.Message } }
    Send-EbiKeys -Keys '^{f}' -WaitMs 300
    Send-EbiKeys -Keys '^a' -WaitMs 100
    Send-EbiKeys -Keys '{BACKSPACE}' -WaitMs 250
    Send-EbiKeys -Keys '^v' -WaitMs 600
    Send-EbiKeys -Keys '{ENTER}' -WaitMs ([int]$In['waitMs'])
    if ([bool]$In['closeAfter']) { Send-EbiKeys -Keys '{ESC}' -WaitMs 150 }
    return @{ ok = $true; hit = $hit; rect = $null }
}
