# modules/browser/browser.read_text.ps1
# The visible text of a registered window's page (Ctrl+A, Ctrl+C, Esc,
# clipboard), optionally archived to a file. Ported from Read-PageText.ps1
# (P1-13). Archiving is the norm, not the exception: text that exists is
# never OCR'd (Plan.md 6.1).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'browser.read_text'
  group      = 'browser'
  summary    = 'Read the page text of a registered window via the clipboard'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('foreground')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    window       = @{ type='session'; sessionKind='window'; required=$true; desc='the window to read' }
    selectWaitMs = @{ type='int'; default=400; desc='wait after Ctrl+A' }
    copyWaitMs   = @{ type='int'; default=400; desc='wait after Ctrl+C' }
    archiveTo    = @{ type='path'; default=''; desc='also write the text here (relative: under the work dir)' }
  }
  outputs    = @{
    text   = @{ type='string'; desc='the page text' }
    length = @{ type='int' }
    path   = @{ type='path'; desc='where it was archived, or empty' }
  }
  failures   = @(
    @{ id = 'foreground_lost'; transient = $true }
    @{ id = 'archive_failed';  transient = $true }
  )
  example    = @{ use = 'browser.read_text'; with = @{ window = 'mainWindow'; archiveTo = 'capture/before_list/{{item.keySafe}}.txt' } }
  notes      = 'An empty clipboard is reported as a warning (empty_text), not a failure: a page can legitimately have no text, and browser.wait_for is the step that waits for one.'
}

function Invoke-Step {
    param($In, $Ctx)
    $archive = Resolve-EbiWorkPath -PathValue ([string]$In['archiveTo']) -WorkDir ([string]$Ctx['WorkDir'])
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would read the page text' + $(if ($archive -ne '') { ' and archive it to ' + $archive } else { ''}))); return @{ ok = $true; text = ''; length = 0; path = $archive } }
    $fg = Set-EbiForeground -HWnd (ConvertTo-EbiHandle $In['window'])
    if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message'] } }
    $text = Read-EbiPageText -SelectWaitMs ([int]$In['selectWaitMs']) -CopyWaitMs ([int]$In['copyWaitMs'])
    $warnings = @()
    if ([string]::IsNullOrWhiteSpace($text)) { $warnings = @( @{ code = 'empty_text'; message = 'the clipboard came back empty'; data = @{} } ) }
    if ($archive -ne '') {
        $w = Write-EbiTextFile -Path $archive -Text $text
        if (-not $w['ok']) { return @{ ok = $false; failure = 'archive_failed'; message = $w['message']; text = $text; length = $text.Length; path = $archive } }
    }
    return @{ ok = $true; text = $text; length = $text.Length; path = $archive; warnings = $warnings }
}
