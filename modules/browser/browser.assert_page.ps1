# modules/browser/browser.assert_page.ps1
# Classify a page text against a page fingerprint (PROFILE-SCHEMA 3.1) and
# FAIL on anything but the expected page. Ported from SnapVerify.ps1
# Get-SnapPageKind with the system-specific markers moved into the profile
# (P1-15). An unknown page is the worst failure there is -- it looks like
# success -- so it is never ok.

. (Join-Path $PSScriptRoot '..\..\kernel\Parse.ps1')   # Get-EbiPageKind

$Manifest = @{
  id         = 'browser.assert_page'
  group      = 'browser'
  summary    = 'Classify page text by fingerprint: ok, loading, empty, expired, unknown'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    text        = @{ type='string'; required=$true; desc='the page text (browser.wait_for / read_text output)' }
    fingerprint = @{ type='map'; required=$true; desc='{ ok: [...all must appear], loading: [...], empty: [...], expired: [...] } (any one appears)' }
  }
  outputs    = @{
    kind    = @{ type='string'; desc='ok | loading | empty | expired | unknown' }
    matched = @{ type='list';   desc='the fingerprint strings that were found' }
  }
  failures   = @(
    @{ id = 'page_loading'; transient = $true  }
    @{ id = 'page_empty';   transient = $false }
    @{ id = 'page_expired'; transient = $false }
    @{ id = 'page_unknown'; transient = $false }
  )
  example    = @{ use = 'browser.assert_page'; with = @{ text = '{{steps.wait.out.text}}'; fingerprint = '{{page.fingerprint}}' } }
  notes      = 'Order of judgment: blank text is loading; then expired, then empty, then loading (any listed string), then ok (ALL listed strings); anything else is unknown. expired / empty / unknown are not transient: retrying the same page changes nothing, a person must look.'
}

function BrowserAssertPage-Classify {
    # The kernel classifier (kernel/Parse.ps1 Get-EbiPageKind), kept under the
    # step prefix so the step file owns a name the tests can call.
    param([string]$Text, $Fingerprint)
    return (Get-EbiPageKind -Text $Text -Fingerprint $Fingerprint)
}

function Invoke-Step {
    param($In, $Ctx)
    $r = BrowserAssertPage-Classify -Text ([string]$In['text']) -Fingerprint $In['fingerprint']
    $kind = [string]$r['kind']
    if ($kind -eq 'ok') { return @{ ok = $true; kind = 'ok'; matched = $r['matched'] } }
    return @{ ok = $false; failure = ('page_' + $kind); message = ('the page is ' + $kind + $(if (@($r['matched']).Count) { ' (saw: ' + (@($r['matched']) -join ', ') + ')' } else { ' (none of the fingerprint strings found)' })); kind = $kind; matched = $r['matched'] }
}
