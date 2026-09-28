# modules/browser/browser.assert_page.ps1
# Classify a page text against a page fingerprint (PROFILE-SCHEMA 3.1) and
# FAIL on anything but the expected page. Ported from SnapVerify.ps1
# Get-SnapPageKind with the system-specific markers moved into the profile
# (P1-15). An unknown page is the worst failure there is -- it looks like
# success -- so it is never ok.

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

function BrowserAssertPage-Strings {
    # PURE. The fingerprint entry for a kind as string[] (missing -> empty).
    param($Fingerprint, [string]$Kind)
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $Fingerprint -or -not ($Fingerprint -is [System.Collections.IDictionary]) -or -not $Fingerprint.Contains($Kind) -or $null -eq $Fingerprint[$Kind]) { return $out.ToArray() }
    $v = $Fingerprint[$Kind]
    if ($v -is [string]) { [void]$out.Add($v); return $out.ToArray() }
    foreach ($s in $v) { if ($null -ne $s -and [string]$s -ne '') { [void]$out.Add([string]$s) } }
    return $out.ToArray()
}

function BrowserAssertPage-Classify {
    <#
      PURE. -> @{ kind; matched }.
        blank text                     -> loading (nothing arrived yet)
        any 'expired' string present   -> expired
        any 'empty' string present     -> empty
        any 'loading' string present   -> loading
        ALL 'ok' strings present       -> ok   (an empty ok list never matches)
        otherwise                      -> unknown
    #>
    param([string]$Text, $Fingerprint)
    $matched = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($Text)) { return @{ kind = 'loading'; matched = $matched.ToArray() } }
    foreach ($kind in @('expired', 'empty', 'loading')) {
        foreach ($s in @(BrowserAssertPage-Strings -Fingerprint $Fingerprint -Kind $kind)) {
            if ($Text.IndexOf($s, [System.StringComparison]::Ordinal) -ge 0) { [void]$matched.Add($s); return @{ kind = $kind; matched = $matched.ToArray() } }
        }
    }
    $okList = @(BrowserAssertPage-Strings -Fingerprint $Fingerprint -Kind 'ok')
    if ($okList.Count -gt 0) {
        $all = $true
        foreach ($s in $okList) { if ($Text.IndexOf($s, [System.StringComparison]::Ordinal) -ge 0) { [void]$matched.Add($s) } else { $all = $false } }
        if ($all) { return @{ kind = 'ok'; matched = $matched.ToArray() } }
    }
    return @{ kind = 'unknown'; matched = $matched.ToArray() }
}

function Invoke-Step {
    param($In, $Ctx)
    $r = BrowserAssertPage-Classify -Text ([string]$In['text']) -Fingerprint $In['fingerprint']
    $kind = [string]$r['kind']
    if ($kind -eq 'ok') { return @{ ok = $true; kind = 'ok'; matched = $r['matched'] } }
    return @{ ok = $false; failure = ('page_' + $kind); message = ('the page is ' + $kind + $(if (@($r['matched']).Count) { ' (saw: ' + (@($r['matched']) -join ', ') + ')' } else { ' (none of the fingerprint strings found)' })); kind = $kind; matched = $r['matched'] }
}
