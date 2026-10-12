# modules/browser/browser.wait_for.ps1
# Poll a registered window's page text until it contains a string, or time
# out. Ported from MqSnap.ps1 Wait-MqPageReady with the MQ specifics removed
# (P1-14): what to wait for is an input, the page kind is browser.assert_page's
# business. The last text read is archived whether or not it matched --
# a timeout with the page text on disk is a diagnosable timeout.
# A page that does not update by itself (a list behind a refresh button)
# takes a refreshRecipe: find:<text> / keys:<SendKeys> / wait:<ms> entries
# run before every read (kernel/Native.ps1 Invoke-EbiKeyRecipe).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'browser.wait_for'
  group      = 'browser'
  summary    = 'Poll the page text until it contains a string, or time out'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('foreground')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    window     = @{ type='session'; sessionKind='window'; required=$true; desc='the window to read' }
    contains   = @{ type='string'; default=''; desc='the text that means the page is ready' }
    containsAny = @{ type='list'; default=@(); desc='or: ready when ANY of these texts is there (e.g. every minute of a time window)' }
    deselectAt = @{ type='map';  default=@{}; desc='@{ x; y } window-relative blank point clicked after the last read, so the Ctrl+A selection does not show in a screenshot taken next' }
    timeoutSec = @{ type='int'; default=12; desc='give up after this many seconds' }
    pollMs     = @{ type='int'; default=800; desc='wait between reads' }
    archiveTo  = @{ type='path'; default=''; desc='write the last text read here (relative: under the work dir)' }
    refreshRecipe = @{ type='list'; default=@(); desc='run before every read: find:<text> | keys:<SendKeys> | wait:<ms> (e.g. a refresh button the page needs)' }
    deselectRecipe = @{ type='list'; default=@(); desc='instead of deselectAt: keys run after the last read to drop the Ctrl+A selection without a click (e.g. find:<a label at the top>)' }
    expectPage = @{ type='list'; default=@(); desc='texts the RIGHT page always shows; a read without them stops at once with wrong_page (a refresh that navigated away must not be repeated for minutes)' }
    settledAfter = @{ type='string'; default=''; desc='ISO time after which the page cannot change any more (the end of the time window): when it is already past, one read decides -- no polling' }
  }
  outputs    = @{
    text      = @{ type='string'; desc='the page text at the end (matched or not)' }
    elapsedMs = @{ type='int' }
    polls     = @{ type='int' }
    path      = @{ type='path'; desc='where the text was archived, or empty' }
  }
  failures   = @(
    @{ id = 'timeout';         transient = $true }
    @{ id = 'foreground_lost'; transient = $true }
    @{ id = 'archive_failed';  transient = $true }
    @{ id = 'recipe_invalid';  transient = $false }
    @{ id = 'input_invalid';   transient = $false }
    @{ id = 'wrong_page';      transient = $false }
  )
  example    = @{ use = 'browser.wait_for'; with = @{ window = 'mainWindow'; contains = '{{item.Correl_ID_S}}'; timeoutSec = 12; archiveTo = 'capture/before_list/{{item.keySafe}}.txt' } }
}

function BrowserWaitFor-Matches {
    # PURE. Ordinal, case-sensitive containment; an empty needle never matches.
    param([string]$Text, [string]$Needle)
    if ([string]::IsNullOrEmpty($Needle) -or [string]::IsNullOrEmpty($Text)) { return $false }
    return ($Text.IndexOf($Needle, [System.StringComparison]::Ordinal) -ge 0)
}

function BrowserWaitFor-StopKey {
    # The operator clicked the console and pressed q or Esc while the step
    # polls: stop now (Ctrl+C would kill the whole run). No console -> never.
    try {
        while ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq [ConsoleKey]::Escape -or $k.KeyChar -eq 'q' -or $k.KeyChar -eq 'Q') { return $true }
        }
    } catch { }
    return $false
}

function BrowserWaitFor-IsPast {
    # PURE-ish. Is the ISO time already past? '' or unreadable -> $false.
    param([string]$Iso)
    if ([string]::IsNullOrWhiteSpace($Iso)) { return $false }
    $t = [datetime]::MinValue
    if (-not [datetime]::TryParse($Iso, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$t)) { return $false }
    return ((Get-Date) -gt $t)
}

function Invoke-Step {
    param($In, $Ctx)
    $needle = [string]$In['contains']
    $any = @(@($In['containsAny']) | Where-Object { -not [string]::IsNullOrEmpty([string]$_) } | ForEach-Object { [string]$_ })
    if ($needle -ne '') { $any = @($needle) + $any }
    if ($any.Count -eq 0) { return @{ ok = $false; failure = 'input_invalid'; message = 'give contains or containsAny'; text = ''; elapsedMs = 0; polls = 0; path = '' } }
    $needle = $any -join ' | '
    $archive = Resolve-EbiWorkPath -PathValue ([string]$In['archiveTo']) -WorkDir ([string]$Ctx['WorkDir'])
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would poll the page text for "{0}" up to {1}s' -f $needle, [int]$In['timeoutSec'])); return @{ ok = $true; text = ''; elapsedMs = 0; polls = 0; path = $archive } }
    $recipe = @(@($In['refreshRecipe']) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($recipe.Count -gt 0) { $t = Test-EbiKeyRecipe -Recipe $recipe; if (-not $t['ok']) { return @{ ok = $false; failure = 'recipe_invalid'; message = $t['message']; text = ''; elapsedMs = 0; polls = 0; path = $archive } } }
    $hWnd = ConvertTo-EbiHandle $In['window']
    $fg = Set-EbiForeground -HWnd $hWnd
    if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message'] } }
    $deadline = (Get-Date).AddSeconds([Math]::Max(1, [int]$In['timeoutSec']))
    $settled = BrowserWaitFor-IsPast -Iso ([string]$In['settledAfter'])
    $expect = @(@($In['expectPage']) | Where-Object { -not [string]::IsNullOrEmpty([string]$_) } | ForEach-Object { [string]$_ })
    $started = Get-Date
    $text = ''; $polls = 0; $hit = $false; $empties = 0; $why = ''
    do {
        $polls++
        if ($recipe.Count -gt 0) {
            $rr = Invoke-EbiKeyRecipe -HWnd $hWnd -Recipe $recipe
            if (-not $rr['ok']) { return @{ ok = $false; failure = $rr['failure']; message = $rr['message']; text = $text; elapsedMs = [int]((Get-Date) - $started).TotalMilliseconds; polls = $polls; path = $archive } }
        } elseif ($polls -gt 1) {
            $fg = Set-EbiForeground -HWnd $hWnd -SettleMs 100
            if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message'] } }
        }
        $text = Read-EbiPageText
        foreach ($n in $any) { if (BrowserWaitFor-Matches -Text $text -Needle $n) { $hit = $true; break } }
        if ($hit) { break }
        if ($expect.Count -gt 0) {
            if ([string]::IsNullOrEmpty($text)) { $empties++ } else { $empties = 0 }
            $missing = @($expect | Where-Object { -not (BrowserWaitFor-Matches -Text $text -Needle $_) })
            if ($empties -ge 2 -or ($text -ne '' -and $missing.Count -gt 0)) { $why = 'wrong_page'; break }
        }
        if ($settled) { $why = 'settled'; break }
        if (BrowserWaitFor-StopKey) { $why = 'stopped'; break }
        Start-Sleep -Milliseconds ([Math]::Max(100, [int]$In['pollMs']))
        if (BrowserWaitFor-StopKey) { $why = 'stopped'; break }
    } while ((Get-Date) -lt $deadline)
    $elapsed = [int]((Get-Date) - $started).TotalMilliseconds
    $dr = @(@($In['deselectRecipe']) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $da = $In['deselectAt']
    if ($why -ne 'wrong_page') {
        if ($dr.Count -gt 0) { [void](Invoke-EbiKeyRecipe -HWnd $hWnd -Recipe $dr) }
        elseif ($da -is [System.Collections.IDictionary] -and $da.Contains('x') -and $da.Contains('y')) { [void](Invoke-EbiDeselect -HWnd $hWnd -X ([int]$da['x']) -Y ([int]$da['y'])) }
    }
    if ($archive -ne '') {
        $w = Write-EbiTextFile -Path $archive -Text $text
        if (-not $w['ok']) { return @{ ok = $false; failure = 'archive_failed'; message = $w['message']; text = $text; elapsedMs = $elapsed; polls = $polls; path = $archive } }
    }
    if ($why -eq 'wrong_page') { return @{ ok = $false; failure = 'wrong_page'; message = ('not the expected page (it should show: ' + ($expect -join ', ') + ') -- a click or a key went somewhere else; put the page back by hand, then r'); text = $text; elapsedMs = $elapsed; polls = $polls; path = $archive } }
    $shown = if ($any.Count -gt 3) { ('"' + $any[0] + '" .. "' + $any[$any.Count - 1] + '" (' + $any.Count + ' texts)') } else { ('"' + $needle + '"') }
    $tail = switch ($why) { 'settled' { ' (the time window is over, so one read decides)' } 'stopped' { ' (stopped by the operator)' } default { '' } }
    if (-not $hit) { return @{ ok = $false; failure = 'timeout'; message = ($shown + ' not seen after ' + $polls + ' read(s) in ' + $elapsed + ' ms' + $tail); text = $text; elapsedMs = $elapsed; polls = $polls; path = $archive } }
    return @{ ok = $true; text = $text; elapsedMs = $elapsed; polls = $polls; path = $archive }
}
