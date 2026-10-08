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
  )
  example    = @{ use = 'browser.wait_for'; with = @{ window = 'mainWindow'; contains = '{{item.Correl_ID_S}}'; timeoutSec = 12; archiveTo = 'capture/before_list/{{item.keySafe}}.txt' } }
}

function BrowserWaitFor-Matches {
    # PURE. Ordinal, case-sensitive containment; an empty needle never matches.
    param([string]$Text, [string]$Needle)
    if ([string]::IsNullOrEmpty($Needle) -or [string]::IsNullOrEmpty($Text)) { return $false }
    return ($Text.IndexOf($Needle, [System.StringComparison]::Ordinal) -ge 0)
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
    $started = Get-Date
    $text = ''; $polls = 0; $hit = $false
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
        Start-Sleep -Milliseconds ([Math]::Max(100, [int]$In['pollMs']))
    } while ((Get-Date) -lt $deadline)
    $elapsed = [int]((Get-Date) - $started).TotalMilliseconds
    $da = $In['deselectAt']
    if ($da -is [System.Collections.IDictionary] -and $da.Contains('x') -and $da.Contains('y')) { [void](Invoke-EbiDeselect -HWnd $hWnd -X ([int]$da['x']) -Y ([int]$da['y'])) }
    if ($archive -ne '') {
        $w = Write-EbiTextFile -Path $archive -Text $text
        if (-not $w['ok']) { return @{ ok = $false; failure = 'archive_failed'; message = $w['message']; text = $text; elapsedMs = $elapsed; polls = $polls; path = $archive } }
    }
    if (-not $hit) { return @{ ok = $false; failure = 'timeout'; message = ('"' + $needle + '" not seen after ' + $polls + ' read(s) in ' + $elapsed + ' ms'); text = $text; elapsedMs = $elapsed; polls = $polls; path = $archive } }
    return @{ ok = $true; text = $text; elapsedMs = $elapsed; polls = $polls; path = $archive }
}
