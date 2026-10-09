# modules/browser/browser.download_each.ps1
# For each term: find it on the page (Ctrl+F, Esc focuses the link the
# match is in), open it (Enter), run a short key recipe that triggers a
# download, wait for the new file in the downloads folder, move it to a
# destination named after the term, and go back. The download half of
# P4-19, for pages whose download is a button behind a link -- there is
# no URL to GET. One step for the whole list because a workflow has no
# inner loop (a transfer of three files has three job logs).
#
# The recipe is data: 'find:<text>' (Ctrl+F the text, Esc), 'keys:<SendKeys>',
# 'wait:<ms>'. Before each term the page text is read: a term the page
# does not show is not_found (transient -- put the list page back and
# retry), never a blind Ctrl+F into the wrong page.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'browser.download_each'
  group      = 'browser'
  summary    = 'Per term: Ctrl+F it, open it, key a download, collect the file, go back'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('foreground')
  provides   = @()
  releases   = @()
  idempotent = $false
  inputs     = @{
    window       = @{ type='session'; sessionKind='window'; required=$true }
    terms        = @{ type='list';   required=$true; desc='exact texts of the links to open, one download each' }
    recipe       = @{ type='list';   default=@(); desc='after opening: find:<text> | keys:<SendKeys> | wait:<ms>, in order' }
    openWaitMs   = @{ type='int';    default=2500; desc='wait after Enter on the link' }
    downloadDir  = @{ type='path';   default=''; desc='where the browser saves; empty = the user''s Downloads' }
    glob         = @{ type='string'; default='*' }
    timeoutSec   = @{ type='int';    default=60 }
    destDir      = @{ type='path';   required=$true }
    nameTemplate = @{ type='string'; default='{term}{ext}'; desc='{term} {ext} {name}: the destination file name' }
    backKeys     = @{ type='string'; default='%{LEFT}'; desc='keys that return to the list; empty = stay' }
    backWaitMs   = @{ type='int';    default=2500 }
    skipExisting = @{ type='bool';   default=$true; desc='a term whose destination file already exists is not downloaded again' }
  }
  outputs    = @{
    files   = @{ type='list'; desc='@{ term; path; skipped } per term' }
    paths   = @{ type='list' }
    fetched = @{ type='int';  desc='downloads actually made' }
  }
  failures   = @(
    @{ id = 'foreground_lost'; transient = $true }
    @{ id = 'not_found';       transient = $true }
    @{ id = 'timeout';         transient = $true }
    @{ id = 'move_failed';     transient = $true }
    @{ id = 'clipboard_error'; transient = $true }
    @{ id = 'recipe_invalid';  transient = $false }
  )
  example    = @{ use = 'browser.download_each'; with = @{ window = 'listWindow'; terms = '{{steps.mine.out.plucked}}'; recipe = @('find:Download job log', 'keys:{TAB}{TAB}', 'keys:{ENTER}'); destDir = 'log/jobs' } }
  notes      = 'Not idempotent by nature (each run downloads again) -- skipExisting makes a rerun cheap, and a flow.checkpoint after it records the item done. The foreground is re-checked before every key burst.'
}

function BrowserDownloadEach-Snapshot {
    param([string]$Dir, [string]$Glob)
    $m = @{}
    foreach ($f in @(Get-ChildItem -LiteralPath $Dir -File -Filter $Glob -ErrorAction SilentlyContinue)) { $m[$f.FullName] = [string]$f.LastWriteTime.Ticks + ':' + [string]$f.Length }
    return $m
}

function BrowserDownloadEach-Already {
    # A finished file in the downloads folder whose name carries the term
    # (job numbers are unique) -- downloaded by hand after a failed try, or
    # by an earlier run that stopped before moving it. Newest first; $null.
    param([string]$Dir, [string]$Glob, [string]$Term)
    $hit = @(Get-ChildItem -LiteralPath $Dir -File -Filter $Glob -ErrorAction SilentlyContinue | Where-Object { $_.Name.IndexOf($Term, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -and $_.Name -notmatch '(?i)\.(crdownload|partial|tmp|download)$' -and $_.Length -gt 0 } | Sort-Object LastWriteTime -Descending)
    if ($hit.Count -gt 0) { return $hit[0] }
    return $null
}

function BrowserDownloadEach-StopKey {
    # q / Esc typed in the console while waiting: give up now. No console -> never.
    try { while ([Console]::KeyAvailable) { $k = [Console]::ReadKey($true); if ($k.Key -eq [ConsoleKey]::Escape -or $k.KeyChar -eq 'q' -or $k.KeyChar -eq 'Q') { return $true } } } catch { }
    return $false
}

function BrowserDownloadEach-WaitNew {
    # A file not in Before (or rewritten since) -- or any file naming the
    # term -- not a partial, whose size held for 1.2 s. $null on timeout.
    param([string]$Dir, [string]$Glob, [hashtable]$Before, [int]$TimeoutSec, [string]$Term = '')
    $deadline = (Get-Date).AddSeconds([Math]::Max(1, $TimeoutSec))
    $size = @{}; $at = @{}
    while ((Get-Date) -lt $deadline) {
        if (BrowserDownloadEach-StopKey) { return $null }
        foreach ($f in @(Get-ChildItem -LiteralPath $Dir -File -Filter $Glob -ErrorAction SilentlyContinue)) {
            if ($f.Name -match '(?i)\.(crdownload|partial|tmp|download)$') { continue }
            $sig = [string]$f.LastWriteTime.Ticks + ':' + [string]$f.Length
            $named = ($Term -ne '' -and $f.Name.IndexOf($Term, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
            if (-not $named -and $Before.Contains($f.FullName) -and $Before[$f.FullName] -eq $sig) { continue }
            if ($size.Contains($f.FullName) -and [int64]$size[$f.FullName] -eq [int64]$f.Length -and $f.Length -gt 0) {
                if (((Get-Date) - [datetime]$at[$f.FullName]).TotalMilliseconds -ge 1200) { return $f }
            } else { $size[$f.FullName] = [int64]$f.Length; $at[$f.FullName] = Get-Date }
        }
        Start-Sleep -Milliseconds 400
    }
    return $null
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $terms = @(@($In['terms']) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
    $dest = Resolve-EbiWorkPath -PathValue ([string]$In['destDir']) -WorkDir $work
    $dl = [string]$In['downloadDir']
    if ([string]::IsNullOrWhiteSpace($dl)) { $dl = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads' } else { $dl = Resolve-EbiWorkPath -PathValue $dl -WorkDir $work }
    $glob = [string]$In['glob']; if ([string]::IsNullOrWhiteSpace($glob)) { $glob = '*' }
    $recipe = @(@($In['recipe']) | ForEach-Object { [string]$_ })
    $tr = Test-EbiKeyRecipe -Recipe $recipe
    if (-not $tr['ok']) { return @{ ok = $false; failure = 'recipe_invalid'; message = $tr['message']; files = @(); paths = @(); fetched = 0 } }
    $files = New-Object System.Collections.ArrayList
    if ($Ctx['DryRun']) {
        foreach ($t in $terms) {
            $Ctx.Log.Info(('would open "{0}", run {1} recipe step(s), wait for a {2} in {3}, move it to {4}' -f $t, $recipe.Count, $glob, $dl, $dest))
            [void]$files.Add(@{ term = $t; path = (Join-Path $dest ($t + '.download')); skipped = $false })
        }
        if ($terms.Count -eq 0) { $Ctx.Log.Info('would download nothing (no terms)') }
        return @{ ok = $true; files = $files.ToArray(); paths = @($files | ForEach-Object { $_['path'] }); fetched = 0 }
    }
    if (-not (Test-Path -LiteralPath $dest)) { New-Item -ItemType Directory -Path $dest -Force | Out-Null }
    $hwnd = ConvertTo-EbiHandle $In['window']
    $fetched = 0
    $i = 0
    foreach ($t in $terms) {
        $i++
        $existing = @(Get-ChildItem -LiteralPath $dest -File -Filter ($t + '*') -ErrorAction SilentlyContinue)
        if ([bool]$In['skipExisting'] -and $existing.Count -gt 0) { $Ctx.Log.Info(('[{0}/{1}] {2}: already in {3}' -f $i, $terms.Count, $t, $dest)); [void]$files.Add(@{ term = $t; path = $existing[0].FullName; skipped = $true }); continue }
        # already in the downloads folder (by hand after a failed try): just take it
        $ready = BrowserDownloadEach-Already -Dir $dl -Glob $glob -Term $t
        if ($null -ne $ready) {
            $name = ([string]$In['nameTemplate']).Replace('{term}', $t).Replace('{ext}', $ready.Extension).Replace('{name}', $ready.Name)
            $target = Join-Path $dest $name
            try { Move-Item -LiteralPath $ready.FullName -Destination $target -Force } catch { return @{ ok = $false; failure = 'move_failed'; message = $_.Exception.Message; files = $files.ToArray(); paths = @($files | ForEach-Object { $_['path'] }); fetched = $fetched } }
            $Ctx.Log.Info(('[{0}/{1}] {2}: found {3} in the downloads folder -> {4}' -f $i, $terms.Count, $t, $ready.Name, $target))
            [void]$files.Add(@{ term = $t; path = $target; skipped = $false }); $fetched++
            continue
        }
        $Ctx.Log.Info(('[{0}/{1}] {2}: opening it on the page and starting the download' -f $i, $terms.Count, $t))
        $fg = Set-EbiForeground -HWnd $hwnd
        if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message']; files = $files.ToArray(); paths = @($files | ForEach-Object { $_['path'] }); fetched = $fetched } }
        $page = Read-EbiPageText
        if ($page.IndexOf($t, [System.StringComparison]::Ordinal) -lt 0) { return @{ ok = $false; failure = 'not_found'; message = ('"' + $t + '" is not on the page in front -- bring the list back and retry'); files = $files.ToArray(); paths = @($files | ForEach-Object { $_['path'] }); fetched = $fetched } }
        $before = BrowserDownloadEach-Snapshot -Dir $dl -Glob $glob
        try {
            Invoke-EbiFindText -Text $t
            Send-EbiKeys -Keys '{ENTER}' -WaitMs ([int]$In['openWaitMs'])
            $rr = Invoke-EbiKeyRecipe -HWnd $hwnd -Recipe $recipe
            if (-not $rr['ok']) { return @{ ok = $false; failure = $(if ($rr['failure'] -eq 'recipe_invalid') { 'recipe_invalid' } else { 'foreground_lost' }); message = $rr['message']; files = $files.ToArray(); paths = @($files | ForEach-Object { $_['path'] }); fetched = $fetched } }
        } catch { return @{ ok = $false; failure = 'clipboard_error'; message = $_.Exception.Message; files = $files.ToArray(); paths = @($files | ForEach-Object { $_['path'] }); fetched = $fetched } }
        $Ctx.Log.Info(('[{0}/{1}] {2}: waiting up to {3}s for the file in {4} (q / Esc in this window stops)' -f $i, $terms.Count, $t, [int]$In['timeoutSec'], $dl))
        $new = BrowserDownloadEach-WaitNew -Dir $dl -Glob $glob -Before $before -TimeoutSec ([int]$In['timeoutSec']) -Term $t
        if ($null -eq $new) {
            # back to the list first, so r starts from the right page
            if (-not [string]::IsNullOrEmpty([string]$In['backKeys'])) { $fg = Set-EbiForeground -HWnd $hwnd; if ($fg['ok']) { Send-EbiKeys -Keys ([string]$In['backKeys']) -WaitMs ([int]$In['backWaitMs']) } }
            return @{ ok = $false; failure = 'timeout'; message = ('no download for "' + $t + '" appeared in ' + $dl + ' -- download it by hand (any name with ' + $t + ' in ' + $dl + ' is taken), then r'); files = $files.ToArray(); paths = @($files | ForEach-Object { $_['path'] }); fetched = $fetched }
        }
        $name = ([string]$In['nameTemplate']).Replace('{term}', $t).Replace('{ext}', $new.Extension).Replace('{name}', $new.Name)
        $target = Join-Path $dest $name
        try { Move-Item -LiteralPath $new.FullName -Destination $target -Force } catch { return @{ ok = $false; failure = 'move_failed'; message = $_.Exception.Message; files = $files.ToArray(); paths = @($files | ForEach-Object { $_['path'] }); fetched = $fetched } }
        $fetched++
        [void]$files.Add(@{ term = $t; path = $target; skipped = $false })
        $Ctx.Log.Info(('downloaded "{0}" -> {1}' -f $t, $target))
        if (-not [string]::IsNullOrEmpty([string]$In['backKeys'])) {
            $fg = Set-EbiForeground -HWnd $hwnd
            if ($fg['ok']) { Send-EbiKeys -Keys ([string]$In['backKeys']) -WaitMs ([int]$In['backWaitMs']) }
        }
    }
    return @{ ok = $true; files = $files.ToArray(); paths = @($files | ForEach-Object { $_['path'] }); fetched = $fetched }
}
