# modules/file/file.download.ps1
# HTTP GET one or more files into a folder (P4-19's download, without the
# browser: a URL that a plain GET can fetch needs no clicking). Windows
# credentials of the logged-on user go along (UseDefaultCredentials), and
# the system proxy is used with them, so an intranet server that accepts
# the browser accepts this too. Each file lands as <name>.download first
# and is renamed when complete: a half-written file never carries the
# real name.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')   # line counts

$Manifest = @{
  id         = 'file.download'
  group      = 'file'
  summary    = 'HTTP GET files (base URL + names, or one URL) into a folder'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    baseUrl   = @{ type='string'; default=''; desc='URL prefix the names are appended to (a trailing / is added when missing)' }
    names     = @{ type='list';   default=@(); desc='file names under baseUrl; each is saved under the same name' }
    url       = @{ type='string'; default=''; desc='one full URL instead of baseUrl + names' }
    saveAs    = @{ type='path';   default=''; desc='with url: the file to write (relative: under the work dir)' }
    dir       = @{ type='path';   default=''; desc='with names: the folder to write into (created when missing)' }
    overwrite = @{ type='bool';   default=$true; desc='false: an existing file is kept and reported as skipped' }
    timeoutSec = @{ type='int';   default=60 }
    countLines = @{ type='bool';  default=$false; desc='count each file''s lines (files[].lines), for verify.pair_files' }
    missingOk  = @{ type='bool';  default=$false; desc='a 404 is not a failure: the file is left out (total counts only what arrived)' }
  }
  outputs    = @{
    paths   = @{ type='list'; desc='every file now on disk, in the order asked' }
    files   = @{ type='list'; desc='@{ name; path; size; skipped; lines; order } per file (lines -1 unless countLines)' }
    total   = @{ type='int';  desc='files downloaded or already there' }
  }
  failures   = @(
    @{ id = 'input_invalid';   transient = $false }
    @{ id = 'download_failed'; transient = $true  }
    @{ id = 'not_found';       transient = $false }
  )
  example    = @{ use = 'file.download'; with = @{ baseUrl = '{{profile.urls.jenkinsReport}}'; names = '{{steps.jk.out.plucked}}'; dir = 'DATA/GFIX/{{item.Excel_NAME}}' } }
  notes      = 'An HTTP 404 is not_found (not transient): the file is not there. Anything else (timeout, 5xx, refused) is download_failed and may be retried. Nothing is downloaded when names is empty and url is blank -- that is input_invalid, not a silent success.'
}

function FileDownload-Get {
    # One GET to a temp name, then the rename. @{ ok; failure; message; size }.
    param([string]$Url, [string]$Dest, [int]$TimeoutSec)
    $tmp = $Dest + '.download'
    $wc = $null
    try {
        $dir = Split-Path -Path $Dest -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $req = [System.Net.WebRequest]::Create($Url)
        $req.UseDefaultCredentials = $true
        try { if ($null -ne $req.Proxy) { $req.Proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials } } catch { }
        $req.Timeout = [Math]::Max(5, $TimeoutSec) * 1000
        $resp = $req.GetResponse()
        try {
            $in = $resp.GetResponseStream()
            $out = [System.IO.File]::Create($tmp)
            try { $in.CopyTo($out) } finally { $out.Dispose(); $in.Dispose() }
        } finally { $resp.Close() }
        if (Test-Path -LiteralPath $Dest) { Remove-Item -LiteralPath $Dest -Force }
        Move-Item -LiteralPath $tmp -Destination $Dest -Force
        return @{ ok = $true; failure = ''; message = ''; size = [int64](Get-Item -LiteralPath $Dest).Length }
    } catch [System.Net.WebException] {
        $code = 0
        try { if ($null -ne $_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode } } catch { }
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        if ($code -eq 404) { return @{ ok = $false; failure = 'not_found'; message = ('404 ' + $Url); size = 0 } }
        return @{ ok = $false; failure = 'download_failed'; message = ($Url + ': ' + $_.Exception.Message); size = 0 }
    } catch {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        return @{ ok = $false; failure = 'download_failed'; message = ($Url + ': ' + $_.Exception.Message); size = 0 }
    }
}

function FileDownload-Lines {
    param([string]$Path, [bool]$Count)
    if (-not $Count) { return -1 }
    return @(Get-EbiTextLines -Text ((ConvertFrom-EbiMixedBytes -Bytes ([System.IO.File]::ReadAllBytes($Path)))['text'])).Count
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $jobs = New-Object System.Collections.ArrayList
    $url = [string]$In['url']
    if (-not [string]::IsNullOrWhiteSpace($url)) {
        $dest = Resolve-EbiWorkPath -PathValue ([string]$In['saveAs']) -WorkDir $work
        if ([string]::IsNullOrWhiteSpace($dest)) { return @{ ok = $false; failure = 'input_invalid'; message = 'url needs saveAs'; paths = @(); files = @(); total = 0 } }
        [void]$jobs.Add(@{ url = $url; dest = $dest; name = (Split-Path -Path $dest -Leaf) })
    } else {
        $names = @(@($In['names']) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        $base = [string]$In['baseUrl']
        $dir = Resolve-EbiWorkPath -PathValue ([string]$In['dir']) -WorkDir $work
        if ($names.Count -eq 0 -or [string]::IsNullOrWhiteSpace($base) -or [string]::IsNullOrWhiteSpace($dir)) {
            $why = ('give url + saveAs, or baseUrl + names + dir (names: {0}, baseUrl: "{1}", dir: "{2}")' -f $names.Count, $base, $dir)
            if ($Ctx['DryRun']) { $Ctx.Log.Info(('would download nothing: ' + $why)); return @{ ok = $true; paths = @(); files = @(); total = 0 } }
            return @{ ok = $false; failure = 'input_invalid'; message = $why; paths = @(); files = @(); total = 0 }
        }
        if (-not $base.EndsWith('/')) { $base = $base + '/' }
        foreach ($n in $names) { [void]$jobs.Add(@{ url = ($base + [string]$n); dest = (Join-Path $dir ([string]$n)); name = [string]$n }) }
    }
    if ($Ctx['DryRun']) {
        foreach ($j in $jobs) { $Ctx.Log.Info(('would GET {0} -> {1}' -f $j['url'], $j['dest'])) }
        return @{ ok = $true; paths = @($jobs | ForEach-Object { $_['dest'] }); files = @($jobs | ForEach-Object { @{ name = $_['name']; path = $_['dest']; size = 0; skipped = $false; lines = -1; order = $_['name'] } }); total = $jobs.Count }
    }
    $files = New-Object System.Collections.ArrayList
    $warn404 = $false
    foreach ($j in $jobs) {
        if (-not [bool]$In['overwrite'] -and (Test-Path -LiteralPath $j['dest'])) {
            [void]$files.Add(@{ name = $j['name']; path = $j['dest']; size = [int64](Get-Item -LiteralPath $j['dest']).Length; skipped = $true; lines = (FileDownload-Lines -Path $j['dest'] -Count ([bool]$In['countLines'])); order = $j['name'] })
            continue
        }
        $r = FileDownload-Get -Url $j['url'] -Dest $j['dest'] -TimeoutSec ([int]$In['timeoutSec'])
        if (-not $r['ok'] -and $r['failure'] -eq 'not_found' -and [bool]$In['missingOk']) {
            $Ctx.Log.Info(('not there (404, allowed): {0}' -f $j['url']))
            $warn404 = $true
            continue
        }
        if (-not $r['ok']) { return @{ ok = $false; failure = $r['failure']; message = $r['message']; paths = @($files | ForEach-Object { $_['path'] }); files = $files.ToArray(); total = $files.Count } }
        $Ctx.Log.Info(('downloaded {0} ({1} bytes)' -f $j['dest'], $r['size']))
        [void]$files.Add(@{ name = $j['name']; path = $j['dest']; size = $r['size']; skipped = $false; lines = (FileDownload-Lines -Path $j['dest'] -Count ([bool]$In['countLines'])); order = $j['name'] })
    }
    $out = @{ ok = $true; paths = @($files | ForEach-Object { $_['path'] }); files = $files.ToArray(); total = $files.Count }
    if ($warn404) { $out['warnings'] = @(@{ code = 'not_there'; message = ('' + ($jobs.Count - $files.Count) + ' file(s) answered 404 and were left out (missingOk)') }) }
    return $out
}
