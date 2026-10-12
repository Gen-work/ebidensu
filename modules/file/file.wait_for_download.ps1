# modules/file/file.wait_for_download.ps1
# Watch a folder until a NEW file appears and its size stops changing
# (P4-13). "New" means not there when the step started (by name), or
# there but written after the step started -- a browser that overwrites
# a same-named file still counts. Browser partials (.crdownload,
# .partial, .tmp, .download) are never returned.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath

$Manifest = @{
  id         = 'file.wait_for_download'
  group      = 'file'
  summary    = 'Wait for a new file in a folder whose size has stopped changing'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    dir        = @{ type='path';   default=''; desc='folder to watch; empty = the user''s Downloads folder' }
    glob       = @{ type='string'; default='*'; desc='only files matching this' }
    timeoutSec = @{ type='int';    default=60 }
    stableMs   = @{ type='int';    default=1200; desc='size must hold this long' }
    pollMs     = @{ type='int';    default=400 }
  }
  outputs    = @{
    path = @{ type='path';   desc='the new file' }
    name = @{ type='string' }
    size = @{ type='int' }
  }
  failures   = @(
    @{ id = 'timeout';        transient = $true  }
    @{ id = 'file_not_found'; transient = $false }
  )
  example    = @{ use = 'file.wait_for_download'; with = @{ glob = '*.log'; timeoutSec = 60 } }
  notes      = 'Steps run one at a time, so this cannot start before the click that triggers the download: it also accepts a file written up to 30 s before it started. file_not_found: the folder itself is missing.'
}

function FileWaitForDownload-Snapshot {
    param([string]$Dir, [string]$Glob)
    $m = @{}
    foreach ($f in @(Get-ChildItem -LiteralPath $Dir -File -Filter $Glob -ErrorAction SilentlyContinue)) { $m[$f.Name] = $f.LastWriteTime }
    return $m
}

function FileWaitForDownload-IsPartial {
    param([string]$Name)
    return ($Name -match '(?i)\.(crdownload|partial|tmp|download)$')
}

function Invoke-Step {
    param($In, $Ctx)
    $dir = [string]$In['dir']
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads' }
    else { $dir = Resolve-EbiWorkPath -PathValue $dir -WorkDir ([string]$Ctx['WorkDir']) }
    $glob = [string]$In['glob']; if ([string]::IsNullOrWhiteSpace($glob)) { $glob = '*' }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would wait up to {0}s for a new {1} in {2}' -f $In['timeoutSec'], $glob, $dir)); return @{ ok = $true; path = (Join-Path $dir 'fixture.download'); name = 'fixture'; size = 0 } }
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) { return @{ ok = $false; failure = 'file_not_found'; message = ('no such folder: ' + $dir); path = ''; name = ''; size = 0 } }
    $since = (Get-Date).AddSeconds(-30)
    $deadline = (Get-Date).AddSeconds([Math]::Max(1, [int]$In['timeoutSec']))
    $seenSize = @{}; $seenAt = @{}
    while ((Get-Date) -lt $deadline) {
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter $glob -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)) {
            if (FileWaitForDownload-IsPartial -Name $f.Name) { continue }
            if ($f.LastWriteTime -lt $since) { continue }
            $k = $f.FullName
            if ($seenSize.Contains($k) -and [int64]$seenSize[$k] -eq [int64]$f.Length -and $f.Length -gt 0) {
                if (((Get-Date) - [datetime]$seenAt[$k]).TotalMilliseconds -ge [int]$In['stableMs']) {
                    return @{ ok = $true; path = $f.FullName; name = $f.Name; size = [int64]$f.Length }
                }
            } else { $seenSize[$k] = [int64]$f.Length; $seenAt[$k] = Get-Date }
        }
        Start-Sleep -Milliseconds ([Math]::Max(100, [int]$In['pollMs']))
    }
    return @{ ok = $false; failure = 'timeout'; message = ('no new {0} settled in {1} within {2}s' -f $glob, $dir, $In['timeoutSec']); path = ''; name = ''; size = 0 }
}
