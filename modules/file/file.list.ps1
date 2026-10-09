# modules/file/file.list.ps1
# The files in a folder, with what pairing and evidence need to know about
# each: size, time, and (optionally) the line count. Sorted by name or by
# time; the order is the arrival order verify.pair_files relies on.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')

$Manifest = @{
  id         = 'file.list'
  group      = 'file'
  summary    = 'List the files in a folder (size, time, optional line count) in order'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    dir        = @{ type='path';   required=$true; desc='relative: under the work dir' }
    alsoDirs   = @{ type='list';   default=@(); desc='other names the folder may have, tried after dir (the first that has matching files wins, else the first that exists)' }
    glob       = @{ type='string'; default='*' }
    orderBy    = @{ type='string'; default='name'; enum=@('name', 'time') }
    countLines = @{ type='bool';   default=$false; desc='read each file and count its lines (CRLF / LF)' }
    mustExist  = @{ type='bool';   default=$true; desc='false: a missing folder lists as empty' }
  }
  outputs    = @{
    files = @{ type='list'; desc='@{ name; path; size; modified (ISO); lines; order } in order' }
    names = @{ type='list' }
    paths = @{ type='list' }
    total = @{ type='int' }
    dir   = @{ type='path'; desc='the folder actually listed (dir or one of alsoDirs)' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $false }
  )
  example    = @{ use = 'file.list'; with = @{ dir = 'DATA/GIFT/{{item.JOB}}'; glob = '*.csv'; countLines = $true } }
  notes      = 'order is the sort key used (the name, or the time as yyyyMMddHHmmss), so a later pairing step can sort the two sides the same way.'
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $glob = [string]$In['glob']; if ([string]::IsNullOrWhiteSpace($glob)) { $glob = '*' }
    $tried = New-Object System.Collections.ArrayList
    foreach ($d in @(@([string]$In['dir']) + @($In['alsoDirs']))) { if (-not [string]::IsNullOrWhiteSpace([string]$d) -and -not ([string]$d).TrimEnd().EndsWith('/') -and -not ([string]$d).TrimEnd().EndsWith('\')) { $r = Resolve-EbiWorkPath -PathValue ([string]$d) -WorkDir $work; if (-not $tried.Contains($r)) { [void]$tried.Add($r) } } }
    $dir = ''
    foreach ($d in $tried) { if ((Test-Path -LiteralPath $d -PathType Container) -and @(Get-ChildItem -LiteralPath $d -File -Filter $glob -ErrorAction SilentlyContinue).Count -gt 0) { $dir = $d; break } }
    if ($dir -eq '') { foreach ($d in $tried) { if (Test-Path -LiteralPath $d -PathType Container) { $dir = $d; break } } }
    if ($dir -eq '') {
        $first = if ($tried.Count -gt 0) { [string]$tried[0] } else { '' }
        if ($Ctx['DryRun'] -or -not [bool]$In['mustExist']) {
            if ($Ctx['DryRun']) { $Ctx.Log.Info(('would list {0} in {1}' -f $glob, ($tried -join ' | '))) }
            return @{ ok = $true; files = @(); names = @(); paths = @(); total = 0; dir = $first }
        }
        return @{ ok = $false; failure = 'file_not_found'; message = ('no such folder (tried: ' + ($tried -join ' ; ') + ') -- create it and put the files in, then r'); files = @(); names = @(); paths = @(); total = 0; dir = $first }
    }
    $items = @(Get-ChildItem -LiteralPath $dir -File -Filter $glob -ErrorAction SilentlyContinue)
    $items = if ([string]$In['orderBy'] -eq 'time') { @($items | Sort-Object LastWriteTime, Name) } else { @($items | Sort-Object Name) }
    $files = New-Object System.Collections.ArrayList
    foreach ($f in $items) {
        $lines = -1
        if ([bool]$In['countLines']) { $lines = @(Get-EbiTextLines -Text ((ConvertFrom-EbiMixedBytes -Bytes ([System.IO.File]::ReadAllBytes($f.FullName)))['text'])).Count }
        $order = if ([string]$In['orderBy'] -eq 'time') { $f.LastWriteTime.ToString('yyyyMMddHHmmss') } else { $f.Name }
        [void]$files.Add(@{ name = $f.Name; path = $f.FullName; size = [int64]$f.Length; modified = $f.LastWriteTime.ToString('s'); lines = $lines; order = $order })
    }
    if ($dir -ne [string]$tried[0]) { $Ctx.Log.Info(('listed ' + $dir + ' (not ' + $tried[0] + ')')) }
    return @{ ok = $true; files = $files.ToArray(); names = @($files | ForEach-Object { $_['name'] }); paths = @($files | ForEach-Object { $_['path'] }); total = $files.Count; dir = $dir }
}
