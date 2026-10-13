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
    createIfMissing = @{ type='bool'; default=$false; desc='no candidate exists: create dir (so the operator can drop the files in) and stop with no_files' }
    requireFiles    = @{ type='bool'; default=$false; desc='no matching file in the folder is a failure (no_files, retryable) instead of an empty list' }
    glob       = @{ type='string'; default='*' }
    orderBy    = @{ type='string'; default='name'; enum=@('name', 'time') }
    countLines = @{ type='bool';   default=$false; desc='read each file and count its lines (CRLF / LF)' }
    mustExist  = @{ type='bool';   default=$true; desc='false: a missing folder lists as empty' }
    expect     = @{ type='int';    default=-1; desc='how many files should be there (the job''s FileCount); -1 = any number. A different count is a warning (count_mismatch), not a failure: the pairing step decides' }
    expectFallback = @{ type='string'; default=''; desc='count used when expect is -1 (the Teams message''s send count); blank or not a number = none' }
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
    @{ id = 'no_files';       transient = $true  }
  )
  example    = @{ use = 'file.list'; with = @{ dir = 'DATA/GIFT/{{item.JOB}}'; glob = '*.csv'; countLines = $true } }
  notes      = 'order is the sort key used (the name, or the time as yyyyMMddHHmmss), so a later pairing step can sort the two sides the same way. A job of several files is one folder with several files: expect says how many are wanted, so a missing second file is named here, before pairing.'
}

function FileList-Expect {
    # PURE. Expect when 0 or more, else a numeric Fallback, else -1.
    param([int]$Expect, [string]$Fallback)
    if ($Expect -ge 0) { return $Expect }
    $t = if ($null -eq $Fallback) { '' } else { $Fallback.Trim() }
    if ($t -match '^\d+$') { return [int]$t }
    return -1
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
    if ($dir -eq '' -and [bool]$In['createIfMissing'] -and $tried.Count -gt 0 -and -not $Ctx['DryRun']) {
        $first = [string]$tried[0]
        try { [void](New-Item -ItemType Directory -Path $first -Force -ErrorAction Stop) } catch { return @{ ok = $false; failure = 'file_not_found'; message = ('cannot create ' + $first + ': ' + $_.Exception.Message); files = @(); names = @(); paths = @(); total = 0; dir = $first } }
        return @{ ok = $false; failure = 'no_files'; message = ('no folder was there (tried: ' + ($tried -join ' ; ') + '); created ' + $first + ' -- put the ' + $glob + ' files in, then r'); files = @(); names = @(); paths = @(); total = 0; dir = $first }
    }
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
    if ($files.Count -eq 0 -and [bool]$In['requireFiles'] -and -not $Ctx['DryRun']) { return @{ ok = $false; failure = 'no_files'; message = ('no ' + $glob + ' file in ' + $dir + ' -- put the files in, then r'); files = @(); names = @(); paths = @(); total = 0; dir = $dir } }
    $out = @{ ok = $true; files = $files.ToArray(); names = @($files | ForEach-Object { $_['name'] }); paths = @($files | ForEach-Object { $_['path'] }); total = $files.Count; dir = $dir }
    $want = FileList-Expect -Expect ([int]$In['expect']) -Fallback ([string]$In['expectFallback'])
    if ($want -ge 0 -and $files.Count -ne $want) {
        $out['warnings'] = @(@{ code = 'count_mismatch'; message = ('expected ' + $want + ' ' + $glob + ' file(s) in ' + $dir + ', found ' + $files.Count + $(if ($files.Count -lt $want) { ' -- put the missing file(s) in and r on the next stop, or go on and the pairing decides' } else { '' })); data = @{ expected = $want; found = $files.Count; dir = $dir } })
    }
    return $out
}
