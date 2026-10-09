# modules/file/file.read_text.ps1
# Read a text file into the template scope (a downloaded job log, a saved
# page text) so verify.parse_text can parse it like a page. Mixed UTF-8 /
# CP932 files are read through kernel/LogText.ps1.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')

$Manifest = @{
  id         = 'file.read_text'
  group      = 'file'
  summary    = 'Read a text file (UTF-8, CP932 or mixed) and return its text'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path     = @{ type='path';   default=''; desc='relative: under the work dir' }
    paths    = @{ type='list';   default=@(); desc='several files, read in order and joined with a line break (after path, if both)' }
    encoding = @{ type='string'; default='mixed'; enum=@('mixed', 'utf8', 'cp932') }
    maxBytes = @{ type='int';    default=2000000; desc='refuse bigger files (the text goes into the trace)' }
  }
  outputs    = @{
    text  = @{ type='string' }
    lines = @{ type='int' }
    path  = @{ type='path'; desc='the (first) file read' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $true  }
    @{ id = 'too_large';      transient = $false }
  )
  example    = @{ use = 'file.read_text'; with = @{ path = 'log/GFIX/{{item.keySafe}}.log' } }
}

function FileReadText-Decode {
    param([byte[]]$Bytes, [string]$Encoding)
    switch ($Encoding) {
        'cp932' { return (Get-EbiCodePage -CodePage 932).GetString($Bytes) }
        'utf8'  { $t = (New-Object System.Text.UTF8Encoding($false)).GetString($Bytes); if ($t.Length -gt 0 -and [int]$t[0] -eq 0xFEFF) { $t = $t.Substring(1) }; return $t }
        default { return (ConvertFrom-EbiMixedBytes -Bytes $Bytes)['text'] }
    }
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $list = New-Object System.Collections.ArrayList
    if (-not [string]::IsNullOrWhiteSpace([string]$In['path'])) { [void]$list.Add((Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir $work)) }
    foreach ($p in @($In['paths'])) { if (-not [string]::IsNullOrWhiteSpace([string]$p)) { [void]$list.Add((Resolve-EbiWorkPath -PathValue ([string]$p) -WorkDir $work)) } }
    $first = if ($list.Count -gt 0) { [string]$list[0] } else { '' }
    if ($list.Count -eq 0) {
        if ($Ctx['DryRun']) { return @{ ok = $true; text = ''; lines = 0; path = '' } }
        return @{ ok = $false; failure = 'file_not_found'; message = 'no path given'; text = ''; lines = 0; path = '' }
    }
    $parts = New-Object System.Collections.ArrayList
    $total = [int64]0
    foreach ($p in $list) {
        if (-not (Test-Path -LiteralPath $p -PathType Leaf)) {
            if ($Ctx['DryRun']) { $Ctx.Log.Info(('would read {0}' -f $p)); continue }
            return @{ ok = $false; failure = 'file_not_found'; message = $p; text = ''; lines = 0; path = $first }
        }
        $total += (Get-Item -LiteralPath $p).Length
        if ($total -gt [int64]$In['maxBytes']) { return @{ ok = $false; failure = 'too_large'; message = ('{0} bytes read so far (max {1})' -f $total, $In['maxBytes']); text = ''; lines = 0; path = $first } }
        $t = FileReadText-Decode -Bytes ([System.IO.File]::ReadAllBytes($p)) -Encoding ([string]$In['encoding'])
        [void]$parts.Add($t.TrimEnd("`r", "`n"))
    }
    $text = ($parts.ToArray() -join "`r`n")
    return @{ ok = $true; text = $text; lines = @(Get-EbiTextLines -Text $text).Count; path = $first }
}
