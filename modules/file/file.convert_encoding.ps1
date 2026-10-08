# modules/file/file.convert_encoding.ps1
# Re-encode a text file (the operator's "open in the editor, convert to
# SJIS"). The reading side understands the mixed files the batch servers
# write -- UTF-8 with the odd SJIS pair -- through kernel/LogText.ps1, so
# nothing turns into a replacement character on the way. Characters the
# target cannot hold are counted and reported, never silently dropped.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')   # ConvertFrom-EbiMixedBytes, Get-EbiCodePage

$Manifest = @{
  id         = 'file.convert_encoding'
  group      = 'file'
  summary    = 'Re-encode a text file (e.g. mixed UTF-8 to Shift_JIS), in place or to a copy'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path    = @{ type='path';   required=$true; desc='the file (relative: under the work dir)' }
    saveAs  = @{ type='path';   default=''; desc='write here instead of replacing the file' }
    from    = @{ type='string'; default='mixed'; enum=@('mixed', 'utf8', 'cp932'); desc='mixed = UTF-8 with stray CP932 pairs (also reads plain UTF-8)' }
    to      = @{ type='string'; default='cp932'; enum=@('cp932', 'utf8', 'utf8bom') }
    newline = @{ type='string'; default='keep'; enum=@('keep', 'crlf', 'lf') }
  }
  outputs    = @{
    path         = @{ type='path' }
    lines        = @{ type='int' }
    foreignPairs = @{ type='int'; desc='CP932 pairs found inside a mixed file' }
    lost         = @{ type='int'; desc='characters the target encoding could not hold (written as ?)' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $false }
    @{ id = 'write_failed';   transient = $true  }
  )
  example    = @{ use = 'file.convert_encoding'; with = @{ path = 'log/GFIXReceive/{{run.mmdd}}.log'; to = 'cp932' } }
  notes      = 'Idempotent when saveAs is given (the source is left as it was, so a rerun converts the same bytes again). In place, a second run would read already-converted bytes; keep the downloaded original and write the converted copy beside it.'
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $src = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir $work
    $dst = if ([string]::IsNullOrWhiteSpace([string]$In['saveAs'])) { $src } else { Resolve-EbiWorkPath -PathValue ([string]$In['saveAs']) -WorkDir $work }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would re-encode {0} ({1} -> {2}) -> {3}' -f $src, $In['from'], $In['to'], $dst)); return @{ ok = $true; path = $dst; lines = 0; foreignPairs = 0; lost = 0 } }
    if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = $src; path = $dst; lines = 0; foreignPairs = 0; lost = 0 } }
    $bytes = [System.IO.File]::ReadAllBytes($src)
    $pairs = 0
    switch ([string]$In['from']) {
        'cp932' { $text = (Get-EbiCodePage -CodePage 932).GetString($bytes) }
        'utf8'  { $text = (New-Object System.Text.UTF8Encoding($false)).GetString($bytes); if ($text.Length -gt 0 -and [int]$text[0] -eq 0xFEFF) { $text = $text.Substring(1) } }
        default { $d = ConvertFrom-EbiMixedBytes -Bytes $bytes; $text = $d['text']; $pairs = [int]$d['foreignPairs'] }
    }
    switch ([string]$In['newline']) {
        'crlf' { $text = [regex]::Replace($text, "\r?\n", "`r`n") }
        'lf'   { $text = $text.Replace("`r`n", "`n") }
    }
    $enc = switch ([string]$In['to']) { 'utf8' { New-Object System.Text.UTF8Encoding($false) } 'utf8bom' { New-Object System.Text.UTF8Encoding($true) } default { Get-EbiCodePage -CodePage 932 } }
    $out = $enc.GetBytes($text)
    $lost = 0
    if ([string]$In['to'] -eq 'cp932') {
        # Round-trip: a character CP932 cannot hold comes back as '?'.
        $back = $enc.GetString($out)
        $n = [Math]::Min($back.Length, $text.Length)
        for ($i = 0; $i -lt $n; $i++) { if ($back[$i] -ne $text[$i]) { $lost++ } }
    }
    try {
        $tmp = $dst + '.enc.tmp'
        $fs = [System.IO.File]::Create($tmp)
        try {
            if ([string]$In['to'] -eq 'utf8bom') { $pre = $enc.GetPreamble(); $fs.Write($pre, 0, $pre.Length) }
            $fs.Write($out, 0, $out.Length)
        } finally { $fs.Dispose() }
        Move-Item -LiteralPath $tmp -Destination $dst -Force
    } catch { return @{ ok = $false; failure = 'write_failed'; message = $_.Exception.Message; path = $dst; lines = 0; foreignPairs = $pairs; lost = $lost } }
    $lines = @(Get-EbiTextLines -Text $text).Count
    $warn = @()
    if ($lost -gt 0) { $warn = @(@{ code = 'characters_lost'; message = ('' + $lost + ' character(s) have no ' + $In['to'] + ' form and were written as ?'); data = @{ lost = $lost } }) }
    $Ctx.Log.Info(('re-encoded {0}: {1} lines, {2} CP932 pair(s) recovered' -f $dst, $lines, $pairs))
    return @{ ok = $true; path = $dst; lines = $lines; foreignPairs = $pairs; lost = $lost; warnings = $warn }
}
