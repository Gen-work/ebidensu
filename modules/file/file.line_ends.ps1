# modules/file/file.line_ends.ps1
# How a text file's lines end (CRLF / LF / CR / mixed / none) and how wide
# its widest line is in the columns a text editor shows (10/28: a long snap
# of a data file must reach the end of the line -- screen.stitch_scroll
# takes needCols from here). Reads the bytes only; nothing is decoded for
# the trace. kernel/Stitch.ps1 Get-EbiLineEndInfo.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Stitch.ps1')

$Manifest = @{
  id         = 'file.line_ends'
  group      = 'file'
  summary    = 'Line ends (CRLF / LF / CR) of a text file and the width of its widest line in editor columns'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path     = @{ type='path';   required=$true; desc='relative: under the work dir' }
    encoding = @{ type='string'; default='sjis'; enum=@('sjis', 'utf8'); desc='sjis: a 2-byte character is 2 columns (also right for plain ASCII); utf8: wide characters 2. A UTF-8 BOM switches to utf8' }
    tabWidth = @{ type='int';    default=4; desc='the editor''s tab width' }
    expect   = @{ type='string'; default=''; enum=@('', 'CRLF', 'LF', 'CR'); desc='the line end the file must have; another one is the alert' }
  }
  outputs    = @{
    eol          = @{ type='string'; desc='CRLF | LF | CR | mixed | none' }
    crlf         = @{ type='int' }
    lf           = @{ type='int' }
    cr           = @{ type='int' }
    lines        = @{ type='int' }
    lastEol      = @{ type='bool';   desc='the last line ends with a line end' }
    longestLine  = @{ type='int';    desc='1-based' }
    longestCols  = @{ type='int';    desc='its width without the line end, columns' }
    longestChars = @{ type='int' }
    needCols     = @{ type='int';    desc='columns a capture must reach to show that line''s end mark' }
    bytes        = @{ type='int' }
    alert        = @{ type='string'; desc='mixed line ends, or not the expected one; else empty' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $true }
  )
  example    = @{ use = 'file.line_ends'; with = @{ path = '{{vars.dataFile}}'; expect = 'LF' } }
}

function FileLineEnds-Alert {
    # PURE. The one line for the gate, or '' when the line ends are as expected.
    param([string]$Eol, [string]$Expect, [int]$CrLf, [int]$Lf, [int]$Cr)
    if ($Eol -eq 'mixed') { return ('mixed line ends: CRLF {0}, LF {1}, CR {2}' -f $CrLf, $Lf, $Cr) }
    if ($Expect -ne '' -and $Eol -ne 'none' -and $Eol -ne $Expect) { return ('line ends are {0}, expected {1}' -f $Eol, $Expect) }
    return ''
}

function Invoke-Step {
    param($In, $Ctx)
    $path = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    $none = @{ eol = 'none'; crlf = 0; lf = 0; cr = 0; lines = 0; lastEol = $false; longestLine = 0; longestCols = 0; longestChars = 0; needCols = 0; bytes = 0; alert = '' }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        if ($Ctx['DryRun']) { $Ctx.Log.Info(('would read the line ends of {0} (not there yet)' -f $path)); $r = $none.Clone(); $r['ok'] = $true; return $r }
        $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'file_not_found'; $r['message'] = ('no file at ' + $path); return $r
    }
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $i = Get-EbiLineEndInfo -Bytes $bytes -Encoding ([string]$In['encoding']) -TabWidth ([int]$In['tabWidth'])
    $alert = FileLineEnds-Alert -Eol ([string]$i['eol']) -Expect ([string]$In['expect']) -CrLf ([int]$i['crlf']) -Lf ([int]$i['lf']) -Cr ([int]$i['cr'])
    $Ctx.Log.Info(('{0}: {1} lines, line ends {2}{3}, widest line {4} = {5} columns (a capture must reach column {6})' -f [System.IO.Path]::GetFileName($path), $i['lines'], $i['eol'], $(if (-not $i['lastEol'] -and [int]$i['lines'] -gt 0) { ' (the last line has none)' } else { '' }), $i['longestLine'], $i['longestCols'], $i['needCols']))
    if ($alert -ne '') { $Ctx.Log.Warn($alert) }
    return @{ ok = $true; eol = [string]$i['eol']; crlf = [int]$i['crlf']; lf = [int]$i['lf']; cr = [int]$i['cr']; lines = [int]$i['lines']; lastEol = [bool]$i['lastEol'];
              longestLine = [int]$i['longestLine']; longestCols = [int]$i['longestCols']; longestChars = [int]$i['longestChars']; needCols = [int]$i['needCols']; bytes = $bytes.Length; alert = $alert }
}
