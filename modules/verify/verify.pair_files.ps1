# modules/verify/verify.pair_files.ps1
# Pair the files of one side with the files of the other (before / after
# of one transfer): same line count, several files in arrival order
# (kernel/LogText.ps1 Select-EbiFilePairs). Never guesses between two
# same-sized files out of order -- that is code=unknown for a person.

. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')

$Manifest = @{
  id         = 'verify.pair_files'
  group      = 'verify'
  summary    = 'Pair two file lists by line count and arrival order'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    left     = @{ type='list'; required=$true; desc='file.list files (name, path, lines, order)' }
    right    = @{ type='list'; required=$true; desc='file.list files of the other side' }
    longOver = @{ type='int';  default=16; desc='a pair with more lines than this is flagged long (one screen shows this many)' }
  }
  outputs    = @{
    code      = @{ type='string'; desc='ok | unknown' }
    pairs     = @{ type='list';   desc='@{ left; right; leftName; rightName; lines; long } (paths)' }
    leftOver  = @{ type='list' }
    rightOver = @{ type='list' }
    reason    = @{ type='string' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.pair_files'; with = @{ left = '{{steps.gift.out.files}}'; right = '{{steps.gfix.out.files}}' } }
  notes      = 'A file whose lines is -1 (file.list without countLines) cannot be paired by size: run file.list with countLines = true.'
}

function Invoke-Step {
    param($In, $Ctx)
    $l = @(@($In['left']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $r = @(@($In['right']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    foreach ($f in @($l + $r)) {
        if ([int]$f['lines'] -lt 0) { return @{ ok = $false; failure = 'input_invalid'; message = ('no line count for ' + $f['name'] + ' (file.list countLines = true)'); code = 'unknown'; pairs = @(); leftOver = @(); rightOver = @(); reason = '' } }
    }
    $p = Select-EbiFilePairs -Left $l -Right $r
    $pathOf = @{}
    foreach ($f in $l) { $pathOf['L|' + [string]$f['name']] = [string]$f['path'] }
    foreach ($f in $r) { $pathOf['R|' + [string]$f['name']] = [string]$f['path'] }
    $pairs = @(foreach ($x in @($p['pairs'])) {
        @{ left = $pathOf['L|' + $x['left']]; right = $pathOf['R|' + $x['right']]; leftName = [string]$x['left']; rightName = [string]$x['right']; lines = [int]$x['lines']; long = ([int]$x['lines'] -gt [int]$In['longOver']) }
    })
    return @{ ok = $true; code = [string]$p['code']; pairs = $pairs; leftOver = @($p['leftOver']); rightOver = @($p['rightOver']); reason = [string]$p['reason'] }
}
