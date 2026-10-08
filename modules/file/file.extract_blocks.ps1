# modules/file/file.extract_blocks.ps1
# Cut the START ... END blocks of the given keys out of a log and write
# them, in log order, to one text file; report which lines of that file
# match the marker patterns (the lines the evidence highlights). The log
# text goes to a file, not to outputs: a day's receive log is hundreds of
# KB and the ledger keeps every output.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath, Write-EbiTextFile
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')

$Manifest = @{
  id         = 'file.extract_blocks'
  group      = 'file'
  summary    = 'Cut keyed START..END blocks out of a log into a file; find marker lines'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    paths        = @{ type='list';   required=$true; desc='log files to search, in order (e.g. the plain and the unzip log); missing ones are skipped' }
    blockKeys    = @{ type='list';   required=$true; desc='block keys to keep' }
    startPattern = @{ type='string'; default='^(?<tag>\S+) START (?<key>\S+)\s*$'; desc='regex with a named group key' }
    endPattern   = @{ type='string'; default='^(?<tag>\S+) END (?<key>\S+)\s*$' }
    afterEnd     = @{ type='int';    default=1; desc='lines after END that belong to the block (the date stamp)' }
    markers      = @{ type='list';   default=@(); desc='regexes; matching lines of the result are reported' }
    encoding     = @{ type='string'; default='mixed'; enum=@('mixed', 'utf8', 'cp932') }
    saveTo       = @{ type='path';   required=$true; desc='where the extracted lines go (UTF-8, no BOM)' }
  }
  outputs    = @{
    path      = @{ type='path' }
    lineCount = @{ type='int' }
    blocks    = @{ type='list'; desc='@{ key; source; complete; lines } per block, in output order' }
    markers   = @{ type='list'; desc='@{ line (1-based in the output); pattern (0-based); text }' }
    missing   = @{ type='list'; desc='keys with no block in any log' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $true  }
    @{ id = 'not_found';      transient = $true  }
    @{ id = 'write_failed';   transient = $true  }
  )
  example    = @{ use = 'file.extract_blocks'; with = @{ paths = @('log/GFIXReceive/{{run.mmdd}}.log'); blockKeys = '{{steps.names.out.plucked}}'; markers = @('file stored'); saveTo = 'capture/gfix/{{item.keySafe}}.receive.txt' } }
  notes      = 'not_found (transient: the log may not hold the transfer yet) when NO key has a block; some keys missing -> ok with a warning per key and the missing list filled. A block without its END line is kept and warned about (block_incomplete).'
}

function FileExtractBlocks-Decode {
    param([byte[]]$Bytes, [string]$Encoding)
    switch ($Encoding) {
        'cp932' { return (Get-EbiCodePage -CodePage 932).GetString($Bytes) }
        'utf8'  { return (New-Object System.Text.UTF8Encoding($false)).GetString($Bytes) }
        default { return (ConvertFrom-EbiMixedBytes -Bytes $Bytes)['text'] }
    }
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $dest = Resolve-EbiWorkPath -PathValue ([string]$In['saveTo']) -WorkDir $work
    $keys = @(@($In['blockKeys']) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
    $srcs = @(@($In['paths']) | ForEach-Object { Resolve-EbiWorkPath -PathValue ([string]$_) -WorkDir $work })
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would cut {0} block(s) from {1} -> {2}' -f $keys.Count, ($srcs -join ', '), $dest)); return @{ ok = $true; path = $dest; lineCount = 0; blocks = @(); markers = @(); missing = @() } }
    $present = @($srcs | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })
    if ($present.Count -eq 0) { return @{ ok = $false; failure = 'file_not_found'; message = ('none of the logs exists: ' + ($srcs -join ', ')); path = $dest; lineCount = 0; blocks = @(); markers = @(); missing = $keys } }
    $found = New-Object System.Collections.ArrayList
    $warn = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($s in $present) {
        $text = FileExtractBlocks-Decode -Bytes ([System.IO.File]::ReadAllBytes($s)) -Encoding ([string]$In['encoding'])
        $r = Get-EbiLogBlocks -Lines @(Get-EbiTextLines -Text $text) -StartPattern ([string]$In['startPattern']) -EndPattern ([string]$In['endPattern']) -Keys $keys -AfterEnd ([int]$In['afterEnd'])
        foreach ($b in @($r['blocks'])) {
            if ($seen.Contains([string]$b['key'])) { [void]$warn.Add(@{ code = 'block_repeated'; message = ('key ' + $b['key'] + ' also has a block in ' + $s + ' (kept the first)') }); continue }
            $seen[[string]$b['key']] = $true
            $b['source'] = $s
            [void]$found.Add($b)
        }
    }
    $missing = @($keys | Where-Object { -not $seen.Contains($_) })
    if ($found.Count -eq 0) { return @{ ok = $false; failure = 'not_found'; message = ('no block for ' + ($keys -join ', ') + ' in ' + ($present -join ', ')); path = $dest; lineCount = 0; blocks = @(); markers = @(); missing = $missing } }
    foreach ($m in $missing) { [void]$warn.Add(@{ code = 'block_missing'; message = ('no block for key ' + $m) ; data = @{ key = $m } }) }
    $all = New-Object System.Collections.ArrayList
    $info = New-Object System.Collections.ArrayList
    foreach ($b in $found) {
        if (-not [bool]$b['complete']) { [void]$warn.Add(@{ code = 'block_incomplete'; message = ('block ' + $b['key'] + ' has no END line') }) }
        foreach ($l in @($b['lines'])) { [void]$all.Add([string]$l) }
        [void]$info.Add(@{ key = [string]$b['key']; source = [string]$b['source']; complete = [bool]$b['complete']; lines = @($b['lines']).Count })
    }
    $lines = [string[]]$all.ToArray()
    $marks = @(foreach ($h in @(Find-EbiLineHits -Lines $lines -Patterns @(@($In['markers']) | ForEach-Object { [string]$_ }))) { @{ line = ([int]$h['index'] + 1); pattern = [int]$h['pattern']; text = [string]$h['text'] } })
    $w = Write-EbiTextFile -Path $dest -Text (($lines -join "`r`n") + "`r`n")
    if (-not $w['ok']) { return @{ ok = $false; failure = 'write_failed'; message = $w['message']; path = $dest; lineCount = 0; blocks = @(); markers = @(); missing = $missing } }
    $Ctx.Log.Info(('{0} block(s), {1} line(s), {2} marker line(s) -> {3}' -f $found.Count, $lines.Count, $marks.Count, $dest))
    return @{ ok = $true; path = $dest; lineCount = $lines.Count; blocks = $info.ToArray(); markers = $marks; missing = $missing; warnings = $warn.ToArray() }
}
