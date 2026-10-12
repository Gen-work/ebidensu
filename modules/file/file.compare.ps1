# modules/file/file.compare.ps1
# Compare pairs of text files line by line, blind to CRLF vs LF (the
# question a diff tool's "same content" answers, asked without the tool).
# Each file is decoded in its own single encoding (UTF-8 if it is valid
# UTF-8 throughout, else Shift_JIS) and the TEXT is compared: the GIFT side
# writes Shift_JIS, the GFIX side UTF-8, and DF calls them the same content.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')

$Manifest = @{
  id         = 'file.compare'
  group      = 'file'
  summary    = 'Compare pairs of text files line by line; ok when every pair is identical'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    pairs     = @{ type='list';   required=$true; desc='@{ left; right } paths (verify.pair_files output)' }
    leftField = @{ type='string'; default='left' }
    rightField = @{ type='string'; default='right' }
  }
  outputs    = @{
    code      = @{ type='string'; desc='ok (all identical) | ng (a pair differs) | unknown (no pairs, or a file missing)' }
    results   = @{ type='list';   desc='@{ left; right; identical; firstDiff; leftLines; rightLines; leftEncoding; rightEncoding } per pair' }
    identical = @{ type='int';    desc='pairs that are identical' }
    reason    = @{ type='string' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'file.compare'; with = @{ pairs = '{{steps.pair.out.pairs}}' } }
  notes      = 'A difference is a verdict (ng), not a failure: the workflow decides (human.gate). firstDiff is 1-based.'
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $lf = [string]$In['leftField']; $rf = [string]$In['rightField']
    $pairs = @(@($In['pairs']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    if ($pairs.Count -eq 0) { return @{ ok = $true; code = 'unknown'; results = @(); identical = 0; reason = 'no pairs to compare' } }
    $res = New-Object System.Collections.ArrayList
    $same = 0; $code = 'ok'; $why = New-Object System.Collections.ArrayList
    foreach ($p in $pairs) {
        $l = Resolve-EbiWorkPath -PathValue ([string]$p[$lf]) -WorkDir $work
        $r = Resolve-EbiWorkPath -PathValue ([string]$p[$rf]) -WorkDir $work
        if (-not (Test-Path -LiteralPath $l -PathType Leaf) -or -not (Test-Path -LiteralPath $r -PathType Leaf)) {
            if ($Ctx['DryRun']) { [void]$res.Add(@{ left = $l; right = $r; identical = $false; firstDiff = 0; leftLines = 0; rightLines = 0 }); continue }
            $code = 'unknown'; [void]$why.Add(('missing: ' + $(if (-not (Test-Path -LiteralPath $l)) { $l } else { $r })))
            [void]$res.Add(@{ left = $l; right = $r; identical = $false; firstDiff = 0; leftLines = 0; rightLines = 0 }); continue
        }
        # each file in its own single encoding (GIFT Shift_JIS, GFIX UTF-8):
        # the TEXT is compared, as the diff tool's "same content" does
        $da = ConvertFrom-EbiTextBytes -Bytes ([System.IO.File]::ReadAllBytes($l))
        $db = ConvertFrom-EbiTextBytes -Bytes ([System.IO.File]::ReadAllBytes($r))
        $c = Compare-EbiTextLines -Left (ConvertTo-EbiJisNeutral -Text $da['text']) -Right (ConvertTo-EbiJisNeutral -Text $db['text'])
        $encNote = if ($da['encoding'] -ne $db['encoding']) { (' [' + $da['encoding'] + ' vs ' + $db['encoding'] + ']') } else { '' }
        if (-not [bool]$c['identical']) {
            # say exactly where: the line, the column, the two characters
            $la = @(Get-EbiTextLines -Text $da['text']); $lb = @(Get-EbiTextLines -Text $db['text'])
            $k = [int]$c['firstDiff'] - 1
            $fd = Get-EbiFirstCharDiff -Left $(if ($k -lt $la.Count) { [string]$la[$k] } else { '' }) -Right $(if ($k -lt $lb.Count) { [string]$lb[$k] } else { '' })
            if ([int]$fd['col'] -gt 0) { $encNote = $encNote + (' col ' + $fd['col'] + ': ' + $(if ($fd['left'] -ne '') { $fd['left'] } else { 'end' }) + ' vs ' + $(if ($fd['right'] -ne '') { $fd['right'] } else { 'end' })) }
        }
        if ([bool]$c['identical']) { $same++; if ($encNote -ne '') { [void]$why.Add(('{0} vs {1}: same text{2}' -f (Split-Path $l -Leaf), (Split-Path $r -Leaf), $encNote)) } } else { if ($code -eq 'ok') { $code = 'ng' }; [void]$why.Add(('{0} vs {1}: first difference at line {2} ({3} vs {4} lines){5}' -f (Split-Path $l -Leaf), (Split-Path $r -Leaf), $c['firstDiff'], $c['leftLines'], $c['rightLines'], $encNote)) }
        [void]$res.Add(@{ left = $l; right = $r; identical = [bool]$c['identical']; firstDiff = [int]$c['firstDiff']; leftLines = [int]$c['leftLines']; rightLines = [int]$c['rightLines']; leftEncoding = $da['encoding']; rightEncoding = $db['encoding'] })
    }
    if ($Ctx['DryRun']) { return @{ ok = $true; code = 'ok'; results = $res.ToArray(); identical = 0; reason = 'dry run' } }
    $reason = if ($code -eq 'ok') { ('' + $same + ' pair(s) identical' + $(if ($why.Count -gt 0) { ' -- ' + ($why -join '; ') } else { '' })) } elseif ($why.Count -gt 0) { $why -join '; ' } else { ('' + $same + ' pair(s) identical') }
    return @{ ok = $true; code = $code; results = $res.ToArray(); identical = $same; reason = $reason }
}
