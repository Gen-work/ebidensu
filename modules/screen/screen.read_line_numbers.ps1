# modules/screen/screen.read_line_numbers.ps1
# Read the line numbers printed at the left of each picture (10/21: the
# HOST screen shots on the send-data sheet show 000001 ... 000019 and
# 005666 ... 005685; the DF pictures must show the same records). Windows
# OCR (en-US), first on the picture as it is, then on a negative of it
# (the HOST shots are light text on black). Only a number that STARTS an
# OCR line counts, so the header ("Columns 00001 00080") and the data
# never do.
# A picture is trusted (ok) only when its numbers run one after another:
# a misread number is outvoted (the offset n - row shared by most lines),
# but the first and the last line must agree with that run.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Image.ps1')    # Invoke-EbiInvertPng
. (Join-Path $PSScriptRoot '..\..\kernel\Ocr.ps1')

$Manifest = @{
  id         = 'screen.read_line_numbers'
  group      = 'screen'
  summary    = 'OCR the line numbers at the left of each picture -> first / last / count per picture'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    paths    = @{ type='list';   required=$true; desc='PNGs (excel.export_pictures paths)' }
    digits   = @{ type='int';    default=6; desc='digits of a line number (HOST: 6); 5..digits are accepted, OCR drops a leading 0 at times' }
    minLines = @{ type='int';    default=3; desc='fewer numbered lines than this = not a numbered listing (ok false)' }
    language = @{ type='string'; default='en-US' }
  }
  outputs    = @{
    ranges    = @{ type='list'; desc='@{ path; ok; first; last; count; reason } per picture, in input order' }
    available = @{ type='bool'; desc='false: OCR could not run on this PC (every range ok false)' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'screen.read_line_numbers'; with = @{ paths = '{{steps.hostPics.out.paths}}' } }
  notes      = 'OCR look-alikes inside a number (O/o/D -> 0, I/l/| -> 1, S -> 5, B -> 8) are folded before parsing.'
}

function ScreenReadLineNumbers-Parse {
    <#
      PURE. OCR lines -> @{ ok; first; last; count; reason }. A line counts
      when it starts with a run of MinDigits..Digits characters that fold
      to digits (look-alikes folded) followed by a space or the end.
    #>
    param($Lines, [int]$Digits = 6, [int]$MinLines = 3)
    $nums = New-Object System.Collections.ArrayList
    $min = [Math]::Max(3, $Digits - 1)
    foreach ($l in @($Lines)) {
        $t = ([string]$l).Trim()
        $m = [regex]::Match($t, ('^([0-9OoDIl|SB]{' + $min + ',' + $Digits + '})(\s|$)'))
        if (-not $m.Success) { continue }
        $tok = $m.Groups[1].Value
        if ($tok -notmatch '[0-9]') { continue }   # a word such as "SOLID" is not a number
        $f = $tok -replace '[OoD]', '0' -replace '[Il|]', '1' -replace 'S', '5' -replace 'B', '8'
        [void]$nums.Add([int]$f)
    }
    $n = $nums.Count
    if ($n -lt $MinLines) { return @{ ok = $false; first = 0; last = 0; count = $n; reason = ('{0} numbered line(s) read, {1} needed' -f $n, $MinLines) } }
    # the offset number - row shared by most lines
    $votes = @{}
    for ($i = 0; $i -lt $n; $i++) { $k = [string]([int]$nums[$i] - $i); if ($votes.Contains($k)) { $votes[$k]++ } else { $votes[$k] = 1 } }
    $best = ''; $bestN = 0
    foreach ($k in $votes.Keys) { if ([int]$votes[$k] -gt $bestN) { $best = $k; $bestN = [int]$votes[$k] } }
    $off = [int]$best
    $firstOk = ([int]$nums[0] -eq $off)
    $lastOk = ([int]$nums[$n - 1] -eq ($off + $n - 1))
    if ($bestN * 10 -lt $n * 8 -or -not $firstOk -or -not $lastOk) {
        return @{ ok = $false; first = [int]$nums[0]; last = [int]$nums[$n - 1]; count = $n; reason = ('the numbers do not run one after another: ' + ((@($nums) | Select-Object -First 8) -join ' ') + $(if ($n -gt 8) { ' ...' } else { '' })) }
    }
    return @{ ok = $true; first = $off; last = ($off + $n - 1); count = $n; reason = $(if ($bestN -lt $n) { ('' + ($n - $bestN) + ' misread number(s) outvoted') } else { '' }) }
}

function Invoke-Step {
    param($In, $Ctx)
    $paths = @(@($In['paths']) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { Resolve-EbiWorkPath -PathValue ([string]$_) -WorkDir ([string]$Ctx['WorkDir']) })
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would read the line numbers of {0} picture(s)' -f $paths.Count)); return @{ ok = $true; ranges = @(); available = $false } }
    $ranges = New-Object System.Collections.ArrayList
    $available = $true
    $warn = New-Object System.Collections.ArrayList
    foreach ($p in $paths) {
        $r = Invoke-EbiOcrFile -Path $p -Language ([string]$In['language'])
        if (-not $r['ok']) {
            $available = $false
            [void]$ranges.Add(@{ path = $p; ok = $false; first = 0; last = 0; count = 0; reason = ('OCR did not run: ' + [string]$r['message']) })
            continue
        }
        $x = ScreenReadLineNumbers-Parse -Lines $r['lines'] -Digits ([int]$In['digits']) -MinLines ([int]$In['minLines'])
        if (-not $x['ok']) {
            $neg = [System.IO.Path]::ChangeExtension($p, '.neg.png')
            $v = Invoke-EbiInvertPng -Path $p -Out $neg
            if ($v['ok']) {
                $r2 = Invoke-EbiOcrFile -Path $neg -Language ([string]$In['language'])
                if ($r2['ok']) {
                    $x2 = ScreenReadLineNumbers-Parse -Lines $r2['lines'] -Digits ([int]$In['digits']) -MinLines ([int]$In['minLines'])
                    if ($x2['ok'] -or [int]$x2['count'] -gt [int]$x['count']) { $x = $x2 }
                }
            }
        }
        $x['path'] = $p
        [void]$ranges.Add($x)
        $Ctx.Log.Info(('{0}: {1}' -f (Split-Path -Leaf $p), $(if ($x['ok']) { ('lines {0:000000}-{1:000000} ({2})' -f $x['first'], $x['last'], $x['count']) } else { ('no line numbers: ' + $x['reason']) })))
    }
    if (-not $available) { [void]$warn.Add(@{ code = 'ocr_unavailable'; message = 'OCR did not run on this PC: no line numbers read' }) }
    return @{ ok = $true; ranges = $ranges.ToArray(); available = $available; warnings = $warn.ToArray() }
}
