# modules/screen/screen.read_line_numbers.ps1
# Read the line numbers printed at the left of each picture (10/21: the
# HOST screen shots on the send-data sheet show 000001 ... 000019 and
# 005666 ... 005685; the DF pictures must show the same records). Windows
# OCR (en-US), first on the picture as it is, then on a negative of it
# (the HOST shots are light text on black). Only a number that STARTS an
# OCR line counts, so the header ("Columns 00001 00080") and the data
# never do.
# Inside one shot the numbers always run one after another, so a number
# that breaks the run while its neighbours keep it is a misread and is
# corrected; a real jump starts a new run (10/21 KJODWWBV: the two HOST
# shots were one Ctrl+G group, one image). A run with more than one
# misread in five lines is not trusted.

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
    ranges    = @{ type='list'; desc='@{ path; ok; first; last; count; reason } per run of consecutive numbers (one per picture, two when one image holds two shots), in input order' }
    available = @{ type='bool'; desc='false: OCR could not run on this PC (every range ok false)' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'screen.read_line_numbers'; with = @{ paths = '{{steps.hostPics.out.paths}}' } }
  notes      = 'OCR look-alikes inside a number (O/o/D -> 0, I/l/| -> 1, S -> 5, B -> 8) are folded before parsing.'
}

function ScreenReadLineNumbers-Numbers {
    # PURE. OCR lines -> the leading line numbers, top to bottom.
    param($Lines, [int]$Digits = 6)
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
    return @($nums)
}

function ScreenReadLineNumbers-Runs {
    <#
      PURE. Leading numbers -> runs @(@{ first; last; count; fixed }) of
      consecutive lines. Inside one HOST shot the numbers ALWAYS run one
      after another, so a number that breaks the run while its neighbours
      keep it is a misread (fixed++); a run's first number that does not
      fit the run under it is too. Runs one to three apart are one listing
      with a line OCR did not see. A real jump (another shot in the same
      image, 10/21: a Ctrl+G group of two shots) starts a new run.
    #>
    param([int[]]$Nums)
    $n = @($Nums).Count
    $runs = New-Object System.Collections.ArrayList
    $i = 0
    while ($i -lt $n) {
        $start = [int]$Nums[$i]; $fixed = 0
        # a misread first line: the next two agree with each other, not with it
        if ($i + 2 -lt $n -and [int]$Nums[$i + 1] -ne $start + 1 -and [int]$Nums[$i + 2] -eq [int]$Nums[$i + 1] + 1) { $start = [int]$Nums[$i + 1] - 1; $fixed++ }
        $last = $start; $cnt = 1; $i++
        while ($i -lt $n) {
            $v = [int]$Nums[$i]
            if ($v -eq $last + 1) { $last = $v; $cnt++; $i++; continue }
            $nextKeeps = ($i + 1 -lt $n -and [int]$Nums[$i + 1] -eq $last + 2)
            $isEnd = ($i + 1 -ge $n)
            if ($nextKeeps -or ($isEnd -and [Math]::Abs($v - ($last + 1)) -lt 100)) { $last = $last + 1; $cnt++; $fixed++; $i++; continue }
            break
        }
        [void]$runs.Add(@{ first = $start; last = $last; count = $cnt; fixed = $fixed })
    }
    # one listing with a line OCR missed: merge runs 1..3 apart
    $merged = New-Object System.Collections.ArrayList
    foreach ($r in $runs) {
        if ($merged.Count -gt 0) {
            $p = $merged[$merged.Count - 1]
            $gap = [int]$r['first'] - [int]$p['last']
            if ($gap -ge 1 -and $gap -le 3) { $p['last'] = $r['last']; $p['count'] = [int]$p['count'] + [int]$r['count']; $p['fixed'] = [int]$p['fixed'] + [int]$r['fixed']; continue }
        }
        [void]$merged.Add($r)
    }
    return $merged.ToArray()
}

function ScreenReadLineNumbers-Parse {
    <#
      PURE. OCR lines -> @(@{ ok; first; last; count; reason }), one per run
      of consecutive line numbers (usually one per picture; two when one
      image holds two shots). A run shorter than MinLines is not trusted;
      no number at all gives one entry with ok false.
    #>
    param($Lines, [int]$Digits = 6, [int]$MinLines = 3)
    $nums = @(ScreenReadLineNumbers-Numbers -Lines $Lines -Digits $Digits)
    if ($nums.Count -eq 0) { return @(@{ ok = $false; first = 0; last = 0; count = 0; reason = 'no line numbers read' }) }
    $out = New-Object System.Collections.ArrayList
    foreach ($r in @(ScreenReadLineNumbers-Runs -Nums ([int[]]$nums))) {
        $ok = ([int]$r['count'] -ge $MinLines -and [int]$r['fixed'] * 5 -le [int]$r['count'])
        $why = if (-not $ok) { ('{0} line(s) {1:000000}-{2:000000}, {3} misread: not trusted' -f $r['count'], $r['first'], $r['last'], $r['fixed']) } elseif ([int]$r['fixed'] -gt 0) { ('' + $r['fixed'] + ' misread number(s) corrected') } else { '' }
        [void]$out.Add(@{ ok = $ok; first = [int]$r['first']; last = [int]$r['last']; count = [int]$r['count']; reason = $why })
    }
    # a stray short run (a number in the page header) must not hide the real ones
    $good = @($out | Where-Object { $_['ok'] })
    if ($good.Count -gt 0) { return $good }
    return $out.ToArray()
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
        $xs = @(ScreenReadLineNumbers-Parse -Lines $r['lines'] -Digits ([int]$In['digits']) -MinLines ([int]$In['minLines']))
        $goodN = @($xs | Where-Object { $_['ok'] } | ForEach-Object { [int]$_['count'] } | Measure-Object -Sum).Sum
        $neg = [System.IO.Path]::ChangeExtension($p, '.neg.png')
        $v = Invoke-EbiInvertPng -Path $p -Out $neg
        if ($v['ok']) {
            # the negative too, every time: on light-on-black shots it usually reads more
            $r2 = Invoke-EbiOcrFile -Path $neg -Language ([string]$In['language'])
            if ($r2['ok']) {
                $xs2 = @(ScreenReadLineNumbers-Parse -Lines $r2['lines'] -Digits ([int]$In['digits']) -MinLines ([int]$In['minLines']))
                $goodN2 = @($xs2 | Where-Object { $_['ok'] } | ForEach-Object { [int]$_['count'] } | Measure-Object -Sum).Sum
                if ([int]$goodN2 -gt [int]$goodN) { $xs = $xs2 }
            }
        }
        foreach ($x in $xs) {
            $x['path'] = $p
            [void]$ranges.Add($x)
            $Ctx.Log.Info(('{0}: {1}' -f (Split-Path -Leaf $p), $(if ($x['ok']) { ('lines {0:000000}-{1:000000} ({2}){3}' -f $x['first'], $x['last'], $x['count'], $(if ($x['reason']) { ', ' + $x['reason'] } else { '' })) } else { ('no line numbers: ' + $x['reason']) })))
        }
    }
    if (-not $available) { [void]$warn.Add(@{ code = 'ocr_unavailable'; message = 'OCR did not run on this PC: no line numbers read' }) }
    return @{ ok = $true; ranges = $ranges.ToArray(); available = $available; warnings = $warn.ToArray() }
}
