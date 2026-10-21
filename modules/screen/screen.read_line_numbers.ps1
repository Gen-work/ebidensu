# modules/screen/screen.read_line_numbers.ps1
# Read the line numbers printed at the left of each picture (10/21: the
# HOST screen shots on the send-data sheet show 000001 ... 000019 and
# 005666 ... 005685; the DF pictures must show the same records). Windows
# OCR (en-US) on the LEFT part of the picture sized for the recognizer
# (10/21 FJODWWBV: wide shots gave nothing), as it is, in black on white
# and as a negative; the reading with the most trusted lines wins. Only a number that STARTS an
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
    minLines = @{ type='int';    default=2; desc='fewer numbered lines than this = not a numbered listing (ok false). 2: a 1-line file''s HOST shot shows only 000001-000002 (10/21 RJODWDBC)' }
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

function ScreenReadLineNumbers-Token {
    # PURE. An OCR line -> its leading line number, or -1.
    param([string]$Text, [int]$Digits = 6)
    $min = [Math]::Max(3, $Digits - 1)
    $m = [regex]::Match(([string]$Text).Trim(), ('^([0-9OoDIl|SB]{' + $min + ',' + $Digits + '})(\s|$)'))
    if (-not $m.Success) { return -1 }
    $tok = $m.Groups[1].Value
    if ($tok -notmatch '[0-9]') { return -1 }
    return [int]($tok -replace '[OoD]', '0' -replace '[Il|]', '1' -replace 'S', '5' -replace 'B', '8')
}

function ScreenReadLineNumbers-ByRows {
    <#
      PURE. OCR line boxes @(@{ text; y }) -> runs @(@{ first; last; count;
      votes; read }) or $null when the boxes carry no positions. Inside one
      shot the lines sit one pitch apart and their numbers run one after
      another, so each number read says "the top row is n - row"; the value
      most numbers agree on wins, and a misread or a line OCR missed (a red
      box over it) no longer breaks the run (10/21 SJODWDBC: read as 2-4,
      then four pieces). Rows more than 2.5 pitches apart start a new run
      (two shots in one grouped image).
    #>
    param($Boxes, [int]$Digits = 6)
    $items = @(@($Boxes) | Where-Object { $_ -is [System.Collections.IDictionary] -and [double]$_['y'] -ge 0 } | ForEach-Object { @{ n = (ScreenReadLineNumbers-Token -Text ([string]$_['text']) -Digits $Digits); y = [double]$_['y'] } } | Where-Object { [int]$_['n'] -ge 0 } | Sort-Object { [double]$_['y'] })
    if ($items.Count -lt 2) { return $null }
    # pitch: neighbours whose numbers differ by 1..3
    $ps = New-Object System.Collections.ArrayList
    for ($i = 1; $i -lt $items.Count; $i++) {
        $dn = [int]$items[$i]['n'] - [int]$items[$i - 1]['n']; $dy = [double]$items[$i]['y'] - [double]$items[$i - 1]['y']
        if ($dn -ge 1 -and $dn -le 3 -and $dy -gt 0) { [void]$ps.Add($dy / $dn) }
    }
    if ($ps.Count -eq 0) { return $null }
    $sorted = @($ps | Sort-Object); $pitch = [double]$sorted[[int][Math]::Floor($sorted.Count / 2)]
    if ($pitch -le 0) { return $null }
    # a new run only where the numbers say so: a far row whose number does
    # not fit the run's own "top row = n - row" (10/21 QJODWDBC: OCR missed
    # rows 5..11 of one shot, and the gap had split it in two)
    $runs = New-Object System.Collections.ArrayList
    $cur = New-Object System.Collections.ArrayList; [void]$cur.Add($items[0])
    $cv = @{}; $cv[[string][int]$items[0]['n']] = 1
    for ($i = 1; $i -lt $items.Count; $i++) {
        $y0 = [double]$cur[0]['y']
        $row = [int][Math]::Round(([double]$items[$i]['y'] - $y0) / $pitch)
        $off = [int]$items[$i]['n'] - $row
        $bestOff = 0; $bn = -1; foreach ($k in $cv.Keys) { if ([int]$cv[$k] -gt $bn) { $bestOff = [int]$k; $bn = [int]$cv[$k] } }
        $fits = ([Math]::Abs($off - $bestOff) -le 2)
        $far = (([double]$items[$i]['y'] - [double]$items[$i - 1]['y']) -gt 2.5 * $pitch)
        if ($far -and -not $fits) {
            [void]$runs.Add($cur.ToArray()); $cur = New-Object System.Collections.ArrayList; $cv = @{}
            [void]$cur.Add($items[$i]); $cv[[string][int]$items[$i]['n']] = 1
            continue
        }
        [void]$cur.Add($items[$i])
        $k2 = [string]$off; if ($cv.Contains($k2)) { $cv[$k2]++ } else { $cv[$k2] = 1 }
    }
    [void]$runs.Add($cur.ToArray())
    $out = New-Object System.Collections.ArrayList
    foreach ($r in $runs) {
        $y0 = [double]$r[0]['y']; $votes = @{}; $rowMax = 0
        foreach ($it in $r) {
            $row = [int][Math]::Round(([double]$it['y'] - $y0) / $pitch); if ($row -gt $rowMax) { $rowMax = $row }
            $k = [string]([int]$it['n'] - $row); if ($votes.Contains($k)) { $votes[$k]++ } else { $votes[$k] = 1 }
        }
        $best = ''; $bn = 0; foreach ($k in $votes.Keys) { if ([int]$votes[$k] -gt $bn) { $best = $k; $bn = [int]$votes[$k] } }
        [void]$out.Add(@{ first = [int]$best; last = [int]$best + $rowMax; count = $rowMax + 1; votes = $bn; read = $r.Count })
    }
    return $out.ToArray()
}

function ScreenReadLineNumbers-Parse {
    <#
      PURE. OCR lines -> @(@{ ok; first; last; count; reason }), one per run
      of consecutive line numbers (usually one per picture; two when one
      image holds two shots). Lines may be strings, or boxes @{ text; y }
      (kernel/Ocr.ps1): with positions the rows decide (ByRows: trusted
      when at least half of the numbers read, and MinLines of them, agree);
      without, the numbers alone (Runs). No number at all gives one entry
      with ok false.
    #>
    param($Lines, [int]$Digits = 6, [int]$MinLines = 2)
    $boxes = @(@($Lines) | Where-Object { $_ -is [System.Collections.IDictionary] })
    if ($boxes.Count -gt 0) {
        $rr = ScreenReadLineNumbers-ByRows -Boxes $boxes -Digits $Digits
        if ($null -ne $rr) {
            $out = New-Object System.Collections.ArrayList
            foreach ($r in @($rr)) {
                $ok = ([int]$r['votes'] -ge $MinLines -and [int]$r['votes'] * 2 -ge [int]$r['read'])
                $why = if ($ok) { $(if ([int]$r['votes'] -lt [int]$r['count']) { ('' + $r['votes'] + ' of ' + $r['count'] + ' rows read alike, the rest by position') } else { '' }) } else { ('{0} row(s) {1:000000}-{2:000000}, only {3} of {4} numbers agree: not trusted' -f $r['count'], $r['first'], $r['last'], $r['votes'], $r['read']) }
                [void]$out.Add(@{ ok = $ok; first = [int]$r['first']; last = [int]$r['last']; count = [int]$r['count']; reason = $why })
            }
            $good = @($out | Where-Object { $_['ok'] })
            if ($good.Count -gt 0) { return $good }
            return $out.ToArray()
        }
        $Lines = @($boxes | ForEach-Object { [string]$_['text'] })
    }
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
        # the left part, sized for the recognizer, as it is / black on white / negative
        $base = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($p), [System.IO.Path]::GetFileNameWithoutExtension($p))
        $vv = New-EbiOcrVariants -Path $p -Base $base
        $tries = if ($vv['ok']) { @($vv['paths']) } else { @($p) }
        $best = $null; $bestN = -1; $ran = $false; $why = ''
        foreach ($t in $tries) {
            $r = Invoke-EbiOcrFile -Path $t -Language ([string]$In['language'])
            if (-not $r['ok']) { $why = [string]$r['message']; continue }
            $ran = $true
            # what OCR read, kept next to the picture (look at it when the numbers come out wrong)
            try { [void](Write-EbiTextFile -Path ([System.IO.Path]::ChangeExtension($t, '.txt')) -Text ((@($r['boxes']) | ForEach-Object { '{0,7:0.0}  {1}' -f [double]$_['y'], $_['text'] }) -join "`r`n")) } catch { }
            $src = if (@($r['boxes']).Count -gt 0) { @($r['boxes']) } else { @($r['lines']) }
            $xs = @(ScreenReadLineNumbers-Parse -Lines $src -Digits ([int]$In['digits']) -MinLines ([int]$In['minLines']))
            $n = [int](@($xs | Where-Object { $_['ok'] } | ForEach-Object { [int]$_['count'] } | Measure-Object -Sum).Sum)
            if ($n -gt $bestN) { $best = $xs; $bestN = $n }
        }
        if (-not $ran) {
            $available = $false
            [void]$ranges.Add(@{ path = $p; ok = $false; first = 0; last = 0; count = 0; reason = ('OCR did not run: ' + $why) })
            $Ctx.Log.Info(('{0}: OCR did not run: {1}' -f (Split-Path -Leaf $p), $why))
            continue
        }
        foreach ($x in $best) {
            $x['path'] = $p
            [void]$ranges.Add($x)
            $Ctx.Log.Info(('{0}: {1}' -f (Split-Path -Leaf $p), $(if ($x['ok']) { ('lines {0:000000}-{1:000000} ({2}){3}' -f $x['first'], $x['last'], $x['count'], $(if ($x['reason']) { ', ' + $x['reason'] } else { '' })) } else { ('no line numbers: ' + $x['reason']) })))
        }
    }
    if (-not $available) { [void]$warn.Add(@{ code = 'ocr_unavailable'; message = 'OCR did not run on this PC: no line numbers read' }) }
    return @{ ok = $true; ranges = $ranges.ToArray(); available = $available; warnings = $warn.ToArray() }
}
