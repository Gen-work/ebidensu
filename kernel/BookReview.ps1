#Requires -Version 5.1
# ============================================================
#  kernel/BookReview.ps1
#
#  The pre-delivery self-review of one evidence workbook, without Excel
#  and without Python: the PowerShell port of the non-OCR checks of
#  tools/review/check_book.py (10/28: the office PC has no Python). The
#  two OCR checks (K-50 GIFT list box, K-53 DF file names) stay in the
#  Python tool. Dot-source only (no param(), ASCII source, no class).
#
#    New-EbiBookModel -Book -Names  -> @{ sheets; cells; items; views }
#    Invoke-EbiBookReview -Model -Names
#                                  -> @(@{ level = NG|WARN; where;
#                                     message; ref }) in check order
#
#  Names (the Japanese ones come from the profile, never from this file):
#    send / giftRecv / gfixRecv / compare  the four sheet names
#    font                                  the one font of the book
#    jobLog                                the GFIX job-log label
#
#  The checks, each with the review that asked for it:
#    view       first sheet active, every sheet at the top, cursor A1/A3
#    fonts      every cell with a value in the four sheets in Names.font
#    recv logs  label bold, label = the log the lines came from (K-36),
#               whole runs only: START first, one START per file, every
#               END its START and the other way round, nothing after the
#               last END + date line (K-37); mojibake
#    yellow     never on an empty row (K-38); WARN when it stops short of
#               / runs well past the text
#    labels     a correl id (B) + STEP name (G) per file set, matching
#               the send sheet; none left where no set starts (K-39/K-48)
#    job logs   the Renaming line yellow (K-45); with several files the
#               job logs in the DF sets' order (K-44)
#    shapes     a picture pasted over a tool picture (K-49); a red box on
#               no picture
# ============================================================

. (Join-Path $PSScriptRoot 'Xlsx.ps1')

function Get-EbiBookReviewPatterns {
    # The RUN line of a receive log, and what mojibake looks like
    # (U+FFFD, the CP932-read-as-UTF-8 kanji, C1 controls).
    return @{
        run      = '^(GIFT|GFIX)Receiver(Unzip)? (START|END) (\S+)'
        mojibake = '[\uFFFD\u7E3A\u7E67\u7E5D\u8B41\u873F\u9A55\u8373\u96B1]|[\u0080-\u009F]'
        correl   = '^[A-Z0-9]{6,}$'
        upload   = "Uploading '[^']*' \(([0-9,]+) bytes\) to "
        rename   = "INFO\s+Renaming '"
    }
}

function New-EbiBookModel {
    # PURE over the parts. Cells of the four named sheets; views and
    # drawing items of every sheet.
    param($Book, $Names)
    $styles = Get-EbiXlsxStyles -Book $Book
    $strings = @(Get-EbiXlsxSharedStrings -Book $Book)
    $sheets = @(Get-EbiXlsxSheets -Book $Book)
    $m = @{ sheets = $sheets; cells = @{}; items = @{}; views = @{} }
    $want = @([string]$Names['send'], [string]$Names['giftRecv'], [string]$Names['gfixRecv'], [string]$Names['compare'])
    foreach ($s in $sheets) {
        $n = [string]$s['name']
        $m['views'][$n] = Get-EbiXlsxView -Book $Book -Part ([string]$s['part'])
        $m['items'][$n] = @(Get-EbiXlsxDrawingItems -Book $Book -Part ([string]$s['part']))
        if ($want -contains $n) { $m['cells'][$n] = @(Get-EbiXlsxCells -Book $Book -Part ([string]$s['part']) -Styles $styles -Strings $strings) }
    }
    return $m
}

function Add-EbiBookFinding {
    param($Out, [string]$Level, [string]$Where, [string]$Message, [string]$Ref = '')
    [void]$Out.Add(@{ level = $Level; where = $Where; message = $Message; ref = $Ref })
}

function Get-EbiBookColumnB {
    # PURE. The column-B cells with a value, in row order: @(@{ row; text; cell }).
    param($Cells)
    return @($Cells | Where-Object { $_['col'] -eq 2 -and $_['hasValue'] } | Sort-Object { [int]$_['row'] } | ForEach-Object { @{ row = [int]$_['row']; text = [string]$_['text']; cell = $_ } })
}

function Test-EbiBookView {
    param($Model, $Out)
    $i = -1
    foreach ($s in @($Model['sheets'])) {
        $i++
        if ([string]$s['state'] -ne 'visible') { continue }
        $n = [string]$s['name']
        $v = $Model['views'][$n]
        if ($null -eq $v -or -not $v['hasView']) { continue }
        if ([bool]$v['tabSelected'] -ne ($i -eq 0)) {
            Add-EbiBookFinding $Out 'NG' $n $(if ($i -eq 0) { 'active sheet on save must be the first sheet' } else { 'this sheet is selected; save with the first sheet active' })
        }
        if (@('A1', 'A3') -notcontains [string]$v['topLeft']) { Add-EbiBookFinding $Out 'NG' $n ('scrolled to ' + [string]$v['topLeft'] + '; scroll to the top before saving') }
        if ([string]$v['active'] -ne '' -and @('A1', 'A3') -notcontains [string]$v['active']) { Add-EbiBookFinding $Out 'NG' $n ('cursor on ' + [string]$v['active'] + '; put it on A3 before saving') }
    }
}

function Test-EbiBookFonts {
    param($Model, $Names, $Out)
    foreach ($key in @('send', 'giftRecv', 'gfixRecv', 'compare')) {
        $sn = [string]$Names[$key]
        if (-not $Model['cells'].Contains($sn)) { Add-EbiBookFinding $Out 'NG' $sn 'sheet missing'; continue }
        $seen = @{}
        # the cells come in sheetData order (row, then column): the first
        # cell of each wrong font is the one named
        foreach ($c in @($Model['cells'][$sn])) {
            if (-not $c['hasValue'] -or [string]$c['font'] -eq [string]$Names['font']) { continue }
            $f = [string]$c['font']
            if ($seen.Contains($f)) { continue }
            $seen[$f] = $true
            Add-EbiBookFinding $Out 'NG' ($sn + '!' + [string]$c['ref']) ('font "' + $f + '" (should be ' + [string]$Names['font'] + ')') 'Check Sheet'
        }
    }
}

function Test-EbiBookRecvLogs {
    param($Model, $Names, $Out)
    $p = Get-EbiBookReviewPatterns
    foreach ($key in @('giftRecv', 'gfixRecv')) {
        $sn = [string]$Names[$key]
        if (-not $Model['cells'].Contains($sn)) { continue }
        $rows = @(Get-EbiBookColumnB -Cells $Model['cells'][$sn])
        $labels = @(for ($i = 0; $i -lt $rows.Count; $i++) { if (([string]$rows[$i]['text']).Trim().EndsWith('.log')) { $i } })
        for ($k = 0; $k -lt $labels.Count; $k++) {
            $li = [int]$labels[$k]
            $label = ([string]$rows[$li]['text']).Trim()
            $endI = if ($k + 1 -lt $labels.Count) { [int]$labels[$k + 1] } else { $rows.Count }
            $block = @(if ($li + 1 -lt $endI) { $rows[($li + 1)..($endI - 1)] })
            $where = $sn + '!B' + $rows[$li]['row']
            if (-not [bool]$rows[$li]['cell']['bold']) { Add-EbiBookFinding $Out 'NG' $where ('log label "' + $label + '" is not bold') 'Check Sheet 10/12' }
            $runs = @(foreach ($b in $block) { $m = [regex]::Match(([string]$b['text']).Trim(), $p['run']); if ($m.Success) { @{ row = [int]$b['row']; side = $m.Groups[1].Value; unzip = $m.Groups[2].Value; kind = $m.Groups[3].Value; file = $m.Groups[4].Value } } })
            if ($runs.Count -eq 0) { continue }
            $expect = $runs[0]['side'] + 'Receive' + $(if ($runs[0]['unzip'] -ne '') { 'Unzip' } else { '' }) + '.log'
            if ($label -ne $expect) { Add-EbiBookFinding $Out 'NG' $where ('label "' + $label + '" but the lines are ' + $runs[0]['side'] + 'Receiver' + $runs[0]['unzip'] + ' -> "' + $expect + '"') 'K-36' }
            if ($block.Count -gt 0 -and -not [regex]::IsMatch(([string]$block[0]['text']).Trim(), $p['run'])) {
                $t = [string]$block[0]['text']
                Add-EbiBookFinding $Out 'NG' ($sn + '!B' + $block[0]['row']) ('block does not start with a START line: "' + $t.Substring(0, [math]::Min(50, $t.Length)) + '"') 'K-37'
            }
            $starts = [ordered]@{}
            foreach ($r in $runs) {
                if ($r['kind'] -ne 'START') { continue }
                if ($starts.Contains($r['file'])) { Add-EbiBookFinding $Out 'NG' ($sn + '!B' + $r['row']) ('START ' + $r['file'] + ' again (B' + $starts[$r['file']] + '): two START lines with one file name') 'K-37' }
                else { $starts[$r['file']] = $r['row'] }
            }
            $ends = [ordered]@{}
            foreach ($r in $runs) { if ($r['kind'] -eq 'END') { $ends[$r['file']] = $r['row'] } }
            foreach ($f in @($ends.Keys)) { if (-not $starts.Contains($f)) { Add-EbiBookFinding $Out 'NG' ($sn + '!B' + $ends[$f]) ('END ' + $f + ' has no START line in this block') 'K-37' } }
            foreach ($f in @($starts.Keys)) { if (-not $ends.Contains($f)) { Add-EbiBookFinding $Out 'NG' ($sn + '!B' + $starts[$f]) ('START ' + $f + ' has no END line: a line of another run?') 'K-37' } }
            if ($ends.Count -gt 0) {
                $lastEnd = (@($ends.Values) | Measure-Object -Maximum).Maximum
                $tail = @($block | Where-Object { [int]$_['row'] -gt $lastEnd })
                if ($tail.Count -gt 1) {
                    $t = [string]$tail[1]['text']
                    Add-EbiBookFinding $Out 'NG' ($sn + '!B' + $tail[1]['row']) ([string]($tail.Count - 1) + ' line(s) after the last END + date line: "' + $t.Substring(0, [math]::Min(50, $t.Length)) + '"') 'K-37'
                }
            }
        }
        foreach ($r in $rows) {
            if ([regex]::IsMatch([string]$r['text'], $p['mojibake'])) {
                $t = [string]$r['text']
                Add-EbiBookFinding $Out 'NG' ($sn + '!B' + $r['row']) ('mojibake? "' + $t.Substring(0, [math]::Min(60, $t.Length)) + '"') 'Check Sheet 10/21'
            }
        }
    }
}

function Test-EbiBookYellow {
    param($Model, $Names, $Out)
    $perCol = 3.0     # log rows: ~3 characters per column at the sheets' widths
    foreach ($key in @('giftRecv', 'gfixRecv')) {
        $sn = [string]$Names[$key]
        if (-not $Model['cells'].Contains($sn)) { continue }
        $byRow = @{}
        foreach ($c in @($Model['cells'][$sn])) { $r = [int]$c['row']; if (-not $byRow.Contains($r)) { $byRow[$r] = New-Object System.Collections.ArrayList }; [void]$byRow[$r].Add($c) }
        foreach ($r in @($byRow.Keys | Sort-Object)) {
            $row = @($byRow[$r] | Sort-Object { [int]$_['col'] })
            $ys = @($row | Where-Object { $_['yellow'] })
            if ($ys.Count -eq 0) { continue }
            $texts = @($row | Where-Object { $_['hasValue'] })
            if ($texts.Count -eq 0) { Add-EbiBookFinding $Out 'NG' ($sn + '!' + [string]$ys[0]['ref']) 'yellow fill on an empty row (often under a picture)' 'K-38'; continue }
            $n = ([string]$texts[0]['text']).TrimEnd().Length
            $cols = [int]($ys | ForEach-Object { [int]$_['col'] } | Measure-Object -Maximum).Maximum - [int]($ys | ForEach-Object { [int]$_['col'] } | Measure-Object -Minimum).Minimum + 1
            if ($cols * $perCol -lt $n - 2 * $perCol) { Add-EbiBookFinding $Out 'WARN' ($sn + '!B' + $r) ('yellow stops before the text ends (' + $cols + ' cols for ' + $n + ' chars)') 'Check Sheet 10/20' }
            elseif ($cols * $perCol -gt $n + 4 * $perCol) { Add-EbiBookFinding $Out 'WARN' ($sn + '!B' + $r) ('yellow runs past the text (' + $cols + ' cols for ' + $n + ' chars)') 'Check Sheet 07/01' }
        }
    }
}

function Get-EbiBookCompareLabelRows {
    # PURE. The compare sheet's label rows (a correl id in B):
    # @(@{ row; vals; stale }). A set's label row is row 3 or the row right
    # above a picture anchored in column B; anywhere else it was left by an
    # earlier run (K-48).
    param($Model, $Names)
    $sn = [string]$Names['compare']
    if (-not $Model['cells'].Contains($sn)) { return @() }
    $p = Get-EbiBookReviewPatterns
    $picRows = @{}
    foreach ($it in @($Model['items'][$sn])) { if ($it['kind'] -eq 'pic' -and [int]$it['row'] -ge 0 -and [int]$it['col'] -eq 1) { $picRows[[int]$it['row'] + 1] = $true } }
    $byRow = @{}
    foreach ($c in @($Model['cells'][$sn])) {
        if (-not $c['hasValue'] -or [int]$c['row'] -lt 3) { continue }
        $r = [int]$c['row']
        if (-not $byRow.Contains($r)) { $byRow[$r] = @{} }
        $byRow[$r][(ConvertTo-EbiColumnName -Col ([int]$c['col']))] = ([string]$c['text']).Trim()
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($r in @($byRow.Keys | Sort-Object)) {
        $vals = $byRow[$r]
        if ($vals.Contains('B') -and [regex]::IsMatch([string]$vals['B'], $p['correl'])) {
            [void]$out.Add(@{ row = $r; vals = $vals; stale = ($picRows.Count -gt 0 -and $r -ne 3 -and -not $picRows.Contains($r + 1)) })
        }
    }
    return $out.ToArray()
}

function Test-EbiBookCompareLabels {
    param($Model, $Names, $Out)
    $send = [string]$Names['send']; $cmp = [string]$Names['compare']
    if (-not $Model['cells'].Contains($send) -or -not $Model['cells'].Contains($cmp)) { return }
    $p = Get-EbiBookReviewPatterns
    $a = @{}; $b = @{}
    foreach ($c in @($Model['cells'][$send])) {
        if (-not $c['hasValue'] -or [int]$c['row'] -lt 3) { continue }
        if ([int]$c['col'] -eq 1) { $a[[int]$c['row']] = ([string]$c['text']).Trim() }
        if ([int]$c['col'] -eq 2) { $b[[int]$c['row']] = ([string]$c['text']).Trim() }
    }
    $sendLabels = @(foreach ($r in @($a.Keys | Sort-Object)) { if ([regex]::IsMatch([string]$a[$r], $p['correl'])) { @{ correl = [string]$a[$r]; step = $(if ($b.Contains($r)) { [string]$b[$r] } else { '' }) } } })
    $cmpRows = New-Object System.Collections.ArrayList
    foreach ($l in @(Get-EbiBookCompareLabelRows -Model $Model -Names $Names)) {
        $r = [int]$l['row']
        if ($l['stale']) {
            $cells = @(foreach ($col in @('B', 'G')) { if ($l['vals'].Contains($col)) { $col + $r } }) -join '/'
            Add-EbiBookFinding $Out 'NG' ($cmp + '!B' + $r) ($cells + ': a label on no set''s row (no picture starts at row ' + ($r + 1) + ') -- left by an earlier run, delete it') 'K-48'
            continue
        }
        [void]$cmpRows.Add($l)
    }
    if ($cmpRows.Count -ne $sendLabels.Count) { Add-EbiBookFinding $Out 'NG' $cmp ([string]$cmpRows.Count + ' label row(s) but ' + $send + ' has ' + $sendLabels.Count + ' file group(s)') 'K-39' }
    for ($k = 0; $k -lt $cmpRows.Count -and $k -lt $sendLabels.Count; $k++) {
        $r = [int]$cmpRows[$k]['row']; $vals = $cmpRows[$k]['vals']
        $correl = [string]$sendLabels[$k]['correl']; $step = [string]$sendLabels[$k]['step']
        if ([string]$vals['B'] -ne $correl) { Add-EbiBookFinding $Out 'NG' ($cmp + '!B' + $r) ('correl "' + [string]$vals['B'] + '", ' + $send + ' group ' + ($k + 1) + ' says "' + $correl + '"') 'K-39' }
        if ($step -eq '') { continue }
        $g = if ($vals.Contains('G')) { [string]$vals['G'] } else { '' }
        if ($g -ne $step) {
            $elsewhere = @(foreach ($col in @($vals.Keys | Sort-Object)) { if ($col -ne 'G' -and [string]$vals[$col] -eq $step) { $col } })
            $msg = if ($elsewhere.Count -gt 0) { 'STEP name in ' + $elsewhere[0] + $r + ', should be in G' + $r } else { 'G' + $r + ' is "' + $g + '", ' + $send + ' says "' + $step + '"' }
            Add-EbiBookFinding $Out 'NG' ($cmp + '!G' + $r) $msg 'K-39'
        }
    }
}

function Get-EbiBookJobLogBlocks {
    # PURE. The GFIX job logs on the GFIX receive sheet, split on 'Start Date
    # and Time': @(@{ row; bytes; renames = @(@{ row; yellow }) }).
    param($Model, $Names)
    $sn = [string]$Names['gfixRecv']
    if (-not $Model['cells'].Contains($sn)) { return @() }
    $p = Get-EbiBookReviewPatterns
    $cells = @($Model['cells'][$sn])
    $rows = @(Get-EbiBookColumnB -Cells $cells)
    $start = -1
    for ($i = 0; $i -lt $rows.Count; $i++) { if (([string]$rows[$i]['text']).Trim() -eq [string]$Names['jobLog']) { $start = $i; break } }
    if ($start -lt 0) { return @() }
    $yellowRows = @{}
    foreach ($c in $cells) { if ($c['yellow'] -and [int]$c['col'] -ge 2 -and [int]$c['col'] -le 5) { $yellowRows[[int]$c['row']] = $true } }
    $blocks = New-Object System.Collections.ArrayList
    for ($i = $start + 1; $i -lt $rows.Count; $i++) {
        $v = [string]$rows[$i]['text']
        if ($v.Trim().EndsWith('.log')) { break }
        if ($v.Contains('Start Date and Time')) { [void]$blocks.Add(@{ row = [int]$rows[$i]['row']; bytes = -1; renames = (New-Object System.Collections.ArrayList) }) }
        if ($blocks.Count -eq 0) { continue }
        $cur = $blocks[$blocks.Count - 1]
        $m = [regex]::Match($v, $p['upload'])
        if ($m.Success -and [long]$cur['bytes'] -lt 0) { $cur['bytes'] = [long]($m.Groups[1].Value.Replace(',', '')) }
        if ([regex]::IsMatch($v, $p['rename'])) { [void]$cur['renames'].Add(@{ row = [int]$rows[$i]['row']; yellow = $yellowRows.Contains([int]$rows[$i]['row']) }) }
    }
    return $blocks.ToArray()
}

function Test-EbiBookJobLogs {
    param($Model, $Names, $Out)
    $sn = [string]$Names['gfixRecv']
    $blocks = @(Get-EbiBookJobLogBlocks -Model $Model -Names $Names)
    foreach ($b in $blocks) { foreach ($r in $b['renames']) { if (-not $r['yellow']) { Add-EbiBookFinding $Out 'NG' ($sn + '!B' + $r['row']) ([string]$Names['jobLog'] + ': the Renaming line is not yellow') 'K-45' } } }
    if ($blocks.Count -lt 2) { return }
    # the DF sets: a set whose only red box is the 'same content' one is an empty file
    $cmp = [string]$Names['compare']
    $starts = @(Get-EbiBookCompareLabelRows -Model $Model -Names $Names | Where-Object { -not $_['stale'] } | ForEach-Object { [int]$_['row'] })
    if ($starts.Count -eq 0) { return }
    $boxes = @($Model['items'][$cmp] | Where-Object { $_['kind'] -eq 'sp' -and ([string]$_['line']).ToUpperInvariant() -eq 'FF0000' -and [int]$_['row'] -ge 0 } | ForEach-Object { [int]$_['row'] + 1 })
    $empty = @()
    for ($k = 0; $k -lt $starts.Count; $k++) {
        $r0 = $starts[$k]; $r1 = if ($k + 1 -lt $starts.Count) { $starts[$k + 1] } else { [int]::MaxValue }
        if (@($boxes | Where-Object { $_ -ge $r0 -and $_ -lt $r1 }).Count -eq 1) { $empty += $k }
    }
    $sizes = @($blocks | ForEach-Object { [long]$_['bytes'] })
    if ($starts.Count -ne $blocks.Count -or @($sizes | Where-Object { $_ -lt 0 }).Count -gt 0 -or @($sizes | Sort-Object -Unique).Count -ne $sizes.Count) { return }
    if ($empty.Count -eq 0 -or $empty.Count -eq $starts.Count) { return }
    $smallest = @(0..($sizes.Count - 1) | Sort-Object { $sizes[$_] } | Select-Object -First $empty.Count | Sort-Object)
    if ((@($smallest) -join ',') -ne (@($empty | Sort-Object) -join ',')) {
        $order = @(foreach ($b in $blocks) { 'B' + $b['row'] + ' ' + ([long]$b['bytes']).ToString('N0', [System.Globalization.CultureInfo]::InvariantCulture) + ' bytes' }) -join ', '
        $want = @(for ($k = 0; $k -lt $starts.Count; $k++) { 'set ' + ($k + 1) + ' ' + $(if ($empty -contains $k) { 'empty' } else { 'data' }) }) -join ', '
        Add-EbiBookFinding $Out 'NG' ($sn + '!B' + $blocks[0]['row']) ([string]$Names['jobLog'] + ' order differs from the DF sets (' + $order + '; DF: ' + $want + ')') 'K-44'
    }
}

function Get-EbiBookOverlap {
    param($A, $B)
    return [math]::Max(0.0, [math]::Min([double]$A[2], [double]$B[2]) - [math]::Max([double]$A[0], [double]$B[0])) * [math]::Max(0.0, [math]::Min([double]$A[3], [double]$B[3]) - [math]::Max([double]$A[1], [double]$B[1]))
}

function Test-EbiBookShapes {
    param($Model, $Out)
    foreach ($s in @($Model['sheets'])) {
        $n = [string]$s['name']
        $items = @($Model['items'][$n])
        foreach ($t in $items) {
            if ($t['kind'] -ne 'pic' -or -not ([string]$t['name']).StartsWith('verifyPic_')) { continue }
            $tb = $t['box']
            $area = [math]::Max(1.0, ([double]$tb[2] - [double]$tb[0]) * ([double]$tb[3] - [double]$tb[1]))
            foreach ($h in $items) {
                if ($h['kind'] -eq 'pic' -and -not ([string]$h['name']).StartsWith('verify') -and [int]$h['z'] -gt [int]$t['z'] -and (Get-EbiBookOverlap -A $h['box'] -B $tb) -ge 0.5 * $area) {
                    $at = if ([int]$t['row'] -ge 0) { 'B' + ([int]$t['row'] + 1) } else { '?' }
                    Add-EbiBookFinding $Out 'NG' ($n + '!' + $at) ('"' + $h['name'] + '" is pasted over "' + $t['name'] + '": the old picture and its boxes (verifyMark_' + ([string]$t['name']).Substring(10) + '_*) stay underneath, measured on the old one -- delete the old picture, re-place the boxes on the new one') 'K-49'
                }
            }
        }
        $pics = @($items | Where-Object { $_['kind'] -eq 'pic' } | ForEach-Object { , $_['box'] })
        foreach ($it in $items) {
            if ($it['kind'] -ne 'sp' -or ([string]$it['line']).ToUpperInvariant() -ne 'FF0000') { continue }
            $b = $it['box']
            $on = $false
            foreach ($pb in $pics) { if ([double]$b[0] -lt [double]$pb[2] -and [double]$b[2] -gt [double]$pb[0] -and [double]$b[1] -lt [double]$pb[3] -and [double]$b[3] -gt [double]$pb[1]) { $on = $true; break } }
            if (-not $on) { Add-EbiBookFinding $Out 'NG' $n ('red box "' + $it['name'] + '" is not on any picture (anchor row ' + $(if ([int]$it['row'] -ge 0) { [int]$it['row'] + 1 } else { '?' }) + ')') 'Check Sheet 08/04' }
        }
    }
}

function Invoke-EbiBookReview {
    # PURE. Every check over the model, in check_book.py's order.
    param($Model, $Names)
    $out = New-Object System.Collections.ArrayList
    Test-EbiBookView -Model $Model -Out $out
    Test-EbiBookFonts -Model $Model -Names $Names -Out $out
    Test-EbiBookRecvLogs -Model $Model -Names $Names -Out $out
    Test-EbiBookYellow -Model $Model -Names $Names -Out $out
    Test-EbiBookCompareLabels -Model $Model -Names $Names -Out $out
    Test-EbiBookJobLogs -Model $Model -Names $Names -Out $out
    Test-EbiBookShapes -Model $Model -Out $out
    return $out.ToArray()
}
