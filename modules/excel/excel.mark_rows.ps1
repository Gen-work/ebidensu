# modules/excel/excel.mark_rows.ps1
# Write values into the rows of a shared sheet that belong to given keys
# (10/20: the WBS actual-start / actual-end columns, AJ / AK, were typed by
# hand after each delivery and sometimes forgotten or typed on the wrong
# row -- K/LJODJDBC got 10/16 while they still wait). A row is the one
# whose key column holds the item's key AND whose other columns equal the
# where values (one job has a send row and a receive row, and one row per
# task). Only blank cells are written unless onlyBlank is false, so a date
# someone typed by hand is never overwritten. Nothing is saved here:
# the workflow shows the rows and saves with excel.close after a gate.
# 10/20 (second day): a where value may be a LIST (any of: the after-fix
# AND the summary task rows), whereMatch checks a column against a regex (the
# owner column: "owner" / "owner<-other"), numberColumns are written as
# numbers (AQ hours 0.1), and a date may be typed as M/d (this year).

. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')         # ConvertTo-EbiHalfWidth
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')     # ConvertTo-EbiColumnNumber
. (Join-Path $PSScriptRoot '..\..\kernel\Worklist.ps1')    # ConvertTo-EbiWorklistDay

$Manifest = @{
  id         = 'excel.mark_rows'
  group      = 'excel'
  summary    = 'Write per-key values (dates) into the matching rows of a sheet, blank cells only'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook    = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet       = @{ type='string'; required=$true }
    firstRow    = @{ type='int';    default=2 }
    keyColumn   = @{ type='string'; required=$true; desc='column letter holding the key (the job name)' }
    items       = @{ type='list';   required=$true; desc='records, one per key (table.select rows)' }
    keyField    = @{ type='string'; required=$true; desc='item field matched against keyColumn' }
    where       = @{ type='map';    default=@{}; desc='column letter -> value the row must also hold (a list: any of them)' }
    whereMatch  = @{ type='map';    default=@{}; desc='column letter -> .NET regex the cell text must match (case-insensitive), e.g. the owner pattern' }
    setFields   = @{ type='map';    default=@{}; desc='column letter -> item field written there' }
    setValues   = @{ type='map';    default=@{}; desc='column letter -> one value written for every item' }
    dateColumns = @{ type='list';   default=@(); desc='letters written as real dates (yyyy-MM-dd / yyyy/M/d / M/d (this year) in, an Excel date out; the cell''s format is kept)' }
    numberColumns = @{ type='list'; default=@(); desc='letters written as numbers (0.1 in, a number out, not the text "0.1")' }
    onlyBlank   = @{ type='bool';   default=$true; desc='false: overwrite cells that already hold a value' }
    highlight   = @{ type='string'; default=''; desc='"#RRGGBB" (or several, comma-separated, first one the sheet does not already use): fill every written cell with it so a person can check them before saving. The workflow then closes WITHOUT saving and writes again un-filled to save (10/22: the fill never reaches the shared file)' }
  }
  outputs    = @{
    written      = @{ type='list'; desc='"KEY row 123: AJ=2026-10-16 AK=2026-10-20" per row written' }
    writtenCount = @{ type='int' }
    kept         = @{ type='list'; desc='"KEY row 123: AJ already 2026-10-16" for cells left as they were' }
    notFound     = @{ type='list'; desc='keys with no matching row' }
    cells        = @{ type='list'; desc='addresses written, e.g. AJ3854' }
    highlight    = @{ type='string'; desc='the fill colour used ("" = none)' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'input_invalid';   transient = $false }
    @{ id = 'read_failed';     transient = $true  }
    @{ id = 'write_failed';    transient = $true  }
  )
  example    = @{ use = 'excel.mark_rows'; with = @{ workbook = 'wbs'; sheet = 'WBS'; firstRow = 7; keyColumn = 'A'; items = '{{steps.pick.out.rows}}'; keyField = 'JOB'; where = @{ B = 'recv'; G = @('after', 'summary') }; whereMatch = @{ P = '^owner' }; setValues = @{ AJ = '{{run.date}}'; AK = '{{run.date}}'; AQ = '0.1' }; dateColumns = @('AJ', 'AK'); numberColumns = @('AQ') } }
  notes      = 'A key with several matching rows (two files listed twice) gets every one written. A cell holding a formula is never written (its value may look blank), reported in kept. A date written into a General cell gets the yyyy/m/d format (what typing a date does). A value that should be a date but does not read as one is written as text and reported in kept. Under DryRun nothing is read or written; the plan is logged.'
}

function ExcelMarkRows-Norm {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [double] -or $Value -is [int]) { return ([string]$Value).Trim() }
    return (ConvertTo-EbiHalfWidth -Value ([string]$Value)).Trim().ToUpperInvariant()
}

function ExcelMarkRows-CellText {
    # A Value2 cell as text for reporting ('' for empty); a serial in a date column as yyyy-MM-dd.
    param($Value, [bool]$IsDate)
    if ($null -eq $Value) { return '' }
    if ($IsDate -and ($Value -is [double] -or $Value -is [int])) { try { return ([datetime]::FromOADate([double]$Value)).ToString('yyyy-MM-dd') } catch { } }
    return ([string]$Value).Trim()
}

function ExcelMarkRows-Day {
    # PURE. A date answer -> yyyy-MM-dd ('' when it is not one). M/d takes the year of Today.
    param([string]$Text, [datetime]$Today = (Get-Date))
    $d = ConvertTo-EbiWorklistDay -Text $Text
    if ($d -ne '') { return $d }
    $t = (ConvertTo-EbiHalfWidth -Value ([string]$Text)).Trim()
    if ($t -match '^(\d{1,2})[/-](\d{1,2})$') {
        try { return ([datetime]::new($Today.Year, [int]$Matches[1], [int]$Matches[2])).ToString('yyyy-MM-dd') } catch { return '' }
    }
    return ''
}

function ExcelMarkRows-WhereOk {
    # PURE. Does row i hold every where value (a list = any of) and match every whereMatch regex?
    param($Grid, [int]$Index, $Where, $WhereMatch)
    foreach ($w in @($Where.Keys)) {
        $cell = ExcelMarkRows-Norm $Grid[([string]$w).ToUpperInvariant()][$Index]
        $want = $Where[$w]
        $alts = if ($want -is [System.Collections.IList] -and -not ($want -is [string])) { @($want) } else { @($want) }
        $any = $false
        foreach ($a in $alts) { if ($cell -eq (ExcelMarkRows-Norm $a)) { $any = $true; break } }
        if (-not $any) { return $false }
    }
    foreach ($w in @($WhereMatch.Keys)) {
        $raw = $Grid[([string]$w).ToUpperInvariant()][$Index]
        $text = if ($null -eq $raw) { '' } else { (ConvertTo-EbiHalfWidth -Value ([string]$raw)).Trim() }
        $pat = [string]$WhereMatch[$w]
        if ($pat -eq '') { continue }
        if (-not [regex]::IsMatch($text, $pat, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) { return $false }
    }
    return $true
}

function ExcelMarkRows-Plan {
    <#
      PURE. Which cells to write.
        Grid    letter -> object[] of cell values, index 0 = FirstRow
                (the key column, every where column, every set column)
      Returns @{ writes (@{ row; letter; value; key }); written; kept;
      notFound } -- written / kept are the report lines.
    #>
    param($Grid, [int]$FirstRow, [string]$KeyColumn, $Items, [string]$KeyField, $Where, $SetFields, $SetValues, $DateColumns = @(), [bool]$OnlyBlank = $true, $WhereMatch = @{}, $NumberColumns = @(), [datetime]$Today = (Get-Date), $Formulas = @{})
    $keyCol = $KeyColumn.ToUpperInvariant()
    $dates = @(@($DateColumns) | ForEach-Object { ([string]$_).ToUpperInvariant() })
    $numbers = @(@($NumberColumns) | ForEach-Object { ([string]$_).ToUpperInvariant() })
    if ($null -eq $WhereMatch) { $WhereMatch = @{} }
    $keys = $Grid[$keyCol]
    $n = @($keys).Count
    $writes = New-Object System.Collections.ArrayList
    $written = New-Object System.Collections.ArrayList
    $kept = New-Object System.Collections.ArrayList
    $notFound = New-Object System.Collections.ArrayList
    $setLetters = New-Object System.Collections.ArrayList
    foreach ($l in @($SetFields.Keys) + @($SetValues.Keys)) { $u = ([string]$l).ToUpperInvariant(); if (-not $setLetters.Contains($u)) { [void]$setLetters.Add($u) } }
    foreach ($it in @($Items)) {
        if (-not ($it -is [System.Collections.IDictionary])) { continue }
        $key = if ($it.Contains($KeyField) -and $null -ne $it[$KeyField]) { ([string]$it[$KeyField]).Trim() } else { '' }
        if ($key -eq '') { continue }
        $nk = ExcelMarkRows-Norm $key
        $hit = $false
        for ($i = 0; $i -lt $n; $i++) {
            if ((ExcelMarkRows-Norm $keys[$i]) -ne $nk) { continue }
            if (-not (ExcelMarkRows-WhereOk -Grid $Grid -Index $i -Where $Where -WhereMatch $WhereMatch)) { continue }
            $hit = $true
            $row = $FirstRow + $i
            $did = New-Object System.Collections.ArrayList
            foreach ($l in $setLetters) {
                $isDate = ($dates -contains $l)
                $want = ''
                $sf = $null; foreach ($k in @($SetFields.Keys)) { if (([string]$k).ToUpperInvariant() -eq $l) { $sf = [string]$SetFields[$k] } }
                if ($null -ne $sf) { $want = if ($it.Contains($sf) -and $null -ne $it[$sf]) { ([string]$it[$sf]).Trim() } else { '' } }
                foreach ($k in @($SetValues.Keys)) { if (([string]$k).ToUpperInvariant() -eq $l) { $want = ([string]$SetValues[$k]).Trim() } }
                if ($want -eq '') { continue }
                if ($isDate) { $d = ExcelMarkRows-Day -Text $want -Today $Today; if ($d -ne '') { $want = $d } }
                $isNum = ($numbers -contains $l)
                # a formula cell (10/22: the WBS lines "broken") is never written,
                # even when it shows blank -- the value would replace the formula
                $fm = if ($null -ne $Formulas -and $Formulas.Contains($l)) { [string]$Formulas[$l][$i] } else { '' }
                if ($fm.StartsWith('=')) { [void]$kept.Add(($key + ' row ' + $row + ': ' + $l + ' has a formula (' + $fm + '), not written')); continue }
                $cur = ExcelMarkRows-CellText -Value $Grid[$l][$i] -IsDate $isDate
                if ($cur -ne '' -and $OnlyBlank) { [void]$kept.Add(($key + ' row ' + $row + ': ' + $l + ' already ' + $cur)); continue }
                [void]$writes.Add(@{ row = $row; letter = $l; value = $want; key = $key; date = $isDate; number = $isNum })
                [void]$did.Add(($l + '=' + $want))
            }
            if ($did.Count -gt 0) { [void]$written.Add(($key + ' row ' + $row + ': ' + ($did.ToArray() -join ' '))) }
        }
        if (-not $hit) { [void]$notFound.Add($key) }
    }
    return @{ writes = $writes.ToArray(); written = $written.ToArray(); kept = $kept.ToArray(); notFound = $notFound.ToArray() }
}

function ExcelMarkRows-ColorOf {
    # PURE. "#RRGGBB" -> Excel's Color (R + G*256 + B*65536), or -1.
    param([string]$Hex)
    $t = ([string]$Hex).Trim().TrimStart('#')
    if ($t -notmatch '^[0-9A-Fa-f]{6}$') { return -1 }
    $r = [Convert]::ToInt32($t.Substring(0, 2), 16); $g = [Convert]::ToInt32($t.Substring(2, 2), 16); $b = [Convert]::ToInt32($t.Substring(4, 2), 16)
    return ($r + $g * 256 + $b * 65536)
}

function ExcelMarkRows-PickColor {
    <#
      PURE. The first candidate ("#RRGGBB", comma-separated) whose colour
      differs from every colour in Used (Excel Color longs) by more than 24
      on some channel -- so the check fill never looks like the sheet's own
      colours (10/22: "different from any color on the sheet"). None fits:
      the first one. -> @{ hex; color }
    #>
    param([string]$Candidates, $Used = @())
    $list = @(([string]$Candidates -split ',') | ForEach-Object { $_.Trim() } | Where-Object { (ExcelMarkRows-ColorOf $_) -ge 0 })
    if ($list.Count -eq 0) { return @{ hex = ''; color = -1 } }
    foreach ($h in $list) {
        $c = ExcelMarkRows-ColorOf $h
        $clash = $false
        foreach ($u in @($Used)) {
            $u = [int64]$u; if ($u -lt 0) { continue }
            $dr = [Math]::Abs(($c -band 255) - ($u -band 255)); $dg = [Math]::Abs((($c -shr 8) -band 255) - (($u -shr 8) -band 255)); $db = [Math]::Abs((($c -shr 16) -band 255) - (($u -shr 16) -band 255))
            if ([Math]::Max($dr, [Math]::Max($dg, $db)) -le 24) { $clash = $true; break }
        }
        if (-not $clash) { return @{ hex = $h; color = $c } }
    }
    return @{ hex = $list[0]; color = (ExcelMarkRows-ColorOf $list[0]) }
}

function ExcelMarkRows-NoteHidden {
    # PURE. Report lines ("KEY row 123: ...") -> the same, with
    # " (hidden by a filter)" on the rows in HiddenRows.
    param($Lines, $HiddenRows = @())
    $set = @{}; foreach ($r in @($HiddenRows)) { $set[[int]$r] = $true }
    return @(foreach ($l in @($Lines)) { $m = [regex]::Match([string]$l, ' row (\d+):'); if ($m.Success -and $set.Contains([int]$m.Groups[1].Value)) { [string]$l + ' (hidden by a filter)' } else { [string]$l } })
}

function Invoke-Step {
    param($In, $Ctx)
    $empty = @{ written = @(); writtenCount = 0; kept = @(); notFound = @(); cells = @(); highlight = '' }
    $letters = New-Object System.Collections.ArrayList
    $wm = if ($In['whereMatch'] -is [System.Collections.IDictionary]) { $In['whereMatch'] } else { @{} }
    foreach ($l in @([string]$In['keyColumn']) + @($In['where'].Keys) + @($wm.Keys) + @($In['setFields'].Keys) + @($In['setValues'].Keys)) {
        $u = ([string]$l).ToUpperInvariant()
        if ((ConvertTo-EbiColumnNumber $u) -le 0) { $r = $empty.Clone(); $r['ok'] = $false; $r['failure'] = 'input_invalid'; $r['message'] = ('"' + $l + '" is not a column letter'); return $r }
        if (-not $letters.Contains($u)) { [void]$letters.Add($u) }
    }
    if (@($In['setFields'].Keys).Count + @($In['setValues'].Keys).Count -eq 0) { $r = $empty.Clone(); $r['ok'] = $false; $r['failure'] = 'input_invalid'; $r['message'] = 'nothing to write: give setFields or setValues'; return $r }
    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would mark {0} key(s) on {1}: {2}' -f @($In['items']).Count, $In['sheet'], (@(@($In['setFields'].Keys) + @($In['setValues'].Keys)) -join ',')))
        $r = $empty.Clone(); $r['ok'] = $true; return $r
    }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { $r = $empty.Clone(); $r['ok'] = $false; $r['failure'] = 'sheet_not_found'; $r['message'] = ('no sheet "' + $In['sheet'] + '"'); return $r }
    $first = [Math]::Max(1, [int]$In['firstRow'])
    $grid = @{}; $formulas = @{}
    try {
        $kc = ConvertTo-EbiColumnNumber ([string]$In['keyColumn'])
        $last = [int]$ws.Cells.Item([int]$ws.Rows.Count, $kc).End(-4162).Row   # xlUp
        # End(xlUp) is Ctrl+Up: with a filter on it stops at the last VISIBLE
        # row, and the rows hidden under it would read as "no row" (10/21)
        try { $ur = $ws.UsedRange; $uLast = [int]$ur.Row + [int]$ur.Rows.Count - 1; if ($uLast -gt $last) { $last = $uLast } } catch { }
        if ($last -lt $first) { $r = $empty.Clone(); $r['ok'] = $true; $r['notFound'] = @(foreach ($it in @($In['items'])) { [string]$it[[string]$In['keyField']] }); return $r }
        $n = $last - $first + 1
        foreach ($l in $letters) {
            $v = $ws.Range(('{0}{1}:{0}{2}' -f $l, $first, $last)).Value2
            $arr = New-Object object[] $n
            for ($i = 1; $i -le $n; $i++) { $arr[$i - 1] = if ($n -eq 1) { $v } else { $v[$i, 1] } }
            $grid[$l] = $arr
        }
        $setL = New-Object System.Collections.ArrayList
        foreach ($l in @($In['setFields'].Keys) + @($In['setValues'].Keys)) { $u = ([string]$l).ToUpperInvariant(); if (-not $setL.Contains($u)) { [void]$setL.Add($u) } }
        foreach ($l in $setL) {
            $f = $ws.Range(('{0}{1}:{0}{2}' -f $l, $first, $last)).Formula
            $arr = New-Object object[] $n
            for ($i = 1; $i -le $n; $i++) { $arr[$i - 1] = if ($n -eq 1) { $f } else { $f[$i, 1] } }
            $formulas[$l] = $arr
        }
    } catch { $r = $empty.Clone(); $r['ok'] = $false; $r['failure'] = 'read_failed'; $r['message'] = $_.Exception.Message; return $r }
    $plan = ExcelMarkRows-Plan -Grid $grid -FirstRow $first -KeyColumn ([string]$In['keyColumn']) -Items $In['items'] -KeyField ([string]$In['keyField']) -Where $In['where'] -SetFields $In['setFields'] -SetValues $In['setValues'] -DateColumns $In['dateColumns'] -OnlyBlank ([bool]$In['onlyBlank']) -WhereMatch $wm -NumberColumns $In['numberColumns'] -Formulas $formulas
    try {
        foreach ($w in @($plan['writes'])) {
            $cell = $ws.Range(('{0}{1}' -f $w['letter'], $w['row']))
            if ([bool]$w['date'] -and [string]$w['value'] -match '^\d{4}-\d{2}-\d{2}$') {
                # a General cell would show the serial (46316); typing a date sets yyyy/m/d
                try { if ([string]$cell.NumberFormat -eq 'General') { $cell.NumberFormat = 'yyyy/m/d' } } catch { }
                $cell.Value2 = ([datetime]::ParseExact([string]$w['value'], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)).ToOADate()
            } elseif ([bool]$w['number'] -and [string]$w['value'] -match '^-?\d+(\.\d+)?$') {
                $cell.Value2 = [double]::Parse([string]$w['value'], [System.Globalization.CultureInfo]::InvariantCulture)
            } else {
                $cell.Value2 = [string]$w['value']
            }
        }
    } catch { $r = $empty.Clone(); $r['ok'] = $false; $r['failure'] = 'write_failed'; $r['message'] = $_.Exception.Message; return $r }
    # 10/20: the log said written, the WBS showed nothing -- read every
    # written cell back, so a write that did not stick is a failure here
    $lost = New-Object System.Collections.ArrayList
    foreach ($w in @($plan['writes'])) {
        $back = $null
        try { $back = $ws.Range(('{0}{1}' -f $w['letter'], $w['row'])).Value2 } catch { }
        if ($null -eq $back -or ([string]$back).Trim() -eq '') { [void]$lost.Add(('{0}{1}' -f $w['letter'], $w['row'])) }
    }
    if ($lost.Count -gt 0) {
        $r = $empty.Clone(); $r['ok'] = $false; $r['failure'] = 'write_failed'
        $r['message'] = ('written but still empty when read back: ' + ($lost.ToArray() -join ', ') + ' (a filter, a protected sheet, or the wrong workbook?) -- nothing saved yet; look at the sheet, then r')
        return $r
    }
    # 10/20: the values were "not there" at the gate -- a filter on the WBS
    # hides the rows written (the write itself is fine). Say which rows.
    $hidden = New-Object System.Collections.ArrayList
    foreach ($w in @($plan['writes'])) {
        if ($hidden.Contains([int]$w['row'])) { continue }
        try { if ([bool]$ws.Rows.Item([int]$w['row']).Hidden) { [void]$hidden.Add([int]$w['row']) } } catch { }
    }
    $written = @(ExcelMarkRows-NoteHidden -Lines $plan['written'] -HiddenRows $hidden.ToArray())
    $cells = @(foreach ($w in @($plan['writes'])) { ('{0}{1}' -f $w['letter'], $w['row']) })
    $hl = ''
    if ([string]$In['highlight'] -ne '' -and $cells.Count -gt 0) {
        # the colours already on the header rows and the rows written (whole
        # sheet would be too many COM calls), conditional formats included
        $used = New-Object System.Collections.ArrayList
        try {
            $ur = $ws.UsedRange; $lastCol = [Math]::Min(150, [int]$ur.Column + [int]$ur.Columns.Count - 1)
            $rows = @(1..([Math]::Max(1, $first - 1))) + @(@($plan['writes']) | ForEach-Object { [int]$_['row'] } | Select-Object -Unique)
            foreach ($r in $rows) { for ($c = 1; $c -le $lastCol; $c++) {
                $col = -1
                try { $col = [int64]$ws.Cells.Item($r, $c).DisplayFormat.Interior.Color } catch { try { $col = [int64]$ws.Cells.Item($r, $c).Interior.Color } catch { } }
                if ($col -ge 0 -and -not $used.Contains($col)) { [void]$used.Add($col) }
            } }
        } catch { }
        $pick = ExcelMarkRows-PickColor -Candidates ([string]$In['highlight']) -Used $used.ToArray()
        if ([int64]$pick['color'] -ge 0) {
            foreach ($a in $cells) { try { $ws.Range($a).Interior.Color = [int64]$pick['color'] } catch { } }
            $hl = [string]$pick['hex']
            $Ctx.Log.Info(('{0} cell(s) filled {1} for checking (the fill is never saved)' -f $cells.Count, $hl))
        }
    }
    Invoke-EbiComRelease $ws
    foreach ($l in $written) { $Ctx.Log.Info(('wrote ' + $l)) }
    $warn = New-Object System.Collections.ArrayList
    foreach ($k in @($plan['notFound'])) { [void]$warn.Add(@{ code = 'row_not_found'; message = ('no ' + $In['sheet'] + ' row for ' + $k + ' (type it by hand)') }) }
    if ($hidden.Count -gt 0) { [void]$warn.Add(@{ code = 'rows_hidden'; message = ('{0} row(s) written are hidden by a filter on {1} ({2}): the values are there, clear the filter to see them' -f $hidden.Count, $In['sheet'], (@($hidden.ToArray()) -join ', ')) }) }
    return @{ ok = $true; written = $written; writtenCount = $written.Count; kept = $plan['kept']; notFound = $plan['notFound']; cells = $cells; highlight = $hl; warnings = $warn.ToArray() }
}
