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
  }
  outputs    = @{
    written      = @{ type='list'; desc='"KEY row 123: AJ=2026-10-16 AK=2026-10-20" per row written' }
    writtenCount = @{ type='int' }
    kept         = @{ type='list'; desc='"KEY row 123: AJ already 2026-10-16" for cells left as they were' }
    notFound     = @{ type='list'; desc='keys with no matching row' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'input_invalid';   transient = $false }
    @{ id = 'read_failed';     transient = $true  }
    @{ id = 'write_failed';    transient = $true  }
  )
  example    = @{ use = 'excel.mark_rows'; with = @{ workbook = 'wbs'; sheet = 'WBS'; firstRow = 7; keyColumn = 'A'; items = '{{steps.pick.out.rows}}'; keyField = 'JOB'; where = @{ B = 'recv'; G = @('after', 'summary') }; whereMatch = @{ P = '^owner' }; setValues = @{ AJ = '{{run.date}}'; AK = '{{run.date}}'; AQ = '0.1' }; dateColumns = @('AJ', 'AK'); numberColumns = @('AQ') } }
  notes      = 'A key with several matching rows (two files listed twice) gets every one written. A value that should be a date but does not read as one is written as text and reported in kept. Under DryRun nothing is read or written; the plan is logged.'
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
    param($Grid, [int]$FirstRow, [string]$KeyColumn, $Items, [string]$KeyField, $Where, $SetFields, $SetValues, $DateColumns = @(), [bool]$OnlyBlank = $true, $WhereMatch = @{}, $NumberColumns = @(), [datetime]$Today = (Get-Date))
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

function Invoke-Step {
    param($In, $Ctx)
    $empty = @{ written = @(); writtenCount = 0; kept = @(); notFound = @() }
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
    $grid = @{}
    try {
        $kc = ConvertTo-EbiColumnNumber ([string]$In['keyColumn'])
        $last = [int]$ws.Cells.Item([int]$ws.Rows.Count, $kc).End(-4162).Row   # xlUp
        if ($last -lt $first) { $r = $empty.Clone(); $r['ok'] = $true; $r['notFound'] = @(foreach ($it in @($In['items'])) { [string]$it[[string]$In['keyField']] }); return $r }
        $n = $last - $first + 1
        foreach ($l in $letters) {
            $v = $ws.Range(('{0}{1}:{0}{2}' -f $l, $first, $last)).Value2
            $arr = New-Object object[] $n
            for ($i = 1; $i -le $n; $i++) { $arr[$i - 1] = if ($n -eq 1) { $v } else { $v[$i, 1] } }
            $grid[$l] = $arr
        }
    } catch { $r = $empty.Clone(); $r['ok'] = $false; $r['failure'] = 'read_failed'; $r['message'] = $_.Exception.Message; return $r }
    $plan = ExcelMarkRows-Plan -Grid $grid -FirstRow $first -KeyColumn ([string]$In['keyColumn']) -Items $In['items'] -KeyField ([string]$In['keyField']) -Where $In['where'] -SetFields $In['setFields'] -SetValues $In['setValues'] -DateColumns $In['dateColumns'] -OnlyBlank ([bool]$In['onlyBlank']) -WhereMatch $wm -NumberColumns $In['numberColumns']
    try {
        foreach ($w in @($plan['writes'])) {
            $cell = $ws.Range(('{0}{1}' -f $w['letter'], $w['row']))
            if ([bool]$w['date'] -and [string]$w['value'] -match '^\d{4}-\d{2}-\d{2}$') {
                $cell.Value2 = ([datetime]::ParseExact([string]$w['value'], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)).ToOADate()
            } elseif ([bool]$w['number'] -and [string]$w['value'] -match '^-?\d+(\.\d+)?$') {
                $cell.Value2 = [double]::Parse([string]$w['value'], [System.Globalization.CultureInfo]::InvariantCulture)
            } else {
                $cell.Value2 = [string]$w['value']
            }
        }
    } catch { $r = $empty.Clone(); $r['ok'] = $false; $r['failure'] = 'write_failed'; $r['message'] = $_.Exception.Message; return $r }
    Invoke-EbiComRelease $ws
    foreach ($l in @($plan['written'])) { $Ctx.Log.Info(('wrote ' + $l)) }
    $warn = @(foreach ($k in @($plan['notFound'])) { @{ code = 'row_not_found'; message = ('no ' + $In['sheet'] + ' row for ' + $k + ' (type it by hand)') } })
    return @{ ok = $true; written = $plan['written']; writtenCount = @($plan['written']).Count; kept = $plan['kept']; notFound = $plan['notFound']; warnings = $warn }
}
