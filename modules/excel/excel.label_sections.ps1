# modules/excel/excel.label_sections.ps1
# Label every picture set on a sheet with the labels of the matching set on
# another sheet. A job that sends two or more files has one set of pictures
# per file (excel.stack_plan: first / separator / last, a blank row between
# sets); the send-data sheet carries the correl id and the STEP name above
# each of its sets (A3 / B3, A66 / B66 ...). The compare sheet needs the same
# labels above each of its sets (10/23 FJODWAF1: only the first set was
# labelled, the second one -- B66 / its STEP -- was typed by hand).
#
# Set k starts at the k-th shape whose role is first or only (the role
# excel.insert_pictures passes through from the plan); its labels go one row
# above that picture. The k-th label row of the source is the k-th row at or
# below fromFirstRow whose keyColumn is not empty.

. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')   # ConvertTo-EbiColumnNumber

$Manifest = @{
  id         = 'excel.label_sections'
  group      = 'excel'
  summary    = 'Write the labels of each picture set (one row above it) from the matching set on another sheet'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook     = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet        = @{ type='string'; required=$true; desc='the sheet whose picture sets get the labels' }
    shapes       = @{ type='list';   required=$true; desc='excel.insert_pictures shapes (row, role) of those sets' }
    fromSheet    = @{ type='string'; required=$true; desc='the sheet the labels come from' }
    keyColumn    = @{ type='string'; default='A'; desc='a non-empty cell here starts a label row on fromSheet' }
    fromFirstRow = @{ type='int';    default=3; desc='first row of fromSheet looked at' }
    cells        = @{ type='map';    required=$true; desc='target column -> source column, e.g. B <- A, G <- B' }
    startAt      = @{ type='int';    default=1; desc='first set labelled (1-based): 2 = leave set 1 to steps that already write it' }
    trim         = @{ type='bool';   default=$true }
    fontName     = @{ type='string'; default='' }
    fontSize     = @{ type='int';    default=0 }
  }
  outputs    = @{
    sets    = @{ type='int';  desc='picture sets found in shapes' }
    written = @{ type='list'; desc='"B66=JJPCRS12" per cell written' }
    missing = @{ type='list'; desc='set numbers with no label row on fromSheet (label them by hand)' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'write_failed';    transient = $true  }
  )
  example    = @{ use = 'excel.label_sections'; with = @{ workbook = 'wb'; sheet = 'compare'; shapes = '{{steps.dfPics.out.shapes}}'; fromSheet = 'data'; cells = @{ B = 'A'; G = 'B' }; startAt = 2 } }
  notes      = 'A set with no label row on fromSheet is left as it is and named in missing (warning label_missing). Cells are written as plain text, like excel.write_cell.'
}

function ExcelLabelSections-Starts {
    # PURE. shapes -> the row of each set's first picture, in order. Without
    # any role (plain picture paths), every picture is its own set.
    param($Shapes)
    $all = @(@($Shapes) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $withRole = @($all | Where-Object { $_.Contains('role') -and [string]$_['role'] -ne '' })
    if ($withRole.Count -eq 0) { return @($all | ForEach-Object { [int]$_['row'] }) }
    return @($all | Where-Object { $_.Contains('role') -and @('first', 'only') -contains [string]$_['role'] } | ForEach-Object { [int]$_['row'] })
}

function ExcelLabelSections-Plan {
    # PURE. set start rows + source label rows (each a map source column ->
    # text) + target<-source columns -> @{ writes = @(@{ set; row; cells });
    # missing = @(set numbers) }. Set k (k >= StartAt) is labelled from the
    # k-th label row, one row above its first picture.
    param($Starts, $Labels, $Cells, [int]$StartAt = 1)
    $writes = New-Object System.Collections.ArrayList
    $missing = New-Object System.Collections.ArrayList
    $starts = @($Starts)
    $labels = @($Labels)
    for ($k = [Math]::Max(1, $StartAt); $k -le $starts.Count; $k++) {
        $row = [int]$starts[$k - 1] - 1
        if ($row -lt 1) { [void]$missing.Add($k); continue }
        if ($k -gt $labels.Count) { [void]$missing.Add($k); continue }
        $src = $labels[$k - 1]
        $out = [ordered]@{}
        foreach ($t in @($Cells.Keys | Sort-Object)) {
            $sc = [string]$Cells[$t]
            $out[[string]$t] = $(if ($src -is [System.Collections.IDictionary] -and $src.Contains($sc) -and $null -ne $src[$sc]) { [string]$src[$sc] } else { '' })
        }
        [void]$writes.Add(@{ set = $k; row = $row; cells = $out })
    }
    return @{ writes = $writes.ToArray(); missing = $missing.ToArray() }
}

function ExcelLabelSections-ReadLabels {
    # The label rows of the source sheet: rows at or below FirstRow whose
    # key column holds text, each as a map source column -> text.
    param($Sheet, [string]$KeyColumn, [int]$FirstRow, [string[]]$Columns, [bool]$Trim)
    $kc = ConvertTo-EbiColumnNumber $KeyColumn
    $last = [int]$Sheet.Cells.Item([int]$Sheet.Rows.Count, $kc).End(-4162).Row   # xlUp
    $out = New-Object System.Collections.ArrayList
    for ($r = [Math]::Max(1, $FirstRow); $r -le $last; $r++) {
        $key = [string]$Sheet.Cells.Item($r, $kc).Text
        if ($key.Trim() -eq '') { continue }
        $m = @{}
        foreach ($c in $Columns) {
            $v = [string]$Sheet.Cells.Item($r, (ConvertTo-EbiColumnNumber $c)).Text
            $m[$c] = $(if ($Trim) { $v.Trim() } else { $v })
        }
        [void]$out.Add($m)
    }
    return $out.ToArray()
}

function Invoke-Step {
    param($In, $Ctx)
    $cells = $In['cells']
    $starts = @(ExcelLabelSections-Starts -Shapes $In['shapes'])
    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would label {0} set(s) on {1} from {2} (from set {3})' -f $starts.Count, $In['sheet'], $In['fromSheet'], [int]$In['startAt']))
        return @{ ok = $true; sets = $starts.Count; written = @(); missing = @() }
    }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $In['sheet'] + '"'); sets = $starts.Count; written = @(); missing = @() } }
    $src = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['fromSheet']
    if ($null -eq $src) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $In['fromSheet'] + '"'); sets = $starts.Count; written = @(); missing = @() } }
    $written = New-Object System.Collections.ArrayList
    try {
        $srcCols = @($cells.Keys | ForEach-Object { [string]$cells[$_] } | Select-Object -Unique)
        $labels = @(ExcelLabelSections-ReadLabels -Sheet $src -KeyColumn ([string]$In['keyColumn']) -FirstRow ([int]$In['fromFirstRow']) -Columns $srcCols -Trim ([bool]$In['trim']))
        $plan = ExcelLabelSections-Plan -Starts $starts -Labels $labels -Cells $cells -StartAt ([int]$In['startAt'])
        foreach ($w in @($plan['writes'])) {
            foreach ($col in @($w['cells'].Keys)) {
                $text = [string]$w['cells'][$col]
                if ($text -eq '') { continue }
                $c = $ws.Cells.Item([int]$w['row'], (ConvertTo-EbiColumnNumber ([string]$col)))
                Set-EbiCellTextPlain -Cell $c -Text $text -FontName ([string]$In['fontName']) -FontSize ([double]$In['fontSize'])
                [void]$written.Add(('{0}{1}={2}' -f $col, $w['row'], $text))
            }
        }
    } catch { return @{ ok = $false; failure = 'write_failed'; message = $_.Exception.Message; sets = $starts.Count; written = $written.ToArray(); missing = @() } }
    $missing = @($plan['missing'])
    $Ctx.Log.Info(('{0} set(s) on {1}: wrote {2}' -f $starts.Count, $In['sheet'], $(if ($written.Count) { $written.ToArray() -join ', ' } else { 'nothing' })))
    $ret = @{ ok = $true; sets = $starts.Count; written = $written.ToArray(); missing = $missing }
    if ($missing.Count -gt 0) { $ret['warnings'] = @(@{ code = 'label_missing'; message = ('set(s) ' + ($missing -join ', ') + ' on ' + $In['sheet'] + ' have no label row on ' + $In['fromSheet'] + ': label them by hand') }) }
    return $ret
}
