# modules/excel/excel.read_rows.ps1
# Read chosen columns of a sheet into records and keep the rows matching
# every condition (the verify.assert ops, kernel/Rules.ps1). Columns are
# named by letter in the call -- a plan sheet's header row is often merged
# or multi-line, the letter is what the operator points at. Each column is
# read in ONE COM call (Range.Value2 over the used rows), so a 5000-row
# plan reads in seconds, not minutes. Date / time columns come back as
# text (yyyy-MM-dd, HH:mm:ss), never as Excel serials.

. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Rules.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')   # ConvertTo-EbiClockText, ConvertTo-EbiColumnNumber

$Manifest = @{
  id         = 'excel.read_rows'
  group      = 'excel'
  summary    = 'Read lettered columns of a sheet into records, keep rows matching conditions'
  tier       = 'core'
  effects    = 'read'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook    = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet       = @{ type='string'; required=$true; desc='sheet name or 1-based index' }
    columns     = @{ type='map';    required=$true; desc='field -> column letter, e.g. { job: A, time: P }' }
    firstRow    = @{ type='int';    default=2; desc='first data row' }
    lastRow     = @{ type='int';    default=0; desc='0 = the last used row' }
    dateFields  = @{ type='list';   default=@(); desc='fields whose cells are dates -> yyyy-MM-dd' }
    timeFields  = @{ type='list';   default=@(); desc='fields whose cells are times -> HH:mm:ss (float artefacts rounded)' }
    where       = @{ type='list';   default=@(); desc='conditions @{ field; op; value } on the read fields' }
    pluck       = @{ type='string'; default=''; desc='field whose values go to plucked' }
  }
  outputs    = @{
    records = @{ type='list'; desc='kept rows: field -> text, plus _row (sheet row number)' }
    matched = @{ type='int' }
    scanned = @{ type='int' }
    plucked = @{ type='list' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'rules_invalid';   transient = $false }
    @{ id = 'read_failed';     transient = $true  }
  )
  example    = @{ use = 'excel.read_rows'; with = @{ workbook = 'plan'; sheet = 'WBS'; columns = @{ job = 'A'; kind = 'B' }; where = @(@{ field = 'kind'; op = 'equals'; value = 'recv' }); pluck = 'job' } }
  notes      = 'No matching row is not a failure (matched = 0): an empty day is an answer. A date cell holding text that looks like a date is passed through as it is.'
}

function ExcelReadRows-Date {
    param($v)
    if ($null -eq $v) { return '' }
    if ($v -is [double] -or $v -is [int]) { try { return ([datetime]::FromOADate([double]$v)).ToString('yyyy-MM-dd') } catch { return [string]$v } }
    if ($v -is [datetime]) { return $v.ToString('yyyy-MM-dd') }
    $t = ([string]$v).Trim()
    $m = [regex]::Match($t, '^(\d{4})[-/](\d{1,2})[-/](\d{1,2})')
    if ($m.Success) { return ('{0}-{1:00}-{2:00}' -f [int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value) }
    return $t
}

function Invoke-Step {
    param($In, $Ctx)
    $cols = $In['columns']
    $conds = @(@($In['where']) | Where-Object { $null -ne $_ })
    $ops = Get-EbiRuleOps
    foreach ($c in $conds) {
        if (-not ($c -is [System.Collections.IDictionary]) -or -not $c.Contains('field') -or -not ($ops -contains [string]$c['op'])) { return @{ ok = $false; failure = 'rules_invalid'; message = ('each condition needs field and op (one of ' + ($ops -join ', ') + ')'); records = @(); matched = 0; scanned = 0; plucked = @() } }
    }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would read sheet {0} columns {1}' -f $In['sheet'], (@($cols.Keys | Sort-Object | ForEach-Object { $_ + '=' + $cols[$_] }) -join ' '))); return @{ ok = $true; records = @(); matched = 0; scanned = 0; plucked = @() } }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $In['sheet'] + '"'); records = @(); matched = 0; scanned = 0; plucked = @() } }
    $first = [Math]::Max(1, [int]$In['firstRow'])
    $last = [int]$In['lastRow']
    try {
        if ($last -le 0) { $ur = $ws.UsedRange; $last = [int]$ur.Row + [int]$ur.Rows.Count - 1 }
        if ($last -lt $first) { return @{ ok = $true; records = @(); matched = 0; scanned = 0; plucked = @() } }
        $data = @{}
        foreach ($f in $cols.Keys) {
            $letter = [string]$cols[$f]
            $v = $ws.Range(('{0}{1}:{0}{2}' -f $letter, $first, $last)).Value2
            $data[[string]$f] = $v
        }
    } catch { return @{ ok = $false; failure = 'read_failed'; message = $_.Exception.Message; records = @(); matched = 0; scanned = 0; plucked = @() } }
    $dates = @(@($In['dateFields']) | ForEach-Object { [string]$_ })
    $times = @(@($In['timeFields']) | ForEach-Object { [string]$_ })
    $n = $last - $first + 1
    $kept = New-Object System.Collections.ArrayList
    for ($i = 1; $i -le $n; $i++) {
        $rec = @{ _row = ($first + $i - 1) }
        $any = $false
        foreach ($f in $data.Keys) {
            $arr = $data[$f]
            $raw = if ($n -eq 1) { $arr } else { $arr[$i, 1] }
            if ($dates -contains $f) { $val = ExcelReadRows-Date $raw }
            elseif ($times -contains $f) { $val = if ($null -eq $raw) { '' } else { ConvertTo-EbiClockText -Value $raw } }
            elseif ($null -eq $raw) { $val = '' }
            else { $val = ([string]$raw).Trim() }
            if ($val -ne '') { $any = $true }
            $rec[$f] = $val
        }
        if (-not $any) { continue }
        $all = $true
        foreach ($c in $conds) { if (-not (Test-EbiRuleHolds -Record $rec -Rule $c)) { $all = $false; break } }
        if ($all) { [void]$kept.Add($rec) }
    }
    $pf = [string]$In['pluck']
    $values = @(if ($pf -ne '') { foreach ($r in $kept) { [string]$r[$pf] } })
    Invoke-EbiComRelease $ws
    $Ctx.Log.Info(('read {0} row(s) of {1}, kept {2}' -f $n, $In['sheet'], $kept.Count))
    return @{ ok = $true; records = $kept.ToArray(); matched = $kept.Count; scanned = $n; plucked = $values }
}
