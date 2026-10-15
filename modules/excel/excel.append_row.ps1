# modules/excel/excel.append_row.ps1
# Append ONE row to a list sheet (the review check sheet: one line per
# deliverable), unless a row already holds the same key -- so a rerun, or
# a deliverable someone already listed by hand, adds nothing (10/15: R/S
# were listed on 10/14 and must not be listed twice). The new row goes
# under the last used cell of `lastBy`; a date column takes the number
# format of the row above, so the sheet looks the same as hand-entered
# rows. Ported from FillCheckSheet.ps1 (the old CheckSheet phase).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Start/Stop-EbiDialogWatch around Save
. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')   # ConvertTo-EbiColumnNumber

$Manifest = @{
  id         = 'excel.append_row'
  group      = 'excel'
  summary    = 'Append one row under the last used row, unless the key is already listed; optionally save'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook    = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet       = @{ type='string'; required=$true }
    cells       = @{ type='map';    required=$true; desc='column letter -> value, e.g. { "B": "2026-10-16", "F": "<file name>" }; empty values are left blank' }
    keyColumn   = @{ type='string'; default=''; desc='column whose text identifies a row (e.g. F, the review target); a row already holding cells[keyColumn] there means "already listed": nothing is written. Empty = always append' }
    lastBy      = @{ type='string'; default='B'; desc='column whose last used cell marks the end of the list (pre-filled number columns would fool a whole-row search)' }
    firstRow    = @{ type='int';    default=2; desc='first data row: the key search starts here' }
    dateColumns = @{ type='list';   default=@(); desc='columns whose value (yyyy-MM-dd or yyyy/MM/dd) is written as a real date, formatted like the row above' }
    save        = @{ type='bool';   default=$true; desc='save the workbook right after writing (a shared list: keep the time it is changed short)' }
    checkOnly   = @{ type='bool';   default=$false; desc='only look: is the key listed? (existing / verdict / line), never write or save -- for a flow that writes the row after a person confirmed (gfixRecv.deliver: after the mail was sent)' }
    dismissDialogs = @{ type='list'; default=@(); desc='title parts of dialogs a save opens (the classification add-in) answered with Esc, as excel.tidy does' }
  }
  outputs    = @{
    row      = @{ type='int';  desc='the row written, or the row that already held the key; 0 in a dry run' }
    added    = @{ type='bool'; desc='true when a row was written' }
    existing = @{ type='bool'; desc='true when the key was already listed (nothing written)' }
    line     = @{ type='string'; desc='one line for a gate: "row N added" / "already listed on row N (date)"' }
    verdict  = @{ type='string'; desc='ok when the key was already listed (the list says it is done: nothing for a person to do), unknown when a row was added (a person finishes the job, e.g. sends the mail) -- the code for a human.gate askWhen [unknown]' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'input_invalid';   transient = $false }
    @{ id = 'read_only';       transient = $false }
    @{ id = 'write_failed';    transient = $true  }
    @{ id = 'save_failed';     transient = $true  }
  )
  example    = @{ use = 'excel.append_row'; with = @{ workbook = 'cs'; sheet = 'Check Sheet_J4'; keyColumn = 'F'; dateColumns = @('B'); cells = @{ B = '{{run.date}}'; C = 'JAVA'; F = 'book_{{item.key}}.xlsx' } } }
  notes      = 'The key is compared as trimmed text, case-insensitively. read_only is not transient here: the session holds the read-only copy, so retrying cannot help -- close the run, have the other person close the book, run again (the row is not listed yet, so the rerun adds it).'
}

function ExcelAppendRow-Plan {
    # PURE. Where the row goes and whether to write at all.
    #   LastRow   last used row of lastBy (0 = empty column)
    #   KeyTexts  hashtable row -> text of keyColumn for firstRow..LastRow
    # -> @{ ok; row; added; existing; message }
    param([int]$LastRow, [int]$FirstRow, [hashtable]$KeyTexts, [string]$Key)
    $k = ([string]$Key).Trim()
    if ($k -ne '' -and $null -ne $KeyTexts) {
        foreach ($r in ($KeyTexts.Keys | Sort-Object)) {
            if ([string]::Equals(([string]$KeyTexts[$r]).Trim(), $k, [System.StringComparison]::OrdinalIgnoreCase)) { return @{ ok = $true; row = [int]$r; added = $false; existing = $true; verdict = 'ok'; message = ('already listed on row ' + $r) } }
        }
    }
    $row = [Math]::Max($LastRow + 1, $FirstRow)
    return @{ ok = $true; row = $row; added = $true; existing = $false; verdict = 'unknown'; line = '' }
}

function ExcelAppendRow-Said {
    # PURE. The gate line for what happened.
    param([string]$Sheet, [int]$Row, [bool]$Existing, [string]$When = '')
    if ($Existing) { return ('already listed on {0} row {1}{2}: nothing added' -f $Sheet, $Row, $(if ($When -ne '') { ' (' + $When + ')' } else { '' })) }
    return ('{0} row {1} added and saved' -f $Sheet, $Row)
}

function ExcelAppendRow-ParseDate {
    # PURE. yyyy-MM-dd / yyyy/MM/dd (time ignored) -> [datetime] or $null.
    param([string]$Text)
    $t = ([string]$Text).Trim()
    if ($t -match '^(\d{4})[-/](\d{1,2})[-/](\d{1,2})') {
        try { return (New-Object DateTime ([int]$Matches[1]), ([int]$Matches[2]), ([int]$Matches[3])) } catch { return $null }
    }
    return $null
}

function Invoke-Step {
    param($In, $Ctx)
    $cells = $In['cells']
    if (-not ($cells -is [System.Collections.IDictionary]) -or $cells.Count -eq 0) { return @{ ok = $false; failure = 'input_invalid'; message = 'cells must be a map of column letter -> value'; row = 0; added = $false; existing = $false } }
    foreach ($c in @($cells.Keys) + @($In['dateColumns']) + @([string]$In['lastBy'])) { if (-not ([string]$c -match '^[A-Za-z]{1,3}$')) { return @{ ok = $false; failure = 'input_invalid'; message = ('"' + $c + '" is not a column letter'); row = 0; added = $false; existing = $false } } }
    $keyCol = [string]$In['keyColumn']
    $key = if ($keyCol -ne '' -and $cells.Contains($keyCol)) { [string]$cells[$keyCol] } else { '' }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would append to {0} (key {1}={2}): {3}' -f $In['sheet'], $keyCol, $key, ((@($cells.Keys | Sort-Object) | ForEach-Object { $_ + '=' + $cells[$_] }) -join ' '))); return @{ ok = $true; row = 0; added = $false; existing = $false; verdict = 'unknown'; line = '' } }
    $wb = $In['workbook']
    try { if ([bool]$wb.ReadOnly -and -not [bool]$In['checkOnly']) { return @{ ok = $false; failure = 'read_only'; message = ([string]$wb.Name + ' is open read-only (someone else has it open): close this run, ask them to close it, run again'); row = 0; added = $false; existing = $false } } } catch { }
    $ws = Get-EbiSheet -Workbook $wb -Sheet $In['sheet']
    if ($null -eq $ws) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $In['sheet'] + '"'); row = 0; added = $false; existing = $false } }
    try {
        $lastCol = ConvertTo-EbiColumnNumber ([string]$In['lastBy'])
        $last = [int]$ws.Cells.Item([int]$ws.Rows.Count, $lastCol).End(-4162).Row   # xlUp
        if ($last -eq 1 -and [string]$ws.Cells.Item(1, $lastCol).Text -eq '') { $last = 0 }
        $texts = @{}
        if ($keyCol -ne '' -and $key -ne '') {
            $kc = ConvertTo-EbiColumnNumber $keyCol
            $kLast = [Math]::Max($last, [int]$ws.Cells.Item([int]$ws.Rows.Count, $kc).End(-4162).Row)
            for ($r = [int]$In['firstRow']; $r -le $kLast; $r++) { $texts[$r] = [string]$ws.Cells.Item($r, $kc).Text }
        }
        $plan = ExcelAppendRow-Plan -LastRow $last -FirstRow ([int]$In['firstRow']) -KeyTexts $texts -Key $key
        $row = [int]$plan['row']
        $when = ''
        if ($plan['existing']) { try { $when = [string]$ws.Cells.Item($row, $lastCol).Text } catch { } }
        $pendingSave = $false
        try { $pendingSave = -not [bool]$wb.Saved } catch { }
        if ([bool]$In['checkOnly']) {
            if ($plan['existing']) { $Ctx.Log.Info(('{0}: {1}' -f $In['sheet'], $plan['message'])); return @{ ok = $true; row = $row; added = $false; existing = $true; verdict = 'ok'; line = (ExcelAppendRow-Said -Sheet ([string]$In['sheet']) -Row $row -Existing $true -When $when) } }
            return @{ ok = $true; row = 0; added = $false; existing = $false; verdict = 'unknown'; line = ('not on ' + [string]$In['sheet'] + ' yet') }
        }
        if ($plan['existing']) {
            $Ctx.Log.Info(('{0}: {1} -- nothing written' -f $In['sheet'], $plan['message']))
            Show-EbiExcelAt -Worksheet $ws -Row $row
            # a retry after save_failed finds its own unsaved row: save it now
            if (-not ([bool]$In['save'] -and $pendingSave)) { return @{ ok = $true; row = $row; added = $false; existing = $true; verdict = 'ok'; line = (ExcelAppendRow-Said -Sheet ([string]$In['sheet']) -Row $row -Existing $true -When $when) } }
        }
        if (-not $plan['existing']) { Show-EbiExcelAt -Worksheet $ws -Row $row -Front }
        $dates = @(@($In['dateColumns']) | ForEach-Object { ([string]$_).ToUpperInvariant() })
        foreach ($c in $(if ($plan['existing']) { @() } else { @($cells.Keys) })) {
            $v = [string]$cells[$c]
            if ($v -eq '') { continue }
            $col = ConvertTo-EbiColumnNumber ([string]$c)
            $cell = $ws.Cells.Item($row, $col)
            $d = if ($dates -contains ([string]$c).ToUpperInvariant()) { ExcelAppendRow-ParseDate $v } else { $null }
            if ($null -ne $d) {
                $fmt = ''
                if ($row -gt 1) { try { $fmt = [string]$ws.Cells.Item($row - 1, $col).NumberFormat } catch { } }
                if ($fmt -eq '' -or $fmt -eq 'General' -or $fmt -eq '@') { $fmt = 'yyyy/m/d' }
                $cell.NumberFormat = $fmt
                $cell.Value2 = $d.ToOADate()
            } else {
                Set-EbiCellTextPlain -Cell $cell -Text $v
            }
        }
    } catch { return @{ ok = $false; failure = 'write_failed'; message = $_.Exception.Message; row = 0; added = $false; existing = $false } }
    if ([bool]$In['save']) {
        $err = ''
        $watch = Start-EbiDialogWatch -TitleParts @($In['dismissDialogs'])
        try { [void]$wb.Save() } catch { $err = $_.Exception.Message }
        $seen = Stop-EbiDialogWatch $watch
        foreach ($t in @($seen['answered'])) { $Ctx.Log.Info(('answered "{0}" with Esc while saving' -f $t)) }
        if ($err -ne '') { return @{ ok = $false; failure = 'save_failed'; message = $err; row = $row; added = (-not $plan['existing']); existing = [bool]$plan['existing'] } }
    }
    if ($plan['existing']) { $Ctx.Log.Info(('{0}: row {1} was already there, unsaved: saved now' -f $In['sheet'], $row)); return @{ ok = $true; row = $row; added = $false; existing = $true; verdict = 'ok'; line = (ExcelAppendRow-Said -Sheet ([string]$In['sheet']) -Row $row -Existing $true -When $when) } }
    $Ctx.Log.Info(('{0}: row {1} added{2}' -f $In['sheet'], $row, $(if ([bool]$In['save']) { ' and saved' } else { '' })))
    return @{ ok = $true; row = $row; added = $true; existing = $false; verdict = 'unknown'; line = (ExcelAppendRow-Said -Sheet ([string]$In['sheet']) -Row $row -Existing $false -When '') }
}
