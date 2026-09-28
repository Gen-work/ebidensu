# modules/flow/flow.checkpoint.ps1
# Mark the current item done for a field (P1-28, WORKFLOW-SCHEMA 7.3): the
# worklist cell gets the logical value translated to the stored code (or a
# named bit set) and the table is flushed atomically. The ledger record is
# the runner's, written around every each-step; this step is the one whose
# replay-on-resume matters, which is why it never reads what it wrote.
# `value` is meant to be {{steps.gate.out.code}} with
# `when: steps.gate.out.action != skip` (5.2): a skipped item is left
# pending by NOT running this step.

. (Join-Path $PSScriptRoot '..\..\kernel\Table.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Worklist.ps1')

$Manifest = @{
  id         = 'flow.checkpoint'
  group      = 'flow'
  summary    = 'Write the verdict (or a named bit) for the current item and flush'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    worklist = @{ type='session'; sessionKind='worklist'; required=$true }
    field    = @{ type='string'; required=$true }
    value    = @{ type='string'; default=''; desc='logical value: ok | ng | unknown | empty' }
    bit      = @{ type='string'; default=''; desc='instead of value: the bit NAME (profile bits map) to set' }
    key      = @{ type='string'; default=''; desc='which row; empty = the current item' }
  }
  outputs    = @{
    stored = @{ type='string'; desc='the code the cell holds now' }
    key    = @{ type='string'; desc='the row that was written' }
  }
  failures   = @(
    @{ id = 'row_not_found';  transient = $false }
    @{ id = 'ambiguous';      transient = $false }
    @{ id = 'column_missing'; transient = $false }
    @{ id = 'bit_unknown';    transient = $false }
    @{ id = 'value_invalid';  transient = $false }
    @{ id = 'write_failed';   transient = $true  }
  )
  example    = @{ use = 'flow.checkpoint'; with = @{ worklist = 'wl'; field = 'before_transferStatus'; value = '{{steps.gate.out.code}}' } }
  notes      = 'value must be a logical verdict (ok / ng / unknown / empty); a stored code is refused (value_invalid) so a workflow cannot hard-wire the profile''s encoding.'
}

function FlowCheckpoint-Row {
    # PURE. The row to write: by key through Key.ps1, else the current item.
    param($Worklist, [string]$Key, $Item, $KeyColumns, $Rules)
    if ($Key -ne '') {
        $m = Find-EbiKeyMatches -Records $Worklist['rows'] -Key $Key -KeyColumns $KeyColumns -Rules $Rules
        $n = @($m['indexes']).Count
        if ($n -eq 0) { return @{ ok = $false; failure = 'row_not_found'; message = ('no row for key "' + $Key + '"'); row = $null } }
        if ($n -gt 1) { return @{ ok = $false; failure = 'ambiguous'; message = ('' + $n + ' rows match key "' + $Key + '" (' + $m['tier'] + ')'); row = $null } }
        return @{ ok = $true; failure = ''; message = ''; row = $m['records'][0] }
    }
    if ($null -eq $Item) { return @{ ok = $false; failure = 'row_not_found'; message = 'no key given and no current item (checkpoint belongs in each)'; row = $null } }
    return @{ ok = $true; failure = ''; message = ''; row = $Item }
}

function Invoke-Step {
    param($In, $Ctx)
    $wl = $In['worklist']
    $field = [string]$In['field']; $value = [string]$In['value']; $bit = [string]$In['bit']
    $keyColumns = @($Ctx['KeyColumns'])
    if ($keyColumns.Count -eq 0) { $keyColumns = @(Get-EbiWorklistKeyColumns -Profile $Ctx['Profile']) }
    if (-not (@($wl['columns']) -contains $field)) { return @{ ok = $false; failure = 'column_missing'; message = ('column "' + $field + '" is not in the worklist'); stored = ''; key = '' } }
    if ($bit -eq '' -and -not ($value -in @('ok', 'ng', 'unknown', ''))) { return @{ ok = $false; failure = 'value_invalid'; message = ('value "' + $value + '" is not a logical verdict (ok / ng / unknown / empty)'); stored = ''; key = '' } }
    $r = FlowCheckpoint-Row -Worklist $wl -Key ([string]$In['key']) -Item $Ctx['Item'] -KeyColumns $keyColumns -Rules (Get-EbiKeyRules -Profile $Ctx['Profile'])
    if (-not $r['ok']) { return @{ ok = $false; failure = $r['failure']; message = $r['message']; stored = ''; key = '' } }
    $row = $r['row']
    $keyText = Get-EbiKeyDisplay -Item $row -KeyColumns $keyColumns
    $spec = Get-EbiWorklistColumnSpec -Profile $Ctx['Profile'] -Field $field
    $before = if ($row.Contains($field) -and $null -ne $row[$field]) { [string]$row[$field] } else { '' }
    if ($bit -ne '') {
        $bv = Get-EbiWorklistBitValue -Name $bit -Spec $spec
        if ($bv -le 0) { return @{ ok = $false; failure = 'bit_unknown'; message = ('bit "' + $bit + '" is not declared for column "' + $field + '"'); stored = $before; key = $keyText } }
        $cur = 0; if ($before.Trim() -match '^\d+$') { $cur = [int]$before.Trim() }
        $stored = [string]($cur -bor $bv)
    } else {
        $stored = ConvertTo-EbiStoredValue -Logical $value -Spec $spec
    }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would checkpoint {0}: {1} = "{2}"' -f $keyText, $field, $stored)); return @{ ok = $true; stored = $stored; key = $keyText } }
    $row[$field] = $stored
    if (-not [string]::IsNullOrWhiteSpace([string]$wl['path'])) {
        $s = Save-EbiWorklist -Worklist $wl
        if (-not $s['ok']) { $row[$field] = $before; return @{ ok = $false; failure = 'write_failed'; message = $s['message']; stored = $before; key = $keyText } }
    }
    return @{ ok = $true; stored = $stored; key = $keyText }
}
