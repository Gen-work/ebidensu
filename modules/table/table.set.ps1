# modules/table/table.set.ps1
# Write one cell of one row and flush (P1-28, ported from MappingStore.ps1
# Update-MappingRows / Set-MappingBit). The row is found through
# kernel/Key.ps1 (no -eq here); a logical verdict value is translated to
# the column's stored code (profile verdict.values, P0-R11); a bit is
# named, and the name -> value map is the profile's (never 1/2/4 in a
# workflow).

. (Join-Path $PSScriptRoot '..\..\kernel\Table.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Worklist.ps1')

$Manifest = @{
  id         = 'table.set'
  group      = 'table'
  summary    = 'Set a cell (value or named bit) of the row for a key, then flush'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    worklist = @{ type='session'; sessionKind='worklist'; required=$true }
    key      = @{ type='string'; default=''; desc='which row (key display form); empty = the current item' }
    field    = @{ type='string'; required=$true }
    value    = @{ type='string'; default=''; desc='logical value to store (translated through verdict.values)' }
    bit      = @{ type='string'; default=''; desc='instead of value: the bit NAME to set in a bitmask column' }
    clear    = @{ type='bool';   default=$false; desc='with bit: clear it instead of setting it' }
  }
  outputs    = @{
    stored  = @{ type='string'; desc='what the cell holds now' }
    updated = @{ type='int';    desc='rows changed (0 or 1)' }
  }
  failures   = @(
    @{ id = 'row_not_found';   transient = $false }
    @{ id = 'ambiguous';       transient = $false }
    @{ id = 'column_missing';  transient = $false }
    @{ id = 'bit_unknown';     transient = $false }
    @{ id = 'write_failed';    transient = $true  }
  )
  example    = @{ use = 'table.set'; with = @{ worklist = 'wl'; field = 'note'; value = 'checked by hand' } }
}

function TableSet-Resolve {
    <#
      PURE. Which row and what to write.
      -> @{ ok; failure; message; row; stored }
    #>
    param($Worklist, [string]$Key, $Item, $KeyColumns, $Rules, [string]$Field, [string]$Value, [string]$Bit, [bool]$Clear, $Spec)
    $columns = @($Worklist['columns'])
    if (-not ($columns -contains $Field)) { return @{ ok = $false; failure = 'column_missing'; message = ('column "' + $Field + '" is not in the worklist'); row = $null; stored = '' } }
    $row = $null
    if ($Key -ne '') {
        $m = Find-EbiKeyMatches -Records $Worklist['rows'] -Key $Key -KeyColumns $KeyColumns -Rules $Rules
        $n = @($m['indexes']).Count
        if ($n -eq 0) { return @{ ok = $false; failure = 'row_not_found'; message = ('no row for key "' + $Key + '"'); row = $null; stored = '' } }
        if ($n -gt 1) { return @{ ok = $false; failure = 'ambiguous'; message = ('' + $n + ' rows match key "' + $Key + '" (' + $m['tier'] + ')'); row = $null; stored = '' } }
        $row = $m['records'][0]
    } elseif ($null -ne $Item) {
        $row = $Item
    } else {
        return @{ ok = $false; failure = 'row_not_found'; message = 'no key given and no current item (outside each?)'; row = $null; stored = '' }
    }
    $stored = ''
    if ($Bit -ne '') {
        $bv = Get-EbiWorklistBitValue -Name $Bit -Spec $Spec
        if ($bv -le 0) { return @{ ok = $false; failure = 'bit_unknown'; message = ('bit "' + $Bit + '" is not declared for column "' + $Field + '"'); row = $row; stored = '' } }
        $cur = 0
        $s = if ($row.Contains($Field) -and $null -ne $row[$Field]) { ([string]$row[$Field]).Trim() } else { '' }
        if ($s -match '^\d+$') { $cur = [int]$s }
        $stored = [string]$(if ($Clear) { $cur -band (-bnot $bv) } else { $cur -bor $bv })
    } else {
        $stored = ConvertTo-EbiStoredValue -Logical $Value -Spec $Spec
    }
    return @{ ok = $true; failure = ''; message = ''; row = $row; stored = $stored }
}

function Invoke-Step {
    param($In, $Ctx)
    $wl = $In['worklist']
    $keyColumns = @($Ctx['KeyColumns'])
    if ($keyColumns.Count -eq 0) { $keyColumns = @(Get-EbiWorklistKeyColumns -Profile $Ctx['Profile']) }
    $field = [string]$In['field']
    $spec = Get-EbiWorklistColumnSpec -Profile $Ctx['Profile'] -Field $field
    $r = TableSet-Resolve -Worklist $wl -Key ([string]$In['key']) -Item $Ctx['Item'] -KeyColumns $keyColumns -Rules (Get-EbiKeyRules -Profile $Ctx['Profile']) -Field $field -Value ([string]$In['value']) -Bit ([string]$In['bit']) -Clear ([bool]$In['clear']) -Spec $spec
    if (-not $r['ok']) { return @{ ok = $false; failure = $r['failure']; message = $r['message']; stored = ''; updated = 0 } }
    $row = $r['row']
    $before = if ($row.Contains($field) -and $null -ne $row[$field]) { [string]$row[$field] } else { '' }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would set {0} = "{1}" (was "{2}")' -f $field, $r['stored'], $before)); return @{ ok = $true; stored = $r['stored']; updated = 0 } }
    $row[$field] = $r['stored']
    if (-not [string]::IsNullOrWhiteSpace([string]$wl['path'])) {
        $s = Save-EbiWorklist -Worklist $wl
        if (-not $s['ok']) { $row[$field] = $before; return @{ ok = $false; failure = 'write_failed'; message = $s['message']; stored = $before; updated = 0 } }
    }
    return @{ ok = $true; stored = $r['stored']; updated = 1 }
}
