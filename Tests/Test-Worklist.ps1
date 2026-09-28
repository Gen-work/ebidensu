#Requires -Version 5.1
# Test-Worklist.ps1 -- kernel/Worklist.ps1 (P1-03): the one row-selection
# implementation shared by the runner's source.select and table.select.
# Pure: fixtures only, no files.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Key.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Worklist.ps1')

Reset-Tests 'Worklist'

# ---------------------------------------------------------------- pendingWhen parsing
foreach ($case in @(
    @{ t = 'empty';        kind = 'empty'; arg = '' },
    @{ t = 'always';       kind = 'always'; arg = '' },
    @{ t = '!= ok';        kind = 'notOk'; arg = '' },
    @{ t = '!=ok';         kind = 'notOk'; arg = '' },
    @{ t = '== ng';        kind = 'isNg';  arg = '' },
    @{ t = 'bit !before';  kind = 'bit';   arg = 'before' },
    @{ t = 'bit ! 3';      kind = 'bit';   arg = '3' },
    @{ t = '  always  ';   kind = 'always'; arg = '' }
)) {
    $p = ConvertFrom-EbiPendingWhen -Text $case['t']
    Assert-True ($p['ok'] -and $p['kind'] -eq $case['kind'] -and $p['arg'] -eq $case['arg']) ('parse: "' + $case['t'] + '" -> ' + $case['kind'])
}
foreach ($bad in @('', '!= done', '== ok', 'bit 3', 'bit !', 'pending', '!= ok && x')) {
    $p = ConvertFrom-EbiPendingWhen -Text $bad
    Assert-True (-not $p['ok'] -and $p['message'] -like '*allowed*' -or $p['message'] -like '*five allowed forms*') ('parse: "' + $bad + '" is refused with the allowed list')
}

# ---------------------------------------------------------------- values translation
$verdict = @{ name = 'status'; role = 'verdict'; values = @{ ok = '1'; ng = '2'; unknown = ''; pending = '0' } }
Assert-Equal '1'  (ConvertTo-EbiStoredValue -Logical 'ok' -Spec $verdict) 'store: ok -> 1 through the values map'
Assert-Equal '0'  (ConvertTo-EbiStoredValue -Logical 'pending' -Spec $verdict) 'store: pending -> 0'
Assert-Equal 'ok' (ConvertTo-EbiStoredValue -Logical 'ok' -Spec $null) 'store: no spec -> identity'
Assert-Equal 'ok' (ConvertTo-EbiStoredValue -Logical 'ok' -Spec @{ name = 'x' }) 'store: no values map -> identity'
Assert-Equal 'ok'      (ConvertFrom-EbiStoredValue -Value '1' -Spec $verdict) 'read: 1 -> ok'
Assert-Equal 'ng'      (ConvertFrom-EbiStoredValue -Value '2' -Spec $verdict) 'read: 2 -> ng'
Assert-Equal 'pending' (ConvertFrom-EbiStoredValue -Value '0' -Spec $verdict) 'read: 0 -> pending'
Assert-Equal ''        (ConvertFrom-EbiStoredValue -Value '' -Spec $verdict) 'read: a blank cell stays blank even though unknown maps to blank'
Assert-Equal ''        (ConvertFrom-EbiStoredValue -Value $null -Spec $verdict) 'read: null is blank'
Assert-Equal '7'       (ConvertFrom-EbiStoredValue -Value '7' -Spec $verdict) 'read: an unmapped code is returned as-is'
Assert-Equal 'ok'      (ConvertFrom-EbiStoredValue -Value 'ok' -Spec $null) 'read: no spec -> identity'
Assert-Equal 'ok'      (ConvertFrom-EbiStoredValue -Value ' 1 ' -Spec $verdict) 'read: surrounding whitespace is ignored'

# ---------------------------------------------------------------- row tests
$mask = @{ name = 'mask'; role = 'bitmask'; bits = @{ before = 1; after = 2; compare = 4 } }
$rows = @(
    @{ id = 'r0'; status = '0'; plain = '';        mask = '0' },
    @{ id = 'r1'; status = '1'; plain = 'ok';      mask = '1' },
    @{ id = 'r2'; status = '2'; plain = 'ng';      mask = '3' },
    @{ id = 'r3'; status = '';  plain = 'unknown'; mask = '7' },
    @{ id = 'r4'; status = '9'; plain = 'x';       mask = 'abc' }
)
function Sel { param([string]$Field, [string]$Pw, $Spec) $r = Select-EbiWorklistRows -Rows $rows -Field $Field -PendingWhen $Pw -Spec $Spec; return (@($r['rows'] | ForEach-Object { $_['id'] }) -join ',') }
Assert-Equal 'r0,r3'          (Sel 'status' 'empty' $verdict)  'select: empty = blank or the pending code'
Assert-Equal 'r0,r2,r3,r4'    (Sel 'status' '!= ok' $verdict)  'select: != ok keeps ng, unknown, pending and unmapped (3.2: ng is still pending)'
Assert-Equal 'r2'             (Sel 'status' '== ng' $verdict)  'select: == ng'
Assert-Equal 'r0,r1,r2,r3,r4' (Sel 'status' 'always' $verdict) 'select: always'
Assert-Equal 'r0'             (Sel 'plain' 'empty' $null)      'select: without a values map, empty = blank or 0 (unknown is a verdict, not empty)'
Assert-Equal 'r0,r2,r3,r4'    (Sel 'plain' '!= ok' $null)      'select: without a values map, the cell IS the logical value'
Assert-Equal 'r2'             (Sel 'plain' '== ng' $null)      'select: == ng on plain logical values'
Assert-Equal 'r0,r4'          (Sel 'mask' 'bit !before' $mask) 'select: bit by name (1 not set; a non-numeric cell counts as 0)'
Assert-Equal 'r0,r1,r4'       (Sel 'mask' 'bit !after' $mask)  'select: bit by name (2 not set)'
Assert-Equal 'r0,r1,r2,r4'    (Sel 'mask' 'bit !4' $mask)      'select: bit by number'
Assert-Equal 'r0,r1,r2,r3,r4' (Sel 'mask' 'bit !nope' $mask)   'select: an unknown bit name is never done (everything pending, nothing hidden)'
Assert-Equal 'r0,r1,r2,r3,r4' (Sel 'status' 'always' $null)    'select: always needs no field'
$r = Select-EbiWorklistRows -Rows $rows -Field '' -PendingWhen '!= ok' -Spec $null
Assert-True (-not $r['ok'] -and $r['message'] -like '*select.field is required*') 'select: a field is required unless always'
$r = Select-EbiWorklistRows -Rows $rows -Field 'status' -PendingWhen 'nope' -Spec $null
Assert-True (-not $r['ok'] -and $r['message'] -like '*five allowed forms*') 'select: a bad pendingWhen is a failure record'
$r = Select-EbiWorklistRows -Rows $rows -Field 'status' -PendingWhen 'always' -Spec $null -Limit 2
Assert-Equal 2 $r['selected'] 'select: limit caps the result'
Assert-Equal 5 $r['total'] 'select: total is the whole table'
$r = Select-EbiWorklistRows -Rows $rows -Field 'nosuch' -PendingWhen '!= ok' -Spec $null
Assert-Equal 5 $r['selected'] 'select: a missing column reads as blank, so everything is pending'
$r = Select-EbiWorklistRows -Rows $rows -Field 'status' -PendingWhen 'always' -Spec $null -Only @('r1', 'r3') -KeyColumns @('id')
Assert-Equal 'r1,r3' (@($r['rows'] | ForEach-Object { $_['id'] }) -join ',') 'select: Only keeps the named keys'
$r = Select-EbiWorklistRows -Rows $rows -Field 'status' -PendingWhen 'always' -Spec $null
Assert-True ([object]::ReferenceEquals($r['rows'][0], $rows[0])) 'select: the same row objects come back (a write through them reaches the table)'
$r = Select-EbiWorklistRows -Rows @() -Field 'status' -PendingWhen 'always' -Spec $null
Assert-True ($r['ok'] -and $r['selected'] -eq 0) 'select: an empty table selects nothing, no error'

# ---------------------------------------------------------------- ordering
$grouped = @(
    @{ k = 'c'; g = 'G2' }, @{ k = 'a'; g = 'G1' }, @{ k = 'd'; g = '' }, @{ k = 'b'; g = 'G1' }, @{ k = 'e'; g = 'G2' }
)
function Ord { param($Res) return (@($Res | ForEach-Object { $_['k'] }) -join ',') }
Assert-Equal 'c,a,d,b,e' (Ord (Sort-EbiWorklistRows -Rows $grouped)) 'sort: nothing asked -> table order'
Assert-Equal 'a,b,c,e,d' (Ord (Sort-EbiWorklistRows -Rows $grouped -GroupBy 'g')) 'sort: groupBy makes groups contiguous, stable inside, blank group last'
Assert-Equal 'a,b,c,e,d' (Ord (Sort-EbiWorklistRows -Rows $grouped -GroupBy 'g' -OrderBy 'zz')) 'sort: an absent orderBy column keeps the table order within the group'
Assert-Equal 'a,b,c,d,e' (Ord (Sort-EbiWorklistRows -Rows $grouped -OrderBy 'key' -KeyColumns @('k'))) 'sort: orderBy key sorts by the key display'
Assert-Equal 'a,b,c,e,d' (Ord (Sort-EbiWorklistRows -Rows $grouped -GroupBy 'g' -OrderBy 'key' -KeyColumns @('k'))) 'sort: groupBy then key'
Assert-Equal 'G1' (Get-EbiWorklistGroupOf -Row $grouped[1] -GroupBy 'g') 'group of: the column value'
Assert-Equal '' (Get-EbiWorklistGroupOf -Row $grouped[1] -GroupBy '') 'group of: no groupBy -> blank'

# ---------------------------------------------------------------- profile readers
$profile = @{ worklist = @{ key = @{ columns = @('A', 'B') }; columns = @( @{ name = 'A' }, $verdict ) }; vocabulary = @{ columns = @{ group = 'JOB' } } }
Assert-Equal 'A,B' ((Get-EbiWorklistKeyColumns -Profile $profile) -join ',') 'profile: key columns'
Assert-Equal 0 @(Get-EbiWorklistKeyColumns -Profile @{}).Count 'profile: no worklist -> no key columns (no throw)'
Assert-Equal 'JOB' (Get-EbiWorklistGroupColumn -Profile $profile) 'profile: group column'
Assert-Equal '' (Get-EbiWorklistGroupColumn -Profile $null) 'profile: null profile -> blank'
Assert-Equal 'status' (Get-EbiWorklistColumnSpec -Profile $profile -Field 'status')['name'] 'profile: a column spec by name'
Assert-True ($null -eq (Get-EbiWorklistColumnSpec -Profile $profile -Field 'nope')) 'profile: an unknown column is null'

$rc = Complete-Tests
exit $rc
