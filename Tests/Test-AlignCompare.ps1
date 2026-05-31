#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'AlignCompare.ps1')

Reset-Tests 'AlignCompare'

# Compare-SheetGrid: identical
$r = Compare-SheetGrid -RowsA 2 -ColsA 2 -FlatA @('a','b','c','d') -RowsB 2 -ColsB 2 -FlatB @('a','b','c','d')
Assert-True $r.Same 'identical grids are Same'

# value diff
$r = Compare-SheetGrid -RowsA 2 -ColsA 2 -FlatA @('a','b','c','d') -RowsB 2 -ColsB 2 -FlatB @('a','X','c','d')
Assert-True (-not $r.Same) 'value diff -> not Same'
Assert-True ($r.Reason -match 'value differs') 'value diff reason'

# dimension diff
$r = Compare-SheetGrid -RowsA 2 -ColsA 2 -FlatA @('a','b','c','d') -RowsB 3 -ColsB 2 -FlatB @('a','b','c','d','e','f')
Assert-True (-not $r.Same) 'dim diff -> not Same'
Assert-True ($r.Reason -match 'dimensions differ') 'dim diff reason'

# Get-MigrationType
Assert-Equal 'Unknown'    (Get-MigrationType -FromSys 'X' -ToSys 'Y' -HostTypes @())            'empty host types -> Unknown'
Assert-Equal 'HostToOpen' (Get-MigrationType -FromSys 'HOST' -ToSys 'IGP' -HostTypes @('HOST')) 'host->open'
Assert-Equal 'OpenToOpen' (Get-MigrationType -FromSys 'IDS' -ToSys 'IGP' -HostTypes @('HOST'))  'open->open'
Assert-Equal 'OpenToHost' (Get-MigrationType -FromSys 'IDS' -ToSys 'HOST' -HostTypes @('HOST')) 'open->host'

# Get-AlignSheetsForMigration (send has 5, recv has 3)
$send = @('S1','S2','S3','S4','S5')
$recv = @('R1','R2','R3')
Assert-Equal 'R1|R2|R3' ((Get-AlignSheetsForMigration -MigrationType 'HostToOpen' -SendSheets $send -RecvSheets $recv) -join '|') 'HostToOpen = 3 recv'
Assert-Equal 'S3|S4|R1|R2|R3' ((Get-AlignSheetsForMigration -MigrationType 'OpenToOpen' -SendSheets $send -RecvSheets $recv) -join '|') 'OpenToOpen = S3,S4 + recv'
Assert-Equal 'R1|R2|R3' ((Get-AlignSheetsForMigration -MigrationType 'Unknown' -SendSheets $send -RecvSheets $recv) -join '|') 'Unknown -> recv only (safe)'

exit (Complete-Tests)
