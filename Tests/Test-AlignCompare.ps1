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
Assert-Equal 'S1|S3|S4' ((Get-AlignSheetsForMigration -MigrationType 'HostToOpen' -SendSheets $send -RecvSheets $recv) -join '|') 'HostToOpen = send data + GIFT/GFIX send results'
Assert-Equal 'S3|S4|R1|R2|R3' ((Get-AlignSheetsForMigration -MigrationType 'OpenToOpen' -SendSheets $send -RecvSheets $recv) -join '|') 'OpenToOpen = S3,S4 + recv'
Assert-Equal 'R1|R2|R3' ((Get-AlignSheetsForMigration -MigrationType 'Unknown' -SendSheets $send -RecvSheets $recv) -join '|') 'Unknown -> recv only (safe)'

# Compare-AlignSheet: equal grid + equal picture count -> Same
$r = Compare-AlignSheet -RowsA 1 -ColsA 1 -FlatA @('') -PicsA 1 -RowsB 1 -ColsB 1 -FlatB @('') -PicsB 1
Assert-True $r.Same 'equal grid + equal pics -> Same'

# Compare-AlignSheet: picture-only sheet, work has no pic yet -> diff (so it syncs)
$r = Compare-AlignSheet -RowsA 1 -ColsA 1 -FlatA @('') -PicsA 0 -RowsB 1 -ColsB 1 -FlatB @('') -PicsB 1
Assert-True (-not $r.Same) 'picture count diff -> not Same'
Assert-True ($r.Reason -match 'picture count differs') 'picture count diff reason'

# Get-AlignSheetKind
Assert-Equal 'SendData'   (Get-AlignSheetKind -SheetName 'S1' -SendDataName 'S1' -SendResultNames @('S3','S4')) 'send-data sheet'
Assert-Equal 'SendResult' (Get-AlignSheetKind -SheetName 'S4' -SendDataName 'S1' -SendResultNames @('S3','S4')) 'send-result sheet'
Assert-Equal 'Other'      (Get-AlignSheetKind -SheetName 'R1' -SendDataName 'S1' -SendResultNames @('S3','S4')) 'other sheet'

# Test-J4SheetPrepared: send-data needs a picture
Assert-True (Test-J4SheetPrepared -Kind 'SendData' -PictureCount 1 -TextRowCount 0).Prepared        'send-data with picture -> prepared'
Assert-True (-not (Test-J4SheetPrepared -Kind 'SendData' -PictureCount 0 -TextRowCount 9).Prepared) 'send-data without picture -> not prepared'

# Test-J4SheetPrepared: send-result needs > MinTextRows rows
Assert-True (Test-J4SheetPrepared -Kind 'SendResult' -PictureCount 0 -TextRowCount 4 -MinTextRows 3).Prepared        'send-result with 4 rows -> prepared'
Assert-True (-not (Test-J4SheetPrepared -Kind 'SendResult' -PictureCount 0 -TextRowCount 3 -MinTextRows 3).Prepared) 'send-result with 3 rows -> not prepared'
Assert-True (-not (Test-J4SheetPrepared -Kind 'SendResult' -PictureCount 0 -TextRowCount 0 -MinTextRows 3).Prepared) 'send-result empty -> not prepared'

# Other sheets are always prepared (no emptiness rule)
Assert-True (Test-J4SheetPrepared -Kind 'Other' -PictureCount 0 -TextRowCount 0).Prepared 'other sheet -> always prepared'

exit (Complete-Tests)
