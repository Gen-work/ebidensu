#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'MappingStore.ps1')

Reset-Tests 'MappingStore'

# -- ConvertTo-TargetIdList --
Assert-Equal 'a|b|c' ((ConvertTo-TargetIdList @('a','b','c')) -join '|') 'array input'
Assert-Equal 'a|b|c' ((ConvertTo-TargetIdList 'a,b,c') -join '|')        'comma string'
Assert-Equal 'a|b'   ((ConvertTo-TargetIdList ' a , , b ') -join '|')    'trim + drop empty'
Assert-Equal 'x|y|z' ((ConvertTo-TargetIdList @('x','y,z')) -join '|')   'mixed array + comma'
Assert-Equal 0       (ConvertTo-TargetIdList $null).Count               'null -> empty'

# -- Ensure-MappingColumns --
$rows = @([pscustomobject]@{ Correl_ID_S = 'JIDSF48S'; GIFT_HM_snap = '1' })
Ensure-MappingColumns -Rows $rows | Out-Null
Assert-True ($rows[0].PSObject.Properties.Name -contains 'isMarked')        'adds missing isMarked'
Assert-True ($rows[0].PSObject.Properties.Name -contains 'ReviewComment')   'adds missing ReviewComment'
Assert-Equal '0' $rows[0].isMarked        'default isMarked = 0'
Assert-Equal '1' $rows[0].GIFT_HM_snap    'existing value preserved'

# -- Test-TargetRow --
$r = [pscustomobject]@{ Correl_ID_S='JIDSF48S'; Correl_ID_M='JIDSF48M'; JOB_NAME='JIDSJ48S'; Excel_NAME='JIDSW48S' }
Assert-True (Test-TargetRow $r @())                    'no targets -> match all'
Assert-True (Test-TargetRow $r @('JIDSF48S'))          'match Correl_ID_S'
Assert-True (Test-TargetRow $r @('JIDSJ48S'))          'match JOB_NAME'
Assert-True (-not (Test-TargetRow $r @('NOPE')))       'no match'

# -- Get-PendingRows (snap-style) --
$snapRows = @(
    [pscustomobject]@{ Correl_ID_S='A'; GIFT_HM_snap='1' },
    [pscustomobject]@{ Correl_ID_S='B'; GIFT_HM_snap='0' },
    [pscustomobject]@{ Correl_ID_S='C'; GIFT_HM_snap=''  }
)
$pend = @(Get-PendingRows -Rows $snapRows -Field 'GIFT_HM_snap')
Assert-Equal 'B|C' (($pend | ForEach-Object { $_.Correl_ID_S }) -join '|') 'pending = 0/empty only'
$pendF = @(Get-PendingRows -Rows $snapRows -Field 'GIFT_HM_snap' -Force $true)
Assert-Equal 3 $pendF.Count 'force -> all pending'
$pendT = @(Get-PendingRows -Rows $snapRows -Field 'GIFT_HM_snap' -Force $true -Targets @('A'))
Assert-Equal 'A' (($pendT | ForEach-Object { $_.Correl_ID_S }) -join '|') 'target filter applied'

# -- Get-PendingRows (bitmask) --
$bitRows = @(
    [pscustomobject]@{ Correl_ID_S='A'; isReplaced='1' },  # gift done
    [pscustomobject]@{ Correl_ID_S='B'; isReplaced='3' },  # gift+gfix done
    [pscustomobject]@{ Correl_ID_S='C'; isReplaced='0' }
)
$pendGfix = @(Get-PendingRows -Rows $bitRows -Field 'isReplaced' -Bit 2)
Assert-Equal 'A|C' (($pendGfix | ForEach-Object { $_.Correl_ID_S }) -join '|') 'bit 2 pending = A,C'

# -- Set-MappingBit --
$br = [pscustomobject]@{ isReplaced='1' }
Set-MappingBit -Row $br -Field 'isReplaced' -Bit 2
Assert-Equal '3' $br.isReplaced 'OR bit 2 into 1 -> 3'
Set-MappingBit -Row $br -Field 'isReplaced' -Bit 2
Assert-Equal '3' $br.isReplaced 'OR is idempotent'

# -- Export-MappingAtomic round-trip --
$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ('mapstore_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
try {
    $csv = Join-Path $tmpDir 'mapping_test.csv'
    $w = @(
        [pscustomobject]@{ Correl_ID_S='JIDSF48S'; JOB_NAME='JIDSJ48S'; GIFT_HM_snap='1' },
        [pscustomobject]@{ Correl_ID_S='JIGPF05S'; JOB_NAME='JIGPJ05S'; GIFT_HM_snap='0' }
    )
    Export-MappingAtomic -Rows $w -Path $csv | Out-Null
    Assert-True (Test-Path -LiteralPath $csv) 'atomic write created file'
    $back = Import-Mapping $csv
    Assert-Equal 2 $back.Count 'round-trip row count'
    Assert-Equal 'JIDSF48S' $back[0].Correl_ID_S 'round-trip value'
    # no leftover temp files
    $leftovers = @(Get-ChildItem -LiteralPath $tmpDir -Filter '*.tmp.*' -ErrorAction SilentlyContinue)
    Assert-Equal 0 $leftovers.Count 'no leftover temp file after success'
} finally {
    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}

exit (Complete-Tests)
