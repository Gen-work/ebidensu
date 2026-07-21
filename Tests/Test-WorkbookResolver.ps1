#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'WorkbookResolver.ps1')

Reset-Tests 'WorkbookResolver'

$root = Join-Path ([System.IO.Path]::GetTempPath()) ('WorkbookResolver_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
try {
    $exact = Join-Path $root 'CJRVWD50.xlsx'
    Set-Content -LiteralPath $exact -Value 'exact' -Encoding UTF8
    Assert-Equal $exact (Find-WorkbookByExcelName -Dir $root -ExcelName 'CJRVWD50') 'prefers exact Excel_NAME.xlsx'

    Remove-Item -LiteralPath $exact -Force
    $prefixed = Join-Path $root 'J4baseline_REQ-000xxxxx_CJRVWD50.xlsx'
    Set-Content -LiteralPath $prefixed -Value 'prefixed' -Encoding UTF8
    Assert-Equal $prefixed (Find-WorkbookByExcelName -Dir $root -ExcelName 'CJRVWD50') 'finds prefixed suffix workbook'

    $sub = Join-Path $root 'JRV'
    New-Item -ItemType Directory -Path $sub -Force | Out-Null
    $nested = Join-Path $sub 'J4baseline_REQ-000xxxxx_LJRVWD64.xlsx'
    Set-Content -LiteralPath $nested -Value 'nested' -Encoding UTF8
    Assert-Equal $null (Find-WorkbookByExcelName -Dir $root -ExcelName 'LJRVWD64') 'does not recurse by default'
    Assert-Equal $nested (Find-WorkbookByExcelName -Dir $root -ExcelName 'LJRVWD64' -Recurse) 'recurse finds nested suffix workbook'

    # Get-PrefixFromFilename: inverse of Get-ExcelFullStem
    Assert-Equal 'J4baseline_REQ-000xxxxx' (Get-PrefixFromFilename -FileName 'J4baseline_REQ-000xxxxx_CJRVWD50.xlsx' -Name 'CJRVWD50') 'recovers prefix before _Excel_NAME'
    Assert-Equal '' (Get-PrefixFromFilename -FileName 'CJRVWD50.xlsx' -Name 'CJRVWD50') 'no prefix when filename equals Excel_NAME'
    Assert-Equal '' (Get-PrefixFromFilename -FileName 'somethingCJRVWD50.xlsx' -Name 'CJRVWD50') 'no prefix without underscore separator'
    Assert-Equal '' (Get-PrefixFromFilename -FileName '' -Name 'CJRVWD50') 'empty filename -> empty prefix'
    # round-trips with Get-ExcelFullStem
    $rtPrefix = 'J4xyz(REQ-00012345_GIFT)'
    $rtStem   = Get-ExcelFullStem -Prefix $rtPrefix -Name 'LJRVWD64'
    Assert-Equal $rtPrefix (Get-PrefixFromFilename -FileName ("{0}.xlsx" -f $rtStem) -Name 'LJRVWD64') 'round-trips Get-ExcelFullStem prefix'

    # Resolve-ExcelPrefix: project config default, legacy row override.
    $rowNoPrefix = [pscustomobject]@{ Excel_NAME = 'CJRVWD50' }
    $rowOverride = [pscustomobject]@{ Excel_NAME = 'CJRVWD50'; Excel_Prefix = 'RowPrefix' }
    Assert-Equal 'ProjectPrefix' (Resolve-ExcelPrefix -Row $rowNoPrefix -DefaultPrefix 'ProjectPrefix') 'uses project-level default prefix'
    Assert-Equal 'RowPrefix' (Resolve-ExcelPrefix -Row $rowOverride -DefaultPrefix 'ProjectPrefix') 'legacy row Excel_Prefix overrides default'

    # Resolve-ExcelPrefixWithDisk: on-disk recovery only when row + default
    # are both blank; config always wins over the disk name.
    $evDir = Join-Path $root 'evidence'
    New-Item -ItemType Directory -Path $evDir -Force | Out-Null
    $evPrefixed = Join-Path $evDir ((Get-ExcelFullStem -Prefix 'J4disk(REQ-000xxxxx)' -Name 'KJODWWB5') + '.xlsx')
    Set-Content -LiteralPath $evPrefixed -Value 'ev' -Encoding UTF8
    $rowBare = [pscustomobject]@{ Excel_NAME = 'KJODWWB5' }
    Assert-Equal 'J4disk(REQ-000xxxxx)' (Resolve-ExcelPrefixWithDisk -Row $rowBare -ExcelName 'KJODWWB5' -EvidenceDir $evDir) 'recovers prefix from the on-disk evidence filename'
    Assert-Equal 'ProjectPrefix' (Resolve-ExcelPrefixWithDisk -Row $rowBare -DefaultPrefix 'ProjectPrefix' -ExcelName 'KJODWWB5' -EvidenceDir $evDir) 'configured prefix wins over the on-disk name'
    Assert-Equal '' ([string](Resolve-ExcelPrefixWithDisk -Row $rowBare -ExcelName 'KJODWWB5' -EvidenceDir '')) 'no EvidenceDir -> empty prefix, no throw'
    Assert-Equal '' ([string](Resolve-ExcelPrefixWithDisk -Row $rowBare -ExcelName 'NOSUCH99' -EvidenceDir $evDir)) 'no on-disk file -> empty prefix'
    Remove-Item -LiteralPath $evPrefixed -Force
    $evBare = Join-Path $evDir 'KJODWWB5.xlsx'
    Set-Content -LiteralPath $evBare -Value 'ev' -Encoding UTF8
    Assert-Equal '' ([string](Resolve-ExcelPrefixWithDisk -Row $rowBare -ExcelName 'KJODWWB5' -EvidenceDir $evDir)) 'bare on-disk filename -> empty prefix'
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

exit (Complete-Tests)
