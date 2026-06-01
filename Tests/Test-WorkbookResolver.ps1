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
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

exit (Complete-Tests)
