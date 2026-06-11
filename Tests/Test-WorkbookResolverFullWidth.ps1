#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'WorkbookResolver.ps1')

Reset-Tests 'WorkbookResolver full-width fallback'

$root = Join-Path ([System.IO.Path]::GetTempPath()) ('WorkbookResolverFullWidth_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
try {
    $fwChar = [char]([int][char]'0' + 0xFEE0)
    $fwStem = 'CJRVWD5' + $fwChar
    $fullWidth = Join-Path $root ("J4baseline_REQ-000xxxxx_{0}.xlsx" -f $fwStem)
    Set-Content -LiteralPath $fullWidth -Value 'full-width' -Encoding UTF8

    $genericFullWidth = Join-Path $root ('report' + $fwChar + '.txt')
    Set-Content -LiteralPath $genericFullWidth -Value 'generic full-width' -Encoding UTF8

    Assert-Equal $fullWidth (Find-WorkbookByExcelName -Dir $root -ExcelName 'CJRVWD50' -FullWidthFallback Accept) 'accepts normalized full-width ASCII workbook candidate'
    Assert-Equal $null (Find-WorkbookByExcelName -Dir $root -ExcelName 'CJRVWD50' -FullWidthFallback Reject) 'rejecting full-width fallback keeps not-found behavior'
    Assert-Equal $genericFullWidth (Resolve-FullWidthFileName -Dir $root -Name 'report0.txt' -Filter '*.txt' -FullWidthFallback Accept) 'generic resolver accepts normalized full-width filename candidate'
    Assert-Equal 'CJRVWD50' ([FullWidthFilenameResolver]::ConvertFullWidthAsciiToHalfWidth($fwStem)) 'class normalizes full-width ASCII to half-width'
    Assert-Equal 'CJRVWD50' (Convert-FullWidthAsciiToHalfWidth $fwStem) 'wrapper normalizes full-width ASCII to half-width'
    Assert-True ([FullWidthFilenameResolver]::ContainsFullWidthAscii($fwStem)) 'class detects full-width ASCII characters'
    Assert-True (Test-ContainsFullWidthAscii $fwStem) 'wrapper detects full-width ASCII characters'
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

exit (Complete-Tests)
