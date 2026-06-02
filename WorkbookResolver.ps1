# ============================================================
#  WorkbookResolver.ps1
#
#  Dot-source helper for resolving evidence/J4 workbook filenames.
#  ASCII source -- no raw Japanese literals.
#
#  Usage pattern in callers:
#    $fullStem = Get-ExcelFullStem -Prefix ([string]$first.Excel_Prefix) -Name ([string]$first.Excel_NAME)
#    $wbPath   = Find-WorkbookByExcelName -Dir $evDir -ExcelName $fullStem
#    $destLeaf = Get-ExcelDestLeaf $fullStem
#
#  When Excel_Prefix is empty the full stem equals Excel_NAME (legacy behaviour).
# ============================================================

# Combines the J4 prefix column with the short name column.
#   Excel_Prefix = 'J4<title>(REQ-000xxxxx_GIFT<suffix>)'  (or '')
#   Excel_NAME   = 'LJRVWD64'
#   -> 'J4<title>(REQ-000xxxxx_GIFT<suffix>)_LJRVWD64'    (when prefix set)
#   -> 'LJRVWD64'                                          (when prefix empty)
function Get-ExcelFullStem {
    param([string]$Prefix, [string]$Name)
    $n = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    if ([string]::IsNullOrWhiteSpace($Prefix)) { return $n }
    return "{0}_{1}" -f $Prefix.TrimEnd('_'), $n
}

# Returns the filename to create when cloning from a template.
#   Get-ExcelDestLeaf 'J4..._LJRVWD64'  -> 'J4..._LJRVWD64.xlsx'
#   Get-ExcelDestLeaf 'LJRVWD64'        -> 'LJRVWD64.xlsx'
function Get-ExcelDestLeaf {
    param([string]$FullStem)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($FullStem)
    return "{0}.xlsx" -f $stem
}

# Inverse of Get-ExcelFullStem: given a real workbook filename and the short
# Excel_NAME, recover the J4 prefix that precedes "_<Excel_NAME>".
#   Get-PrefixFromFilename 'J4title(REQ-...)_LJRVWD64.xlsx' 'LJRVWD64'
#       -> 'J4title(REQ-...)'
#   Get-PrefixFromFilename 'LJRVWD64.xlsx' 'LJRVWD64' -> ''   (no prefix)
# Returns '' when the filename does not carry a recoverable prefix.
function Get-PrefixFromFilename {
    param([string]$FileName, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($FileName) -or [string]::IsNullOrWhiteSpace($Name)) { return '' }
    $stem  = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $short = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    if ($stem -eq $short) { return '' }
    $suffix = '_' + $short
    if ($stem.EndsWith($suffix)) {
        return $stem.Substring(0, $stem.Length - $suffix.Length)
    }
    return ''
}

# Finds the evidence/J4 workbook file for a given stem (full or short).
# Search order:
#   1. Exact match: <Dir>\<FullStem>.xlsx
#   2. Wildcard:    <Dir>\*_<FullStem>.xlsx  (handles any extra prefix on disk)
#   3. Wildcard:    <Dir>\*<FullStem>.xlsx
# Returns newest LastWriteTime when multiple wildcard hits.
function Find-WorkbookByExcelName([string]$Dir, [string]$ExcelName, [switch]$Recurse) {
    if ([string]::IsNullOrWhiteSpace($Dir) -or [string]::IsNullOrWhiteSpace($ExcelName) -or -not (Test-Path -LiteralPath $Dir)) { return $null }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($ExcelName)
    $leaf = if ($ExcelName -match '\.xlsx$') { $ExcelName } else { ("{0}.xlsx" -f $stem) }
    $exact = Join-Path $Dir $leaf
    if (Test-Path -LiteralPath $exact) { return (Resolve-Path -LiteralPath $exact).Path }

    $hits = @()
    foreach ($pattern in @(("*_{0}.xlsx" -f $stem), ("*{0}.xlsx" -f $stem))) {
        if ($Recurse.IsPresent) {
            $hits += @(Get-ChildItem -LiteralPath $Dir -Filter $pattern -File -Recurse -ErrorAction SilentlyContinue)
        } else {
            $hits += @(Get-ChildItem -LiteralPath $Dir -Filter $pattern -File -ErrorAction SilentlyContinue)
        }
    }
    $hits = @($hits | Sort-Object FullName -Unique)
    if ($hits.Count -eq 0) { return $null }
    return ($hits | Sort-Object @{Expression='LastWriteTime';Descending=$true}, FullName | Select-Object -First 1).FullName
}
