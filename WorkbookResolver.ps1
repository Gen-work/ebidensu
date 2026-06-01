# ============================================================
#  WorkbookResolver.ps1
#
#  Dot-source helper for resolving evidence/J4 workbook filenames from
#  mapping Excel_NAME values.  ASCII source -- no raw Japanese.
#
#  Excel_NAME supports three forms:
#    Full stem   : J4<title>(REQ-000xxxxx_GIFT<suffix>)_LJRVWD64
#                  -> exact match <full>.xlsx; template output also uses full name
#    Suffix form : _LJRVWD64   (leading _ = suffix-only shorthand)
#                  -> strips leading _ for wildcard; template output -> LJRVWD64.xlsx
#    Short stem  : LJRVWD64    (legacy/plain, existing behaviour unchanged)
#                  -> exact LJRVWD64.xlsx or wildcard *_LJRVWD64.xlsx
# ============================================================

function Find-WorkbookByExcelName([string]$Dir, [string]$ExcelName, [switch]$Recurse) {
    if ([string]::IsNullOrWhiteSpace($Dir) -or [string]::IsNullOrWhiteSpace($ExcelName) -or -not (Test-Path -LiteralPath $Dir)) { return $null }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($ExcelName)
    $leaf = if ($ExcelName -match '\.xlsx$') { $ExcelName } else { ("{0}.xlsx" -f $stem) }

    # 1) Exact match (works for full-name and short-stem forms)
    $exact = Join-Path $Dir $leaf
    if (Test-Path -LiteralPath $exact) { return (Resolve-Path -LiteralPath $exact).Path }

    # 2) For suffix form (_XYZ), strip leading _ and try exact without prefix
    $searchStem = if ($stem -match '^_(.+)$') { $Matches[1] } else { $stem }
    if ($searchStem -ne $stem) {
        $exact2 = Join-Path $Dir ("{0}.xlsx" -f $searchStem)
        if (Test-Path -LiteralPath $exact2) { return (Resolve-Path -LiteralPath $exact2).Path }
    }

    # 3) Wildcard: *_<searchStem>.xlsx then *<searchStem>.xlsx
    #    Uses searchStem (leading _ already stripped) to avoid double __ in the pattern.
    $hits = @()
    foreach ($pattern in @(("*_{0}.xlsx" -f $searchStem), ("*{0}.xlsx" -f $searchStem))) {
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

# Returns the filename to create when cloning a template.
#   Full stem  (no leading _) : "<full stem>.xlsx"  (includes any J4 prefix)
#   Suffix form (_LJRVWD64)   : "LJRVWD64.xlsx"     (leading _ stripped)
#   Short stem  (LJRVWD64)    : "LJRVWD64.xlsx"     (unchanged)
function Get-ExcelDestLeaf {
    param([string]$ExcelName)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($ExcelName)
    if ($stem -match '^_(.+)$') { $stem = $Matches[1] }
    return "{0}.xlsx" -f $stem
}
