# ============================================================
#  WorkbookResolver.ps1
#
#  Dot-source helper for resolving evidence/J4 workbook filenames from
#  mapping Excel_NAME values. Source files can have a prefix plus a final _<Excel_NAME>.xlsx
#  suffix while mapping stores only Excel_NAME.
# ============================================================

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
