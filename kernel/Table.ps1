#Requires -Version 5.1
# ============================================================
#  kernel/Table.ps1
#
#  CSV in and out for the worklist (P1-24), ported from MappingStore.ps1
#  Import-Mapping / Export-MappingAtomic. Dot-source only (no param(),
#  ASCII source, no class).
#
#    Read-EbiCsv  -Path            -> @{ ok; columns; rows; message; exists }
#                                     rows are hashtables column -> string,
#                                     columns in header order
#    Write-EbiCsvAtomic -Path -Columns -Rows
#                                  -> @{ ok; message; tmp }
#                                     UTF-8 WITH BOM (Excel opens it as
#                                     UTF-8 only with the BOM), CRLF, every
#                                     field quoted, temp file + Move with
#                                     retries (the CSV is often open in
#                                     Excel or an editor on the office PC)
#    ConvertTo-EbiCsvLine -Values   PURE, the quoting rule
#    New-EbiWorklist -Path -Columns -Rows
#                                  -> the Session resource of kind
#                                     'worklist' (kernel/Worklist.ps1)
#
#  Encoding is fixed here rather than by the Export-Csv -Encoding switch,
#  whose meaning of UTF8 differs between Windows PowerShell 5.1 (BOM) and
#  pwsh 7 (no BOM).
# ============================================================

function ConvertTo-EbiCsvLine {
    # PURE. Values -> one CSV line, every field quoted, quotes doubled.
    param($Values)
    $parts = New-Object System.Collections.ArrayList
    foreach ($v in @($Values)) {
        $s = if ($null -eq $v) { '' } else { [string]$v }
        [void]$parts.Add('"' + $s.Replace('"', '""') + '"')
    }
    return ($parts.ToArray() -join ',')
}

function ConvertTo-EbiCsvRowHashtable {
    # One Import-Csv object -> hashtable in header order.
    param($Obj, $Columns)
    $row = @{}
    foreach ($c in @($Columns)) {
        $name = [string]$c
        $v = $null
        $p = $Obj.PSObject.Properties[$name]
        if ($null -ne $p) { $v = $p.Value }
        $row[$name] = $(if ($null -eq $v) { '' } else { [string]$v })
    }
    return $row
}

function Read-EbiCsv {
    param([string]$Path)
    $empty = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($Path)) { return @{ ok = $false; columns = @(); rows = $empty.ToArray(); message = 'no path given'; exists = $false } }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{ ok = $false; columns = @(); rows = $empty.ToArray(); message = ('file not found: ' + $Path); exists = $false } }
    try {
        $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding($false)))   # a BOM is skipped by the decoder
        if ([string]::IsNullOrWhiteSpace($text)) { return @{ ok = $false; columns = @(); rows = $empty.ToArray(); message = ('empty file: ' + $Path); exists = $true } }
        $objs = @(ConvertFrom-Csv -InputObject $text)
        $columns = New-Object System.Collections.ArrayList
        $header = ($text -split "\r?\n")[0]
        $headObj = ConvertFrom-Csv -InputObject ($header + "`r`n" + $header)   # the header parsed as a row gives the names in order
        if ($null -ne $headObj) { foreach ($p in $headObj.PSObject.Properties) { [void]$columns.Add([string]$p.Name) } }
        $rows = New-Object System.Collections.ArrayList
        foreach ($o in $objs) { if ($null -ne $o) { [void]$rows.Add((ConvertTo-EbiCsvRowHashtable -Obj $o -Columns $columns.ToArray())) } }
        return @{ ok = $true; columns = $columns.ToArray(); rows = $rows.ToArray(); message = ''; exists = $true }
    } catch {
        return @{ ok = $false; columns = @(); rows = $empty.ToArray(); message = ('cannot read {0}: {1}' -f $Path, $_.Exception.Message); exists = $true }
    }
}

function Write-EbiCsvAtomic {
    param([string]$Path, $Columns, $Rows, [int]$Retries = 5, [int]$BaseDelayMs = 300)
    if ([string]::IsNullOrWhiteSpace($Path)) { return @{ ok = $false; message = 'no path given'; tmp = '' } }
    $cols = @(@($Columns) | ForEach-Object { [string]$_ })
    if ($cols.Count -eq 0) { return @{ ok = $false; message = 'no columns to write'; tmp = '' } }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((ConvertTo-EbiCsvLine -Values $cols)).Append("`r`n")
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $vals = New-Object System.Collections.ArrayList
        foreach ($c in $cols) { [void]$vals.Add($(if (($row -is [System.Collections.IDictionary]) -and $row.Contains($c) -and $null -ne $row[$c]) { [string]$row[$c] } else { '' })) }
        [void]$sb.Append((ConvertTo-EbiCsvLine -Values $vals.ToArray())).Append("`r`n")
    }
    $dir = Split-Path -Path $Path -Parent
    if ([string]::IsNullOrEmpty($dir)) { $dir = '.' }
    $tmp = Join-Path $dir ('.' + (Split-Path -Leaf $Path) + '.tmp.' + $PID)
    try {
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [System.IO.File]::WriteAllText($tmp, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))
    } catch { return @{ ok = $false; message = ('cannot write temp file {0}: {1}' -f $tmp, $_.Exception.Message); tmp = $tmp } }
    $lastErr = ''
    for ($i = 0; $i -lt [Math]::Max(1, $Retries); $i++) {
        try {
            Move-Item -LiteralPath $tmp -Destination $Path -Force -ErrorAction Stop
            return @{ ok = $true; message = ''; tmp = '' }
        } catch {
            $lastErr = $_.Exception.Message
            if ($i -lt ($Retries - 1)) { Start-Sleep -Milliseconds ([int]($BaseDelayMs * [Math]::Pow(2, $i))) }
        }
    }
    return @{ ok = $false; message = ("could not replace '{0}' after {1} tries (open in Excel?); the data is safe in '{2}': {3}" -f $Path, $Retries, $tmp, $lastErr); tmp = $tmp }
}

function New-EbiWorklist {
    param([string]$Path, $Columns, $Rows)
    return @{ path = $Path; columns = @(@($Columns) | ForEach-Object { [string]$_ }); rows = @($Rows) }
}

function Save-EbiWorklist {
    # The atomic flush every writing step does before returning (STEP-CONTRACT 3.2).
    param($Worklist, [string]$Path = '')
    $p = if ($Path -ne '') { $Path } else { [string]$Worklist['path'] }
    if ([string]::IsNullOrWhiteSpace($p)) { return @{ ok = $false; message = 'the worklist has no path (loaded from rows?) and none was given'; tmp = '' } }
    return (Write-EbiCsvAtomic -Path $p -Columns $Worklist['columns'] -Rows $Worklist['rows'])
}
