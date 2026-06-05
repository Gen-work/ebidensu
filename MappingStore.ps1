# ============================================================
#  MappingStore.ps1
#
#  Single source of truth for reading/writing mapping_<Owner>.csv.
#  Dot-source only -- this file has NO param() block, so it is safe
#  under the project's dot-source rule (see CLAUDE.md).
#
#  Why this exists:
#    JenkinsSnap / DfSnap / GfixLogDownload / ReplaceEvidence each used
#    to hand-roll Import-Csv/Export-Csv + per-row re-read/re-write. That
#    drifted (TargetIds parsing, missing columns, full-file rewrite per
#    row, silent write failures when Excel held the CSV open). Everyone
#    now funnels through these helpers.
#
#  Encoding:
#    CSV is read/written with -Encoding UTF8. Under PS 5.1 that means
#    UTF-8 *with BOM*, which is exactly what Excel needs to open the
#    Japanese columns without mojibake. (The .ps1/.psd1/.jsonl files stay
#    BOM-less; only the CSV carries a BOM -- see Check-Encoding.ps1.)
#
#  Source stays ASCII-only on purpose: no raw Japanese literals here, so
#  the no-BOM .ps1 cannot be mis-decoded by PS 5.1 on a JP-locale host.
# ============================================================

# Status columns every mapping should carry. Identity/data columns
# (Correl_ID_*, JOB_NAME, Excel_NAME, FROM_*, TO_*, IF, Amount, ...) are
# produced by Generate-HostOpenMapping.ps1 and are NEVER defaulted here.
# Excel_Prefix is the J4 filename prefix entered by the operator
# (e.g. J4<title>(REQ-000xxxxx_GIFT<suffix>)).  Full evidence filename =
# "{Excel_Prefix}_{Excel_NAME}.xlsx" when prefix is set, else "{Excel_NAME}.xlsx".
function Get-MappingStatusColumns {
    return @(
        @{ Name = 'Excel_Prefix';         Default = '' },
        @{ Name = 'Excel_snap';           Default = '0' },
        @{ Name = 'GIFT_HM_snap';         Default = '0' },
        @{ Name = 'GIFT_MQ_snap';         Default = '0' },
        @{ Name = 'GIFT_Jenkins_snap';    Default = '0' },
        @{ Name = 'GIFT_noGfixfile_snap'; Default = '0' },
        @{ Name = 'GFIX_HM_snap';         Default = '0' },
        @{ Name = 'GFIX_Jenkins_snap';    Default = '0' },
        @{ Name = 'GFIX_log';             Default = '0' },
        @{ Name = 'DF_snap';              Default = '0' },
        @{ Name = 'isReplaced';           Default = '0' },
        @{ Name = 'isMarked';             Default = '0' },
        @{ Name = 'isReviewed';           Default = '0' },
        @{ Name = 'ReviewComment';        Default = '' },
        @{ Name = 'isDelivered';          Default = '0' },
        @{ Name = 'DeliverComment';       Default = '' },
        @{ Name = 'isFilesDelivered';     Default = '0' }
    )
}

# StrictMode-safe property read. Returns '' for missing/null.
function Get-RowProp {
    param($Row, [string]$Name)
    if ($null -eq $Row) { return '' }
    if ($Row.PSObject.Properties.Name -contains $Name) {
        $v = $Row.$Name
        if ($null -eq $v) { return '' }
        return [string]$v
    }
    return ''
}

# Accepts: string[]  /  "a,b,c"  /  "a"  /  @("a","b,c").
# Splits on comma, trims, drops blanks. Always returns string[].
function ConvertTo-TargetIdList {
    param([object]$TargetIds)
    $out = [System.Collections.Generic.List[string]]::new()
    if ($null -eq $TargetIds) { return ,@() }
    foreach ($item in @($TargetIds)) {
        if ($null -eq $item) { continue }
        foreach ($part in ([string]$item -split ',')) {
            $v = $part.Trim()
            if (-not [string]::IsNullOrWhiteSpace($v)) { $out.Add($v) }
        }
    }
    return $out.ToArray()
}

# True when no targets given, or row matches any target on
# Correl_ID_S / Correl_ID_M / JOB_NAME / Excel_NAME.
function Test-TargetRow {
    param($Row, [string[]]$Targets)
    if ($null -eq $Targets -or @($Targets).Count -eq 0) { return $true }
    $ids = @(
        (Get-RowProp $Row 'Correl_ID_S'),
        (Get-RowProp $Row 'Correl_ID_M'),
        (Get-RowProp $Row 'JOB_NAME'),
        (Get-RowProp $Row 'Excel_NAME')
    )
    foreach ($t in $Targets) {
        if ($ids -contains $t) { return $true }
    }
    return $false
}

function Import-Mapping {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Mapping not found: $Path" }
    return @(Import-Csv -LiteralPath $Path -Encoding UTF8)
}

# Adds any missing status columns (and optional caller -Extra columns)
# with their defaults. Mutates the rows in place; also returns them.
function Ensure-MappingColumns {
    param([object[]]$Rows, [object[]]$Extra = @())
    $cols = @(Get-MappingStatusColumns)
    foreach ($e in @($Extra)) { if ($null -ne $e) { $cols += $e } }
    foreach ($r in @($Rows)) {
        foreach ($c in $cols) {
            if (-not ($r.PSObject.Properties.Name -contains $c.Name)) {
                $r | Add-Member -NotePropertyName $c.Name -NotePropertyValue $c.Default -Force
            }
        }
    }
    return $Rows
}

function Test-BitDone {
    param([string]$Value, [int]$Bit)
    $v = 0; try { $v = [int]$Value } catch { $v = 0 }
    return (($v -band $Bit) -eq $Bit)
}

# A snap-style field is "done" when it is non-empty and not '0'.
function Test-SnapDone {
    param([string]$Value)
    return ((-not [string]::IsNullOrEmpty($Value)) -and ($Value -ne '0'))
}

# Filters rows by target, then returns those still pending for -Field.
# -Bit > 0 treats the field as a bitmask (done = (val -band Bit) -eq Bit);
# otherwise snap-style (done = non-empty and not '0'). -Force = all targets.
function Get-PendingRows {
    param(
        [object[]]$Rows,
        [string]$Field,
        [bool]$Force = $false,
        [string[]]$Targets = @(),
        [int]$Bit = 0
    )
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($r in @($Rows)) {
        if (-not (Test-TargetRow $r $Targets)) { continue }
        $cur = Get-RowProp $r $Field
        if ($Bit -gt 0) { $done = Test-BitDone $cur $Bit }
        else            { $done = Test-SnapDone $cur }
        if ($Force -or -not $done) { $out.Add($r) }
    }
    return $out.ToArray()
}

# Applies -Updates (field -> value) to every row whose -KeyField equals
# -KeyValue. Adds the column if absent. Returns the count updated.
function Update-MappingRows {
    param(
        [object[]]$Rows,
        [string]$KeyField,
        [string]$KeyValue,
        [hashtable]$Updates
    )
    $n = 0
    foreach ($r in @($Rows)) {
        if ((Get-RowProp $r $KeyField) -eq $KeyValue) {
            foreach ($k in $Updates.Keys) {
                if (-not ($r.PSObject.Properties.Name -contains $k)) {
                    $r | Add-Member -NotePropertyName $k -NotePropertyValue ([string]$Updates[$k]) -Force
                } else {
                    $r.$k = [string]$Updates[$k]
                }
            }
            $n++
        }
    }
    return $n
}

# OR a bit into a single row's bitmask field (creates the column if absent).
function Set-MappingBit {
    param([object]$Row, [string]$Field, [int]$Bit)
    if (-not ($Row.PSObject.Properties.Name -contains $Field)) {
        $Row | Add-Member -NotePropertyName $Field -NotePropertyValue '0' -Force
    }
    $cur = 0; try { $cur = [int]$Row.$Field } catch { $cur = 0 }
    $Row.$Field = [string]($cur -bor $Bit)
}

# Atomic CSV write: serialise to a temp file in the same directory, then
# Move-Item -Force over the target. If the target is locked (e.g. Excel has
# it open) the move is retried with exponential backoff and then throws a
# clear, actionable error -- it NEVER fails silently. On final failure the
# temp file is kept so no data is lost.
function Export-MappingAtomic {
    param(
        [object[]]$Rows,
        [string]$Path,
        [int]$Retries = 5,
        [int]$BaseDelayMs = 300
    )
    if ($null -eq $Rows) { throw 'Export-MappingAtomic: -Rows is null' }
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'Export-MappingAtomic: -Path is empty' }

    $dir = Split-Path -Parent $Path
    if ([string]::IsNullOrEmpty($dir)) { $dir = '.' }
    $tmp = Join-Path $dir ('.' + (Split-Path -Leaf $Path) + '.tmp.' + $PID)

    @($Rows) | Export-Csv -LiteralPath $tmp -Encoding UTF8 -NoTypeInformation -Force

    $lastErr = $null
    for ($i = 0; $i -lt $Retries; $i++) {
        try {
            Move-Item -LiteralPath $tmp -Destination $Path -Force
            return $true
        } catch {
            $lastErr = $_
            if ($i -lt ($Retries - 1)) {
                Start-Sleep -Milliseconds ([int]($BaseDelayMs * [Math]::Pow(2, $i)))
            }
        }
    }
    throw ("Could not write mapping '{0}' after {1} tries (is it open in Excel/SAKURA?). " +
           "Your data is safe in the temp file '{2}'. Close the program and rename it back. " +
           "Last error: {3}") -f $Path, $Retries, $tmp, $lastErr
}
