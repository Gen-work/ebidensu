# ============================================================
#  AlignCompare.ps1
#
#  PURE comparison + migration-type logic for the Align/Precheck phase
#  (spec 6) -- NO Excel COM, so it is unit-testable (Tests\Test-AlignCompare).
#  Align.ps1 reads each sheet's UsedRange via COM, flattens it to trimmed
#  row-major strings, and feeds it here.
#  Dot-source only (no param() block).
# ============================================================

# Compare two flattened (row-major, trimmed) sheet grids.
# Returns @{ Same; Reason }. Format differences are out of scope (spec 6:
# "format diffs can be TODO"); this compares dimensions + values only.
function Compare-SheetGrid {
    param(
        [int]$RowsA, [int]$ColsA, [string[]]$FlatA,
        [int]$RowsB, [int]$ColsB, [string[]]$FlatB
    )
    if ($RowsA -ne $RowsB -or $ColsA -ne $ColsB) {
        return [pscustomobject]@{ Same = $false; Reason = ("dimensions differ ({0}x{1} vs {2}x{3})" -f $RowsA, $ColsA, $RowsB, $ColsB) }
    }
    $a = @($FlatA); $b = @($FlatB)
    if ($a.Count -ne $b.Count) {
        return [pscustomobject]@{ Same = $false; Reason = ("cell count differs ({0} vs {1})" -f $a.Count, $b.Count) }
    }
    for ($i = 0; $i -lt $a.Count; $i++) {
        if ([string]$a[$i] -ne [string]$b[$i]) {
            return [pscustomobject]@{ Same = $false; Reason = ("value differs at cell #{0} ('{1}' vs '{2}')" -f $i, $a[$i], $b[$i]) }
        }
    }
    return [pscustomobject]@{ Same = $true; Reason = '' }
}

# Classify the migration from system-type values. -HostTypes is the set of
# system-type literals that count as "Host" (the host/mainframe side); fill
# it via Align.ps1 -HostSystemTypes / VerifyConfig once the real values are
# known. Returns HostToOpen / OpenToOpen / OpenToHost / HostToHost, or
# 'Unknown' when -HostTypes is empty (caller should warn + fall back).
function Get-MigrationType {
    param([string]$FromSys, [string]$ToSys, [string[]]$HostTypes)
    if ($null -eq $HostTypes -or @($HostTypes).Count -eq 0) { return 'Unknown' }
    $fromHost = (@($HostTypes) -contains $FromSys)
    $toHost   = (@($HostTypes) -contains $ToSys)
    if ($fromHost -and -not $toHost) { return 'HostToOpen' }
    if (-not $fromHost -and -not $toHost) { return 'OpenToOpen' }
    if (-not $fromHost -and $toHost) { return 'OpenToHost' }
    return 'HostToHost'
}

# Compare two sheets that may carry pictures (e.g. the send-data screenshot
# sheet). Picture count is a cheap proxy for "the evidence image is present":
# we cannot diff picture bytes over COM, so a differing picture COUNT (e.g. work
# has 0, J4 has 1) is a difference, and an equal count + equal grid is "same"
# (treated as already-done -> skipped). Delegates the grid check to
# Compare-SheetGrid.
function Compare-AlignSheet {
    param(
        [int]$RowsA, [int]$ColsA, [string[]]$FlatA, [int]$PicsA,
        [int]$RowsB, [int]$ColsB, [string[]]$FlatB, [int]$PicsB
    )
    if ($PicsA -ne $PicsB) {
        return [pscustomobject]@{ Same = $false; Reason = ("picture count differs ({0} vs {1})" -f $PicsA, $PicsB) }
    }
    return Compare-SheetGrid -RowsA $RowsA -ColsA $ColsA -FlatA $FlatA -RowsB $RowsB -ColsB $ColsB -FlatB $FlatB
}

# Classify a sheet by the no-content rule it should be checked against.
#   'SendData'   : the send-data screenshot sheet (must hold >=1 picture).
#   'SendResult' : the GIFT/GFIX send-result sheets (must hold > MinTextRows text rows).
#   'Other'      : no emptiness rule (always considered prepared).
function Get-AlignSheetKind {
    param([string]$SheetName, [string]$SendDataName, [string[]]$SendResultNames)
    if ($SheetName -eq $SendDataName) { return 'SendData' }
    if (@($SendResultNames) -contains $SheetName) { return 'SendResult' }
    return 'Other'
}

# Decide whether a J4 baseline sheet actually holds evidence worth copying.
# A J4 workbook that the host team has not filled in yet (only the blank
# template) must NOT overwrite the work sheet -- report "no contents" and skip.
#   SendData   : needs at least one picture.
#   SendResult : needs more than MinTextRows rows containing text.
# Returns @{ Prepared; Reason }.
function Test-J4SheetPrepared {
    param(
        [string]$Kind,
        [int]$PictureCount,
        [int]$TextRowCount,
        [int]$MinTextRows = 3
    )
    switch ($Kind) {
        'SendData' {
            if ($PictureCount -lt 1) {
                return [pscustomobject]@{ Prepared = $false; Reason = 'no picture in send-data sheet (template only)' }
            }
        }
        'SendResult' {
            if ($TextRowCount -le $MinTextRows) {
                return [pscustomobject]@{ Prepared = $false; Reason = ("only {0} text row(s) (<= {1})" -f $TextRowCount, $MinTextRows) }
            }
        }
    }
    return [pscustomobject]@{ Prepared = $true; Reason = '' }
}

# Which sheets Align should compare for a migration type.
# Returns a list of sheet names from this workbook that should be checked
# against the J4 baseline. Recv sheets (operator evidence) are included for
# comparison/read scope; Align.ps1 replaces only the sheets returned here.
#
# HostToOpen  : host team owns send data + GIFT/GFIX send result sheets.
# OpenToOpen  : compare send-result sheets (S[2]/S[3]) + all recv sheets.
# OpenToHost  : operator owns send side; compare send-result sheets + recv sheets.
# HostToHost  : host team owns all send sheets.
# Unknown     : recv sheets only (legacy safest fallback until HostTypes are set).
function Get-AlignSheetsForMigration {
    param([string]$MigrationType, [string[]]$SendSheets, [string[]]$RecvSheets)
    $send = @($SendSheets)
    $recv = @($RecvSheets)
    switch ($MigrationType) {
        'HostToOpen' { return @($send[0], $send[2], $send[3]) }
        'OpenToOpen' { return @($send[2], $send[3]) + $recv }
        'OpenToHost' { return @($send[2], $send[3]) + $recv }
        'HostToHost' { return $send }
        default      { return $recv }   # Unknown -> recv only (safe)
    }
}
