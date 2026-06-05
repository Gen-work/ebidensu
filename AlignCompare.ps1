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

# Which sheets Align should compare for a migration type.
# Returns a list of sheet names from this workbook that should be checked
# against the J4 baseline. Recv sheets (operator evidence) are included for
# comparison (read) but Align.ps1 must never overwrite them (-Apply is send-only).
#
# HostToOpen  : host team owns the J4 send side; operator compares recv sheets.
# OpenToOpen  : compare send-result sheets (S[2]/S[3]) + all recv sheets.
# OpenToHost  : operator owns send side; compare send-result sheets + recv sheets.
# HostToHost  : host team owns all send sheets.
# Unknown     : recv sheets only (safest non-destructive fallback).
function Get-AlignSheetsForMigration {
    param([string]$MigrationType, [string[]]$SendSheets, [string[]]$RecvSheets)
    $send = @($SendSheets)
    $recv = @($RecvSheets)
    switch ($MigrationType) {
        'HostToOpen' { return $recv }
        'OpenToOpen' { return @($send[2], $send[3]) + $recv }
        'OpenToHost' { return @($send[2], $send[3]) + $recv }
        'HostToHost' { return $send }
        default      { return $recv }   # Unknown -> recv only (safe)
    }
}
