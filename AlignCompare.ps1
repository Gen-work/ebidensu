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

# Which sheets Align should compare/sync for a migration type.
# -SendSheets = the 5 send-side names; -RecvSheets = the 3 receive-side names
# (from ProjectLabels). Known rule: Host->Open touches only the 3 receive
# sheets. Open->Open / Open->Host additionally touch the GIFT/GFIX send-result
# sheets (the "check 3 and 4" note). These defaults are config-overridable and
# should be confirmed against the real workflow.
function Get-AlignSheetsForMigration {
    param([string]$MigrationType, [string[]]$SendSheets, [string[]]$RecvSheets)
    $send = @($SendSheets)
    $recv = @($RecvSheets)
    switch ($MigrationType) {
        'HostToOpen' { return $recv }
        'OpenToOpen' { return @($send[2], $send[3]) + $recv }   # GIFT/GFIX soushin kekka + 3 recv
        'OpenToHost' { return @($send[2], $send[3]) + $recv }
        'HostToHost' { return @($send) + $recv }
        default      { return $recv }   # Unknown -> safest (Host->Open scope) + caller warns
    }
}
