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
# -SendSheets = the 5 send-side names (RecvSheets arg kept for signature compat but NEVER synced).
#
# Recv sheets (GIFT/GFIX jushin kekka, Df compare) hold the operator's own captured
# evidence. Align must never overwrite them from J4.
#
# HostToOpen  : host team owns send[0]/send[2]/send[3] in J4; operator fetches updates.
# OpenToOpen  : operator owns both sides; only send-result sheets kept for future
#               coworker-alignment use.
# OpenToHost  : operator owns send side; sync send-result sheets from J4 baseline.
# HostToHost  : host team owns all send sheets.
# Unknown     : fall back to the two send-result sheets (safest non-destructive guess).
function Get-AlignSheetsForMigration {
    param([string]$MigrationType, [string[]]$SendSheets, [string[]]$RecvSheets)
    $send = @($SendSheets)
    switch ($MigrationType) {
        'HostToOpen' { return @($send[0], $send[2], $send[3]) }   # soushin-data + GIFT/GFIX soushin-kekka
        'OpenToOpen' { return @($send[2], $send[3]) }              # GIFT/GFIX soushin-kekka only
        'OpenToHost' { return @($send[2], $send[3]) }
        'HostToHost' { return @($send) }
        default      { return @($send[2], $send[3]) }              # Unknown -> send-result sheets only
    }
}
