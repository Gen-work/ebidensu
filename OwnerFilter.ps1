# ============================================================
#  OwnerFilter.ps1
#
#  PURE owner-cell matching + job-by-owner filtering -- no Excel COM,
#  so it is unit-testable without Excel (Tests\Test-OwnerFilter.ps1).
#  Generate-HostOpenMapping.ps1 feeds it the WBS owner cells it reads.
#  Dot-source only (no param() block; ASCII source, arrows via [char]).
# ============================================================

# Decide whether a WBS owner cell represents the current operator (-Owner).
# Count only these patterns (others, incl. reverse direction, are NOT owned):
#   1) exact owner              : Owner
#   2) owner followed by <-...  : Owner<-Other   (owner is the receiver)
#   3) ... followed by ->owner  : Other->Owner   (owner is the receiver)
function Test-OwnerMatch([string]$OwnerCell, [string]$OwnerInput) {
    if ([string]::IsNullOrWhiteSpace($OwnerCell) -or [string]::IsNullOrWhiteSpace($OwnerInput)) {
        return $false
    }

    $cell  = $OwnerCell.Trim()
    $owner = $OwnerInput.Trim()

    # NOTE: the arrows are built from [char] code points on purpose. Raw
    # arrow literals in a BOM-less .ps1 get mis-decoded by PS 5.1 on a
    # JP-locale host, which silently breaks owner-matching.
    $arrowLeft  = [char]0x2190   # leftwards arrow  (owner<-...)
    $arrowRight = [char]0x2192   # rightwards arrow (...->owner)
    if ($cell -eq $owner) { return $true }
    if ($cell.StartsWith($owner + $arrowLeft)) { return $true }
    if ($cell.EndsWith($arrowRight + $owner)) { return $true }

    return $false
}

# Filter a list of candidate JOB_NAMEs against a WBS job->owner-cell map so
# that incremental -Add explicit selectors compose with the owner filter.
#
# Decision per candidate:
#   - present in $JobOwnerMap with a matching owner cell  -> KEEP
#   - present in $JobOwnerMap with a NON-matching owner   -> EXCLUDE (other
#                                                            owner's job)
#   - absent from $JobOwnerMap                            -> KEEP, because the
#       WBS cannot judge it (a temp / not-yet-in-WBS job); owner filtering
#       must not silently drop these. Reported via -Unknown for visibility.
#
# $JobOwnerMap : hashtable JOB_NAME -> owner cell string (case-insensitive
#                keys; build it from the WBS col A / col P scan).
# Returns a hashtable: @{ Kept = @(...); Excluded = @(...); Unknown = @(...) }
# (Excluded/Unknown carry the JOB_NAME plus the owner cell for warnings.)
function Select-JobsByOwner {
    param(
        [string[]]$Jobs,
        $JobOwnerMap,
        [string]$OwnerInput
    )

    $kept     = [System.Collections.Generic.List[string]]::new()
    $excluded = [System.Collections.Generic.List[object]]::new()
    $unknown  = [System.Collections.Generic.List[string]]::new()
    $seen     = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($raw in @($Jobs)) {
        if ($null -eq $raw) { continue }
        $job = ([string]$raw).Trim()
        if ([string]::IsNullOrWhiteSpace($job)) { continue }
        if (-not $seen.Add($job)) { continue }   # de-dup, keep first-seen order

        $hasOwner = ($null -ne $JobOwnerMap) -and $JobOwnerMap.ContainsKey($job)
        if (-not $hasOwner) {
            $unknown.Add($job)
            $kept.Add($job)            # cannot judge -> keep
            continue
        }

        $ownerCell = [string]$JobOwnerMap[$job]
        if (Test-OwnerMatch $ownerCell $OwnerInput) {
            $kept.Add($job)
        } else {
            $excluded.Add([pscustomobject]@{ Job = $job; OwnerCell = $ownerCell })
        }
    }

    return @{
        Kept     = @($kept)
        Excluded = @($excluded)
        Unknown  = @($unknown)
    }
}
