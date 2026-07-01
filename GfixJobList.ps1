# ============================================================
#  GfixJobList.ps1
#
#  PURE parser for the GoAnywhere "completed jobs" LIST page text (the
#  result of Ctrl+A / Ctrl+C on that page). No COM, no SendKeys, no
#  mapping I/O. Dot-source only (no param() block).
#
#  Why this exists:
#    GfixLogDownload used to find its target row by Ctrl+F-searching the
#    IF_NO (project name) text. When two job runs share the same IF_NO
#    (one IF_NO commonly feeds more than one downstream SS_CODE receive
#    job), Find-in-page cannot tell them apart and both correls end up
#    opening the SAME physical row -- one correl's real log is then never
#    downloaded at all. This parser turns the raw list text into
#    structured rows keyed by the unique JobNo (job number) column, so
#    GfixLogDownload can download every candidate job for a needed IF_NO
#    and let GfixLog.ps1's content match (Find-GfixLogForCorrel) decide
#    which correl each log actually belongs to.
#
#  Row shape (tab-separated columns as copied from the GoAnywhere table):
#    JobNo  ProjectName  Folder  Status  User  StartTime  EndTime  Seconds  Method
#  Status is frequently empty (still occupies its own tab-delimited slot).
#  A data row is identified purely by JobNo matching ^\d{6,}$ -- this lets
#  the parser skip the Japanese header row, blank lines, and page-chrome
#  footer lines (row count / pager / copyright) without needing any
#  Japanese literals in this ASCII-source file.
# ============================================================

function ConvertFrom-GfixJobListText {
    param([string]$Text)
    $rows = [System.Collections.Generic.List[object]]::new()
    if ([string]::IsNullOrEmpty($Text)) { return $rows.ToArray() }
    $lines = [regex]::Split($Text, '\r?\n')
    foreach ($line in $lines) {
        $t = $line.Trim()
        if ([string]::IsNullOrEmpty($t)) { continue }
        $parts = $t -split "`t"
        if ($parts.Count -lt 8) { continue }
        $jobNo = $parts[0].Trim()
        if ($jobNo -notmatch '^\d{6,}$') { continue }
        $method = ''
        if ($parts.Count -gt 8) { $method = $parts[8].Trim() }
        $rows.Add([pscustomobject]@{
            JobNo       = $jobNo
            ProjectName = $parts[1].Trim()
            Folder      = $parts[2].Trim()
            Status      = $parts[3].Trim()
            User        = $parts[4].Trim()
            StartTime   = $parts[5].Trim()
            EndTime     = $parts[6].Trim()
            Seconds     = $parts[7].Trim()
            Method      = $method
        })
    }
    return $rows.ToArray()
}

# Filters parsed job-list rows down to the ones whose ProjectName carries
# the given normalized IF_NO substring (same normalization GfixLogDownload
# already applies: dashes -> underscores). -ReceiveOnly (default $true)
# additionally requires Folder to mention "Receive", since the SEND-side
# (FTP) sibling job commonly shares a similar/overlapping project name and
# GfixLogDownload only ever wants receive-side GFIX logs.
function Get-GfixJobListRowsForIf {
    param(
        [object[]]$Rows,
        [string]$IfNorm,
        [bool]$ReceiveOnly = $true
    )
    $out = [System.Collections.Generic.List[object]]::new()
    if ([string]::IsNullOrWhiteSpace($IfNorm)) { return $out.ToArray() }
    foreach ($r in @($Rows)) {
        if (-not ([string]$r.ProjectName).Contains($IfNorm)) { continue }
        if ($ReceiveOnly -and (([string]$r.Folder) -notmatch 'Receive')) { continue }
        $out.Add($r)
    }
    return $out.ToArray()
}

# Pure planning step for GfixLogDownload (mirrors EvidencePlan.ps1's split from
# its COM executor): given pending correls (each an object with .CorrelIdS and
# a pre-normalized .IfNorm, e.g. IF "5001-001" -> "5001_001") and the parsed
# job-list rows, decides which correls have NO candidate job today (HardMiss --
# nothing to navigate to, caller marks GFIX_log=2 immediately) and which
# distinct job numbers must actually be downloaded (NeededJobNumbers -- deduped
# across correls/IF_NOs sharing the same job, first-seen order). Content
# matching (Find-GfixLogForCorrel, GfixLog.ps1) -- not this function -- is what
# ultimately decides which correl a downloaded log belongs to.
function Get-GfixLogDownloadPlan {
    param(
        [object[]]$PendingRows,
        [object[]]$JobListRows,
        [bool]$ReceiveOnly = $true
    )
    $hardMiss = [System.Collections.Generic.List[object]]::new()
    $neededJobNumbers = [System.Collections.Generic.List[string]]::new()
    $seen = @{}
    $ifGroups = [ordered]@{}
    foreach ($p in @($PendingRows)) {
        $ifNorm = [string]$p.IfNorm
        if (-not $ifGroups.Contains($ifNorm)) { $ifGroups[$ifNorm] = [System.Collections.Generic.List[object]]::new() }
        $ifGroups[$ifNorm].Add($p)
    }
    foreach ($ifNorm in @($ifGroups.Keys)) {
        $rowsForIf = @(Get-GfixJobListRowsForIf -Rows $JobListRows -IfNorm $ifNorm -ReceiveOnly $ReceiveOnly)
        if ($rowsForIf.Count -eq 0) {
            foreach ($p in $ifGroups[$ifNorm]) {
                $hardMiss.Add([pscustomobject]@{ CorrelIdS = [string]$p.CorrelIdS; IfNorm = $ifNorm })
            }
            continue
        }
        foreach ($jr in $rowsForIf) {
            if (-not $seen.ContainsKey($jr.JobNo)) {
                $seen[$jr.JobNo] = $true
                $neededJobNumbers.Add($jr.JobNo)
            }
        }
    }
    return [pscustomobject]@{
        HardMiss         = $hardMiss.ToArray()
        NeededJobNumbers = $neededJobNumbers.ToArray()
    }
}
