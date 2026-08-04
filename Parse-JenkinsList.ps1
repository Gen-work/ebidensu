# ============================================================
# Parse-JenkinsList.ps1
#   Extracts (name, datetime, size) for each file from the Ctrl+A text of a
#   Jenkins file-list page. With -CorrelId, resolves that correl's entry --
#   the NEWEST one when several match, since duplicates are reruns of the
#   same transfer -- and, with -ExpectedTime, reports whether it falls inside
#   -ToleranceMinutes.
#   (ASCII source per the project encoding policy; the Japanese column word
#   this parser matches on is built from [char] below.)
# ============================================================
param(
    [Parameter(Mandatory)][string]$Text,
    [string]$CorrelId = '',
    [datetime]$ExpectedTime = [datetime]::MinValue,
    [int]$ToleranceMinutes = 30
)
$ErrorActionPreference = 'Stop'

$files = @()
foreach ($line in ($Text -split "`r?`n")) {
    $line = $line.Trim()
    # Row shape: "JIGPLB1S 2026/05/15 13:45:21 189.90 KB <sansho>", where
    # <sansho> ("reference") is the trailing link text of every listed file.
    $refWord = [char]0x53C2 + [char]0x7167   # sansho (reference)
    if ($line -match ('^(\S+)\s+(\d{4}/\d{2}/\d{2})\s+(\d{1,2}:\d{2}:\d{2})\s+(.+?)\s+' + $refWord + '$')) {
        $dt = $null
        try {
            $dt = [datetime]::ParseExact(
                ("{0} {1}" -f $Matches[2], $Matches[3]),
                'yyyy/MM/dd H:mm:ss',
                [System.Globalization.CultureInfo]::InvariantCulture)
        } catch {}
        $files += [PSCustomObject]@{
            Name     = $Matches[1]
            DateTime = $dt
            Size     = $Matches[4]
        }
    }
}

if (-not $CorrelId) { return $files }

# CorrelId may carry the transfer-batch stamp ("<correl>.<YYMMDD>.<8-digit>")
# that the listed file name may or may not include (see SnapVerify.ps1's
# Test-SnapCorrelIdMatch, which this legacy standalone parser deliberately
# duplicates rather than depending on).
function ConvertTo-JenkinsListBaseId([string]$Id) {
    if ($Id -match '^(?<base>.+)\.\d{6}\.\d{8}$') { return [string]$Matches['base'] }
    return $Id
}
$correlBase = ConvertTo-JenkinsListBaseId $CorrelId
$candidates = @($files | Where-Object {
    $_.Name -eq $CorrelId -or (ConvertTo-JenkinsListBaseId ([string]$_.Name)) -eq $correlBase
})
if ($candidates.Count -eq 0) {
    return [PSCustomObject]@{ Found = $false; Reason = "file not in list" }
}
# Several entries for one correl are reruns of the same transfer (or its plain
# and batch-stamped spellings): the latest one is the evidence. Taking the
# first LISTED entry, as this did, routinely picked an older run. Undated
# entries sort last -- they say nothing about which run they are.
$target = @($candidates | Sort-Object -Property `
    @{ Expression = { $null -ne $_.DateTime }; Descending = $true }, `
    @{ Expression = { if ($null -ne $_.DateTime) { $_.DateTime } else { [datetime]::MinValue } }; Descending = $true })[0]

if ($ExpectedTime -eq [datetime]::MinValue) {
    return [PSCustomObject]@{
        Found = $true; FileDateTime = $target.DateTime; Size = $target.Size
    }
}

$diff    = ($target.DateTime - $ExpectedTime).TotalMinutes
$absDiff = [Math]::Abs($diff)
[PSCustomObject]@{
    Found        = $true
    FileDateTime = $target.DateTime
    Size         = $target.Size
    DiffMinutes  = [Math]::Round($diff, 1)
    IsInRange    = ($absDiff -le $ToleranceMinutes)
    Reason       = if ($absDiff -le $ToleranceMinutes) { "OK" }
                   else { "time off by $([Math]::Round($diff,1)) min" }
}
