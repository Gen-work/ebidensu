# ============================================================
# Parse-JenkinsList.ps1
#   Jenkins ファイル一覧ページの Ctrl+A テキストから
#   各ファイルの (name, datetime, size) を抽出。
#   ExpectedTime 指定時は ToleranceMinutes 内に収まるか判定。
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
    # "JIGPLB1S 2026/05/15 13:45:21 189.90 KB 参照"
    $refWord = [char]0x53C2 + [char]0x7167   # 参照
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

$target = $files | Where-Object { $_.Name -eq $CorrelId } | Select-Object -First 1
if (-not $target) {
    return [PSCustomObject]@{ Found = $false; Reason = "file not in list" }
}

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
