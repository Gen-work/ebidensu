# ============================================================
# Parse-GiftMq.ps1
#   GIFT/MQ Transfer status inquiry ページの Ctrl+A テキスト解析。
#   呼び出し側で frame_main を click してから Read-PageText.ps1 を実行。
# ============================================================
param(
    [Parameter(Mandatory)][string]$Text
)
$ErrorActionPreference = 'Stop'

$m = [regex]::Match($Text, 'Number of records\s+(\d+)')
$numRec = if ($m.Success) { [int]$m.Groups[1].Value } else { -1 }

# データ行抽出
# 例: "1 JSSS004R JHM102R JIDSQ48S 2026/05/15 00:48:38 TXT 2026/05/15 00:48:38 0 0"
$rows = @()
foreach ($line in ($Text -split "`r?`n")) {
    $line = $line.Trim()
    if ($line -match '^(\d+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2})\s+(\S+)\s+(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2})\s+(\d+)\s+(\d+)\s*$') {
        $sendDt = $null; $recvDt = $null
        try {
            $sendDt = [datetime]::ParseExact($Matches[5], 'yyyy/MM/dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
            $recvDt = [datetime]::ParseExact($Matches[7], 'yyyy/MM/dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
        } catch {}
        $rows += [PSCustomObject]@{
            No        = [int]$Matches[1]
            SendNode  = $Matches[2]
            RecvNode  = $Matches[3]
            CorrelId  = $Matches[4]
            SendDate  = $sendDt
            Tmode     = $Matches[6]
            RecvDate  = $recvDt
            Rtncd     = [int]$Matches[8]
            Rsncd     = [int]$Matches[9]
        }
    }
}

[PSCustomObject]@{
    NumRecords = $numRec
    Rows       = $rows
}
