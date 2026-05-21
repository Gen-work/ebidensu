# Read-ClipboardJson.ps1
# ============================================================
# bookmarklet が clipboard に書き込んだ JSON を読み取り、
# PSCustomObject に変換して返す。
# bookmarklet 側のフォーマット例:
#   {"source":"GIFT","numRecords":2,"records":[...]}
#
# 使用例:
#   $data = & .\Read-ClipboardJson.ps1 -ExpectedSource "GIFT" -TimeoutSec 10
#   if ($data) { Write-Host ("N = {0}" -f $data.numRecords) }
# ============================================================
param(
    [string]$ExpectedSource = '',   # "GIFT" / "HM" / "MQ" / "Jenkins"
    [int]$TimeoutSec = 10,          # bookmarklet 起動を待つ最大秒数
    [int]$PollMs = 250
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

$deadline = (Get-Date).AddSeconds($TimeoutSec)
$prevText = $null

while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds $PollMs
    try {
        $text = [System.Windows.Forms.Clipboard]::GetText()
    } catch { $text = $null }

    if ([string]::IsNullOrWhiteSpace($text)) { continue }
    if ($text -eq $prevText) { continue }
    $prevText = $text

    # JSON っぽいか軽くチェック
    $trimmed = $text.Trim()
    if (-not $trimmed.StartsWith('{')) { continue }

    try {
        $obj = $trimmed | ConvertFrom-Json -ErrorAction Stop
    } catch { continue }

    # source タグ照合（指定があれば）
    if ($ExpectedSource -and $obj.source -ne $ExpectedSource) {
        Write-Host ("  [INFO] source mismatch: got '{0}', expected '{1}'. waiting..." `
            -f $obj.source, $ExpectedSource) -ForegroundColor DarkGray
        continue
    }

    return $obj
}

Write-Warning ("Timeout: no valid JSON in clipboard after {0}s." -f $TimeoutSec)
return $null
