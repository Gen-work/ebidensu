# Read-ClipboardJson.ps1
# ============================================================
# Reads the JSON a bookmarklet wrote to the clipboard and returns it
# as a PSCustomObject.
# Bookmarklet-side format example:
#   {"source":"GIFT","numRecords":2,"records":[...]}
#
# Usage:
#   $data = & .\Read-ClipboardJson.ps1 -ExpectedSource "GIFT" -TimeoutSec 10
#   if ($data) { Write-Host ("N = {0}" -f $data.numRecords) }
# ============================================================
param(
    [string]$ExpectedSource = '',   # "GIFT" / "HM" / "MQ" / "Jenkins"
    [int]$TimeoutSec = 10,          # max seconds to wait for bookmarklet
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

    # quick check that it looks like JSON
    $trimmed = $text.Trim()
    if (-not $trimmed.StartsWith('{')) { continue }

    try {
        $obj = $trimmed | ConvertFrom-Json -ErrorAction Stop
    } catch { continue }

    # match the source tag (when one was requested)
    if ($ExpectedSource -and $obj.source -ne $ExpectedSource) {
        Write-Host ("  [INFO] source mismatch: got '{0}', expected '{1}'. waiting..." `
            -f $obj.source, $ExpectedSource) -ForegroundColor DarkGray
        continue
    }

    return $obj
}

Write-Warning ("Timeout: no valid JSON in clipboard after {0}s." -f $TimeoutSec)
return $null
