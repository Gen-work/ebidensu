# ============================================================
# Fix-Encoding.ps1
#   Ensures every .ps1 / .psd1 in this folder has UTF-8 BOM.
#   PS 5.1 requires BOM for Import-PowerShellDataFile and for
#   Japanese/Chinese string literals to be parsed correctly.
#
#   Usage (run once after copying files from repo/Claude):
#     .\Fix-Encoding.ps1
#     .\Fix-Encoding.ps1 -DryRun      <- preview only, no writes
# ============================================================
param([switch]$DryRun)

$bom   = [byte[]](0xEF, 0xBB, 0xBF)
$added = 0
$ok    = 0

Get-ChildItem $PSScriptRoot -Include '*.ps1','*.psd1' -File |
    Sort-Object Name |
    ForEach-Object {
        $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
        $hasBom = ($bytes.Length -ge 3) -and
                  ($bytes[0] -eq 0xEF) -and ($bytes[1] -eq 0xBB) -and ($bytes[2] -eq 0xBF)
        if ($hasBom) {
            Write-Host ("  ok      {0}" -f $_.Name) -ForegroundColor DarkGray
            $ok++
        } else {
            Write-Host ("  +BOM    {0}" -f $_.Name) -ForegroundColor Yellow
            if (-not $DryRun) {
                $newBytes = New-Object byte[] ($bom.Length + $bytes.Length)
                [Array]::Copy($bom,   0, $newBytes, 0,           $bom.Length)
                [Array]::Copy($bytes, 0, $newBytes, $bom.Length, $bytes.Length)
                [System.IO.File]::WriteAllBytes($_.FullName, $newBytes)
            }
            $added++
        }
    }

Write-Host ''
if ($DryRun) {
    Write-Host ("[DRY-RUN] {0} files need BOM, {1} already ok." -f $added, $ok) -ForegroundColor Cyan
} else {
    Write-Host ("[DONE] BOM added to {0} file(s). {1} already ok." -f $added, $ok) -ForegroundColor Green
}
