# ============================================================
# Fix-Encoding.ps1
#   Normalises every .ps1 / .psd1 / .json in this folder to the
#   encoding policy documented in CLAUDE.md / enforced by
#   Check-Encoding.ps1:
#     .ps1          -> UTF-8, NO BOM, ASCII-only source
#                      (strip a stray BOM; CANNOT auto-fix raw
#                       Japanese -> build it via [char] by hand)
#     .psd1         -> UTF-8; BOM only when it holds raw (non-ASCII)
#                       text, because Import-PowerShellDataFile needs
#                       the BOM to read Japanese correctly
#     .json/.jsonl  -> UTF-8, NO BOM
#
#   (Earlier versions of this script added a BOM to *every* file,
#    which violated the .ps1 "no BOM" rule -- fixed.)
#
#   Usage (run once after copying files from repo / a Claude paste):
#     .\Fix-Encoding.ps1
#     .\Fix-Encoding.ps1 -DryRun      <- preview only, no writes
# ============================================================
param([switch]$DryRun)

$ErrorActionPreference = 'Stop'
$bom     = [byte[]](0xEF, 0xBB, 0xBF)
$changed = 0
$ok      = 0
$warn    = 0

function Test-HasBom([byte[]]$bytes) {
    return ($bytes.Length -ge 3) -and ($bytes[0] -eq 0xEF) -and ($bytes[1] -eq 0xBB) -and ($bytes[2] -eq 0xBF)
}
function Test-IsAscii([byte[]]$bytes, [int]$start) {
    for ($i = $start; $i -lt $bytes.Length; $i++) { if ($bytes[$i] -gt 127) { return $false } }
    return $true
}
function Write-Bytes([string]$path, [byte[]]$bytes) {
    if (-not $DryRun) { [System.IO.File]::WriteAllBytes($path, $bytes) }
}
function Strip-Bom([byte[]]$bytes) {
    $out = New-Object byte[] ($bytes.Length - 3)
    [Array]::Copy($bytes, 3, $out, 0, $out.Length)
    return $out
}
function Add-Bom([byte[]]$bytes) {
    $out = New-Object byte[] ($bom.Length + $bytes.Length)
    [Array]::Copy($bom,   0, $out, 0,           $bom.Length)
    [Array]::Copy($bytes, 0, $out, $bom.Length, $bytes.Length)
    return $out
}

# Robust enumeration: -Include without -Recurse is unreliable, so filter
# by extension on the piped items instead.
Get-ChildItem -LiteralPath $PSScriptRoot -File |
    Where-Object { $_.Extension -in '.ps1', '.psd1', '.json', '.jsonl' } |
    Sort-Object Name |
    ForEach-Object {
        $name  = $_.Name
        $ext   = $_.Extension
        $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
        $hasBom = Test-HasBom $bytes

        switch ($ext) {
            '.ps1' {
                if ($hasBom) {
                    Write-Host ("  -BOM    {0}" -f $name) -ForegroundColor Yellow
                    Write-Bytes $_.FullName (Strip-Bom $bytes); $changed++
                    $bytes = Strip-Bom $bytes
                }
                if (-not (Test-IsAscii $bytes 0)) {
                    Write-Host ("  WARN    {0}: non-ASCII .ps1 (migrate Japanese to [char] by hand)" -f $name) -ForegroundColor Red
                    $warn++
                } elseif (-not $hasBom) {
                    Write-Host ("  ok      {0}" -f $name) -ForegroundColor DarkGray; $ok++
                }
            }
            '.psd1' {
                $body   = if ($hasBom) { Strip-Bom $bytes } else { $bytes }
                $ascii  = Test-IsAscii $body 0
                if (-not $ascii -and -not $hasBom) {
                    Write-Host ("  +BOM    {0}: holds raw text, needs BOM" -f $name) -ForegroundColor Yellow
                    Write-Bytes $_.FullName (Add-Bom $bytes); $changed++
                } elseif ($ascii -and $hasBom) {
                    Write-Host ("  -BOM    {0}: pure ASCII, BOM not needed" -f $name) -ForegroundColor Yellow
                    Write-Bytes $_.FullName (Strip-Bom $bytes); $changed++
                } else {
                    Write-Host ("  ok      {0}" -f $name) -ForegroundColor DarkGray; $ok++
                }
            }
            default {
                # .json / .jsonl -> no BOM
                if ($hasBom) {
                    Write-Host ("  -BOM    {0}" -f $name) -ForegroundColor Yellow
                    Write-Bytes $_.FullName (Strip-Bom $bytes); $changed++
                } else {
                    Write-Host ("  ok      {0}" -f $name) -ForegroundColor DarkGray; $ok++
                }
            }
        }
    }

Write-Host ''
if ($DryRun) {
    Write-Host ("[DRY-RUN] {0} file(s) would change, {1} ok, {2} warning(s)." -f $changed, $ok, $warn) -ForegroundColor Cyan
} else {
    Write-Host ("[DONE] {0} file(s) changed, {1} ok, {2} warning(s)." -f $changed, $ok, $warn) -ForegroundColor Green
}
if ($warn -gt 0) {
    Write-Host 'NOTE: non-ASCII .ps1 files cannot be auto-fixed; rebuild Japanese via [char] (see CLAUDE.md).' -ForegroundColor Yellow
}
