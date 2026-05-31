# ============================================================
#  Check-Encoding.ps1   (read-only diagnostic)
#
#  Verifies the repo's encoding policy and flags mojibake risks.
#
#  Policy (per project decision + PS 5.1 reality):
#    .ps1   -> UTF-8, NO BOM. Runtime-critical Japanese must be built
#              from [char] codes (see ProjectLabels.ps1), so the source
#              stays ASCII and a BOM-less file cannot be mis-decoded by
#              PS 5.1 on a JP-locale host.
#    .psd1  -> if it contains raw Japanese, it MUST keep a BOM, because
#              Import-PowerShellDataFile cannot evaluate [char] and would
#              mis-decode a BOM-less JP file. (Pure-ASCII .psd1 = no BOM.)
#    .csv   -> UTF-8 WITH BOM (Excel-friendly; handled by MappingStore).
#    .json / .jsonl -> UTF-8, NO BOM.
#
#  Exit code: 1 if any hard violation (replacement char U+FFFD, BOM on a
#  .ps1, or a key Japanese label failed to construct); else 0. Raw
#  non-ASCII in .ps1 is reported as a WARNING (migration debt), not a
#  hard failure.
# ============================================================

param(
    [string]$Root = $PSScriptRoot,
    [switch]$ListNonAscii
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

if ([string]::IsNullOrWhiteSpace($Root)) { $Root = (Get-Location).Path }

$errors   = [System.Collections.Generic.List[string]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()

function Test-HasBom([byte[]]$bytes) {
    return ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
}

function Test-HasRawJapanese([string]$text) {
    foreach ($ch in $text.ToCharArray()) {
        $c = [int]$ch
        if ($c -gt 0x7F) { return $true }
    }
    return $false
}

Write-Host ''
Write-Host '===== Check-Encoding =====' -ForegroundColor Green
Write-Host ("  Root: {0}" -f $Root)
Write-Host ''

$files = @(Get-ChildItem -LiteralPath $Root -File -Recurse -Include '*.ps1','*.psd1','*.json','*.jsonl' -ErrorAction SilentlyContinue)

foreach ($f in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $hasBom = Test-HasBom $bytes
    $text   = [System.Text.Encoding]::UTF8.GetString($bytes)
    $ext    = $f.Extension.ToLower()
    $rel    = $f.FullName.Substring($Root.Length).TrimStart('\','/')

    # U+FFFD replacement char = decoding already went wrong somewhere.
    if ($text.Contains([char]0xFFFD)) {
        $errors.Add(("{0}: contains U+FFFD replacement character (already corrupted)" -f $rel))
    }

    switch ($ext) {
        '.ps1' {
            if ($hasBom) { $errors.Add(("{0}: .ps1 has a BOM (policy: no BOM)" -f $rel)) }
            if (Test-HasRawJapanese $text) {
                $warnings.Add(("{0}: .ps1 has raw non-ASCII -> mojibake risk under no-BOM PS 5.1; move runtime strings to [char] (ProjectLabels.ps1)" -f $rel))
                if ($ListNonAscii) {
                    $n = 0; $ln = 0
                    foreach ($line in ($text -split "`n")) {
                        $ln++
                        if (Test-HasRawJapanese $line) {
                            Write-Host ("      {0}:{1}: {2}" -f $rel, $ln, $line.Trim()) -ForegroundColor DarkGray
                            $n++; if ($n -ge 6) { break }
                        }
                    }
                }
            }
        }
        '.psd1' {
            if (Test-HasRawJapanese $text) {
                if (-not $hasBom) {
                    $errors.Add(("{0}: .psd1 has raw Japanese but NO BOM -> Import-PowerShellDataFile will mojibake on JP host. Add a BOM or remove Japanese." -f $rel))
                }
            } elseif ($hasBom) {
                $warnings.Add(("{0}: pure-ASCII .psd1 has an unneeded BOM" -f $rel))
            }
        }
        default {
            if ($hasBom) { $warnings.Add(("{0}: {1} has a BOM (policy: no BOM)" -f $rel, $ext)) }
        }
    }
}

# ---- Pack-LlmContext paste residue ----
# A truncated context paste leaves a "`" / "``n" / "--- File: X ---" tail
# that breaks the PowerShell parser. Pack-LlmContext.ps1 itself legitimately
# emits the separator, so skip it.
foreach ($f in $files) {
    if ($f.Name -eq 'Pack-LlmContext.ps1') { continue }
    if ($f.Extension.ToLower() -ne '.ps1' -and $f.Extension.ToLower() -ne '.psd1') { continue }
    $rel  = $f.FullName.Substring($Root.Length).TrimStart('\','/')
    $text = [System.Text.Encoding]::UTF8.GetString([System.IO.File]::ReadAllBytes($f.FullName))
    if ($text -match '(?m)^\s*--- File:') {
        $errors.Add(("{0}: contains a '--- File:' Pack-LlmContext separator (truncated paste)" -f $rel))
    }
    if ($text -match '(?m)^``n\s*$') {
        $errors.Add(("{0}: contains a stray '``n' Pack-LlmContext marker (truncated paste)" -f $rel))
    }
}

# ---- self-test: key Japanese labels must construct cleanly ----
Write-Host '[*] Label self-test (ProjectLabels.ps1):' -ForegroundColor Cyan
$labelsOk = $true
try {
    . (Join-Path $Root 'ProjectLabels.ps1')
    $L = Get-ProjectLabels
    foreach ($key in @('SheetSoshinData','SheetGiftRecv','SheetGfixRecv','SheetDfCompare','GfixLogLabel','GiftNoGfixHeader')) {
        $val = [string]$L[$key]
        $bad = ([string]::IsNullOrEmpty($val) -or $val.Contains([char]0xFFFD))
        if ($bad) { $labelsOk = $false }
        $flag = if ($bad) { '[BAD]' } else { '[ok]' }
        Write-Host ("    {0,-18} {1} {2}" -f $key, $flag, $val)
    }
} catch {
    $labelsOk = $false
    $errors.Add(("ProjectLabels.ps1 self-test threw: {0}" -f $_.Exception.Message))
}
if (-not $labelsOk) { $errors.Add('Japanese label self-test failed') }

# ---- report ----
Write-Host ''
if ($warnings.Count -gt 0) {
    Write-Host ("Warnings ({0}):" -f $warnings.Count) -ForegroundColor Yellow
    foreach ($w in $warnings) { Write-Host ("  [WARN] {0}" -f $w) -ForegroundColor Yellow }
}
if ($errors.Count -gt 0) {
    Write-Host ("Errors ({0}):" -f $errors.Count) -ForegroundColor Red
    foreach ($e in $errors) { Write-Host ("  [ERR ] {0}" -f $e) -ForegroundColor Red }
    Write-Host ''
    Write-Host '===== Check-Encoding: FAIL =====' -ForegroundColor Red
    exit 1
}
Write-Host '===== Check-Encoding: OK =====' -ForegroundColor Green
exit 0
