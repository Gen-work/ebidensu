# ============================================================
#  OcrTool.ps1 - standalone Windows-OCR command line tool
#
#  Thin CLI over the dot-source libs so OCR is reusable outside the
#  SendVsGift phase:
#    OcrWindows.ps1          - Windows.Media.Ocr engine (zero installs)
#    SendMetadata.ps1        - word-box spacing rebuild (the ja recognizer
#                              drops spaces between tokens)
#    EvidenceImageExport.ps1 - optional: export embedded workbook pictures
#                              (Ctrl+G groups are flattened) before OCR
#
#  Usage:
#    .\OcrTool.ps1 -ListLanguages
#    .\OcrTool.ps1 -Path shot.png
#    .\OcrTool.ps1 -Path C:\caps\ -OutDir C:\caps\txt
#    .\OcrTool.ps1 -Path 'C:\caps\*.png' -Json -OutFile result.json
#    .\OcrTool.ps1 -Workbook evidence.xlsx -Sheet <send sheet> -OutDir C:\tmp\pics
#
#  Programmatic reuse: dot-source OcrWindows.ps1 + SendMetadata.ps1 and
#  call Invoke-WinOcrFile / ConvertTo-SendTextLines directly (this file
#  has a param() block - never dot-source it).
# ============================================================

param(
    [string[]]$Path = @(),
    [string]$Language = 'ja',
    [string]$Workbook = '',
    [string]$Sheet = '',
    [string]$OutDir = '',
    [string]$OutFile = '',
    [switch]$Json,
    [switch]$NoSpacing,
    [switch]$ListLanguages,
    [switch]$Diag
)

$ErrorActionPreference = 'Stop'

# capture switches BEFORE dot-sourcing (see CLAUDE.md dot-source safety rule)
$jsonFlag      = [bool]$Json.IsPresent
$noSpacingFlag = [bool]$NoSpacing.IsPresent
$listFlag      = [bool]$ListLanguages.IsPresent
$diagFlag      = [bool]$Diag.IsPresent

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

. (Join-Path $PSScriptRoot 'OcrWindows.ps1')
. (Join-Path $PSScriptRoot 'SendMetadata.ps1')

if ($listFlag) {
    if (-not (Test-WinOcrAvailable)) {
        Write-Host ("[ERROR] Windows OCR unavailable: {0}" -f (Get-WinOcrInitError)) -ForegroundColor Red
        exit 1
    }
    Write-Host 'Installed OCR recognizer languages:' -ForegroundColor Cyan
    foreach ($t in (Get-WinOcrLanguageTags)) { Write-Host ("  {0}" -f $t) }
    exit 0
}

$imageExt = @('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff')

function Get-OcrImageList([string[]]$Items) {
    $files = @()
    foreach ($raw in @($Items)) {
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        if (Test-Path -LiteralPath $raw) {
            $it = Get-Item -LiteralPath $raw
            if ($it.PSIsContainer) {
                $files += @(Get-ChildItem -LiteralPath $it.FullName -File |
                    Where-Object { $imageExt -contains $_.Extension.ToLower() } |
                    Sort-Object Name | ForEach-Object { $_.FullName })
            } else {
                $files += $it.FullName
            }
        } else {
            # wildcard pattern
            $files += @(Get-Item -Path $raw -ErrorAction SilentlyContinue |
                Where-Object { -not $_.PSIsContainer } | Sort-Object Name |
                ForEach-Object { $_.FullName })
        }
    }
    # no comma protection: the caller wraps this in @(...) and a
    # comma-protected return would nest (one element = whole list)
    return @($files | Select-Object -Unique)
}

$images = @(Get-OcrImageList $Path)

# Optional workbook source: export the embedded pictures first, then OCR
# them like any other image file.
if (-not [string]::IsNullOrWhiteSpace($Workbook)) {
    . (Join-Path $PSScriptRoot 'EvidenceImageExport.ps1')
    if (-not (Test-Path -LiteralPath $Workbook)) { throw "workbook not found: $Workbook" }
    $wbFull = (Resolve-Path -LiteralPath $Workbook).Path
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($wbFull)
    $exportDir = $OutDir
    if ([string]::IsNullOrWhiteSpace($exportDir)) {
        $exportDir = Join-Path (Split-Path -Parent $wbFull) ('ocr_images_' + $stem)
    }
    $excel = $null
    $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($wbFull, 0, $true)
        $sheetNames = @()
        if (-not [string]::IsNullOrWhiteSpace($Sheet)) { $sheetNames += $Sheet }
        else { foreach ($s in $wb.Worksheets) { $sheetNames += [string]$s.Name } }
        foreach ($sn in $sheetNames) {
            $safe = ($sn -replace '[\\/:*?"<>|]', '_')
            $pngs = @(Export-SheetPicturesToPng $wb $sn $exportDir ($stem + '_' + $safe))
            if ($pngs.Count -gt 0) {
                Write-Host ("[EXPORT] {0}: {1} picture(s) -> {2}" -f $sn, $pngs.Count, $exportDir) -ForegroundColor Green
                $images += $pngs
            }
        }
    } finally {
        if ($wb) { try { $wb.Close($false) } catch {} }
        if ($excel) {
            try { $excel.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
        }
    }
}

if ($images.Count -eq 0) {
    Write-Host 'No images to OCR. Use -Path <file|dir|wildcard> and/or -Workbook <xlsx>.' -ForegroundColor Yellow
    exit 1
}

if (-not (Test-WinOcrAvailable)) {
    Write-Host ("[ERROR] Windows OCR unavailable: {0}" -f (Get-WinOcrInitError)) -ForegroundColor Red
    exit 1
}

# -Diag: per-image engine sweep (every installed language + user-profile
# engine) with pixel size vs the engine's MaxImageDimension. Use this
# when OCR returns nothing to tell an image problem from an engine one.
if ($diagFlag) {
    foreach ($img in $images) {
        Write-Host ''
        Write-Host ("===== DIAG {0} =====" -f $img) -ForegroundColor Cyan
        try {
            $d = Invoke-WinOcrDiag -Path $img
            $dimNote = ''
            if ($d.MaxImageDimension -gt 0 -and -not [string]::IsNullOrWhiteSpace($d.PixelSize)) {
                $parts = $d.PixelSize -split 'x'
                $over = ([int]$parts[0] -gt $d.MaxImageDimension) -or ([int]$parts[1] -gt $d.MaxImageDimension)
                if ($over) { $dimNote = '  ** EXCEEDS MaxImageDimension **' }
            }
            Write-Host ("  pixel size : {0}  (engine MaxImageDimension: {1}){2}" -f $d.PixelSize, $d.MaxImageDimension, $dimNote)
            foreach ($a in $d.Attempts) {
                if ([string]::IsNullOrWhiteSpace($a.Error)) {
                    $color = if ($a.Chars -gt 0) { 'Green' } elseif ($a.Lines -gt 0) { 'Yellow' } else { 'DarkYellow' }
                    Write-Host ("  {0,-14} engine={1,-6} lines={2,-4} words={3,-5} chars={4,-6} rawChars={5,-6} sample: {6}" -f `
                        $a.Language, $a.Engine, $a.Lines, $a.Words, $a.Chars, $a.RawChars, $a.Sample) -ForegroundColor $color
                    if ($a.Lines -gt 0 -and $a.Chars -eq 0) {
                        Write-Host ("                 ** Lines/Words enumerate but every .Text is EMPTY (WinRT projection issue); line type: {0} **" -f $a.LineType) -ForegroundColor Yellow
                    }
                } else {
                    Write-Host ("  {0,-14} ERROR: {1}" -f $a.Language, $a.Error) -ForegroundColor Red
                }
            }
        } catch {
            Write-Host ("  [ERROR] {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    }
    exit 0
}

$results = @()
foreach ($img in $images) {
    try {
        $res = Invoke-WinOcrFile -Path $img -LanguageTag $Language
        $lines = if ($noSpacingFlag) {
            @(@($res.Lines) | ForEach-Object { [string]$_.Text })
        } else {
            @(ConvertTo-SendTextLines $res.Lines)
        }
        $results += [pscustomobject]@{
            Path     = [string]$res.Path
            Language = [string]$res.LanguageTag
            Lines    = @($lines)
            Text     = ($lines -join "`r`n")
        }
    } catch {
        Write-Host ("[WARN] OCR failed: {0} ({1})" -f $img, $_.Exception.Message) -ForegroundColor Yellow
    }
}

if ($jsonFlag) {
    $payload = ($results | ConvertTo-Json -Depth 5)
    if (-not [string]::IsNullOrWhiteSpace($OutFile)) {
        [System.IO.File]::WriteAllText($OutFile, $payload, (New-Object System.Text.UTF8Encoding($false)))
        Write-Host ("[OK] wrote {0} result(s): {1}" -f $results.Count, $OutFile) -ForegroundColor Green
    } else {
        $payload
    }
} else {
    $sb = New-Object System.Text.StringBuilder
    foreach ($r in $results) {
        Write-Host ''
        Write-Host ("===== {0} ({1}) =====" -f $r.Path, $r.Language) -ForegroundColor Cyan
        if (@($r.Lines).Count -eq 0) {
            Write-Host '(no text recognized - image resolution too low?)' -ForegroundColor Yellow
        }
        foreach ($l in $r.Lines) { Write-Host $l }
        [void]$sb.AppendLine(('===== {0} =====' -f $r.Path))
        [void]$sb.AppendLine($r.Text)
        # per-image .txt next to the source when an OutDir is given
        if (-not [string]::IsNullOrWhiteSpace($OutDir) -and [string]::IsNullOrWhiteSpace($Workbook)) {
            if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
            $txt = Join-Path $OutDir ([System.IO.Path]::GetFileNameWithoutExtension($r.Path) + '.txt')
            [System.IO.File]::WriteAllText($txt, $r.Text, (New-Object System.Text.UTF8Encoding($false)))
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($OutFile)) {
        [System.IO.File]::WriteAllText($OutFile, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
        Write-Host ("[OK] wrote combined text: {0}" -f $OutFile) -ForegroundColor Green
    }
}
