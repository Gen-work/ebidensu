#Requires -Version 5.1
# ============================================================
# Convert-PomBook.ps1 - bring old-layout POM evidence workbooks into the
# delivered POM layout (GFIX result sheets twice: before / after the fix).
#
# The plan comes from PomBook.ps1 (pure); this script performs it with Excel
# COM, so pictures, red boxes and yellow rows travel with Worksheet.Copy.
#
# The original is NEVER modified: each book is copied to -OutDir first and
# only the copy is changed. A book already in the POM layout is skipped; a
# book with neither layout is reported and left alone.
#
# Usage:
#   .\Convert-PomBook.ps1 -Path C:\work\pom                 # every *.xlsx
#   .\Convert-PomBook.ps1 -Path C:\work\pom\J4..._YJDSWM39_POM....xlsx
#   .\Convert-PomBook.ps1 -Path C:\work\pom -DryRun         # plan only, no Excel
#   .\Convert-PomBook.ps1 -Path C:\work\pom -OutDir C:\work\pom\new
# Default -OutDir: <folder of the first input>\pom_converted.
#
# Has param(): call with -File or &, never dot-source.
# ============================================================

param(
    [Parameter(Mandatory = $true)][string[]]$Path,
    [string]$OutDir = '',
    [string]$Filter = '*.xlsx',
    [switch]$DryRun
)

$dryRunFlag = [bool]$DryRun.IsPresent
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'ProjectLabels.ps1')
. (Join-Path $PSScriptRoot 'PomBook.ps1')
. (Join-Path $PSScriptRoot 'ExcelHelpers.ps1')

# ---- collect the books ----
$books = New-Object System.Collections.ArrayList
foreach ($p in $Path) {
    if (Test-Path -LiteralPath $p -PathType Container) {
        Get-ChildItem -LiteralPath $p -Filter $Filter -File |
            Where-Object { $_.Name -notlike '~$*' } |
            ForEach-Object { [void]$books.Add($_.FullName) }
    } elseif (Test-Path -LiteralPath $p -PathType Leaf) {
        [void]$books.Add((Resolve-Path -LiteralPath $p).Path)
    } else {
        Write-Host ('  [warn ] not found: ' + $p) -ForegroundColor Yellow
    }
}
if ($books.Count -eq 0) { Write-Host 'No workbook to convert.'; exit 2 }

if (-not $OutDir) { $OutDir = Join-Path (Split-Path $books[0] -Parent) 'pom_converted' }
if (-not $dryRunFlag -and -not (Test-Path -LiteralPath $OutDir)) {
    [void](New-Item -ItemType Directory -Path $OutDir)
}

# Sheet names without opening Excel: read xl/workbook.xml from the zip.
function Get-PomBookSheetNamesFromZip([string]$file) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($file)
    try {
        $entry = $zip.GetEntry('xl/workbook.xml')
        $sr = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        try { [xml]$x = $sr.ReadToEnd() } finally { $sr.Dispose() }
        return @($x.workbook.sheets.sheet | ForEach-Object { $_.name })
    } finally { $zip.Dispose() }
}

# The new before/after compare sheet: the look (column widths, font, A1
# file-name formula, B3 data name) of GIFT-vs-GFIX compare, nothing else.
function New-PomBookCompareSheet($wb, $action) {
    $last = $wb.Worksheets.Item($wb.Worksheets.Count)
    if ($action.template) {
        $src = Get-SheetByName $wb $action.template
        $src.Copy([System.Reflection.Missing]::Value, $last)
        $ws = $wb.Worksheets.Item($wb.Worksheets.Count)
        # Shapes lists top-level items only: a Ctrl+G group (picture + red boxes,
        # as on the JJDSWPA1 sample's compare sheet) goes with its group.
        for ($i = $ws.Shapes.Count; $i -ge 1; $i--) { $ws.Shapes.Item($i).Delete() }
        $keep = @{ 'A1' = $ws.Range('A1').Formula; 'B3' = $ws.Range('B3').Formula }
        [void]$ws.Cells.ClearContents()
        foreach ($k in $keep.Keys) { $ws.Range($k).Formula = $keep[$k] }
    } else {
        $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $last)
        $ws.Range('A1').Formula = '=RIGHT(CELL("filename",A1),LEN(CELL("filename",A1))-FIND("]",CELL("filename",A1)))'
    }
    $ws.Name = $action.name
    return $ws
}

$P = Get-PomBookLabels
$excel = $null
$summary = New-Object System.Collections.ArrayList
try {
    foreach ($book in $books) {
        $name = Split-Path $book -Leaf
        $plan = Get-PomBookPlan -SheetNames (Get-PomBookSheetNamesFromZip $book)
        if ($plan.status -eq 'already') {
            Write-Host ('  [skip ] ' + $name + ' : already in the POM layout')
            [void]$summary.Add(@($name, 'already')); continue
        }
        if ($plan.status -eq 'unknown') {
            Write-Host ('  [warn ] ' + $name + ' : missing ' + ($plan.missing -join ', ') + ' -- left alone') -ForegroundColor Yellow
            [void]$summary.Add(@($name, 'unknown')); continue
        }

        $out = Join-Path $OutDir $name
        if ($dryRunFlag) {
            Write-Host ('  [plan ] ' + $name)
            foreach ($a in $plan.actions) {
                switch ($a.op) {
                    'delete'     { Write-Host ('            delete  ' + $a.sheet) }
                    'rename'     { Write-Host ('            rename  ' + $a.sheet + ' -> ' + $a.name) }
                    'copyAfter'  { Write-Host ('            copy    ' + $a.sheet + ' -> ' + $a.name) }
                    'newCompare' { Write-Host ('            new     ' + $a.name) }
                }
            }
            $order = Get-PomBookResultOrder -SheetNames (Get-PomBookSheetNamesFromZip $book) -Actions $plan.actions
            Write-Host ('            order   ' + ($order -join ' | '))
            [void]$summary.Add(@($name, 'plan')); continue
        }

        if (Test-Path -LiteralPath $out) {
            Write-Host ('  [skip ] ' + $name + ' : ' + $out + ' already exists (delete it to convert again)') -ForegroundColor Yellow
            [void]$summary.Add(@($name, 'exists')); continue
        }
        Copy-Item -LiteralPath $book -Destination $out

        if ($null -eq $excel) {
            $excel = New-ExcelApp
            # Visible so the office save add-in's dialog (classification) can be answered.
            $excel.Visible = $true
            $excel.DisplayAlerts = $false
        }
        $wb = Open-Workbook $excel $out
        $ok = $false
        try {
            foreach ($a in $plan.actions) {
                switch ($a.op) {
                    'delete' {
                        (Get-SheetByName $wb $a.sheet).Delete()
                    }
                    'rename' {
                        (Get-SheetByName $wb $a.sheet).Name = $a.name
                    }
                    'copyAfter' {
                        $src = Get-SheetByName $wb $a.sheet
                        $src.Copy([System.Reflection.Missing]::Value, $src)
                        $copy = $wb.Worksheets.Item($src.Index + 1)
                        $copy.Name = $a.name
                    }
                    'newCompare' {
                        [void](New-PomBookCompareSheet $wb $a)
                    }
                }
            }
            $wb.Worksheets.Item(1).Activate()
            $wb.Save()
            $ok = $true
        } finally {
            Close-Workbook $wb $false
        }
        if ($ok) {
            Write-Host ('  [ok   ] ' + $name + ' -> ' + $out) -ForegroundColor Green
            [void]$summary.Add(@($name, 'ok'))
        }
    }
} finally {
    if ($null -ne $excel) { Close-ExcelApp $excel }
}

Write-Host ''
Write-Host ('Done: ' + (@($summary | Where-Object { $_[1] -eq 'ok' }).Count) + ' converted, ' +
    (@($summary | Where-Object { $_[1] -eq 'already' }).Count) + ' already POM, ' +
    (@($summary | Where-Object { $_[1] -eq 'unknown' -or $_[1] -eq 'exists' }).Count) + ' left alone' +
    $(if ($dryRunFlag) { ' (dry run, nothing written)' } else { '; output in ' + $OutDir }))
Write-Host ('Rerun evidence goes into ' + $P.GfixSendAfter + ' / ' + $P.GfixRecvAfter + ' (now a copy of the first run).')
