# ============================================================
#  ProcessTimeParse.ps1
#
#  PURE library for the ProcessTime phase -- NO Excel COM, NO OCR calls,
#  NO mapping I/O. Dot-source only (no param() block).
#
#  ProcessTime.ps1 extracts each correl's HM batch processing start/end
#  time (and derives the duration) from the GIFT/GFIX receive-result
#  evidence screenshot already inserted into the evidence workbook
#  (sheet SheetGiftRecv / SheetGfixRecv, see ProjectLabels.ps1). Two
#  sources feed the same shape of result, cheapest/most-accurate first:
#    1. the archived Ctrl+A page text saved by HmSnap.ps1 at snap time
#       (WorkDir\snap\GIFT_HM\<correl>.txt / GFIX_HM), re-parsed with
#       SnapVerify.ps1's ConvertFrom-HmPageText (exact, TAB-anchored).
#    2. OCR of the evidence picture itself (ProcessTime.ps1's
#       Export-CorrelPicture + Windows OCR), when no archived text is
#       available (e.g. only the delivered J4 workbook is on hand). OCR
#       output never carries real TAB characters, so this file's own
#       ConvertFrom-ProcessTimeOcrLines is a looser, anchor-based reader
#       (find two datetime tokens + a status literal on one line) rather
#       than reusing ConvertFrom-HmPageText's strict field-split.
#
#  Both tiers produce objects carrying StartTime/EndTime/Status, so
#  Get-NewestProcessTimeRow (newest-by-StartTime, matching this project's
#  established newest-wins convention -- see SnapVerify.Test-HmAbend)
#  works the same regardless of which tier produced the rows.
#
#  Convention: functions return plain arrays -- never return ,@(...)
#  because callers wrap calls in @() and that nests in PS 5.1.
# ============================================================

# Japanese status literals (source stays ASCII; see SnapVerify.ps1's
# identical constants -- duplicated here on purpose so this file has no
# cross-file dependency and stays independently unit-testable).
$script:PT_Normal = [char]0x6B63 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # seijo-shuuryo (normal end)
$script:PT_Abend  = [char]0x7570 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # ijo-shuuryo (abend)

# ---------------------------------------------------------------------------
# Get-ProcessDurationText
#   Formats the elapsed time between two [datetime] values as HH:mm:ss
#   (HH is not clamped to 24 -- a run can span more than a day). Returns ''
#   when either value is missing or EndTime is before StartTime (garbled
#   OCR read; never invent a negative duration).
# ---------------------------------------------------------------------------
function Get-ProcessDurationText {
    param($StartTime, $EndTime)
    if ($null -eq $StartTime -or $null -eq $EndTime) { return '' }
    $span = $EndTime - $StartTime
    if ($span.TotalSeconds -lt 0) { return '' }
    $totalHours = [int][Math]::Floor($span.TotalHours)
    return ('{0:00}:{1:00}:{2:00}' -f $totalHours, $span.Minutes, $span.Seconds)
}

# ---------------------------------------------------------------------------
# ConvertFrom-ProcessTimeOcrLines
#   Anchor-based reader for OCR'd HM page text: a real table column split
#   is unreliable once OCR has collapsed variable-width whitespace, so
#   instead of splitting fields this scans each reconstructed line for two
#   'yyyy/MM/dd HH:mm:ss' tokens (start, end, in that order) and an
#   optional normal-end/abend literal, matching this project's general OCR
#   philosophy of anchoring on distinctive substrings rather than trusting
#   column position (see GfixLog.ps1 / Compare-SendRecordCheck).
#
#   $Lines should already be RECONSTRUCTED rows (one table row per string),
#   e.g. via SendMetadata.ps1's ConvertTo-SendRowLines against the OCR
#   result's word boxes -- a raw OcrResult.Text line split can fragment one
#   wide HM row across several lines.
#
#   Returns plain array of PSCustomObject { StartTime; EndTime; Status; RawLine }.
# ---------------------------------------------------------------------------
function ConvertFrom-ProcessTimeOcrLines {
    param([string[]]$Lines)

    $dtToken = '\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2}'
    $dtFmt   = 'yyyy/MM/dd HH:mm:ss'
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $out     = [System.Collections.Generic.List[object]]::new()

    foreach ($line in @($Lines)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $dtMatches = @([regex]::Matches($line, $dtToken))
        if ($dtMatches.Count -lt 2) { continue }

        $startTime = $null; $endTime = $null
        $startText = ($dtMatches[0].Value -replace '\s+', ' ').Trim()
        $endText   = ($dtMatches[1].Value -replace '\s+', ' ').Trim()
        try { $startTime = [datetime]::ParseExact($startText, $dtFmt, $culture) } catch {}
        try { $endTime   = [datetime]::ParseExact($endText,   $dtFmt, $culture) } catch {}
        if ($null -eq $startTime -or $null -eq $endTime) { continue }

        $status = ''
        if ($line.Contains($script:PT_Normal)) { $status = $script:PT_Normal }
        elseif ($line.Contains($script:PT_Abend)) { $status = $script:PT_Abend }

        $out.Add([PSCustomObject]@{
            StartTime = $startTime
            EndTime   = $endTime
            Status    = $status
            RawLine   = $line
        })
    }

    return $out.ToArray()
}

# ---------------------------------------------------------------------------
# Get-NewestProcessTimeRow
#   Returns the row with the latest StartTime (newest-wins, matching
#   Test-HmAbend / GIFT_MQ's established convention), or $null when -Rows
#   is empty. Works uniformly on rows from ConvertFrom-HmPageText (archived
#   text tier) or ConvertFrom-ProcessTimeOcrLines (OCR tier) -- both carry
#   a StartTime property.
# ---------------------------------------------------------------------------
function Get-NewestProcessTimeRow {
    param([object[]]$Rows)
    $rows = @($Rows | Where-Object { $null -ne $_ -and $null -ne $_.StartTime })
    if ($rows.Count -eq 0) { return $null }
    return @($rows | Sort-Object StartTime -Descending)[0]
}
