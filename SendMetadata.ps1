# ============================================================
#  SendMetadata.ps1
#
#  PURE send-side evidence metadata helpers for the SendVsGift
#  Stage 2 OCR flow -- NO Excel COM, NO WinRT, NO file I/O.
#  Dot-source only (no param() block). Unit-tested by
#  Tests\Test-SendMetadata.ps1.
#
#  Pipeline position:
#    EvidenceImageExport.ps1  -> PNGs of the send-sheet screenshots
#    OcrWindows.ps1           -> OCR lines (Text + word boxes)
#    SendMetadata.ps1 (here)  -> parse lines into a record parallel to
#                                gift_metadata.csv and compare the two
#
#  NOTE on spacing: the Japanese Windows OCR recognizer drops spaces
#  between tokens. Get-SendLineTextFromWords rebuilds them from the
#  word bounding boxes (X/Width) so fixed-position record text can be
#  approximated and first/last tokens stay separable.
# ============================================================

function Get-SendFirstToken {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $parts = @($Text.Trim() -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts.Count -eq 0) { return '' }
    return $parts[0]
}

# Rebuilds one text line from OCR word objects (Text/X/Width), inserting
# spaces sized from the running average character width. A gap wider than
# GapRatio * charWidth becomes round(gap / charWidth) spaces (min 1).
function Get-SendLineTextFromWords {
    param($Words, [double]$GapRatio = 0.5)
    $ws = @(@($Words) | Where-Object { $null -ne $_ } | Sort-Object { [double]$_.X })
    if ($ws.Count -eq 0) { return '' }
    $sb = New-Object System.Text.StringBuilder
    $prevEnd = $null
    $charW = 0.0
    foreach ($w in $ws) {
        $text = [string]$w.Text
        $x = 0.0
        $width = 0.0
        try { $x = [double]$w.X } catch {}
        try { $width = [double]$w.Width } catch {}
        $len = [Math]::Max(1, $text.Length)
        if ($width -gt 0) { $charW = $width / $len }
        if ($null -ne $prevEnd -and $charW -gt 0) {
            $gap = $x - $prevEnd
            if ($gap -ge ($GapRatio * $charW)) {
                $n = [int][Math]::Round($gap / $charW)
                if ($n -lt 1) { $n = 1 }
                [void]$sb.Append(' ' * $n)
            }
        }
        [void]$sb.Append($text)
        $prevEnd = $x + $width
    }
    return $sb.ToString()
}

# Normalizes OCR output to plain trimmed text lines. Accepts plain strings
# or objects shaped like OcrWindows.ps1 lines (.Text and optional .Words).
# When word boxes exist they win, so rebuilt spacing replaces the raw Text.
function ConvertTo-SendTextLines {
    param($OcrLines, [double]$GapRatio = 0.5)
    $out = @()
    foreach ($ln in @($OcrLines)) {
        if ($null -eq $ln) { continue }
        $text = ''
        if ($ln -is [string]) {
            $text = $ln
        } else {
            $words = $null
            if ($ln.PSObject.Properties.Name -contains 'Words') { $words = @($ln.Words) }
            if ($null -ne $words -and $words.Count -gt 0) {
                $text = Get-SendLineTextFromWords $words $GapRatio
            } elseif ($ln.PSObject.Properties.Name -contains 'Text') {
                $text = [string]$ln.Text
            }
        }
        $text = ([string]$text).TrimEnd()
        if (-not [string]::IsNullOrWhiteSpace($text)) { $out += $text }
    }
    return ,@($out)
}

# Detects the standard 0-byte screenshot pattern.
# TODO(Stage 2): tighten the default once representative 0-byte SEND
# screenshots are available; override via SendVsGift.ZeroBytePattern.
function Test-SendZeroByteText {
    param([string[]]$TextLines, [string]$Pattern = '')
    if ([string]::IsNullOrWhiteSpace($Pattern)) {
        $Pattern = '(^|[^0-9])0\s*(byte|bytes)([^a-z]|$)'
    }
    foreach ($t in @($TextLines)) {
        if ([string]$t -match $Pattern) { return $true }
    }
    return $false
}

# Host list screens print a leading row number on each record line; the
# largest leading integer is the best row-count guess. 0 = unknown.
function Get-SendRowNumberGuess {
    param([string[]]$TextLines)
    $max = 0
    foreach ($t in @($TextLines)) {
        $m = [regex]::Match([string]$t, '^\s*(\d{1,9})\b')
        if (-not $m.Success) { continue }
        $v = 0
        if ([int]::TryParse($m.Groups[1].Value, [ref]$v) -and $v -gt $max) { $max = $v }
    }
    return $max
}

# Assembles one send_metadata.csv row, field-parallel to gift_metadata.csv.
# Confidence is heuristic: the Windows OCR API exposes no per-word score,
# so we score by how many comparison fields were actually parsed.
function Build-SendMetadataRecord {
    param(
        [string]$CorrelIdS,
        [string]$ExcelName,
        [int]$ImageCount,
        [string[]]$TextLines,
        [string]$ZeroBytePattern = ''
    )
    $lines = @(@($TextLines) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $first = if ($lines.Count -gt 0) { [string]$lines[0] } else { '' }
    $last  = if ($lines.Count -gt 0) { [string]$lines[$lines.Count - 1] } else { '' }
    $zero  = Test-SendZeroByteText $lines $ZeroBytePattern
    $rows  = Get-SendRowNumberGuess $lines

    $conf = 0.0
    if ($lines.Count -gt 0) { $conf += 0.4 }
    if ($rows -gt 0 -or $zero) { $conf += 0.3 }
    if ((Get-SendFirstToken $first) -ne '') { $conf += 0.3 }

    [pscustomobject]@{
        CorrelIdS        = [string]$CorrelIdS
        ExcelName        = [string]$ExcelName
        ImageCount       = [int]$ImageCount
        OcrLineCount     = [int]$lines.Count
        ZeroByte         = [bool]$zero
        RowNumberGuess   = [int]$rows
        FirstRecord      = $first
        LastRecord       = $last
        FirstRecordToken = Get-SendFirstToken $first
        LastRecordToken  = Get-SendFirstToken $last
        Confidence       = [Math]::Round($conf, 2)
        MetadataVersion  = '2'
    }
}

# Compares a send record (Build-SendMetadataRecord) against one
# gift_metadata.csv row. Each check is match / mismatch / unknown;
# absence of OCR evidence is "unknown", never "mismatch".
# Verdict: any mismatch -> mismatch; zero-byte agreement or at least
# MinMatches matches -> match; otherwise unknown (manual review stays).
function Compare-SendGiftMetadata {
    param($Send, $Gift, [int]$MinMatches = 2)
    $checks = @()

    $giftZero = $false
    try {
        if ($Gift.PSObject.Properties.Name -contains 'SizeBytes') { $giftZero = ([long]$Gift.SizeBytes -eq 0) }
    } catch {}
    $sendZero = $false
    if ($Send.PSObject.Properties.Name -contains 'ZeroByte') {
        $zv = [string]$Send.ZeroByte
        $sendZero = ($zv -eq 'True' -or $zv -eq 'true' -or $zv -eq '1')
    }
    $zeroMatched = $false
    if ($sendZero -or $giftZero) {
        $st = 'unknown'
        if ($sendZero -and $giftZero) { $st = 'match'; $zeroMatched = $true }
        elseif ($sendZero -and -not $giftZero) { $st = 'mismatch' }
        $checks += [pscustomobject]@{ Name = 'ZeroByte'; Send = [string]$sendZero; Gift = [string]$giftZero; Status = $st }
    }

    $sendRows = 0
    try {
        if ($Send.PSObject.Properties.Name -contains 'RowNumberGuess') { $sendRows = [int]$Send.RowNumberGuess }
    } catch {}
    $giftRows = 0
    try {
        if ($Gift.PSObject.Properties.Name -contains 'MaxRowNumber') { $giftRows = [int]$Gift.MaxRowNumber }
    } catch {}
    $st = 'unknown'
    if ($sendRows -gt 0 -and $giftRows -gt 0) {
        $st = if ($sendRows -eq $giftRows) { 'match' } else { 'mismatch' }
    }
    $checks += [pscustomobject]@{ Name = 'RowNumber'; Send = [string]$sendRows; Gift = [string]$giftRows; Status = $st }

    foreach ($tok in @('FirstRecordToken', 'LastRecordToken')) {
        $sv = ''
        $gv = ''
        if ($Send.PSObject.Properties.Name -contains $tok) { $sv = [string]$Send.$tok }
        if ($Gift.PSObject.Properties.Name -contains $tok) { $gv = [string]$Gift.$tok }
        $st = 'unknown'
        if ($sv -ne '' -and $gv -ne '') {
            $st = if ($sv -eq $gv) { 'match' } else { 'mismatch' }
        }
        $checks += [pscustomobject]@{ Name = $tok; Send = $sv; Gift = $gv; Status = $st }
    }

    $matchCount    = @($checks | Where-Object { $_.Status -eq 'match' }).Count
    $mismatchCount = @($checks | Where-Object { $_.Status -eq 'mismatch' }).Count
    $unknownCount  = @($checks | Where-Object { $_.Status -eq 'unknown' }).Count

    $verdict = 'unknown'
    if ($mismatchCount -gt 0) { $verdict = 'mismatch' }
    elseif ($zeroMatched -or $matchCount -ge $MinMatches) { $verdict = 'match' }

    [pscustomobject]@{
        Checks        = @($checks)
        Verdict       = $verdict
        MatchCount    = [int]$matchCount
        MismatchCount = [int]$mismatchCount
        UnknownCount  = [int]$unknownCount
    }
}
