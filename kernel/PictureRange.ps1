#Requires -Version 5.1
# ============================================================
#  kernel/PictureRange.ps1
#
#  PURE. "The DF pictures show the same line ranges as the HOST pictures."
#  (gfix-recv, leader feedback 2026-10-12, docs/TODO.md). The HOST team's
#  pictures on the send-data sheet are fixed emulator shots with a 6-digit
#  line number at the left of every line. Read those numbers (OCR of the
#  number strip, done elsewhere), make sure they are what they claim to be
#  (consecutive, as many as the strip has text lines), map them onto the
#  CSV the DF window shows (the HOST data carries ONE header line the CSV
#  does not: HOST 1-17 -> DF 1-16, HOST n..m -> DF n-1..m-1), and plan one
#  DF capture per HOST picture: its line range, the window height that
#  shows exactly that many lines, the keys that bring that range on
#  screen. Anything that does not add up -> code 'unknown' with the reason,
#  never a guess (the gate asks the operator).
#
#  Dot-source only (no param(), ASCII source). No I/O, no OCR, no Win32:
#  the strip crop, the OCR call and the window sizing are the glue's job
#  (office-PC calibration: strip box, DF row pitch, DF key recipe).
#
#    ConvertFrom-EbiLineNumberText  OCR text -> line numbers + rejected lines
#    Test-EbiLineNumberRun          numbers consecutive? as many as the bands?
#    ConvertTo-EbiDfRange           HOST first..last -> DF first..last
#    Get-EbiDfRangePlan             HOST pictures -> DF capture plan / unknown
#
#  Hashtable reads are by index ($h['k']), never by dot: Set-StrictMode.
# ============================================================

function ConvertFrom-EbiLineNumberText {
    <#
      PURE. The OCR text of a line-number strip -> the numbers it shows.
        Text    the OCR output, one strip line per text line
        Digits  how many digits a line number has (6 on the HOST shots);
                0 = any length
      Spaces INSIDE a number are OCR artefacts and are removed ('0000 01'
      -> 000001); anything else that is not digits-only is NOT repaired
      (the ja recognizer reads 9 as 3 -- a wrong digit still parses, so
      the run checks below are what catch it) but listed in 'rejected' so
      the caller can warn. Blank lines are skipped silently.
      Returns @{ numbers = int[]; rejected = string[] }.
    #>
    param([string]$Text, [int]$Digits = 6)
    $numbers = New-Object System.Collections.ArrayList
    $rejected = New-Object System.Collections.ArrayList
    if ($null -eq $Text) { $Text = '' }
    foreach ($raw in ($Text -split "`r?`n")) {
        $t = ([string]$raw).Trim()
        if ($t -eq '') { continue }
        $compact = ($t -replace '\s+', '')
        $okShape = ($compact -match '^[0-9]+$') -and ($Digits -le 0 -or $compact.Length -eq $Digits)
        if ($okShape) { [void]$numbers.Add([int]$compact) } else { [void]$rejected.Add($t) }
    }
    return @{ numbers = [int[]]$numbers.ToArray(); rejected = [string[]]$rejected.ToArray() }
}

function Test-EbiLineNumberRun {
    <#
      PURE. Are these line numbers one picture's run?
        Numbers  as read, top to bottom
        Bands    how many text lines the strip really holds (ink bands);
                 0 = not known, skip that check
      A run is strictly consecutive (+1 each line) and, when Bands is
      given, exactly Bands long: a skipped or doubled OCR line and a
      misread digit both break one of the two. Returns
      @{ ok; first; last; count; reason; message } with reason one of
      '' / 'no_numbers' / 'not_consecutive' / 'count_mismatch'.
    #>
    param([int[]]$Numbers, [int]$Bands = 0)
    $n = @($Numbers)
    if ($n.Count -eq 0) {
        return @{ ok = $false; first = 0; last = 0; count = 0; reason = 'no_numbers'; message = 'no line number read' }
    }
    for ($i = 1; $i -lt $n.Count; $i++) {
        if ([int]$n[$i] -ne ([int]$n[$i - 1] + 1)) {
            $msg = ('line {0} reads {1:D6} after {2:D6} (expected {3:D6})' -f ($i + 1), [int]$n[$i], [int]$n[$i - 1], ([int]$n[$i - 1] + 1))
            return @{ ok = $false; first = [int]$n[0]; last = [int]$n[$n.Count - 1]; count = $n.Count; reason = 'not_consecutive'; message = $msg }
        }
    }
    if ($Bands -gt 0 -and $n.Count -ne $Bands) {
        $msg = ('{0} line numbers read but the strip has {1} text lines' -f $n.Count, $Bands)
        return @{ ok = $false; first = [int]$n[0]; last = [int]$n[$n.Count - 1]; count = $n.Count; reason = 'count_mismatch'; message = $msg }
    }
    return @{ ok = $true; first = [int]$n[0]; last = [int]$n[$n.Count - 1]; count = $n.Count; reason = ''; message = '' }
}

function ConvertTo-EbiDfRange {
    <#
      PURE. A HOST picture's line range -> the DF (CSV) line range showing
      the same records. The HOST data has HeaderLines header lines at the
      top that the CSV does not, so every number shifts down by that much
      and the first picture starts at CSV line 1 (HOST 1-17 -> DF 1-16;
      HOST 878-896 -> DF 877-895). HeaderLines = 1 was confirmed on ONE
      workbook (FJODWCP2) -- it is a parameter, not a law.
      Returns @{ first; last; lines } (lines <= 0 when the range is empty).
    #>
    param([int]$First, [int]$Last, [int]$HeaderLines = 1)
    $f = [Math]::Max(1, $First - $HeaderLines)
    $l = $Last - $HeaderLines
    return @{ first = $f; last = $l; lines = ($l - $f + 1) }
}

function Get-EbiDfRangePlan {
    <#
      PURE. The DF capture plan that mirrors the HOST pictures.
        HostPictures  @(@{ text; bands; name }) in sheet order: 'text' the
                      OCR output of the picture's line-number strip,
                      'bands' how many text lines the strip holds (0 = not
                      counted), 'name' for messages (optional)
        DfLines       the DF (CSV) file's line count; 0 = not known
        HeaderLines   HOST header lines the CSV lacks (ConvertTo-EbiDfRange)
        ScreenLines   data lines the DF window shows at BaseHeight (16 at
                      429 px, the sample evidence)
        BaseHeight    the DF window height showing ScreenLines lines
        RowPitch      px per extra line (office-PC calibration); 0 = not
                      calibrated -> only ScreenLines-line pictures can be
                      sized
        KeysFirst     keys that show line 1 at the top (^{HOME})
        KeysPageDown  one page down ({PGDN}); a picture starting at line
                      k*lines+1 is KeysFirst + k pages
        KeysEnd       keys that show the last lines (the END view)
        Digits        line-number width for the OCR parse
      Returns @{ ok = $true; code = 'ok' | 'unknown'; reason; message;
                 ranges = @(@{ index; name; hostFirst; hostLast; first;
                 last; lines; height; keys }); warnings = string[] }.
      'unknown' names the first thing that did not add up (reason:
      no_host_pictures / no_numbers / not_consecutive / count_mismatch /
      empty_range / ranges_overlap / df_lines_mismatch /
      row_pitch_uncalibrated / no_key_recipe); ranges still holds what
      was derived so the gate can show it. Rejected OCR lines are
      warnings, not failures: a wrong digit shows up as not_consecutive.
    #>
    param(
        $HostPictures,
        [int]$DfLines = 0,
        [int]$HeaderLines = 1,
        [int]$ScreenLines = 16,
        [int]$BaseHeight = 429,
        [int]$RowPitch = 0,
        [string]$KeysFirst = '^{HOME}',
        [string]$KeysPageDown = '{PGDN}',
        [string]$KeysEnd = '{END}{PGUP}{DOWN}{DOWN}',
        [int]$Digits = 6
    )
    $pics = @(@($HostPictures) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $ranges = New-Object System.Collections.ArrayList
    $warnings = New-Object System.Collections.ArrayList
    $unknown = {
        param([string]$Reason, [string]$Message)
        return @{ ok = $true; code = 'unknown'; reason = $Reason; message = $Message; ranges = $ranges.ToArray(); warnings = [string[]]$warnings.ToArray() }
    }
    if ($pics.Count -eq 0) { return (& $unknown 'no_host_pictures' 'no HOST picture to copy the ranges from') }
    $i = 0
    foreach ($p in $pics) {
        $i++
        $name = if ($p.Contains('name') -and -not [string]::IsNullOrWhiteSpace([string]$p['name'])) { [string]$p['name'] } else { ('HOST picture {0}' -f $i) }
        $text = if ($p.Contains('text')) { [string]$p['text'] } else { '' }
        $bands = if ($p.Contains('bands') -and $null -ne $p['bands']) { [int]$p['bands'] } else { 0 }
        $parsed = ConvertFrom-EbiLineNumberText -Text $text -Digits $Digits
        foreach ($r in @($parsed['rejected'])) { [void]$warnings.Add(('{0}: strip line not a line number: "{1}"' -f $name, $r)) }
        $run = Test-EbiLineNumberRun -Numbers $parsed['numbers'] -Bands $bands
        if (-not $run['ok']) { return (& $unknown $run['reason'] ('{0}: {1}' -f $name, $run['message'])) }
        $df = ConvertTo-EbiDfRange -First $run['first'] -Last $run['last'] -HeaderLines $HeaderLines
        $entry = @{ index = $i; name = $name; hostFirst = [int]$run['first']; hostLast = [int]$run['last']
                    first = [int]$df['first']; last = [int]$df['last']; lines = [int]$df['lines']; height = 0; keys = '' }
        [void]$ranges.Add($entry)
        if ([int]$df['lines'] -le 0) { return (& $unknown 'empty_range' ('{0}: HOST {1:D6}-{2:D6} leaves no DF line' -f $name, $run['first'], $run['last'])) }
        if ($ranges.Count -ge 2) {
            $prev = $ranges[$ranges.Count - 2]
            if ([int]$entry['first'] -le [int]$prev['last']) {
                return (& $unknown 'ranges_overlap' ('{0}: DF {1}-{2} overlaps the previous picture ({3}-{4})' -f $name, $entry['first'], $entry['last'], $prev['first'], $prev['last']))
            }
        }
    }
    if ($DfLines -gt 0) {
        $lastEntry = $ranges[$ranges.Count - 1]
        if ([int]$lastEntry['last'] -ne $DfLines) {
            return (& $unknown 'df_lines_mismatch' ('the last HOST picture ends at {0:D6} (DF {1}) but the DF file has {2} lines' -f $lastEntry['hostLast'], $lastEntry['last'], $DfLines))
        }
    }
    foreach ($e in $ranges) {
        $lines = [int]$e['lines']
        if ($lines -eq $ScreenLines) { $e['height'] = $BaseHeight }
        elseif ($RowPitch -gt 0) { $e['height'] = $BaseHeight + $RowPitch * ($lines - $ScreenLines) }
        else { return (& $unknown 'row_pitch_uncalibrated' ('{0}: {1} lines a picture but the DF row pitch is not calibrated (only {2}-line pictures can be sized)' -f $e['name'], $lines, $ScreenLines)) }
        if ([int]$e['first'] -eq 1) { $e['keys'] = $KeysFirst }
        elseif ($DfLines -gt 0 -and [int]$e['last'] -eq $DfLines) { $e['keys'] = $KeysEnd }
        elseif ((([int]$e['first'] - 1) % $lines) -eq 0) {
            $pages = ([int]$e['first'] - 1) / $lines
            $e['keys'] = $KeysFirst + ($KeysPageDown * $pages)
        }
        else { return (& $unknown 'no_key_recipe' ('{0}: no key sequence shows DF lines {1}-{2} ({3} a page) -- a line jump is needed' -f $e['name'], $e['first'], $e['last'], $lines)) }
    }
    return @{ ok = $true; code = 'ok'; reason = ''; message = ''; ranges = $ranges.ToArray(); warnings = [string[]]$warnings.ToArray() }
}
