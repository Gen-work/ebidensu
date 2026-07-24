# ============================================================
#  OldSnapVerify.ps1
#
#  PURE library for the ProcessTime "old-snap 9->3 hand-verification"
#  feature -- NO Excel COM, NO GDI/System.Drawing, NO OCR, NO mapping I/O,
#  NO file I/O. Dot-source only (no param() block), so ProcessTime.ps1 can
#  dot-source it safely per CLAUDE.md's dot-source rule. ASCII source;
#  Japanese labels built from [char] code points (same convention as the
#  sibling audit-column module ProcessTimeCheck.ps1) so the file is
#  codepage-agnostic and independently unit-testable without ProjectLabels.
#
#  Background (docs/ProcessTime-OldSnap-Verify-Plan.md): the ja
#  Windows.Media.Ocr recognizer misreads MS Gothic digit '9' as '3' on the
#  HM batch page when the capture resolution is low. New snaps carry an
#  exact Ctrl+A page-text file (snap\<Stage>_HM\<correl>.txt) that is IMMUNE
#  to the misread, but a finite backlog of OLD snaps have only the low-res
#  PNG and fall back to OCR. This module supplies the pure decision + path
#  helpers used to triage that backlog: auto-confirm the confident majority
#  and route the rest to a fast human check (via the D1 correl-id -> snap
#  hyperlink). The COM/GDI wiring lives in ProcessTime.ps1 and is
#  static-checked only.
#
#  Convention: functions return plain values -- never return ,@(...) --
#  matching ProcessTimeParse.ps1 / ProcessTimeCheck.ps1.
# ============================================================

# ---------------------------------------------------------------------------
# Resolve-OldSnapImagePath
#   Builds (no I/O) the path to the standalone per-correl HM snap PNG for one
#   side: <WorkDir>\<DirPattern for side>\<correl>.png. <DirPattern>'s {0} is
#   the side stage ('GIFT'/'GFIX'); default 'snap\{0}_HM' matches the layout
#   HmSnap.ps1 / ProcessTime.ps1 already use. Returns $null when WorkDir or
#   CorrelId is blank or Side is not GIFT/GFIX; the caller does the Test-Path
#   existence check (this stays pure) and marks NoSnap when the file is
#   absent. Uses [IO.Path]::Combine (Windows path semantics on the office PC;
#   see the note in the body for why not Join-Path).
# ---------------------------------------------------------------------------
function Resolve-OldSnapImagePath {
    param(
        [string]$WorkDir,
        [string]$Side,
        [string]$CorrelId,
        [string]$DirPattern = 'snap\{0}_HM'
    )
    if ([string]::IsNullOrWhiteSpace($WorkDir)) { return $null }
    if ([string]::IsNullOrWhiteSpace($CorrelId)) { return $null }
    $s = ([string]$Side).Trim().ToUpperInvariant()
    if ($s -ne 'GIFT' -and $s -ne 'GFIX') { return $null }
    if ([string]::IsNullOrWhiteSpace($DirPattern)) { $DirPattern = 'snap\{0}_HM' }
    $sub = ($DirPattern -f $s)
    # [System.IO.Path]::Combine (not Join-Path) so a rooted Windows path like
    # 'C:\work' does not trigger Join-Path's drive-qualifier resolution, which
    # throws on a non-Windows CI host ("A drive with the name 'C' does not
    # exist"). Combine is pure string joining; on the office PC it yields the
    # normal all-backslash Windows path.
    return [System.IO.Path]::Combine($WorkDir, $sub, ("{0}.png" -f $CorrelId))
}

# ---------------------------------------------------------------------------
# ConvertTo-OldSnapDurationSeconds
#   Parses an 'HH:mm:ss' duration string (HH not clamped to 24 -- a run can
#   span more than a day, matching Get-ProcessDurationText) into a whole
#   number of seconds. Returns $null when blank or not the expected shape.
# ---------------------------------------------------------------------------
function ConvertTo-OldSnapDurationSeconds {
    param([string]$Duration)
    if ([string]::IsNullOrWhiteSpace($Duration)) { return $null }
    $m = [regex]::Match($Duration.Trim(), '^(\d+):([0-5]?\d):([0-5]?\d)$')
    if (-not $m.Success) { return $null }
    return ([int]$m.Groups[1].Value * 3600 + [int]$m.Groups[2].Value * 60 + [int]$m.Groups[3].Value)
}

# ---------------------------------------------------------------------------
# Test-OldSnapDurationArithmetic  (deterministic check, plan section 4.3)
#   The PS-side mirror of the output workbook's J column: does the written
#   duration equal (End - Start) to the second? Start/End are the formatted
#   'yyyy/MM/dd HH:mm:ss' stamps ProcessTime writes (Format-ProcessTimeStamp);
#   Duration is 'HH:mm:ss'. Returns:
#     $true   values are self-consistent (duration == end - start)
#     $false  values disagree (a likely OCR misread -- flag the row)
#     $null   cannot decide (a field is blank / unparseable -- a partial row;
#             leave it to the conservative default, never auto-confirm on it)
# ---------------------------------------------------------------------------
function Test-OldSnapDurationArithmetic {
    param([string]$Start, [string]$End, [string]$Duration)
    $durSec = ConvertTo-OldSnapDurationSeconds $Duration
    if ($null -eq $durSec) { return $null }
    # [string[]] cast is REQUIRED: a plain @(...) is object[], and the
    # DateTime.TryParseExact(string, string[], ...) overload is only selected
    # for a real string[] -- otherwise the binder falls back to the single-
    # format (string) overload and every parse fails.
    [string[]]$fmts = @('yyyy/MM/dd HH:mm:ss', 'yyyy/MM/dd H:mm:ss', 'yyyy-MM-dd HH:mm:ss')
    $ci = [System.Globalization.CultureInfo]::InvariantCulture
    $none = [System.Globalization.DateTimeStyles]::None
    [datetime]$st = [datetime]::MinValue
    [datetime]$en = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($Start, $fmts, $ci, $none, [ref]$st)) { return $null }
    if (-not [datetime]::TryParseExact($End,   $fmts, $ci, $none, [ref]$en)) { return $null }
    $span = $en - $st
    if ($span.TotalSeconds -lt 0) { return $false }
    return ([int][Math]::Round($span.TotalSeconds) -eq $durSec)
}

# ---------------------------------------------------------------------------
# Repair-ProcessTimeStartFromStamp  (deterministic 3<->9 correction, plan 4.3)
#   When the clean 14-digit data-creation datestamp (yyyyMMddHHmmss) is
#   present and its HH:mm differs from the OCR'd start time's HH:mm by ONLY a
#   3<->9 swap (every differing digit is a 3<->9 pair, at least one differs),
#   the ja OCR read is the wrong one -- adopt the datestamp's HH:mm for the
#   start time, keeping the OCR'd seconds unchanged. A difference that is NOT
#   purely 3<->9 is left untouched (some other discrepancy; not our bug).
#   Pure string logic, no [datetime]. Returns a hashtable:
#     Repaired      $true when Start was corrected
#     SwapDetected  $true when a pure 3<->9 HH:mm disagreement was found
#     Start         the (possibly corrected) 'HH:mm:ss' string
#   StartTimeOfDay must be a bare 'HH:mm:ss' (the time-of-day part); callers
#   that hold a full 'yyyy/MM/dd HH:mm:ss' stamp pass its last token.
# ---------------------------------------------------------------------------
function Repair-ProcessTimeStartFromStamp {
    param([string]$StartTimeOfDay, [string]$Datestamp)
    $res = @{ Repaired = $false; SwapDetected = $false; Start = $StartTimeOfDay }
    if ([string]::IsNullOrWhiteSpace($StartTimeOfDay)) { return $res }
    $sm = [regex]::Match($StartTimeOfDay.Trim(), '^(\d{2}):(\d{2}):(\d{2})$')
    if (-not $sm.Success) { return $res }
    $stamp = ([string]$Datestamp).Trim()
    if ($stamp -notmatch '^\d{14}$') { return $res }

    $startHHMM = $sm.Groups[1].Value + $sm.Groups[2].Value   # 4 digits: HHmm
    $ss        = $sm.Groups[3].Value
    $stampHHMM = $stamp.Substring(8, 4)                       # HHmm of yyyyMMddHHmmss
    if ($startHHMM -eq $stampHHMM) { return $res }            # already agree -> nothing to do

    $swap = $false
    for ($i = 0; $i -lt 4; $i++) {
        $a = $startHHMM[$i]; $b = $stampHHMM[$i]
        if ($a -eq $b) { continue }
        if (($a -eq '3' -and $b -eq '9') -or ($a -eq '9' -and $b -eq '3')) { $swap = $true; continue }
        return $res   # a non-3/9 difference -> not the 9->3 case; leave untouched
    }
    if (-not $swap) { return $res }
    $res.SwapDetected = $true
    $res.Repaired     = $true
    $res.Start        = ('{0}:{1}:{2}' -f $stampHHMM.Substring(0, 2), $stampHHMM.Substring(2, 2), $ss)
    return $res
}

# ---------------------------------------------------------------------------
# Get-OldSnapVerifyVerdict  (the triage decision -- plan section 4.2 step 5)
#   Decides one output row's kenshou (Verify) verdict from the row's OCR source,
#   whether a snap image exists, and the deterministic + (optional) image
#   checks. CONSERVATIVE BIAS: never auto-confirm unless every enabled check
#   is satisfied -- when in doubt, flag (a false flag costs one human glance;
#   a false auto-confirm ships a wrong value). Returns a stable verdict KEY;
#   Get-OldSnapVerifyLabel maps it to the display string.
#
#   Inputs:
#     Source        the row's per-side OCR source tag (ProcessTime.ps1's
#                   Resolve-ProcessTimeSide): 'archived' (from the .txt tier
#                   -- immune to 9->3, trusted), 'ocr' / 'ocr:<tier>' /
#                   'ocr-partial[:<tier>]' (OCR-derived -- in scope), 'none'.
#                   NOTE 'ocr:snap-png' is still OCR of a PNG, so it is
#                   triaged, NOT trusted -- only 'archived'/'txt' is immune.
#     SnapExists    a standalone/exported snap image exists for this row.
#     ArithmeticOk  Test-OldSnapDurationArithmetic result ($true/$false/$null).
#     DatestampSwap a pure 3<->9 datestamp-vs-start disagreement was found
#                   (Repair-ProcessTimeStartFromStamp .SwapDetected).
#     PixelResult   the D2 per-digit image check: 'ok' / 'ng' / '' (unknown /
#                   not run). Only consulted when PixelEnabled.
#     PixelEnabled  D2 image comparison is turned on (Phase 0 passed).
#
#   Verdict keys: 'Txt' | 'OcrOk' | 'NeedsCheck' | 'NoSnap'.
# ---------------------------------------------------------------------------
function Get-OldSnapVerifyVerdict {
    param(
        [string]$Source,
        [bool]$SnapExists,
        $ArithmeticOk = $null,
        [bool]$DatestampSwap = $false,
        [string]$PixelResult = '',
        [bool]$PixelEnabled = $false
    )
    $src = ([string]$Source).Trim().ToLowerInvariant()

    # The .txt tier is copied page text, not OCR -> immune to 9->3. Trust it.
    if ($src -eq 'archived' -or $src -eq 'txt') { return 'Txt' }

    # OCR-derived (or unknown source): needs a snap image to verify against.
    if (-not $SnapExists) { return 'NoSnap' }

    # A partial / missing read is never auto-confirmed.
    if ($src -like 'ocr-partial*' -or $src -eq 'none' -or [string]::IsNullOrWhiteSpace($src)) { return 'NeedsCheck' }

    # Deterministic checks (robust floor, ship regardless of D2).
    if ($ArithmeticOk -eq $false) { return 'NeedsCheck' }
    if ($DatestampSwap)           { return 'NeedsCheck' }

    # D2 image check (only when Phase 0 passed and it is enabled): require an
    # explicit 'ok'; 'ng' or an unknown/failed localization both flag.
    if ($PixelEnabled) {
        if ($PixelResult -ne 'ok') { return 'NeedsCheck' }
    }

    return 'OcrOk'
}

# ---------------------------------------------------------------------------
# Get-OldSnapVerifyLabel
#   Maps a verdict key to the operator-facing kenshou (Verify) cell string. ASCII tokens
#   ('txt', 'OCR-OK') are literals; the Japanese ones are built from [char]:
#     NeedsCheck -> 'you-kaku-nin' (needs check) U+8981 U+78BA U+8A8D
#     NoSnap     -> 'gazou-nashi'  (no image)    U+753B U+50CF + na U+306A shi U+3057
#   An unknown key yields '' (nothing written).
# ---------------------------------------------------------------------------
function Get-OldSnapVerifyLabel {
    param([string]$Verdict)
    switch ([string]$Verdict) {
        'Txt'        { return 'txt' }
        'OcrOk'      { return 'OCR-OK' }
        'NeedsCheck' { return ([string][char]0x8981 + [char]0x78BA + [char]0x8A8D) }
        'NoSnap'     { return ([string][char]0x753B + [char]0x50CF + [char]0x306A + [char]0x3057) }
        default      { return '' }
    }
}

# ---------------------------------------------------------------------------
# Get-OldSnapVerifyColumnSpec
#   The data-driven spec (header / width / number format) for the single kenshou (Verify)
#   (Verify) column appended after the A..H data and any I/J/K audit columns.
#   Its worksheet column INDEX is positional (it depends on whether the audit
#   columns are emitted), so the COM writer supplies it; keeping the header /
#   width / format here matches ProcessTimeCheck.ps1's data-driven pattern.
#     Header 'kenshou' (verify) U+691C U+8A3C. NumberFormat '@' = text, so a
#     value like 'OCR-OK' is never coerced. Width mirrors the audit columns.
# ---------------------------------------------------------------------------
function Get-OldSnapVerifyColumnSpec {
    $header = [string][char]0x691C + [char]0x8A3C
    return @{ Header = $header; Width = 12.0; NumberFormat = '@' }
}
