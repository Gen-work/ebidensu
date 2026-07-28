#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'OldSnapVerify.ps1')

Reset-Tests 'OldSnapVerify'

# Japanese labels rebuilt here from [char] so the test is codepage-agnostic
# too and independently pins the exact code points the module emits.
$lblNeeds  = [string][char]0x8981 + [char]0x78BA + [char]0x8A8D   # you-kaku-nin (needs check)
$lblNoSnap = [string][char]0x753B + [char]0x50CF + [char]0x306A + [char]0x3057   # gazou-nashi (no image)
$hdrVerify = [string][char]0x691C + [char]0x8A3C                  # kenshou (verify)

# ---------------------------------------------------------------------------
# Resolve-OldSnapImagePath : pure path build (no I/O).
# ---------------------------------------------------------------------------
$p = Resolve-OldSnapImagePath -WorkDir 'C:\work' -Side 'GIFT' -CorrelId 'JIGPC06S'
Assert-True ($p -like '*GIFT_HM*') 'GIFT side resolves under the GIFT_HM snap dir'
Assert-True ($p -like '*JIGPC06S.png') 'the file name is <correl>.png'
$pg = Resolve-OldSnapImagePath -WorkDir 'C:\work' -Side 'gfix' -CorrelId 'ABC'
Assert-True ($pg -like '*GFIX_HM*') 'side is case-insensitive and maps gfix -> GFIX_HM'
Assert-Equal $null (Resolve-OldSnapImagePath -WorkDir '' -Side 'GIFT' -CorrelId 'X') 'blank WorkDir -> null'
Assert-Equal $null (Resolve-OldSnapImagePath -WorkDir 'C:\w' -Side 'GIFT' -CorrelId '') 'blank correl -> null'
Assert-Equal $null (Resolve-OldSnapImagePath -WorkDir 'C:\w' -Side 'DF' -CorrelId 'X') 'unknown side -> null'
$pp = Resolve-OldSnapImagePath -WorkDir 'C:\w' -Side 'GIFT' -CorrelId 'X' -DirPattern 'archive\{0}'
Assert-True ($pp -like '*archive*GIFT*X.png') 'a custom DirPattern is honored'

# ---------------------------------------------------------------------------
# ConvertTo-OldSnapDurationSeconds : HH:mm:ss -> seconds, >24h allowed.
# ---------------------------------------------------------------------------
Assert-Equal 45   (ConvertTo-OldSnapDurationSeconds '00:00:45') '45s'
Assert-Equal 94   (ConvertTo-OldSnapDurationSeconds '00:01:34') '1m34s = 94s'
Assert-Equal 90190 (ConvertTo-OldSnapDurationSeconds '25:03:10') 'over 24h is not clamped'
Assert-Equal $null (ConvertTo-OldSnapDurationSeconds '') 'blank -> null'
Assert-Equal $null (ConvertTo-OldSnapDurationSeconds 'nope') 'garbage -> null'

# ---------------------------------------------------------------------------
# Test-OldSnapDurationArithmetic : duration == end - start (to the second).
# ---------------------------------------------------------------------------
Assert-Equal $true  (Test-OldSnapDurationArithmetic -Start '2026/07/23 03:59:07' -End '2026/07/23 03:59:52' -Duration '00:00:45') 'consistent 45s span'
Assert-Equal $false (Test-OldSnapDurationArithmetic -Start '2026/07/23 03:59:07' -End '2026/07/23 03:59:52' -Duration '00:00:44') 'off-by-one duration flags false'
Assert-Equal $false (Test-OldSnapDurationArithmetic -Start '2026/07/23 03:59:52' -End '2026/07/23 03:59:07' -Duration '00:00:45') 'end before start -> false'
Assert-Equal $null  (Test-OldSnapDurationArithmetic -Start '' -End '2026/07/23 03:59:52' -Duration '00:00:45') 'blank start -> null (undecidable)'
Assert-Equal $null  (Test-OldSnapDurationArithmetic -Start '2026/07/23 03:59:07' -End '2026/07/23 03:59:52' -Duration '') 'blank duration -> null'
Assert-Equal $true  (Test-OldSnapDurationArithmetic -Start '2026/07/23 23:59:30' -End '2026/07/24 00:59:30' -Duration '01:00:00') 'spanning midnight is consistent'

# ---------------------------------------------------------------------------
# Repair-ProcessTimeStartFromStamp : adopt datestamp HH:mm on a pure 3<->9 swap.
# ---------------------------------------------------------------------------
# ja OCR read '03:53:07' but the clean datestamp says 035907 -> minute 53 vs 59
# is a single 3<->9 swap: correct to 03:59:07 (seconds unchanged).
$r1 = Repair-ProcessTimeStartFromStamp -StartTimeOfDay '03:53:07' -Datestamp '20260723035907'
Assert-True $r1.Repaired 'a 9->3 minute misread is repaired'
Assert-True $r1.SwapDetected 'the swap is detected'
Assert-Equal '03:59:07' $r1.Start 'start adopts the datestamp HH:mm, keeps OCR seconds'

# Already-agreeing HH:mm -> nothing to do.
$r2 = Repair-ProcessTimeStartFromStamp -StartTimeOfDay '03:59:07' -Datestamp '20260723035907'
Assert-True (-not $r2.Repaired) 'agreeing HH:mm is not repaired'
Assert-True (-not $r2.SwapDetected) 'no swap detected when they agree'
Assert-Equal '03:59:07' $r2.Start 'start unchanged when already correct'

# A non-3/9 disagreement is left untouched (not our bug to fix).
$r3 = Repair-ProcessTimeStartFromStamp -StartTimeOfDay '03:58:07' -Datestamp '20260723035907'
Assert-True (-not $r3.Repaired) 'a non-3/9 minute difference (8 vs 9) is not repaired'
Assert-Equal '03:58:07' $r3.Start 'start unchanged on a non-3/9 difference'

# Multiple 3<->9 digits across the hour+minute all corrected: real time
# 19:39:20 mis-OCR'd as 13:33:20 (both 9s read as 3) -> HHmm 1333 vs stamp
# 1939 differs only at the two 3<->9 positions.
$r4 = Repair-ProcessTimeStartFromStamp -StartTimeOfDay '13:33:20' -Datestamp '20260723193920'
Assert-True $r4.Repaired 'both hour and minute 9->3 swaps repaired'
Assert-Equal '19:39:20' $r4.Start 'HH:mm adopts 1939, seconds kept'

# Malformed inputs are safe.
$r5 = Repair-ProcessTimeStartFromStamp -StartTimeOfDay '' -Datestamp '20260723035907'
Assert-True (-not $r5.Repaired) 'blank start -> not repaired'
$r6 = Repair-ProcessTimeStartFromStamp -StartTimeOfDay '03:53:07' -Datestamp 'not-14-digits'
Assert-True (-not $r6.Repaired) 'bad datestamp -> not repaired'

# ---------------------------------------------------------------------------
# Get-OldSnapVerifyVerdict : the conservative triage decision.
# ---------------------------------------------------------------------------
Assert-Equal 'Txt' (Get-OldSnapVerifyVerdict -Source 'archived' -SnapExists $false) 'archived (.txt tier) is trusted regardless of snap presence'
Assert-Equal 'Txt' (Get-OldSnapVerifyVerdict -Source 'txt' -SnapExists $true) 'txt source is trusted'
Assert-Equal 'NoSnap' (Get-OldSnapVerifyVerdict -Source 'ocr' -SnapExists $false) 'OCR row with no snap image -> NoSnap'
Assert-Equal 'NeedsCheck' (Get-OldSnapVerifyVerdict -Source 'ocr:snap-png' -SnapExists $true -ArithmeticOk $false) 'snap-png OCR is triaged, not trusted; arithmetic mismatch flags'
Assert-Equal 'NeedsCheck' (Get-OldSnapVerifyVerdict -Source 'ocr-partial' -SnapExists $true) 'a partial read is never auto-confirmed'
Assert-Equal 'NeedsCheck' (Get-OldSnapVerifyVerdict -Source 'none' -SnapExists $true) 'none source with a snap -> NeedsCheck'
Assert-Equal 'NeedsCheck' (Get-OldSnapVerifyVerdict -Source 'ocr' -SnapExists $true -ArithmeticOk $true -DatestampSwap $true) 'a 3<->9 datestamp swap flags even when arithmetic is ok'
Assert-Equal 'OcrOk' (Get-OldSnapVerifyVerdict -Source 'ocr' -SnapExists $true -ArithmeticOk $true) 'OCR row passing every deterministic check (D2 off) auto-confirms'
Assert-Equal 'OcrOk' (Get-OldSnapVerifyVerdict -Source 'ocr' -SnapExists $true -ArithmeticOk $null) 'an undecidable arithmetic (null) does NOT by itself flag when D2 off'
# D2 enabled: require an explicit image 'ok'.
Assert-Equal 'OcrOk' (Get-OldSnapVerifyVerdict -Source 'ocr' -SnapExists $true -ArithmeticOk $true -PixelResult 'ok' -PixelEnabled $true) 'D2 ok + deterministic pass -> auto-confirm'
Assert-Equal 'NeedsCheck' (Get-OldSnapVerifyVerdict -Source 'ocr' -SnapExists $true -ArithmeticOk $true -PixelResult 'ng' -PixelEnabled $true) 'D2 ng flags'
Assert-Equal 'NeedsCheck' (Get-OldSnapVerifyVerdict -Source 'ocr' -SnapExists $true -ArithmeticOk $true -PixelResult '' -PixelEnabled $true) 'D2 enabled but no image result -> conservative flag'

# ---------------------------------------------------------------------------
# Get-OldSnapVerifyLabel : verdict key -> display string.
# ---------------------------------------------------------------------------
Assert-Equal 'txt'     (Get-OldSnapVerifyLabel 'Txt') 'Txt -> txt'
Assert-Equal 'OCR-OK'  (Get-OldSnapVerifyLabel 'OcrOk') 'OcrOk -> OCR-OK'
Assert-Equal $lblNeeds (Get-OldSnapVerifyLabel 'NeedsCheck') 'NeedsCheck -> you-kaku-nin'
Assert-Equal $lblNoSnap (Get-OldSnapVerifyLabel 'NoSnap') 'NoSnap -> gazou-nashi'
Assert-Equal '' (Get-OldSnapVerifyLabel 'Bogus') 'unknown key -> empty'

# ---------------------------------------------------------------------------
# Get-OldSnapVerifyColumnSpec : header / width / format.
# ---------------------------------------------------------------------------
$spec = Get-OldSnapVerifyColumnSpec
Assert-Equal $hdrVerify $spec.Header 'verify column header is kenshou'
Assert-Equal '@' $spec.NumberFormat 'verify column is text-formatted'
Assert-True ($spec.Width -gt 0) 'verify column has a positive width'

# ---------------------------------------------------------------------------
# Resolve-OldSnapExportImageDir : the per-correl ProcessTime export folder
# (D1's fallback image source when there is no standalone snap PNG).
# ---------------------------------------------------------------------------
# Asserted with -like (not an exact string) for the same reason as
# Resolve-OldSnapImagePath above: [IO.Path]::Combine emits the HOST's
# separator, so a Linux CI run yields '/' where the office PC yields '\'.
$expDir = Resolve-OldSnapExportImageDir -WorkDir 'C:\work' -CorrelId 'JIGPC06S'
Assert-True ($expDir -like '*C:\work*') 'export dir starts at the work folder'
Assert-True ($expDir -like '*ProcessTime*JIGPC06S') 'export dir defaults to snap\ProcessTime\<correl>'
Assert-True ((Resolve-OldSnapExportImageDir -WorkDir 'C:\work' -CorrelId 'JIGPC06S' -ExportRoot 'alt') -like '*alt*JIGPC06S') `
    'a custom export root is honored'
Assert-True ((Resolve-OldSnapExportImageDir -WorkDir 'C:\work' -CorrelId 'JIGPC06S' -ExportRoot '') -like '*ProcessTime*JIGPC06S') `
    'a blank export root falls back to the default layout'
Assert-True ($null -eq (Resolve-OldSnapExportImageDir -WorkDir '' -CorrelId 'JIGPC06S')) 'blank WorkDir -> null'
Assert-True ($null -eq (Resolve-OldSnapExportImageDir -WorkDir 'C:\work' -CorrelId '')) 'blank correl -> null'

# ---------------------------------------------------------------------------
# Select-OldSnapFallbackImageName : which exported picture to hyperlink.
# ---------------------------------------------------------------------------
$names = @(
    'GIFT_JIGPC06S_belowlabel_01.png',
    'GIFT_JIGPC06S_02.png',
    'GIFT_JIGPC06S_01.png',
    'GIFT_JIGPC06S_01_pre.png',
    'GIFT_JIGPC06S_01.ocr.txt',
    'GFIX_JIGPC06S_01.png'
)
Assert-Equal 'GIFT_JIGPC06S_01.png' (Select-OldSnapFallbackImageName -Side 'GIFT' -CorrelId 'JIGPC06S' -Names $names) `
    'the trusted section tier, lowest index, wins'
Assert-Equal 'GFIX_JIGPC06S_01.png' (Select-OldSnapFallbackImageName -Side 'GFIX' -CorrelId 'JIGPC06S' -Names $names) `
    'the other side picks its own picture'
Assert-Equal 'GIFT_JIGPC06S_01.png' (Select-OldSnapFallbackImageName -Side 'gift' -CorrelId 'JIGPC06S' -Names $names) `
    'the side argument is case-insensitive'

Assert-Equal 'GIFT_JIGPC06S_belowlabel_01.png' `
    (Select-OldSnapFallbackImageName -Side 'GIFT' -CorrelId 'JIGPC06S' -Names @('GIFT_JIGPC06S_abovelabel_01.png', 'GIFT_JIGPC06S_belowlabel_01.png')) `
    'below-label outranks above-label (the OCR tiers own trust order)'
Assert-Equal 'GIFT_JIGPC06S_wholesheet_01.png' `
    (Select-OldSnapFallbackImageName -Side 'GIFT' -CorrelId 'JIGPC06S' -Names @('GIFT_JIGPC06S_abovelabel_01.png', 'GIFT_JIGPC06S_wholesheet_01.png')) `
    'whole-sheet outranks above-label'
Assert-Equal 'GIFT_JIGPC06S_belowlabel_01.png' `
    (Select-OldSnapFallbackImageName -Side 'GIFT' -CorrelId 'JIGPC06S' -Names @('GIFT_JIGPC06S_belowlabel_02.png', 'GIFT_JIGPC06S_belowlabel_01.png')) `
    'within a tier the lowest picture index wins'

Assert-True ($null -eq (Select-OldSnapFallbackImageName -Side 'GIFT' -CorrelId 'JIGPC06S' -Names @('GIFT_JIGPC06S_01_pre.png'))) `
    'the preprocessed OCR derivative is never linked -- a human must see the original pixels'
Assert-True ($null -eq (Select-OldSnapFallbackImageName -Side 'GIFT' -CorrelId 'JIGPC06S' -Names @('GIFT_JIGPC06S_snapocr_pre.png'))) `
    'the snap-OCR preprocessed dump is never linked either'
Assert-True ($null -eq (Select-OldSnapFallbackImageName -Side 'GIFT' -CorrelId 'JIGPC06S' -Names @('GIFT_JIGPC06SX_01.png'))) `
    'a longer correl id sharing this prefix is not accepted'
Assert-True ($null -eq (Select-OldSnapFallbackImageName -Side 'GIFT' -CorrelId 'JIGPC06S' -Names @())) 'no candidates -> null'
Assert-True ($null -eq (Select-OldSnapFallbackImageName -Side 'GIFT' -CorrelId 'JIGPC06S' -Names $null)) 'null candidate list -> null'
Assert-True ($null -eq (Select-OldSnapFallbackImageName -Side 'DF' -CorrelId 'JIGPC06S' -Names $names)) 'an unknown side -> null'
Assert-True ($null -eq (Select-OldSnapFallbackImageName -Side 'GIFT' -CorrelId '' -Names $names)) 'a blank correl -> null'
Assert-Equal 'GIFT_JIGPC06S.png' (Select-OldSnapFallbackImageName -Side 'GIFT' -CorrelId 'JIGPC06S' -Names @('GIFT_JIGPC06S.png')) `
    'an un-indexed section picture is still accepted'

# ---------------------------------------------------------------------------
# Resolve-OldSnapPromotionMarkerPath : marks a snap PNG as a promoted copy.
# ---------------------------------------------------------------------------
Assert-Equal 'C:\work\snap\GIFT_HM\A.png.promoted.json' `
    (Resolve-OldSnapPromotionMarkerPath -SnapPath 'C:\work\snap\GIFT_HM\A.png') `
    'the marker sits beside the snap it describes'
Assert-True ($null -eq (Resolve-OldSnapPromotionMarkerPath -SnapPath '')) 'blank snap path -> null'
Assert-True ($null -eq (Resolve-OldSnapPromotionMarkerPath -SnapPath $null)) 'null snap path -> null'

# A fallback image makes the row verifiable rather than 'no image': the
# deterministic checks never depended on WHICH image exists.
Assert-Equal 'OcrOk' (Get-OldSnapVerifyVerdict -Source 'ocr:section' -SnapExists $true -ArithmeticOk $true) `
    'a row backed only by the exported evidence picture is triaged like any other OCR row'
Assert-Equal 'NeedsCheck' (Get-OldSnapVerifyVerdict -Source 'ocr:section' -SnapExists $true -ArithmeticOk $true -PixelEnabled $true -PixelResult '') `
    'with the pixel check on, a fallback-image row (no calibrated snap -> no pixel result) still flags'

exit (Complete-Tests)
