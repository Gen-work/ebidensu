#Requires -Version 5.1
# Test-PictureRange.ps1 -- kernel/PictureRange.ps1: the DF capture plan that
# mirrors the HOST pictures (gfix-recv, docs/TODO.md). Pure: OCR text in,
# ranges / heights / keys or 'unknown' out. The office-PC halves (strip
# crop, OCR call, window sizing) are not here.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path $repoRoot 'kernel/PictureRange.ps1')

Reset-Tests 'PictureRange'

function New-Strip {
    # the OCR text of a strip showing lines From..To, one number per line
    param([int]$From, [int]$To)
    return ((@($From..$To) | ForEach-Object { '{0:D6}' -f $_ }) -join "`n")
}

# ---------------------------------------------------------------- parse
$p = ConvertFrom-EbiLineNumberText -Text (New-Strip 1 3)
Assert-Equal 3 @($p['numbers']).Count 'parse: three clean lines -> three numbers'
Assert-Equal 1 $p['numbers'][0] 'parse: 000001 -> 1'
Assert-Equal 0 @($p['rejected']).Count 'parse: nothing rejected'

$p = ConvertFrom-EbiLineNumberText -Text "0000 01`r`n00 00 02`r`n"
Assert-Equal '1,2' (@($p['numbers']) -join ',') 'parse: spaces inside a number are OCR artefacts and removed; CRLF and a trailing blank line ok'

$p = ConvertFrom-EbiLineNumberText -Text "000001`nDATA`n00002`n000003`n  `n"
Assert-Equal '1,3' (@($p['numbers']) -join ',') 'parse: a word and a 5-digit line are not numbers'
Assert-Equal 'DATA|00002' (@($p['rejected']) -join '|') 'parse: rejected lines are listed as read (no repair)'

$p = ConvertFrom-EbiLineNumberText -Text "12`n0034" -Digits 0
Assert-Equal '12,34' (@($p['numbers']) -join ',') 'parse: Digits 0 accepts any length'

$p = ConvertFrom-EbiLineNumberText -Text ''
Assert-Equal 0 @($p['numbers']).Count 'parse: empty text -> no numbers'
$p = ConvertFrom-EbiLineNumberText -Text $null
Assert-Equal 0 @($p['numbers']).Count 'parse: null text -> no numbers, no throw'

# ---------------------------------------------------------------- run check
$r = Test-EbiLineNumberRun -Numbers @(1, 2, 3, 4) -Bands 4
Assert-True ([bool]$r['ok']) 'run: consecutive and as many as the bands -> ok'
Assert-Equal 1 $r['first'] 'run: first'
Assert-Equal 4 $r['last'] 'run: last'
Assert-Equal 4 $r['count'] 'run: count'

$r = Test-EbiLineNumberRun -Numbers @(1, 2, 3, 4)
Assert-True ([bool]$r['ok']) 'run: Bands 0 = the band count is not checked'

$r = Test-EbiLineNumberRun -Numbers @()
Assert-Equal 'no_numbers' $r['reason'] 'run: no numbers -> no_numbers'

$r = Test-EbiLineNumberRun -Numbers @(1, 2, 4, 5) -Bands 4
Assert-Equal 'not_consecutive' $r['reason'] 'run: a gap -> not_consecutive'
Assert-True ($r['message'] -like '*line 3 reads 000004 after 000002 (expected 000003)*') 'run: the message names the line and the numbers'

$r = Test-EbiLineNumberRun -Numbers @(1, 2, 2, 3) -Bands 4
Assert-Equal 'not_consecutive' $r['reason'] 'run: a doubled OCR line -> not_consecutive'

$r = Test-EbiLineNumberRun -Numbers @(878, 879, 880, 831) -Bands 4
Assert-Equal 'not_consecutive' $r['reason'] 'run: a 9 read as 3 (881 -> 831) breaks the run instead of being guessed at'

$r = Test-EbiLineNumberRun -Numbers @(1, 2, 3) -Bands 4
Assert-Equal 'count_mismatch' $r['reason'] 'run: a skipped OCR line -> count_mismatch'
Assert-True ($r['message'] -like '*3 line numbers read but the strip has 4 text lines*') 'run: count_mismatch message'

# ---------------------------------------------------------------- HOST -> DF
$d = ConvertTo-EbiDfRange -First 1 -Last 17
Assert-Equal '1-16' ('{0}-{1}' -f $d['first'], $d['last']) 'range: HOST 1-17 -> DF 1-16 (the header line)'
Assert-Equal 16 $d['lines'] 'range: 16 lines'
$d = ConvertTo-EbiDfRange -First 878 -Last 896
Assert-Equal '877-895' ('{0}-{1}' -f $d['first'], $d['last']) 'range: HOST 878-896 -> DF 877-895'
$d = ConvertTo-EbiDfRange -First 1 -Last 19
Assert-Equal '1-18' ('{0}-{1}' -f $d['first'], $d['last']) 'range: a 19-line HOST picture -> 18 DF lines'
$d = ConvertTo-EbiDfRange -First 1 -Last 17 -HeaderLines 0
Assert-Equal '1-17' ('{0}-{1}' -f $d['first'], $d['last']) 'range: HeaderLines 0 -> no shift'
$d = ConvertTo-EbiDfRange -First 1 -Last 1
Assert-Equal 0 $d['lines'] 'range: a header-only picture leaves no DF line'

# ---------------------------------------------------------------- the plan
# FJODWCP2 (docs/TODO.md): HOST shows 000001-000019 and 000878-000896, 19 lines a picture
$host2 = @(
    @{ name = 'send pic 1'; text = (New-Strip 1 19);   bands = 19 }
    @{ name = 'send pic 2'; text = (New-Strip 878 896); bands = 19 }
)
$plan = Get-EbiDfRangePlan -HostPictures $host2 -DfLines 895 -RowPitch 26
Assert-Equal 'ok' $plan['code'] 'plan: FJODWCP2 -> ok'
Assert-Equal 2 @($plan['ranges']).Count 'plan: two DF pictures'
$a = $plan['ranges'][0]; $b = $plan['ranges'][1]
Assert-Equal '1-18' ('{0}-{1}' -f $a['first'], $a['last']) 'plan: picture 1 DF 1-18'
Assert-Equal '877-895' ('{0}-{1}' -f $b['first'], $b['last']) 'plan: picture 2 DF 877-895'
Assert-Equal 18 $a['lines'] 'plan: 18 lines a picture'
Assert-Equal (429 + 26 * 2) $a['height'] 'plan: height = base + pitch x (lines - screenLines)'
Assert-Equal '^{HOME}' $a['keys'] 'plan: picture 1 starts at line 1 -> KeysFirst'
Assert-Equal '{END}{PGUP}{DOWN}{DOWN}' $b['keys'] 'plan: the last picture ends at the DF end -> KeysEnd'
Assert-Equal '1,19' ('{0},{1}' -f $a['hostFirst'], $a['hostLast']) 'plan: the HOST numbers are kept for the gate'
Assert-Equal 'send pic 2' $b['name'] 'plan: the picture name is kept'
Assert-Equal 0 @($plan['warnings']).Count 'plan: no warnings'

# the usual 16-line workbook: no calibration needed
$host16 = @(@{ text = (New-Strip 1 17); bands = 17 })
$plan = Get-EbiDfRangePlan -HostPictures $host16 -DfLines 16
Assert-Equal 'ok' $plan['code'] 'plan: a 16-line picture needs no row pitch'
Assert-Equal 429 $plan['ranges'][0]['height'] 'plan: 16 lines -> the base height'
Assert-Equal '^{HOME}' $plan['ranges'][0]['keys'] 'plan: a single picture from line 1 -> KeysFirst even though it also ends at the DF end'

# a 17..32-line file as HOST shows it: 1-17 and 18-33 -> DF 1-16 and 17-32
$host32 = @(@{ text = (New-Strip 1 17); bands = 17 }, @{ text = (New-Strip 18 33); bands = 16 })
$plan = Get-EbiDfRangePlan -HostPictures $host32 -DfLines 32
Assert-Equal 'ok' $plan['code'] 'plan: two 16-line pictures -> ok without a pitch'
Assert-Equal '17-32' ('{0}-{1}' -f $plan['ranges'][1]['first'], $plan['ranges'][1]['last']) 'plan: second picture DF 17-32'
Assert-Equal '{END}{PGUP}{DOWN}{DOWN}' $plan['ranges'][1]['keys'] 'plan: ends at the DF end -> KeysEnd'

# a middle picture reachable by whole pages (a HOST set that does not reach the DF end: DfLines unknown)
$hostMid = @(@{ text = (New-Strip 1 17); bands = 17 }, @{ text = (New-Strip 34 49); bands = 16 })
$plan = Get-EbiDfRangePlan -HostPictures $hostMid
Assert-Equal 'ok' $plan['code'] 'plan: a middle picture at line 33 = 2 pages of 16 -> ok (DfLines unknown, so no end check)'
Assert-Equal '^{HOME}{PGDN}{PGDN}' $plan['ranges'][1]['keys'] 'plan: KeysFirst + two pages down'

# a middle picture NOT on a page boundary -> no recipe yet
$hostOff = @(@{ text = (New-Strip 1 17); bands = 17 }, @{ text = (New-Strip 40 55); bands = 16 })
$plan = Get-EbiDfRangePlan -HostPictures $hostOff
Assert-Equal 'no_key_recipe' $plan['reason'] 'plan: DF 39-54 is not a page boundary -> no_key_recipe'
Assert-Equal 2 @($plan['ranges']).Count 'plan: the derived ranges are still returned for the gate'

# unknowns
$plan = Get-EbiDfRangePlan -HostPictures @()
Assert-Equal 'unknown' $plan['code'] 'plan: no pictures -> unknown'
Assert-Equal 'no_host_pictures' $plan['reason'] 'plan: reason no_host_pictures'

$plan = Get-EbiDfRangePlan -HostPictures @(@{ text = ''; bands = 17 })
Assert-Equal 'no_numbers' $plan['reason'] 'plan: empty OCR -> no_numbers'

$bad = @(@{ name = 'send pic 1'; text = (New-Strip 1 19); bands = 19 }, @{ name = 'send pic 2'; text = ((New-Strip 878 880) + "`n000831"); bands = 4 })
$plan = Get-EbiDfRangePlan -HostPictures $bad -DfLines 895 -RowPitch 26
Assert-Equal 'not_consecutive' $plan['reason'] 'plan: a misread digit in picture 2 -> not_consecutive'
Assert-True ($plan['message'] -like 'send pic 2:*') 'plan: the message names the picture'
Assert-Equal 1 @($plan['ranges']).Count 'plan: picture 1 was derived before the failure'

$plan = Get-EbiDfRangePlan -HostPictures @(@{ text = (New-Strip 1 16); bands = 17 })
Assert-Equal 'count_mismatch' $plan['reason'] 'plan: 16 numbers for 17 bands -> count_mismatch'

$plan = Get-EbiDfRangePlan -HostPictures $host2 -DfLines 900 -RowPitch 26
Assert-Equal 'df_lines_mismatch' $plan['reason'] 'plan: the DF file has 900 lines but HOST ends at 896 -> df_lines_mismatch'
Assert-True ($plan['message'] -like '*000896*895*900*') 'plan: the message gives HOST end, DF end and the DF count'

$plan = Get-EbiDfRangePlan -HostPictures $host2 -DfLines 895
Assert-Equal 'row_pitch_uncalibrated' $plan['reason'] 'plan: 18-line pictures with no pitch -> row_pitch_uncalibrated'

$plan = Get-EbiDfRangePlan -HostPictures $host2 -RowPitch 26
Assert-Equal 'no_key_recipe' $plan['reason'] 'plan: DfLines 0 skips the end check, so the last picture is not known as the END view -> no_key_recipe'
Assert-Equal 2 @($plan['ranges']).Count 'plan: ...both ranges were still derived'

$over = @(@{ text = (New-Strip 1 19); bands = 19 }, @{ text = (New-Strip 15 33); bands = 19 })
$plan = Get-EbiDfRangePlan -HostPictures $over -RowPitch 26
Assert-Equal 'ranges_overlap' $plan['reason'] 'plan: overlapping HOST pictures -> ranges_overlap'

$plan = Get-EbiDfRangePlan -HostPictures @(@{ text = '000001'; bands = 1 })
Assert-Equal 'empty_range' $plan['reason'] 'plan: a header-only picture -> empty_range'

$noisy = @(@{ text = ("000001`nxx`n" + (New-Strip 2 17)); bands = 17 })
$plan = Get-EbiDfRangePlan -HostPictures $noisy -DfLines 16
Assert-Equal 'ok' $plan['code'] 'plan: a rejected strip line does not fail the plan'
Assert-Equal 1 @($plan['warnings']).Count 'plan: ...it is a warning'
Assert-True ($plan['warnings'][0] -like '*"xx"*') 'plan: the warning quotes the line'

$plan = Get-EbiDfRangePlan -HostPictures @('not a map', $host16[0]) -DfLines 16
Assert-Equal 'ok' $plan['code'] 'plan: non-map entries are ignored'

exit (Complete-Tests)
