#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'SnapVerify.ps1')

Reset-Tests 'SnapVerify'

# ---------------------------------------------------------------------------
# Fixture helpers (ASCII source; Japanese via [char])
# ---------------------------------------------------------------------------
$t      = "`t"
$nl     = "`n"
$normal = [char]0x6B63 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # 正常終了
$abend  = [char]0x7570 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # 異常終了
$diam   = [char]0x25C6  # ◆

# HM page title for page-kind detection
$hmTitle = [char]0x30D0 + [char]0x30C3 + [char]0x30C1 + [char]0x51E6 `
         + [char]0x7406 + [char]0x72B6 + [char]0x6CC1 + [char]0x4E00 + [char]0x89A7  # バッチ処理状況一覧

# Appendix A sample (real HM page, 3 data rows):
#   row1: 2026/06/12 11:05:40  正常終了  JIDSK01S  (latest)
#   row2: 2026/06/12 10:35:40  異常終了  JIDSK01S  (middle)
#   row3: 2026/06/12 07:51:20  正常終了  JIDSK01S  (oldest)
$hmRow1 = "2026/06/12 11:05:40${t}2026/06/12 11:06:57${t}00:01:17${t}IDSLA013${t}K${t}${normal}${t}20260424040558${t}36,117${t}${diam}${t}JIDSK01S"
$hmRow2 = "2026/06/12 10:35:40${t}2026/06/12 10:35:43${t}00:00:03${t}IDSLA013${t}K${t}${abend}${t}20260424040558${t}17${t}${t}${diam}${t}JIDSK01S"
$hmRow3 = "2026/06/12 07:51:20${t}2026/06/12 07:51:41${t}00:00:21${t}IDSLA013${t}K${t}${normal}${t}20260612075111${t}5,764${t}${diam}${t}JIDSK01S"

# Full HM sample (includes header/menu text that must be ignored by the parser)
$hmStartHdr = [char]0x958B + [char]0x59CB + [char]0x65E5 + [char]0x6642  # 開始日時
$hmEndHdr   = [char]0x7D42 + [char]0x4E86 + [char]0x65E5 + [char]0x6642  # 終了日時
$hmTableHdr = "${hmStartHdr}${t}${hmEndHdr}${t}..." # table header row (non-data, must be skipped)

$hmFullText = @"
 ${hmTitle}			IDSXA041
test
${hmTableHdr}
${hmRow1}
${hmRow2}
${hmRow3}
"@

# Appendix B-1: normal MQ result (3 records, same CorrelId JIDSK05S)
$mqRef = [char]0x53C2 + [char]0x7167  # 参照 (not used in MQ, just defined)
$mqFullText = @"
Transfer status inquiry results
Number of records 3
No${t}Send node${t}Recv node${t}Correlid${t}Send date${t}Tmode${t}Recv date${t}Rtncd${t}Rsncd${t}
Msgid${t}Reccnt${t}File size
1${t}JSSS004R${t}JHM102R${t}JIDSK05S${t}2026/06/12 07:54:22${t}TXT${t}2026/06/12 07:54:23${t}0${t}0${t}
A2009999A0000000001${t}11002${t}4986849
2${t}JSSS004R${t}JHM102R${t}JIDSK05S${t}2026/06/12 10:32:41${t}TXT${t}2026/06/12 10:32:42${t}0${t}0${t}
A2009999A0000000002${t}12334${t}5752860
3${t}JSSS004R${t}JHM102R${t}JIDSK05S${t}2026/06/12 11:01:53${t}TXT${t}2026/06/12 11:01:57${t}0${t}0${t}
A2009999A0000000003${t}12334${t}5752860
"@

# Appendix B-2: No Data!
$mqNoDataText = 'No Data!'

# Appendix B-3: outer-frame focus error
$mqOuterText  = @"
GIFT System

<Transfer status>

Inquiry

<Documents>

Download
"@

# Jenkins file list sample
$ref = [char]0x53C2 + [char]0x7167  # 参照
$jkRow1 = "JIDSK01S 2026/06/12 10:35:21 189.90 KB ${ref}"
$jkRow2 = "JIDSK05S 2026/06/12 11:01:53 512.00 KB ${ref}"
$jkFullText = "File list${nl}${jkRow1}${nl}${jkRow2}${nl}"

# ===========================================================================
# ConvertFrom-HmPageText
# ===========================================================================

$rows = @(ConvertFrom-HmPageText $hmFullText)

Assert-Equal 3 $rows.Count 'HM: parses 3 data rows from full page text'

# Row ordering is preserved (same as source line order: row1=newest is first)
Assert-Equal $normal $rows[0].Status 'HM row[0]: status = 正常終了'
Assert-Equal $abend  $rows[1].Status 'HM row[1]: status = 異常終了'
Assert-Equal $normal $rows[2].Status 'HM row[2]: status = 正常終了'

Assert-Equal 'JIDSK01S' $rows[0].CorrelId 'HM row[0]: CorrelId extracted from last field'
Assert-Equal 'JIDSK01S' $rows[1].CorrelId 'HM row[1]: CorrelId from abend row (extra empty field)'
Assert-Equal 'IDSLA013' $rows[0].BatchId  'HM row[0]: BatchId from field[3]'

Assert-True ($null -ne $rows[0].StartTime) 'HM row[0]: StartTime parsed'
Assert-Equal '2026/06/12 11:05:40' ($rows[0].StartTime.ToString('yyyy/MM/dd HH:mm:ss')) 'HM row[0]: StartTime value'
Assert-Equal '2026/06/12 10:35:40' ($rows[1].StartTime.ToString('yyyy/MM/dd HH:mm:ss')) 'HM row[1]: StartTime value (abend)'

# Header/menu lines must not be parsed as data rows
$nonData = @(ConvertFrom-HmPageText ($hmTableHdr + $nl + 'plain text line'))
Assert-Equal 0 $nonData.Count 'HM: header/non-data lines produce 0 rows'

# Single row (only the abend line)
$singleRows = @(ConvertFrom-HmPageText $hmRow2)
Assert-Equal 1 $singleRows.Count 'HM: single abend row parses correctly'
Assert-Equal $abend $singleRows[0].Status 'HM: single abend row status correct'

# ===========================================================================
# Test-HmAbend
# ===========================================================================

$hmRows = @(ConvertFrom-HmPageText $hmFullText)

# Expected = abend time (10:35:40), Tolerance=30 -> window [10:05:40, 11:05:40]
# Both abend (10:35) and newest-normal (11:05) are in window -> ok with warning
$exp1 = [datetime]::ParseExact('2026/06/12 10:35:40', 'yyyy/MM/dd HH:mm:ss',
        [System.Globalization.CultureInfo]::InvariantCulture)
$v1 = Test-HmAbend -Rows $hmRows -CorrelId 'JIDSK01S' -Expected $exp1 -ToleranceMin 30
Assert-Equal 'ok'   $v1.Verdict  'HmAbend: newest in window is normal -> ok'
Assert-True ($v1.Warnings.Count -gt 0) 'HmAbend: warning added for earlier window abend'
Assert-True ($v1.Warnings[0] -like '*retried*') 'HmAbend: warning says retried'

# Expected = abend time, Tolerance=10 -> window [10:25, 10:45]; only abend row in window -> ng
# Both outside rows (11:05 normal, 07:51 normal) are normal -> no outside-abend warnings
$v2 = Test-HmAbend -Rows $hmRows -CorrelId 'JIDSK01S' -Expected $exp1 -ToleranceMin 10
Assert-Equal 'ng'   $v2.Verdict        'HmAbend: newest in narrow window is abend -> ng'
Assert-Equal 0      $v2.Warnings.Count 'HmAbend: outside normal rows do not generate warnings'

# Expected = oldest row time (07:51:20), Tolerance=30 -> only that normal row in window -> ok
# The 10:35:40 abend is outside this window -> generates a historic-abend warning
$exp3 = [datetime]::ParseExact('2026/06/12 07:51:20', 'yyyy/MM/dd HH:mm:ss',
        [System.Globalization.CultureInfo]::InvariantCulture)
$v3 = Test-HmAbend -Rows $hmRows -CorrelId 'JIDSK01S' -Expected $exp3 -ToleranceMin 30
Assert-Equal 'ok'   $v3.Verdict  'HmAbend: only normal row in window -> ok'
Assert-True ($v3.Warnings.Count -ge 1) 'HmAbend: historic abend outside window generates a warning'

# No time check (Expected=$null): abend row present -> ask
$vNoTime = Test-HmAbend -Rows $hmRows -CorrelId 'JIDSK01S' -Expected $null
Assert-Equal 'ask'  $vNoTime.Verdict 'HmAbend: no time check + abend row present -> ask'

# No rows for the correl -> ask
$vEmpty = Test-HmAbend -Rows @() -CorrelId 'JIDSK01S' -Expected $exp1
Assert-Equal 'ask'  $vEmpty.Verdict  'HmAbend: no rows -> ask'

# Window with zero matching rows -> ask
$expFar = [datetime]::ParseExact('2026/06/11 00:00:00', 'yyyy/MM/dd HH:mm:ss',
          [System.Globalization.CultureInfo]::InvariantCulture)
$vFar = Test-HmAbend -Rows $hmRows -CorrelId 'JIDSK01S' -Expected $expFar -ToleranceMin 30
Assert-Equal 'ask'  $vFar.Verdict    'HmAbend: no rows in window -> ask'
Assert-True ($vFar.Warnings.Count -gt 0) 'HmAbend: out-of-window abend promoted to warning'

# ===========================================================================
# ConvertFrom-MqPageText
# ===========================================================================

$mqParsed = ConvertFrom-MqPageText $mqFullText
Assert-Equal 3  $mqParsed.NumRecords 'MQ: NumRecords = 3'
Assert-Equal 3  $mqParsed.Rows.Count 'MQ: 3 data rows parsed'

$r0 = $mqParsed.Rows[0]
Assert-Equal 1          $r0.No       'MQ row[0]: No = 1'
Assert-Equal 'JSSS004R' $r0.SendNode 'MQ row[0]: SendNode'
Assert-Equal 'JIDSK05S' $r0.CorrelId 'MQ row[0]: CorrelId'
Assert-Equal 0          $r0.Rtncd    'MQ row[0]: Rtncd = 0'
Assert-Equal 0          $r0.Rsncd    'MQ row[0]: Rsncd = 0'
Assert-True ($null -ne $r0.RecvDate)  'MQ row[0]: RecvDate parsed'
Assert-Equal '2026/06/12 07:54:23' ($r0.RecvDate.ToString('yyyy/MM/dd HH:mm:ss')) 'MQ row[0]: RecvDate value'

# No Data text -> 0 rows, NumRecords=-1
$mqEmptyParsed = ConvertFrom-MqPageText $mqNoDataText
Assert-Equal -1 $mqEmptyParsed.NumRecords 'MQ No Data: NumRecords = -1 (no marker)'
Assert-Equal 0  $mqEmptyParsed.Rows.Count 'MQ No Data: 0 rows'

# ===========================================================================
# Test-MqRecord
# ===========================================================================

# Expected = newest RecvDate (11:01:57), Tolerance=30 -> ok
$mqExp = [datetime]::ParseExact('2026/06/12 11:01:57', 'yyyy/MM/dd HH:mm:ss',
         [System.Globalization.CultureInfo]::InvariantCulture)
$mv1 = Test-MqRecord -Parsed $mqParsed -CorrelId 'JIDSK05S' -Expected $mqExp -ToleranceMin 30
Assert-Equal 'ok' $mv1.Verdict 'MQ: newest row in window, Rtncd=0 -> ok'
Assert-Equal 3    $mv1.MatchedRow.No 'MQ: newest row (No=3) selected'

# IsNoData -> ng
$mvNoData = Test-MqRecord -Parsed $mqEmptyParsed -CorrelId 'JIDSK05S' -Expected $null -IsNoData $true
Assert-Equal 'ng' $mvNoData.Verdict 'MQ: IsNoData -> ng'
Assert-True ($mvNoData.Reason -like '*No Data*') 'MQ: No Data reason string'

# No matching CorrelId -> ng
$mvMiss = Test-MqRecord -Parsed $mqParsed -CorrelId 'JXXXXX99' -Expected $null
Assert-Equal 'ng' $mvMiss.Verdict 'MQ: no row for correl -> ng'

# Window miss: Expected is far in the past -> newest RecvDate outside tolerance -> ng
$mqExpFar = [datetime]::ParseExact('2026/06/11 11:00:00', 'yyyy/MM/dd HH:mm:ss',
            [System.Globalization.CultureInfo]::InvariantCulture)
$mvWin = Test-MqRecord -Parsed $mqParsed -CorrelId 'JIDSK05S' -Expected $mqExpFar -ToleranceMin 30
Assert-Equal 'ng' $mvWin.Verdict 'MQ: RecvDate far outside window -> ng'

# Non-zero Rtncd: inject a row with Rtncd=1
$ngRow = [PSCustomObject]@{
    No = 1; SendNode = 'S'; RecvNode = 'R'; CorrelId = 'JTEST01S'
    SendDate = $mqExp; Tmode = 'TXT'; RecvDate = $mqExp; Rtncd = 1; Rsncd = 0
}
$ngParsed = @{ NumRecords = 1; Rows = @($ngRow) }
$mvRtn = Test-MqRecord -Parsed $ngParsed -CorrelId 'JTEST01S' -Expected $null
Assert-Equal 'ng' $mvRtn.Verdict 'MQ: non-zero Rtncd -> ng'
Assert-True ($mvRtn.Reason -like '*Rtncd=1*') 'MQ: Rtncd reason string'

# ===========================================================================
# ConvertFrom-JenkinsListText
# ===========================================================================

$jkFiles = @(ConvertFrom-JenkinsListText $jkFullText)
Assert-Equal 2           $jkFiles.Count       'Jenkins: 2 file rows parsed'
Assert-Equal 'JIDSK01S'  $jkFiles[0].Name     'Jenkins: file[0] name'
Assert-Equal '189.90 KB' $jkFiles[0].Size     'Jenkins: file[0] size'
Assert-True ($null -ne $jkFiles[0].DateTime)  'Jenkins: file[0] DateTime parsed'
Assert-Equal '2026/06/12 10:35:21' ($jkFiles[0].DateTime.ToString('yyyy/MM/dd HH:mm:ss')) 'Jenkins: file[0] DateTime value'

# Empty text -> 0 files
$jkEmpty = @(ConvertFrom-JenkinsListText '')
Assert-Equal 0 $jkEmpty.Count 'Jenkins: empty text -> 0 files'

# ===========================================================================
# Test-JenkinsFile
# ===========================================================================

$jkExp = [datetime]::ParseExact('2026/06/12 10:35:21', 'yyyy/MM/dd HH:mm:ss',
         [System.Globalization.CultureInfo]::InvariantCulture)

# File found, in window -> ok
$jv1 = Test-JenkinsFile -Files $jkFiles -CorrelId 'JIDSK01S' -Expected $jkExp -ToleranceMin 30
Assert-Equal 'ok'  $jv1.Verdict  'Jenkins: file found in window -> ok'
Assert-Equal 'JIDSK01S' $jv1.File.Name 'Jenkins: returned matched file'

# File not found -> ng
$jv2 = Test-JenkinsFile -Files $jkFiles -CorrelId 'JXXXXX99' -Expected $null
Assert-Equal 'ng'  $jv2.Verdict  'Jenkins: file not found -> ng'

# File found but time outside window -> ng
$jkExpFar = [datetime]::ParseExact('2026/06/11 10:35:21', 'yyyy/MM/dd HH:mm:ss',
            [System.Globalization.CultureInfo]::InvariantCulture)
$jv3 = Test-JenkinsFile -Files $jkFiles -CorrelId 'JIDSK01S' -Expected $jkExpFar -ToleranceMin 30
Assert-Equal 'ng'  $jv3.Verdict  'Jenkins: file time outside window -> ng'

# No time check (Expected=$null) + file found -> ok
$jv4 = Test-JenkinsFile -Files $jkFiles -CorrelId 'JIDSK01S' -Expected $null
Assert-Equal 'ok'  $jv4.Verdict  'Jenkins: no time check, file found -> ok'

# NoGfix mode ($ExpectExists=$false): file NOT found -> ok
$jv5 = Test-JenkinsFile -Files $jkFiles -CorrelId 'JXXXXX99' -Expected $null -ExpectExists $false
Assert-Equal 'ok'  $jv5.Verdict  'Jenkins NoGfix: no file found as expected -> ok'

# NoGfix mode: file found -> ng (unexpected past data)
$jv6 = Test-JenkinsFile -Files $jkFiles -CorrelId 'JIDSK01S' -Expected $null -ExpectExists $false
Assert-Equal 'ng'  $jv6.Verdict  'Jenkins NoGfix: unexpected file found -> ng'
Assert-True ($jv6.Reason -like '*past data*') 'Jenkins NoGfix: reason mentions past data'

# ===========================================================================
# Get-SnapPageKind
# ===========================================================================

# HM page (contains title)
Assert-Equal 'HmResult' (Get-SnapPageKind -Phase 'Hm' -Text $hmFullText) 'PageKind: HM full text -> HmResult'

# HM page (minimal, only header)
$hmMinimal = "${hmStartHdr}`t${hmEndHdr}`tother"
Assert-Equal 'HmResult' (Get-SnapPageKind -Phase 'Hm' -Text $hmMinimal) 'PageKind: HM header line -> HmResult'

# MQ result page
Assert-Equal 'MqResult' (Get-SnapPageKind -Phase 'Mq' -Text $mqFullText) 'PageKind: MQ result text -> MqResult'

# MQ No Data
Assert-Equal 'MqNoData' (Get-SnapPageKind -Phase 'Mq' -Text $mqNoDataText) 'PageKind: No Data! -> MqNoData'
Assert-Equal 'MqNoData' (Get-SnapPageKind -Phase 'Mq' -Text "  No Data!  ") 'PageKind: No Data! with whitespace -> MqNoData'

# Outer frame
Assert-Equal 'OuterFrame' (Get-SnapPageKind -Phase 'Mq' -Text $mqOuterText) 'PageKind: GIFT System text -> OuterFrame'

# Jenkins result page
Assert-Equal 'JenkinsResult' (Get-SnapPageKind -Phase 'Jenkins' -Text $jkFullText) 'PageKind: Jenkins file list -> JenkinsResult'

# Empty
Assert-Equal 'Empty'   (Get-SnapPageKind -Phase 'Hm' -Text '')    'PageKind: empty string -> Empty'
Assert-Equal 'Empty'   (Get-SnapPageKind -Phase 'Hm' -Text '   ') 'PageKind: whitespace -> Empty'

# Unknown
Assert-Equal 'Unknown' (Get-SnapPageKind -Phase 'Hm' -Text 'some completely unrecognized page format') 'PageKind: unrecognized -> Unknown'

# ===========================================================================
# Resolve-SnapRunTime
# ===========================================================================

$refNow = [datetime]::ParseExact('2026/06/12 10:00:00', 'yyyy/MM/dd HH:mm:ss',
          [System.Globalization.CultureInfo]::InvariantCulture)

# Empty input -> use Now
$rt1 = Resolve-SnapRunTime -TimeInput '' -DefaultTolerance 30 -Now $refNow
Assert-True $rt1.Ok                         'RunTime: empty input -> Ok'
Assert-Equal 'fixed'    $rt1.TimeMode       'RunTime: empty input -> TimeMode=fixed'
Assert-Equal '2026/06/12 10:00:00' ($rt1.Time.ToString('yyyy/MM/dd HH:mm:ss')) 'RunTime: empty input -> Time=Now'
Assert-Equal 30         $rt1.ToleranceMinutes 'RunTime: default tolerance kept'

# 'n' -> no time check
$rt2 = Resolve-SnapRunTime -TimeInput 'n' -DefaultTolerance 30 -Now $refNow
Assert-True $rt2.Ok                         'RunTime: n -> Ok'
Assert-Equal 'none'     $rt2.TimeMode       'RunTime: n -> TimeMode=none'
Assert-True ($null -eq $rt2.Time)           'RunTime: n -> Time=null'

$rt2N = Resolve-SnapRunTime -TimeInput 'N' -DefaultTolerance 30 -Now $refNow
Assert-Equal 'none' $rt2N.TimeMode          'RunTime: N (uppercase) -> none'

# Explicit datetime
$rt3 = Resolve-SnapRunTime -TimeInput '2026/06/12 10:35:40' -DefaultTolerance 30 -Now $refNow
Assert-True $rt3.Ok                         'RunTime: explicit datetime -> Ok'
Assert-Equal 'fixed'    $rt3.TimeMode       'RunTime: explicit datetime -> TimeMode=fixed'
Assert-Equal '2026/06/12 10:35:40' ($rt3.Time.ToString('yyyy/MM/dd HH:mm:ss')) 'RunTime: explicit datetime -> Time parsed'

# Tolerance override
$rt4 = Resolve-SnapRunTime -TimeInput '' -ToleranceInput '15' -DefaultTolerance 30 -Now $refNow
Assert-Equal 15 $rt4.ToleranceMinutes       'RunTime: tolerance override to 15'

# Invalid input -> Ok=$false
$rt5 = Resolve-SnapRunTime -TimeInput 'bad-date' -DefaultTolerance 30 -Now $refNow
Assert-True (-not $rt5.Ok)                  'RunTime: bad date -> Ok=false'
Assert-True ($rt5.Error -ne '')             'RunTime: bad date -> Error message set'

# ===========================================================================
# ConvertTo-ExpectedDateTime
# ===========================================================================

# Empty / whitespace -> $null (no time window)
Assert-True ($null -eq (ConvertTo-ExpectedDateTime -Value ''))    'ExpectedDT: empty -> null'
Assert-True ($null -eq (ConvertTo-ExpectedDateTime -Value '   ')) 'ExpectedDT: whitespace -> null'

# Full datetime
$edt1 = ConvertTo-ExpectedDateTime -Value '2026/06/12 10:35:40'
Assert-True ($null -ne $edt1) 'ExpectedDT: full datetime parses'
Assert-Equal '2026/06/12 10:35:40' ($edt1.ToString('yyyy/MM/dd HH:mm:ss')) 'ExpectedDT: full datetime value'

# Date-only fallback format
$edt2 = ConvertTo-ExpectedDateTime -Value '2026/06/12'
Assert-True ($null -ne $edt2) 'ExpectedDT: date-only parses via fallback'
Assert-Equal '2026/06/12 00:00:00' ($edt2.ToString('yyyy/MM/dd HH:mm:ss')) 'ExpectedDT: date-only value'

# Unparseable -> $null (treated as no-time, never throws)
Assert-True ($null -eq (ConvertTo-ExpectedDateTime -Value 'not-a-date')) 'ExpectedDT: garbage -> null'

# ===========================================================================
# Set-EmptyRunTimeCells
# ===========================================================================

# Mix of empty / existing / missing-column rows
$rowEmpty   = [PSCustomObject]@{ Correl_ID_S = 'A'; Expected_Time = '' }
$rowKeep    = [PSCustomObject]@{ Correl_ID_S = 'B'; Expected_Time = '2026/06/01 09:00:00' }
$rowBlank   = [PSCustomObject]@{ Correl_ID_S = 'C'; Expected_Time = '   ' }
$rowNoCol   = [PSCustomObject]@{ Correl_ID_S = 'D' }   # column absent entirely
$setRows    = @($rowEmpty, $rowKeep, $rowBlank, $rowNoCol)

$filled = Set-EmptyRunTimeCells -Rows $setRows -Field 'Expected_Time' -Value '2026/06/17 12:00:00'
Assert-Equal 3 $filled 'SetEmptyTime: fills empty/blank/missing cells only (3 of 4)'
Assert-Equal '2026/06/17 12:00:00' $rowEmpty.Expected_Time 'SetEmptyTime: empty cell filled'
Assert-Equal '2026/06/01 09:00:00' $rowKeep.Expected_Time  'SetEmptyTime: existing value kept'
Assert-Equal '2026/06/17 12:00:00' $rowBlank.Expected_Time 'SetEmptyTime: whitespace cell filled'
Assert-Equal '2026/06/17 12:00:00' $rowNoCol.Expected_Time 'SetEmptyTime: missing column added + filled'

# Re-running fills nothing (all rows now have values)
$filled2 = Set-EmptyRunTimeCells -Rows $setRows -Field 'Expected_Time' -Value '2099/01/01 00:00:00'
Assert-Equal 0 $filled2 'SetEmptyTime: second pass fills nothing'
Assert-Equal '2026/06/17 12:00:00' $rowEmpty.Expected_Time 'SetEmptyTime: second pass does not overwrite'

exit (Complete-Tests)
