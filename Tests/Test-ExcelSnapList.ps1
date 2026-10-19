#Requires -Version 5.1
# Test-ExcelSnapList.ps1 -- the pure halves of excel.snap_list and
# excel.place_picture (10/16: picture 1 snapped from the overview list,
# on to the source-file column when the job's rows repeat). The COM halves
# need Excel and are checked on the office PC.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path $repoRoot 'kernel/Json.ps1')
. (Join-Path $repoRoot 'kernel/Registry.ps1')

Reset-Tests 'ExcelSnapList'

$reg = New-EbiRegistry -ModulesRoot (Join-Path $repoRoot 'modules')
foreach ($u in @('excel.snap_list', 'excel.place_picture', 'excel.insert_pictures')) {
    $r = . Import-EbiStep -Registry $reg -Use $u
    if (-not $r['ok']) { throw ('cannot load ' + $u + ': ' + $r['message']) }
}

# excel.insert_pictures replace (10/19): only this step's own shapes under this tag
Assert-True (ExcelInsertPictures-IsOwnShape -Name 'verifyPic_jenkins_1' -Tag 'jenkins') 'replace: our picture'
Assert-True (ExcelInsertPictures-IsOwnShape -Name 'verifyMark_jenkins_1_2' -Tag 'jenkins') 'replace: our box'
Assert-True (-not (ExcelInsertPictures-IsOwnShape -Name 'verifyPic_df_1' -Tag 'jenkins')) 'replace: another tag is left alone'
Assert-True (-not (ExcelInsertPictures-IsOwnShape -Name 'Picture 3' -Tag 'jenkins')) 'replace: a picture put there by hand is left alone'
Assert-True (-not (ExcelInsertPictures-IsOwnShape -Name 'verifyPic_jenkins_1' -Tag '')) 'replace: no tag -> nothing is ours'

# Show-EbiExcelAt -Dwell (10/19): EBI_WATCH_MS, default 700, 0 = off
$keep = $env:EBI_WATCH_MS
$env:EBI_WATCH_MS = ''; Assert-Equal 700 (Get-EbiWatchDwellMs) 'watch: default 700 ms'
$env:EBI_WATCH_MS = '0'; Assert-Equal 0 (Get-EbiWatchDwellMs) 'watch: 0 turns it off'
$env:EBI_WATCH_MS = 'x'; Assert-Equal 700 (Get-EbiWatchDwellMs) 'watch: junk -> default'
$env:EBI_WATCH_MS = $keep

# the overview's two header rows as the 10/16 picture shows them (index 0 unused)
$lay = (Read-EbiJson -Path (Join-Path $repoRoot 'profiles/gfix-recv/layout.json'))['value']['evidence']['list']
$send = $lay['job']['group']
$group = @('', 'from', '', '', 'to', '', '', 'ID', 'IF', 'IF', 'Source', 'T', 'Z', 'TS', $send, '', '', '')
$sub   = @('', $lay['left']['sub'], 'biz', 'other', $lay['left']['sub'], 'biz', 'other', 'cid', 'ifno', 'branch', "Source IF Name`n(RCVFILE)", 'type', 'zip', 'stamp', $lay['job']['sub'], 'lib', 'dir', ($lay['extendTo']['sub'] + "`n(SNDFILE)"))
$rc = ExcelSnapList-ResolveColumns -GroupTexts $group -SubTexts $sub -Wanted @{ left = $lay['left']; job = $lay['job']; extendTo = $lay['extendTo'] }
Assert-True $rc['ok'] ('columns: resolved with the profile labels: ' + $rc['message'])
Assert-Equal 1 $rc['columns']['left'] 'columns: left = the "from" group''s system column (not the "to" one with the same sub label)'
Assert-Equal 14 $rc['columns']['job'] 'columns: job = the send group''s job column'
Assert-Equal 17 $rc['columns']['extendTo'] 'columns: extendTo = the source-file column, its "(SNDFILE)" line ignored'
$rc2 = ExcelSnapList-ResolveColumns -GroupTexts $group -SubTexts $sub -Wanted @{ left = $lay['left']; job = @{ group = 'nosuch'; sub = 'job' } }
Assert-True (-not $rc2['ok'] -and $rc2['message'] -like 'header not found: job*') ('columns: a missing header is named: ' + $rc2['message'])

# the real header of GFIX一覧.xlsx sheet GFIX送受信一覧(BIX), rows 4 / 5 (10/16 upload; labels only)
$realG = (Read-EbiJson -Path (Join-Path $here 'fixtures/gfix-list-bix-header.json'))['value']
$rr = ExcelSnapList-ResolveColumns -GroupTexts $realG['group'] -SubTexts $realG['sub'] -Wanted @{ left = $lay['left']; job = $lay['job']; extendTo = $lay['extendTo'] }
Assert-True ($rr['ok'] -and $rr['columns']['left'] -eq 2 -and $rr['columns']['job'] -eq 15 -and $rr['columns']['extendTo'] -eq 18) ('real header: B (from) .. O (send job), on to R (source file): ' + $rr['message'] + ' ' + (($rr['columns'].Keys | Sort-Object | ForEach-Object { $_ + '=' + $rr['columns'][$_] }) -join ' '))
Assert-Equal 'GFIX送受信一覧(BIX)' $lay['sheet'] 'profile: the BIX sheet is the source'

# Kase 10/13: repeat = the same in EVERY column from left to job
$jjdswm51 = @(@('Host', 'JDS', '-', 'Open', 'JPC', '-', 'JJPCRS13', 'IF6000', '001', 'JDS_PRD1_IF6000_FTP', 'TEXT', 'No', 'us', 'JJDSJM51'),
              @('Host', 'JDS', '-', 'Open', 'JPC', '-', 'JJPCRS13', 'IF6000', '001', 'JDS_PRD1_IF6000_FTP', 'TEXT', 'No', 'us', 'JJDSJM51'))
Assert-True (ExcelSnapList-HasRepeats -Rows $jjdswm51) 'repeat: JJDSWM51''s two rows are the same from "from" to the job -> run on to SNDFILE'
$diffIf = @(@('Host', 'JDS', 'IF6000', 'JJDSJM51'), @('Host', 'JDS', 'IF6001', 'JJDSJM51'))
Assert-True (-not (ExcelSnapList-HasRepeats -Rows $diffIf)) 'repeat: a different IF number is not a repeat'
Assert-True (-not (ExcelSnapList-HasRepeats -Rows @(,@('Host', 'JOD', 'CJODJSPQ')))) 'repeat: one row never repeats'
Assert-True (ExcelSnapList-HasRepeats -Rows @(@('a ', 'b'), @('a', 'b'))) 'repeat: cells compared trimmed'

Assert-Equal 'A' (ExcelSnapList-ColumnLetter 1) 'letter: 1'
Assert-Equal 'O' (ExcelSnapList-ColumnLetter 15) 'letter: 15'
Assert-Equal 'AA' (ExcelSnapList-ColumnLetter 27) 'letter: 27'
Assert-Equal '' (ExcelSnapList-NormalizeLabel $null) 'label: null'
Assert-Equal 'ab' (ExcelSnapList-NormalizeLabel " a`r`nb ") 'label: spaces and line breaks out'

Assert-True (ExcelPlacePicture-InCell -ShapeLeft 60 -ShapeTop 30 -CellLeft 60 -CellTop 30 -CellWidth 50 -CellHeight 15) 'in cell: the top-left corner on the cell corner'
Assert-True (ExcelPlacePicture-InCell -ShapeLeft 60.2 -ShapeTop 29.8 -CellLeft 60 -CellTop 30 -CellWidth 50 -CellHeight 15) 'in cell: half a point off still counts'
Assert-True (-not (ExcelPlacePicture-InCell -ShapeLeft 60 -ShapeTop 46 -CellLeft 60 -CellTop 30 -CellWidth 50 -CellHeight 15)) 'in cell: a picture starting on the next row is not taken away'

exit (Complete-Tests)
