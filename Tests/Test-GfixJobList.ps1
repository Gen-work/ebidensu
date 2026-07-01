#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'GfixJobList.ps1')

Reset-Tests 'GfixJobList'

$t  = "`t"
$nl = "`n"

# Header row (Japanese column names) -- must be ignored (JobNo column is not numeric)
$header = "JobNoHeader${t}ProjectName${t}Folder${t}Status${t}User${t}Start${t}End${t}Seconds${t}Method"

# Data rows lifted from a real Ctrl+A capture, incl. the duplicate-IF_NO
# case that caused JIDSC02S / JIDSC03S to both fetch job 1000002995601.
$row1 = "1000002996220${t}JOD_VER1_IF5058_001_SSH${t}/igp2usr/Receive${t}${t}admin${t}2026-07-01 11:19:13${t}2026-07-01 11:19:28${t}14.74${t}Trigger"
$row2 = "1000002996219${t}JOD_VER1_IF5058_001_SSH${t}/igp2usr/Receive${t}${t}admin${t}2026-07-01 11:19:13${t}2026-07-01 11:19:16${t}2.82${t}Trigger"
$row3 = "1000002996216${t}JOD_VER1_IF5058_FTP${t}/jodusr/Send${t}${t}admin${t}2026-07-01 11:19:09${t}2026-07-01 11:19:13${t}4.43${t}Trigger"
$row4 = "1000002995602${t}JOD_VER1_IF5001_001_SSH${t}/idsusr/Receive${t}${t}admin${t}2026-07-01 11:02:36${t}2026-07-01 11:02:41${t}5.24${t}Trigger"
$row5 = "1000002995601${t}JOD_VER1_IF5001_001_SSH${t}/idsusr/Receive${t}${t}admin${t}2026-07-01 11:02:34${t}2026-07-01 11:02:39${t}4.09${t}Trigger"

$footer1 = 'displaying 1 - 5 of 5'
$footer2 = '(c) 2008-2025 Fortra.'

$fullText = @"
GoAnywhere
$header

$row1

$row2

$row3

$row4

$row5

$footer1
$footer2
"@

# -- ConvertFrom-GfixJobListText --
# (Convert-From helpers return plain arrays -- always wrap the call in @()
# at the call site, per the project's PS 5.1 single/zero-element convention.)
$rows = @(ConvertFrom-GfixJobListText $fullText)
Assert-Equal 5 $rows.Count 'parses exactly the 5 numeric-JobNo data rows (header/footer/blank skipped)'
Assert-Equal '1000002996220' $rows[0].JobNo    'row1 JobNo'
Assert-Equal 'JOD_VER1_IF5058_001_SSH' $rows[0].ProjectName 'row1 ProjectName'
Assert-Equal '/igp2usr/Receive' $rows[0].Folder 'row1 Folder'
Assert-Equal ''  $rows[0].Status                'row1 Status (empty tab-slot preserved)'
Assert-Equal 'admin' $rows[0].User              'row1 User'
Assert-Equal 'Trigger' $rows[0].Method          'row1 Method (last column)'

$empty = @(ConvertFrom-GfixJobListText '')
Assert-Equal 0 $empty.Count 'empty text -> zero rows'

$nullText = @(ConvertFrom-GfixJobListText $null)
Assert-Equal 0 $nullText.Count 'null text -> zero rows'

# -- Get-GfixJobListRowsForIf --
$if5058 = @(Get-GfixJobListRowsForIf -Rows $rows -IfNorm 'IF5058_001')
Assert-Equal 2 $if5058.Count 'IF5058_001: matches only the 2 receive-side (SSH) rows, not the FTP sibling'
Assert-True (@($if5058 | ForEach-Object { $_.JobNo }) -contains '1000002996220') 'IF5058_001 includes job 1000002996220'
Assert-True (@($if5058 | ForEach-Object { $_.JobNo }) -contains '1000002996219') 'IF5058_001 includes job 1000002996219'

$if5001 = @(Get-GfixJobListRowsForIf -Rows $rows -IfNorm 'IF5001_001')
Assert-Equal 2 $if5001.Count 'IF5001_001: duplicate IF_NO -> both distinct job numbers returned (not just one)'
Assert-True (@($if5001 | ForEach-Object { $_.JobNo }) -contains '1000002995602') 'IF5001_001 includes job 1000002995602'
Assert-True (@($if5001 | ForEach-Object { $_.JobNo }) -contains '1000002995601') 'IF5001_001 includes job 1000002995601'

$ifMissing = @(Get-GfixJobListRowsForIf -Rows $rows -IfNorm 'IF9999_001')
Assert-Equal 0 $ifMissing.Count 'IF_NO absent from the list -> zero rows (caller marks GFIX_log=2)'

$ifSendOnly = @(Get-GfixJobListRowsForIf -Rows $rows -IfNorm 'IF5058' -ReceiveOnly $false)
Assert-Equal 3 $ifSendOnly.Count 'ReceiveOnly=$false includes the FTP/Send sibling too (all 3 IF5058 rows)'

$blankIf = @(Get-GfixJobListRowsForIf -Rows $rows -IfNorm '')
Assert-Equal 0 $blankIf.Count 'blank IfNorm -> zero rows'

# -- Get-GfixLogDownloadPlan --
# Three correls: JIDSC02S/JIDSC03S both off the duplicate IF5001_001 (the real
# bug scenario), and JIGPK01S off IF5058_001 (single-match, unaffected today).
# A fourth, JIDSX99S, needs an IF_NO absent from the list entirely (hard miss).
$pendingPlan = @(
    [pscustomobject]@{ CorrelIdS = 'JIDSC02S'; IfNorm = 'IF5001_001' },
    [pscustomobject]@{ CorrelIdS = 'JIDSC03S'; IfNorm = 'IF5001_001' },
    [pscustomobject]@{ CorrelIdS = 'JIGPK01S'; IfNorm = 'IF5058_001' },
    [pscustomobject]@{ CorrelIdS = 'JIDSX99S'; IfNorm = 'IF9999_001' }
)
$plan = Get-GfixLogDownloadPlan -PendingRows $pendingPlan -JobListRows $rows -ReceiveOnly $true

$hardMiss = @($plan.HardMiss)
Assert-Equal 1 $hardMiss.Count 'plan: exactly 1 hard-miss correl (IF9999_001 has no list row)'
Assert-Equal 'JIDSX99S' $hardMiss[0].CorrelIdS 'plan: hard-miss correl is JIDSX99S'
Assert-Equal 'IF9999_001' $hardMiss[0].IfNorm  'plan: hard-miss carries its IfNorm for the error message'

$needed = @($plan.NeededJobNumbers)
Assert-Equal 4 $needed.Count 'plan: 4 distinct job numbers needed (2 for IF5001_001 + 2 for IF5058_001, both duplicate IF_NOs in this fixture) -- not just 1'
Assert-True ($needed -contains '1000002995602') 'plan: needed includes duplicate-IF_NO job 1000002995602'
Assert-True ($needed -contains '1000002995601') 'plan: needed includes duplicate-IF_NO job 1000002995601 (this is the one the old Ctrl+F-by-project-name search could never reach for a second correl)'
Assert-True ($needed -contains '1000002996220') 'plan: needed includes IF5058_001 job 1000002996220 (JIGPK01S only needs it, but its IF_NO also happens to have a duplicate row here)'

# A second correl group sharing an ALREADY-planned job number must not
# duplicate it in NeededJobNumbers (download-once-per-job-number dedup).
$pendingDup = @(
    [pscustomobject]@{ CorrelIdS = 'JIDSC02S'; IfNorm = 'IF5001_001' },
    [pscustomobject]@{ CorrelIdS = 'JIDSC03S'; IfNorm = 'IF5001_001' }
)
$planDup = Get-GfixLogDownloadPlan -PendingRows $pendingDup -JobListRows $rows -ReceiveOnly $true
Assert-Equal 0 @($planDup.HardMiss).Count 'plan(dup only): no hard miss'
Assert-Equal 2 @($planDup.NeededJobNumbers).Count 'plan(dup only): still exactly 2 job numbers (deduped), not 4'

exit (Complete-Tests)
