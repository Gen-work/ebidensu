#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'ProcessTimeParse.ps1')

Reset-Tests 'ProcessTimeParse'

$normal = [char]0x6B63 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # seijo-shuuryo (normal end)
$abend  = [char]0x7570 + [char]0x5E38 + [char]0x7D42 + [char]0x4E86  # ijo-shuuryo (abend)

# ---------------------------------------------------------------------------
# Get-ProcessDurationText
# ---------------------------------------------------------------------------
$s1 = [datetime]'2026/07/16 09:00:00'
$e1 = [datetime]'2026/07/16 09:01:23'
Assert-Equal '00:01:23' (Get-ProcessDurationText $s1 $e1) 'duration under an hour'

$s2 = [datetime]'2026/07/15 09:00:00'
$e2 = [datetime]'2026/07/16 12:15:30'
Assert-Equal '27:15:30' (Get-ProcessDurationText $s2 $e2) 'duration over 24h is not clamped/wrapped'

Assert-Equal '' (Get-ProcessDurationText $null $e1) 'null start yields empty duration'
Assert-Equal '' (Get-ProcessDurationText $s1 $null) 'null end yields empty duration'
Assert-Equal '' (Get-ProcessDurationText $e1 $s1) 'end before start yields empty duration (never negative)'

# ---------------------------------------------------------------------------
# ConvertFrom-ProcessTimeOcrLines
# ---------------------------------------------------------------------------
$rows = @(ConvertFrom-ProcessTimeOcrLines @(
    "2026/07/16 09:00:00   2026/07/16 09:01:23   JIDSM01S   $normal",
    'no timestamps on this line at all',
    "2026/07/16   09:10:00    2026/07/16   09:12:45   JIDSM01S   $abend"
))
Assert-Equal 2 $rows.Count 'two rows carry two datetime tokens each'
Assert-Equal '2026/07/16 09:00:00' $rows[0].StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'first row start time parsed'
Assert-Equal '2026/07/16 09:01:23' $rows[0].EndTime.ToString('yyyy/MM/dd HH:mm:ss') 'first row end time parsed'
Assert-Equal $normal $rows[0].Status 'first row status literal found'
Assert-Equal $abend  $rows[1].Status 'second row (extra internal whitespace) status literal found'

$noMatch = @(ConvertFrom-ProcessTimeOcrLines @('nothing here', ''))
Assert-Equal 0 $noMatch.Count 'lines without two datetime tokens are skipped'

$badDate = @(ConvertFrom-ProcessTimeOcrLines @('2026/13/40 09:00:00   2026/07/16 09:01:23   x'))
Assert-Equal 0 $badDate.Count 'an unparseable datetime token drops the row instead of throwing'

# ---------------------------------------------------------------------------
# Get-NewestProcessTimeRow
# ---------------------------------------------------------------------------
$older = [PSCustomObject]@{ StartTime = [datetime]'2026/07/16 08:00:00'; EndTime = [datetime]'2026/07/16 08:05:00' }
$newer = [PSCustomObject]@{ StartTime = [datetime]'2026/07/16 09:00:00'; EndTime = [datetime]'2026/07/16 09:05:00' }
$picked = Get-NewestProcessTimeRow -Rows @($older, $newer)
Assert-Equal $newer.StartTime.ToString('yyyy/MM/dd HH:mm:ss') $picked.StartTime.ToString('yyyy/MM/dd HH:mm:ss') 'newest-by-StartTime wins among multiple rows'

Assert-True ($null -eq (Get-NewestProcessTimeRow -Rows @())) 'empty input returns null'
Assert-True ($null -eq (Get-NewestProcessTimeRow -Rows @($null))) 'array of only nulls returns null'

exit (Complete-Tests)
