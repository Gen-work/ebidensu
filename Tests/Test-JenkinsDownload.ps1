#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'JenkinsDownload.ps1')

Reset-Tests 'JenkinsDownload'

$files = @(
    [pscustomobject]@{ Name = 'JIDSF48S' },
    [pscustomobject]@{ Name = 'JIDSF48S.dat' },
    [pscustomobject]@{ Name = 'JIDSJ48S.log' },
    [pscustomobject]@{ Name = 'OTHER' }
)
# -PreferNewest $false keeps the original contract: every correl match.
$selected = @(Select-JenkinsDownloadFiles -Files $files -CorrelId 'JIDSF48S' -JobName 'JIDSJ48S' -PreferNewest $false)
Assert-Equal 'JIDSF48S|JIDSF48S.dat' (($selected | ForEach-Object { $_.Name }) -join '|') 'prefer Correl_ID_S exact/prefix matches'

# Name-only entries (no DateTime column at all) must not throw under
# Set-StrictMode -- the newest-wins path reads the timestamp defensively.
$undated = @(Select-JenkinsDownloadFiles -Files $files -CorrelId 'JIDSF48S' -JobName 'JIDSJ48S')
Assert-Equal 1 $undated.Count 'newest-wins narrows several correl matches to one'
Assert-Equal 'JIDSF48S' $undated[0].Name 'undated entries tie-break deterministically by name'

$jobOnly = @(Select-JenkinsDownloadFiles -Files $files -CorrelId 'NO_CORREL' -JobName 'JIDSJ48S')
Assert-Equal 'JIDSJ48S.log' (($jobOnly | ForEach-Object { $_.Name }) -join '|') 'fallback to JOB_NAME prefix when no correl file matches'

# Mapping's Correl_ID_S carries the transfer-batch stamp, but the Jenkins
# folder lists the file under its plain (shorter) name -- StartsWith($correl)
# alone cannot match this direction; the reverse StartsWith($name) must.
$stampedFiles = @([pscustomobject]@{ Name = 'JIDSU86S' })
$stampedSelected = @(Select-JenkinsDownloadFiles -Files $stampedFiles -CorrelId 'JIDSU86S.260729.10515511' -JobName '')
Assert-Equal 'JIDSU86S' (($stampedSelected | ForEach-Object { $_.Name }) -join '|') 'batch-stamped Correl_ID_S still matches a plainly-listed file'

# -- newest-wins among reruns of one transfer -------------------------------
# The operator's report: a correl with several Jenkins entries got the OLD file
# downloaded, so the newest had to be fetched and swapped in by hand.
$reruns = @(
    [pscustomobject]@{ Name = 'JIDSU86S';                   DateTime = [datetime]'2026/07/24 09:50:03' },
    [pscustomobject]@{ Name = 'JIDSU86S.260729.10515511';   DateTime = [datetime]'2026/07/29 10:51:55' },
    [pscustomobject]@{ Name = 'JIDSU86S.260726.08300100';   DateTime = [datetime]'2026/07/26 08:30:01' },
    [pscustomobject]@{ Name = 'OTHER';                      DateTime = [datetime]'2026/07/30 11:00:00' }
)
$newest = @(Select-JenkinsDownloadFiles -Files $reruns -CorrelId 'JIDSU86S' -JobName '')
Assert-Equal 1 $newest.Count 'several reruns of one transfer collapse to one download'
Assert-Equal 'JIDSU86S.260729.10515511' $newest[0].Name 'the newest rerun is the one downloaded'

$allReruns = @(Select-JenkinsDownloadFiles -Files $reruns -CorrelId 'JIDSU86S' -JobName '' -PreferNewest $false)
Assert-Equal 3 $allReruns.Count 'PreferNewest $false still returns every match'

# A batch-stamped mapping id picks the newest among the same candidates.
$fromStamped = @(Select-JenkinsDownloadFiles -Files $reruns -CorrelId 'JIDSU86S.260726.08300100' -JobName '')
Assert-Equal 'JIDSU86S.260729.10515511' $fromStamped[0].Name 'a stamped mapping id also resolves to the newest listed entry'

# An entry whose timestamp did not parse must never outrank a dated one.
$mixed = @(
    [pscustomobject]@{ Name = 'JIDSK11S.badstamp'; DateTime = $null },
    [pscustomobject]@{ Name = 'JIDSK11S';          DateTime = [datetime]'2026/07/24 09:50:03' }
)
$mixedPick = @(Select-JenkinsDownloadFiles -Files $mixed -CorrelId 'JIDSK11S' -JobName '')
Assert-Equal 'JIDSK11S' $mixedPick[0].Name 'an undated entry never beats a dated one'

# Sort helper on its own.
$sorted = @(Sort-JenkinsFilesNewestFirst $reruns)
Assert-Equal 'OTHER' $sorted[0].Name 'sort helper puts the newest first'
Assert-True ($null -eq (Get-JenkinsFileTime ([pscustomobject]@{ Name = 'x' }))) 'a missing DateTime column reads as $null, not an error'

# The JOB_NAME fallback is deliberately NOT narrowed: those are normally
# different correls of one job, not reruns of one transfer.
$jobFiles = @(
    [pscustomobject]@{ Name = 'JIDSJ48S_a'; DateTime = [datetime]'2026/07/24 09:00:00' },
    [pscustomobject]@{ Name = 'JIDSJ48S_b'; DateTime = [datetime]'2026/07/25 09:00:00' }
)
$jobPick = @(Select-JenkinsDownloadFiles -Files $jobFiles -CorrelId 'NO_CORREL' -JobName 'JIDSJ48S')
Assert-Equal 2 $jobPick.Count 'the JOB_NAME fallback still returns every match'

$url = ConvertTo-JenkinsDownloadUri -FolderUrl 'https://jenkins.example/job/JRV/ws/out?view=1#top' -FileName 'JIDSF48S data.txt'
Assert-Equal 'https://jenkins.example/job/JRV/ws/out/JIDSF48S%20data.txt' $url 'builds file URL from folder URL and clears query/fragment'

exit (Complete-Tests)
