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
$selected = @(Select-JenkinsDownloadFiles -Files $files -CorrelId 'JIDSF48S' -JobName 'JIDSJ48S')
Assert-Equal 'JIDSF48S|JIDSF48S.dat' (($selected | ForEach-Object { $_.Name }) -join '|') 'prefer Correl_ID_S exact/prefix matches'

$jobOnly = @(Select-JenkinsDownloadFiles -Files $files -CorrelId 'NO_CORREL' -JobName 'JIDSJ48S')
Assert-Equal 'JIDSJ48S.log' (($jobOnly | ForEach-Object { $_.Name }) -join '|') 'fallback to JOB_NAME prefix when no correl file matches'

$url = ConvertTo-JenkinsDownloadUri -FolderUrl 'https://jenkins.example/job/JRV/ws/out?view=1#top' -FileName 'JIDSF48S data.txt'
Assert-Equal 'https://jenkins.example/job/JRV/ws/out/JIDSF48S%20data.txt' $url 'builds file URL from folder URL and clears query/fragment'

exit (Complete-Tests)
