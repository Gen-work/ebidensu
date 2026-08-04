#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Capture at dot-source time: here $MyInvocation.MyCommand is ExternalScriptInfo (has .Path).
# Inside a function it becomes FunctionInfo (no .Path) -- so cache it now.
$_JD_ScriptDir = Split-Path $MyInvocation.MyCommand.Path

function ConvertTo-JenkinsDownloadUri {
    param(
        [Parameter(Mandatory)][string]$FolderUrl,
        [Parameter(Mandatory)][string]$FileName
    )

    $baseText = $FolderUrl.Trim()
    if ([string]::IsNullOrWhiteSpace($baseText)) { throw 'FolderUrl is empty.' }

    $builder = [System.UriBuilder]::new($baseText)
    $builder.Query = ''
    $builder.Fragment = ''

    $path = $builder.Path
    if (-not $path.EndsWith('/')) { $path += '/' }
    $builder.Path = $path + [System.Uri]::EscapeDataString($FileName)
    return $builder.Uri.AbsoluteUri
}

# The listed timestamp of one parsed entry, or $null. Read through
# PSObject.Properties so a caller passing name-only objects (and Set-StrictMode)
# cannot turn a missing column into a terminating error.
function Get-JenkinsFileTime {
    param($File)
    if ($null -eq $File) { return $null }
    if (-not $File.PSObject.Properties['DateTime']) { return $null }
    return $File.DateTime
}

# Newest first. Entries whose listed timestamp did not parse sort LAST: they
# carry no evidence about which run they are, so they must never outrank a
# dated entry. Name breaks a remaining tie so the order is deterministic.
function Sort-JenkinsFilesNewestFirst {
    param([object[]]$Files)
    return @(@($Files) | Sort-Object -Property `
        @{ Expression = { $null -ne (Get-JenkinsFileTime $_) }; Descending = $true }, `
        @{ Expression = { $t = Get-JenkinsFileTime $_; if ($null -ne $t) { $t } else { [datetime]::MinValue } }; Descending = $true }, `
        @{ Expression = { [string]$_.Name }; Descending = $false })
}

# Picks the Jenkins-listed file(s) to download for one correl.
#
# When several entries match the SAME correl they are reruns of one transfer
# (the plain and batch-stamped spellings of one file, or genuine retries), and
# only the latest one is the evidence. -PreferNewest (the default) therefore
# returns just the newest; it used to return every match, so the folder ended
# up holding the old file too and the operator downloaded the right one by
# hand. Pass -PreferNewest:$false to get every match as before.
#
# The JOB-NAME fallback is deliberately NOT narrowed: those matches are
# normally different correls of the same job, not reruns of one transfer.
function Select-JenkinsDownloadFiles {
    param(
        [Parameter(Mandatory)][array]$Files,
        [Parameter(Mandatory)][string]$CorrelId,
        [string]$JobName = '',
        [bool]$PreferNewest = $true
    )

    $correl = $CorrelId.Trim()
    $job = $JobName.Trim()
    if ([string]::IsNullOrWhiteSpace($correl)) { return @() }

    # Correl_ID_S sometimes carries the transfer-batch stamp
    # ("<correl>.<YYMMDD>.<8-digit>") that the Jenkins-listed file name may or
    # may not include. StartsWith($correl) alone only covers the direction
    # where the LISTED name is the longer (stamped) one; check both
    # directions so a stamped mapping id still finds a plainly-named file.
    $selected = @($Files | Where-Object {
        $name = [string]$_.Name
        $name -eq $correl -or
        $name.StartsWith($correl, [System.StringComparison]::OrdinalIgnoreCase) -or
        $correl.StartsWith($name, [System.StringComparison]::OrdinalIgnoreCase)
    })

    if ($selected.Count -gt 0) {
        if ($PreferNewest -and $selected.Count -gt 1) {
            return @((Sort-JenkinsFilesNewestFirst $selected)[0])
        }
        return $selected
    }
    if ([string]::IsNullOrWhiteSpace($job)) { return @() }

    return @($Files | Where-Object {
        $name = [string]$_.Name
        $name -eq $job -or
        $name.StartsWith($job, [System.StringComparison]::OrdinalIgnoreCase)
    })
}

function Invoke-JenkinsFileDownload {
    param(
        [Parameter(Mandatory)][string]$WorkDir,
        [Parameter(Mandatory)][ValidateSet('GiftRecv','GfixRecv')][string]$Mode,
        [Parameter(Mandatory)][string]$FolderUrl,
        [Parameter(Mandatory)][string]$PageText,
        [Parameter(Mandatory)][string]$CorrelId,
        [string]$JobName = '',
        [switch]$Force,
        [string]$ParserScript = '',
        [bool]$PreferNewest = $true
    )

    if ([string]::IsNullOrWhiteSpace($ParserScript)) {
        $ParserScript = Join-Path $_JD_ScriptDir 'Parse-JenkinsList.ps1'
    }
    if (-not (Test-Path -LiteralPath $ParserScript)) { throw "Parser not found: $ParserScript" }

    $allFiles = @(& $ParserScript -Text $PageText)
    # Both sets, so the caller can SAY which older entries were passed over
    # instead of silently downloading one of several.
    $allMatches = @(Select-JenkinsDownloadFiles -Files $allFiles -CorrelId $CorrelId -JobName $JobName -PreferNewest $false)
    $matches    = @(Select-JenkinsDownloadFiles -Files $allFiles -CorrelId $CorrelId -JobName $JobName -PreferNewest $PreferNewest)

    $keptNames = @{}
    foreach ($m in $matches) { $keptNames[[string]$m.Name] = $true }
    $superseded = @($allMatches | Where-Object { -not $keptNames.ContainsKey([string]$_.Name) })

    $dataKind = if ($Mode -eq 'GiftRecv') { 'GIFT' } else { 'GFIX' }
    $dataDir = Join-Path (Join-Path $WorkDir 'DATA') $dataKind
    Ensure-Dir $dataDir

    $result = [ordered]@{
        DataKind = $dataKind
        DataDir = $dataDir
        Found = $allFiles.Count
        Matched = $matches.Count
        # Everything that matched this correl before the newest-wins narrowing,
        # and the older entries it passed over ({ Name; DateTime } each).
        Candidates = $allMatches.Count
        Superseded = @($superseded | ForEach-Object {
            [pscustomobject]@{ Name = [string]$_.Name; DateTime = $_.DateTime }
        })
        Downloaded = 0
        Skipped = 0
        Failed = 0
        Files = @()
    }

    foreach ($file in $matches) {
        $name = [string]$file.Name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        $destPath = Join-Path $dataDir $name
        $fileResult = [ordered]@{ Name = $name; Path = $destPath; Status = ''; Message = '' }

        if ((Test-Path -LiteralPath $destPath) -and -not $Force.IsPresent) {
            $fileResult.Status = 'skip'
            $fileResult.Message = 'exists'
            $result.Skipped++
            $result.Files += [pscustomobject]$fileResult
            continue
        }

        $url = ConvertTo-JenkinsDownloadUri -FolderUrl $FolderUrl -FileName $name
        try {
            Invoke-WebRequest -Uri $url -OutFile $destPath -UseBasicParsing -UseDefaultCredentials
            $fileResult.Status = 'ok'
            $fileResult.Message = $url
            $result.Downloaded++
        } catch {
            $fileResult.Status = 'fail'
            $fileResult.Message = $_.Exception.Message
            $result.Failed++
            if (Test-Path -LiteralPath $destPath) {
                Remove-Item -LiteralPath $destPath -Force -ErrorAction SilentlyContinue
            }
        }
        $result.Files += [pscustomobject]$fileResult
    }

    return [pscustomobject]$result
}
