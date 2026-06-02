#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Capture at dot-source time: here $MyInvocation.MyCommand is ExternalScriptInfo (has .Path).
# Inside a function it becomes FunctionInfo (no .Path) — so cache it now.
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

function Select-JenkinsDownloadFiles {
    param(
        [Parameter(Mandatory)][array]$Files,
        [Parameter(Mandatory)][string]$CorrelId,
        [string]$JobName = ''
    )

    $correl = $CorrelId.Trim()
    $job = $JobName.Trim()
    if ([string]::IsNullOrWhiteSpace($correl)) { return @() }

    $selected = @($Files | Where-Object {
        $name = [string]$_.Name
        $name -eq $correl -or
        $name.StartsWith($correl, [System.StringComparison]::OrdinalIgnoreCase)
    })

    if ($selected.Count -gt 0 -or [string]::IsNullOrWhiteSpace($job)) { return $selected }

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
        [string]$ParserScript = ''
    )

    if ([string]::IsNullOrWhiteSpace($ParserScript)) {
        $ParserScript = Join-Path $_JD_ScriptDir 'Parse-JenkinsList.ps1'
    }
    if (-not (Test-Path -LiteralPath $ParserScript)) { throw "Parser not found: $ParserScript" }

    $allFiles = @(& $ParserScript -Text $PageText)
    $matches = @(Select-JenkinsDownloadFiles -Files $allFiles -CorrelId $CorrelId -JobName $JobName)

    $dataKind = if ($Mode -eq 'GiftRecv') { 'GIFT' } else { 'GFIX' }
    $dataDir = Join-Path (Join-Path $WorkDir 'DATA') $dataKind
    Ensure-Dir $dataDir

    $result = [ordered]@{
        DataKind = $dataKind
        DataDir = $dataDir
        Found = $allFiles.Count
        Matched = $matches.Count
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
