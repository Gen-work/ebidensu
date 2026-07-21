# Pure helpers for discovering mapping JOB_NAME input and applying template order.

function ConvertFrom-MappingIdText {
    param([string[]]$Text)

    $seen = New-Object 'System.Collections.Generic.HashSet[System.String]'
    $jobs = New-Object 'System.Collections.Generic.List[System.String]'
    foreach ($line in @($Text)) {
        foreach ($match in [regex]::Matches(([string]$line).ToUpperInvariant(), '(?<![A-Z0-9])[A-Z]JOD[WJ][A-Z0-9]{3}(?![A-Z0-9])')) {
            $value = $match.Value
            if ($value[4] -eq 'W') { $value = $value.Substring(0, 4) + 'J' + $value.Substring(5) }
            if ($seen.Add($value)) { [void]$jobs.Add($value) }
        }
    }
    return @($jobs)
}

function Get-MappingIdInput {
    param(
        [string]$WorkDir,
        [scriptblock]$OcrImage
    )

    $snap = Join-Path $WorkDir 'snap'
    $txt = Join-Path (Join-Path $snap 'GFIX_HM') 'ID.txt'
    if (Test-Path -LiteralPath $txt -PathType Leaf) {
        $jobs = @(ConvertFrom-MappingIdText -Text @(Get-Content -LiteralPath $txt -ErrorAction Stop))
        if ($jobs.Count -gt 0) { return [pscustomobject]@{ Jobs=$jobs; Source=$txt; Kind='text' } }
    }

    foreach ($folder in @('GIFT_HM', 'GFIX_HM')) {
        $png = Join-Path (Join-Path $snap $folder) 'ID.png'
        if (-not (Test-Path -LiteralPath $png -PathType Leaf)) { continue }
        if ($null -eq $OcrImage) { continue }
        $jobs = @(ConvertFrom-MappingIdText -Text @(& $OcrImage $png))
        if ($jobs.Count -gt 0) { return [pscustomobject]@{ Jobs=$jobs; Source=$png; Kind='ocr' } }
    }
    return $null
}

function Sort-MappingJobsByTemplateOrder {
    param([string[]]$Jobs, [string[]]$TemplateJobs)

    $rank = @{}
    $i = 0
    foreach ($job in @(ConvertFrom-MappingIdText -Text $TemplateJobs)) {
        if (-not $rank.ContainsKey($job)) { $rank[$job] = $i; $i++ }
    }
    $original = @{}
    $i = 0
    foreach ($job in @($Jobs)) { if (-not $original.ContainsKey($job)) { $original[$job] = $i }; $i++ }
    return @($Jobs | Sort-Object @{ Expression={ if ($rank.ContainsKey($_)) { 0 } else { 1 } } },
                                  @{ Expression={ if ($rank.ContainsKey($_)) { $rank[$_] } else { $original[$_] } } })
}
