# ============================================================
#  WorkbookResolver.ps1
#
#  Dot-source helper for resolving evidence/J4 workbook filenames.
#  ASCII source -- no raw Japanese literals.
#  NOTE: return filesystem ProviderPath values (not provider-qualified .Path) so
#  Excel COM can open UNC paths directly.
#
#  Usage pattern in callers:
#    $prefix   = Resolve-ExcelPrefix -Row $first -DefaultPrefix $ExcelPrefix
#    $fullStem = Get-ExcelFullStem -Prefix $prefix -Name ([string]$first.Excel_NAME)
#    $wbPath   = Find-WorkbookByExcelName -Dir $evDir -ExcelName $fullStem
#    $destLeaf = Get-ExcelDestLeaf $fullStem
#
#  When the effective prefix is empty the full stem equals Excel_NAME (legacy behaviour).
# ============================================================

# Combines the effective J4 prefix with the short name column.
#   Prefix     = 'J4<title>(REQ-000xxxxx_GIFT<suffix>)'  (or '')
#   Excel_NAME   = 'LJRVWD64'
#   -> 'J4<title>(REQ-000xxxxx_GIFT<suffix>)_LJRVWD64'    (when prefix set)
#   -> 'LJRVWD64'                                          (when prefix empty)
function Get-ExcelFullStem {
    param([string]$Prefix, [string]$Name)
    $n = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    if ([string]::IsNullOrWhiteSpace($Prefix)) { return $n }
    return "{0}_{1}" -f $Prefix.TrimEnd('_'), $n
}

# Returns the filename to create when cloning from a template.
#   Get-ExcelDestLeaf 'J4..._LJRVWD64'  -> 'J4..._LJRVWD64.xlsx'
#   Get-ExcelDestLeaf 'LJRVWD64'        -> 'LJRVWD64.xlsx'

# Resolve the effective workbook prefix. A project-level prefix should normally
# come from verify_config.json / VerifyConfig.psd1. The legacy per-row
# Excel_Prefix column is still honored as an override for old mappings or rare
# per-workbook exceptions.
function Resolve-ExcelPrefix {
    param([object]$Row, [string]$DefaultPrefix = '')
    $rowPrefix = ''
    if ($null -ne $Row -and ($Row.PSObject.Properties.Name -contains 'Excel_Prefix')) {
        $rowPrefix = [string]$Row.Excel_Prefix
    }
    if (-not [string]::IsNullOrWhiteSpace($rowPrefix)) { return $rowPrefix }
    return [string]$DefaultPrefix
}

function Get-ExcelDestLeaf {
    param([string]$FullStem)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($FullStem)
    return "{0}.xlsx" -f $stem
}

# Inverse of Get-ExcelFullStem: given a real workbook filename and the short
# Excel_NAME, recover the J4 prefix that precedes "_<Excel_NAME>".
#   Get-PrefixFromFilename 'J4title(REQ-...)_LJRVWD64.xlsx' 'LJRVWD64'
#       -> 'J4title(REQ-...)'
#   Get-PrefixFromFilename 'LJRVWD64.xlsx' 'LJRVWD64' -> ''   (no prefix)
# Returns '' when the filename does not carry a recoverable prefix.
function Get-PrefixFromFilename {
    param([string]$FileName, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($FileName) -or [string]::IsNullOrWhiteSpace($Name)) { return '' }
    $stem  = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $short = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    if ($stem -eq $short) { return '' }
    $suffix = '_' + $short
    if ($stem.EndsWith($suffix)) {
        return $stem.Substring(0, $stem.Length - $suffix.Length)
    }
    return ''
}

function Resolve-WorkbookProviderPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction Stop
    if ($resolved.ProviderPath) { return $resolved.ProviderPath }
    return $resolved.Path
}

class FullWidthFilenameResolver {
    static [string] ConvertFullWidthAsciiToHalfWidth([string]$Value) {
        if ([string]::IsNullOrEmpty($Value)) { return $Value }

        $chars = [System.Collections.Generic.List[char]]::new()
        foreach ($ch in $Value.ToCharArray()) {
            $code = [int][char]$ch
            if ($code -eq 0x3000) {
                [void]$chars.Add([char]0x20)
            } elseif ($code -ge 0xFF01 -and $code -le 0xFF5E) {
                [void]$chars.Add([char]($code - 0xFEE0))
            } else {
                [void]$chars.Add($ch)
            }
        }
        return (-join $chars.ToArray())
    }

    static [bool] ContainsFullWidthAscii([string]$Value) {
        if ([string]::IsNullOrEmpty($Value)) { return $false }
        foreach ($ch in $Value.ToCharArray()) {
            $code = [int][char]$ch
            if ($code -eq 0x3000 -or ($code -ge 0xFF01 -and $code -le 0xFF5E)) { return $true }
        }
        return $false
    }

    static [bool] IsCandidate([string]$RequestedName, [string]$CandidateName, [bool]$AllowSuffixMatch) {
        if ([string]::IsNullOrWhiteSpace($RequestedName) -or [string]::IsNullOrWhiteSpace($CandidateName)) { return $false }
        if (-not [FullWidthFilenameResolver]::ContainsFullWidthAscii($CandidateName)) { return $false }

        $requestedStem = [FullWidthFilenameResolver]::ConvertFullWidthAsciiToHalfWidth([System.IO.Path]::GetFileNameWithoutExtension($RequestedName))
        $candidateStem = [FullWidthFilenameResolver]::ConvertFullWidthAsciiToHalfWidth([System.IO.Path]::GetFileNameWithoutExtension($CandidateName))
        if ($candidateStem -eq $requestedStem) { return $true }
        if ($AllowSuffixMatch) {
            return ($candidateStem.EndsWith('_' + $requestedStem) -or $candidateStem.EndsWith($requestedStem))
        }
        return $false
    }

    static [System.IO.FileInfo[]] FindCandidates([string]$Dir, [string]$Name, [string]$Filter, [bool]$Recurse, [bool]$AllowSuffixMatch) {
        if ([string]::IsNullOrWhiteSpace($Dir) -or [string]::IsNullOrWhiteSpace($Name)) { return [System.IO.FileInfo[]]@() }
        if ([string]::IsNullOrWhiteSpace($Filter)) { $Filter = '*' }

        $dirInfo = [System.IO.DirectoryInfo]::new($Dir)
        if (-not $dirInfo.Exists) { return [System.IO.FileInfo[]]@() }

        $option = [System.IO.SearchOption]::TopDirectoryOnly
        if ($Recurse) { $option = [System.IO.SearchOption]::AllDirectories }

        $matches = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
        foreach ($item in $dirInfo.GetFiles($Filter, $option)) {
            if ([FullWidthFilenameResolver]::IsCandidate($Name, $item.Name, $AllowSuffixMatch)) {
                [void]$matches.Add($item)
            }
        }

        $sorted = @($matches.ToArray() | Sort-Object @{Expression='LastWriteTime';Descending=$true}, FullName)
        return [System.IO.FileInfo[]]$sorted
    }
}

function Convert-FullWidthAsciiToHalfWidth {
    param([string]$Value)
    return [FullWidthFilenameResolver]::ConvertFullWidthAsciiToHalfWidth($Value)
}

function Test-ContainsFullWidthAscii {
    param([string]$Value)
    return [FullWidthFilenameResolver]::ContainsFullWidthAscii($Value)
}

function Find-FullWidthFilenameCandidates {
    param(
        [string]$Dir,
        [string]$Name,
        [string]$Filter = '*',
        [switch]$Recurse,
        [switch]$AllowSuffixMatch
    )
    return [FullWidthFilenameResolver]::FindCandidates($Dir, $Name, $Filter, $Recurse.IsPresent, $AllowSuffixMatch.IsPresent)
}

function Confirm-FullWidthFilenameCandidate {
    param(
        [string]$RequestedPath,
        [System.IO.FileInfo]$Candidate,
        [string]$ItemKind = 'file'
    )
    Write-Host ("[WARN] {0} not found: {1}" -f $ItemKind, $RequestedPath) -ForegroundColor Yellow
    Write-Host ("[WARN] Possible full-width filename found: {0}" -f $Candidate.FullName) -ForegroundColor Yellow
    try {
        $resp = Read-Host ("Use this {0} instead? [y/N]" -f $ItemKind)
        return ($resp -match '^(?i:y|yes)$')
    } catch {
        return $false
    }
}

function Resolve-FullWidthFileName {
    param(
        [string]$Dir,
        [string]$Name,
        [string]$Filter = '*',
        [switch]$Recurse,
        [switch]$AllowSuffixMatch,
        [string]$RequestedPath = '',
        [string]$ItemKind = 'file',
        [ValidateSet('Prompt','Accept','Reject')][string]$FullWidthFallback = 'Prompt'
    )
    $candidates = @(Find-FullWidthFilenameCandidates -Dir $Dir -Name $Name -Filter $Filter -Recurse:$Recurse.IsPresent -AllowSuffixMatch:$AllowSuffixMatch.IsPresent)
    if ([string]::IsNullOrWhiteSpace($RequestedPath)) { $RequestedPath = Join-Path $Dir $Name }

    foreach ($candidate in $candidates) {
        if ($FullWidthFallback -eq 'Accept' -or ($FullWidthFallback -eq 'Prompt' -and (Confirm-FullWidthFilenameCandidate -RequestedPath $RequestedPath -Candidate $candidate -ItemKind $ItemKind))) {
            return $candidate.FullName
        }
        if ($FullWidthFallback -eq 'Reject') { break }
    }
    return $null
}

function Get-FullWidthWorkbookCandidates {
    param([string]$Dir, [string]$ExcelName, [switch]$Recurse)
    return Find-FullWidthFilenameCandidates -Dir $Dir -Name $ExcelName -Filter '*.xlsx' -Recurse:$Recurse.IsPresent -AllowSuffixMatch
}

function Confirm-FullWidthWorkbookCandidate {
    param([string]$RequestedPath, [System.IO.FileInfo]$Candidate)
    return Confirm-FullWidthFilenameCandidate -RequestedPath $RequestedPath -Candidate $Candidate -ItemKind 'workbook'
}

# Finds the evidence/J4 workbook file for a given stem (full or short).
# Search order:
#   1. Exact match: <Dir>\<FullStem>.xlsx
#   2. Wildcard:    <Dir>\*_<FullStem>.xlsx  (handles any extra prefix on disk)
#   3. Wildcard:    <Dir>\*<FullStem>.xlsx
#   4. Full-width ASCII fallback: scan workbook names, normalize full-width
#      ASCII forms to half-width, warn, and ask before using the candidate.
#      Pass -FullWidthFallback Accept/Reject in tests or batch callers to avoid
#      an interactive prompt.
# Returns newest LastWriteTime when multiple wildcard hits.
function Find-WorkbookByExcelName {
    param(
        [string]$Dir,
        [string]$ExcelName,
        [switch]$Recurse,
        [ValidateSet('Prompt','Accept','Reject')][string]$FullWidthFallback = 'Prompt'
    )
    if ([string]::IsNullOrWhiteSpace($Dir) -or [string]::IsNullOrWhiteSpace($ExcelName) -or -not (Test-Path -LiteralPath $Dir)) { return $null }
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($ExcelName)
    $leaf = if ($ExcelName -match '\.xlsx$') { $ExcelName } else { ("{0}.xlsx" -f $stem) }
    $exact = Join-Path $Dir $leaf
    if (Test-Path -LiteralPath $exact) { return (Resolve-WorkbookProviderPath $exact) }

    $hits = @()
    foreach ($pattern in @(("*_{0}.xlsx" -f $stem), ("*{0}.xlsx" -f $stem))) {
        if ($Recurse.IsPresent) {
            $hits += @(Get-ChildItem -LiteralPath $Dir -Filter $pattern -File -Recurse -ErrorAction SilentlyContinue)
        } else {
            $hits += @(Get-ChildItem -LiteralPath $Dir -Filter $pattern -File -ErrorAction SilentlyContinue)
        }
    }
    $hits = @($hits | Sort-Object FullName -Unique)
    if ($hits.Count -gt 0) { return ($hits | Sort-Object @{Expression='LastWriteTime';Descending=$true}, FullName | Select-Object -First 1).FullName }

    return (Resolve-FullWidthFileName -Dir $Dir -Name $leaf -Filter '*.xlsx' -Recurse:$Recurse.IsPresent -AllowSuffixMatch -RequestedPath $exact -ItemKind 'workbook' -FullWidthFallback $FullWidthFallback)
}
