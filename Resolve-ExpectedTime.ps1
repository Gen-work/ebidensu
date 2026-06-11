# ============================================================
# Resolve-ExpectedTime.ps1
#   Reads the Expected_Time column of the mapping CSV, creating it when
#   missing. Asks the user (keep / recent / manual input) and returns a
#   datetime. Shortcuts: edit the existing value / recent (now - 1h).
#
# Usage:
#   $dt = & .\Resolve-ExpectedTime.ps1 -CorrelId "JIGPLB1S" `
#           -MappingPath "work\mapping_owner.csv"
# ============================================================
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$CorrelId,
    [Parameter(Mandatory)][string]$MappingPath,
    [string]$TimeColumn = 'Expected_Time',
    [double]$DefaultLookbackHours = 1.0,
    [string]$IdColumn = 'Correl_ID_S',
    [string]$TimeFormat = 'yyyy/MM/dd HH:mm:ss'
)
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $MappingPath)) { throw "Mapping not found: $MappingPath" }
$mapping = @(Import-Csv -LiteralPath $MappingPath)
if ($mapping.Count -eq 0) { throw "Mapping is empty: $MappingPath" }

# add the column to every row when missing
$cols = $mapping[0].PSObject.Properties.Name
if ($cols -notcontains $TimeColumn) {
    Write-Host ("[INFO] Adding column '{0}' to mapping." -f $TimeColumn) -ForegroundColor DarkGray
    foreach ($r in $mapping) {
        $r | Add-Member -NotePropertyName $TimeColumn -NotePropertyValue '' -ErrorAction SilentlyContinue
    }
}

$row = $mapping | Where-Object { $_.$IdColumn -eq $CorrelId } | Select-Object -First 1
if (-not $row) { throw "Correl '$CorrelId' not in column '$IdColumn'." }

$existing = $row.$TimeColumn
$recent   = (Get-Date).AddHours(-$DefaultLookbackHours).ToString($TimeFormat)

Write-Host ""
Write-Host ("Correl  : {0}" -f $CorrelId) -ForegroundColor Cyan
if ([string]::IsNullOrWhiteSpace($existing)) {
    Write-Host "Current : (none)" -ForegroundColor DarkGray
    Write-Host ("  [Enter] use recent ({0})" -f $recent)
    Write-Host ("  or type: {0}" -f $TimeFormat)
    $resp = Read-Host "Time"
    if ([string]::IsNullOrWhiteSpace($resp)) { $resp = $recent }
} else {
    Write-Host ("Current : {0}" -f $existing)
    Write-Host  "  [Enter] keep current"
    Write-Host ("  r       use recent ({0})" -f $recent)
    Write-Host ("  or type: {0}" -f $TimeFormat)
    $resp = Read-Host "Time"
    if     ([string]::IsNullOrWhiteSpace($resp)) { $resp = $existing }
    elseif ($resp -eq 'r')                       { $resp = $recent }
}

# parse
$dt = $null
try {
    $dt = [datetime]::ParseExact($resp, $TimeFormat, [System.Globalization.CultureInfo]::InvariantCulture)
} catch {
    throw "Invalid time format: '$resp' (expected $TimeFormat)"
}

# write back (only when the value changed)
if ($row.$TimeColumn -ne $resp) {
    $row.$TimeColumn = $resp
    $mapping | Export-Csv -LiteralPath $MappingPath -NoTypeInformation -Encoding UTF8
    Write-Host ("  [SAVED] {0} = {1}" -f $TimeColumn, $resp) -ForegroundColor Green
}

return $dt
