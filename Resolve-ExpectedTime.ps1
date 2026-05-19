# ============================================================
# Resolve-ExpectedTime.ps1
#   mapping CSV の Expected_Time 列を読み、無ければ作る。
#   ユーザーに確認 (keep / recent / 手入力) し datetime を返す。
#   既存値の編集 / recent (now - 1h) のショートカットあり。
#
# 使用例:
#   $dt = & .\Resolve-ExpectedTime.ps1 -CorrelId "JIGPLB1S" `
#           -MappingPath "work\mapping_owner.csv"
# ============================================================
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$CorrelId,
    [Parameter(Mandatory)][string]$MappingPath,
    [string]$TimeColumn = 'Expected_Time',
    [double]$DefaultLookbackHours = 1.0,
    [string]$IdColumn = 'Correl_ID_S'
)
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $MappingPath)) { throw "Mapping not found: $MappingPath" }
$mapping = @(Import-Csv -LiteralPath $MappingPath)
if ($mapping.Count -eq 0) { throw "Mapping is empty: $MappingPath" }

# 列が無ければ全行に追加
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
$recent   = (Get-Date).AddHours(-$DefaultLookbackHours).ToString('yyyy/MM/dd HH:mm:ss')

Write-Host ""
Write-Host ("Correl  : {0}" -f $CorrelId) -ForegroundColor Cyan
if ([string]::IsNullOrWhiteSpace($existing)) {
    Write-Host "Current : (none)" -ForegroundColor DarkGray
    Write-Host ("  [Enter] use recent ({0})" -f $recent)
    Write-Host  "  or type: yyyy/MM/dd HH:mm:ss"
    $resp = Read-Host "Time"
    if ([string]::IsNullOrWhiteSpace($resp)) { $resp = $recent }
} else {
    Write-Host ("Current : {0}" -f $existing)
    Write-Host  "  [Enter] keep current"
    Write-Host ("  r       use recent ({0})" -f $recent)
    Write-Host  "  or type: yyyy/MM/dd HH:mm:ss"
    $resp = Read-Host "Time"
    if     ([string]::IsNullOrWhiteSpace($resp)) { $resp = $existing }
    elseif ($resp -eq 'r')                       { $resp = $recent }
}

# 解析
$dt = $null
try {
    $dt = [datetime]::ParseExact($resp, 'yyyy/MM/dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
} catch {
    throw "Invalid time format: '$resp' (expected yyyy/MM/dd HH:mm:ss)"
}

# 書き戻し（値が変わった場合のみ）
if ($row.$TimeColumn -ne $resp) {
    $row.$TimeColumn = $resp
    $mapping | Export-Csv -LiteralPath $MappingPath -NoTypeInformation -Encoding UTF8
    Write-Host ("  [SAVED] {0} = {1}" -f $TimeColumn, $resp) -ForegroundColor Green
}

return $dt