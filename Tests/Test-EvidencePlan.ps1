#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'MappingStore.ps1')
. (Join-Path (Split-Path $here -Parent) 'EvidencePlan.ps1')

Reset-Tests 'EvidencePlan'

function Kinds([object[]]$plan) { return (@($plan | ForEach-Object { $_.Kind }) -join ',') }
# StrictMode-safe: ops are hashtables; not every op has a 'Folder' key.
function Folders([object[]]$plan) {
    return @($plan | ForEach-Object { if ($_.ContainsKey('Folder')) { [string]$_.Folder } else { '' } })
}
function FirstOp([object[]]$plan, [string]$folder) {
    foreach ($op in @($plan)) { if ($op.Kind -eq 'picture' -and $op.Folder -eq $folder) { return $op } }
    return $null
}

# -- Select-ValidCorrelIds --
$valid = Select-ValidCorrelIds @('JIGPF48S', '#VALUE!', '', '  ', 'JIDSF48S.260729.10515511')
Assert-Equal 'JIGPF48S|JIDSF48S.260729.10515511' ($valid -join '|') 'drops invalid values and keeps plain/timestamped correl order'

# -- DF plan (spec 7) --
$df = Build-DfEvidencePlan -SnapRoot 'X' -CorrelOrder @('A','B')
Assert-Equal 'text,picture,blank,text,picture,blank' (Kinds $df) 'DF: text,pic,blank per correl'
$dfPic = FirstOp $df 'DF'
Assert-Equal 'X\DF\A.png' $dfPic.Path 'DF: snap path = snap\DF\<correl>.png'
Assert-True $dfPic.Required 'DF: picture required'

# A DF snap captured under the timestamped mapping id satisfies a workbook
# whose Soushin-data sheet still contains the plain correl id.
$aliasRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('evidence_alias_' + [guid]::NewGuid().ToString('N'))
try {
    $aliasDfDir = Join-Path $aliasRoot 'DF'
    New-Item -ItemType Directory -Path $aliasDfDir -Force | Out-Null
    $aliasPng = Join-Path $aliasDfDir 'JIGPU86S.260729.10515511.png'
    Set-Content -LiteralPath $aliasPng -Value 'fake' -Encoding ASCII
    $aliasPlan = Build-DfEvidencePlan -SnapRoot $aliasRoot -CorrelOrder @('JIGPU86S')
    $aliasPic = FirstOp $aliasPlan 'DF'
    Assert-Equal $aliasPng $aliasPic.Path 'DF: plain correl resolves timestamped snap filename'
} finally {
    Remove-Item -LiteralPath $aliasRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# -- GIFT plan (spec 8) --
$gift = Build-GiftEvidencePlan -SnapRoot 'X' -JobName 'J' -CorrelOrder @('A')
$first = $gift[0]
Assert-Equal 'picture' $first.Kind 'GIFT: first op is picture'
Assert-Equal 'excel'   $first.Folder 'GIFT: first picture is excel snap'
Assert-Equal 'X\excel\J.png' $first.Path 'GIFT: excel snap named by JOB_NAME'

# HM before MQ before Jenkins
$giftFolders = Folders $gift
$idxHm   = [array]::IndexOf($giftFolders, 'GIFT_HM')
$idxMq   = [array]::IndexOf($giftFolders, 'GIFT_MQ')
$idxJk   = [array]::IndexOf($giftFolders, 'GIFT_Jenkins')
$idxNo   = [array]::IndexOf($giftFolders, 'GIFT_noGfixfile')
Assert-True ($idxHm -lt $idxMq) 'GIFT: HM before MQ'
Assert-True ($idxMq -lt $idxJk) 'GIFT: MQ before Jenkins (Jenkins is its own trailing section)'
Assert-True ($idxJk -lt $idxNo) 'GIFT: Jenkins before NoGfix'

# NoGfix has a header and its picture is OPTIONAL
$hasNoGfixHeader = @($gift | Where-Object { $_.Kind -eq 'header' -and $_.LabelKey -eq 'GiftNoGfixHeader' }).Count
Assert-Equal 1 $hasNoGfixHeader 'GIFT: NoGfix header present'
$noPic = FirstOp $gift 'GIFT_noGfixfile'
Assert-True (-not $noPic.Required) 'GIFT: NoGfix picture is optional'

# -- GFIX plan (spec 9) --
$gfix = Build-GfixEvidencePlan -SnapRoot 'X' -JobName 'J' -CorrelOrder @('A') -ToCode 'IDS'
Assert-Equal 'excel' $gfix[0].Folder 'GFIX: starts with excel snap'
$logOps = @($gfix | Where-Object { $_.Kind -eq 'log' })
Assert-Equal 1 $logOps.Count 'GFIX: one log op per correl'
Assert-Equal 'IDS' $logOps[0].ToCode 'GFIX: log op carries TO_code'
Assert-True $logOps[0].Required 'GFIX: log required'
$gfixMixed = Build-GfixEvidencePlan -SnapRoot 'X' -JobName 'J' -CorrelOrder @('JIDSM48S','JIGPM48S','JIGPMO5S') -ToCode 'IDS' -CorrelToCode @{ JIDSM48S='IDS'; JIGPM48S='IGP'; JIGPMO5S='IGP' }
$mixedLogs = @($gfixMixed | Where-Object { $_.Kind -eq 'log' })
Assert-Equal 'IDS|IGP|IGP' (@($mixedLogs | ForEach-Object { $_.ToCode }) -join '|') 'GFIX: log op uses per-correl TO_code map'
# SS_CODE override: explicit per-correl map wins; correls absent from the map
# carry '' so GfixLog infers SS from Correl_ID_S downstream.
Assert-Equal '' $logOps[0].SsCode 'GFIX: log op SS_CODE empty by default (infer downstream)'
$gfixSs = Build-GfixEvidencePlan -SnapRoot 'X' -JobName 'J' -CorrelOrder @('JIDSF48S','JIGPLO5S') -ToCode 'IDS' -CorrelToSs @{ JIDSF48S='F' }
$ssLogs = @($gfixSs | Where-Object { $_.Kind -eq 'log' })
Assert-Equal 'F' $ssLogs[0].SsCode 'GFIX: log op carries per-correl SS_CODE override'
Assert-Equal '' $ssLogs[1].SsCode 'GFIX: log op SS_CODE empty when correl not in map'
# bold GFIX-log header must come immediately before the log op
$kindsArr = @($gfix | ForEach-Object { $_.Kind })
$logIdx = [array]::IndexOf($kindsArr, 'log')
Assert-Equal 'header' $kindsArr[$logIdx - 1] 'GFIX: header precedes log'
$gfixFolders = Folders $gfix
$idxGHm = [array]::IndexOf($gfixFolders, 'GFIX_HM')
$idxGJk = [array]::IndexOf($gfixFolders, 'GFIX_Jenkins')
Assert-True ($idxGHm -lt $idxGJk) 'GFIX: HM/log block before Jenkins block'

# -- Format-EvidencePlan smoke --
$lines = Format-EvidencePlan $df
Assert-True ($lines.Count -ge 4) 'Format-EvidencePlan returns rendered lines'

# -- Missing-file detection --
$miss = Get-PlanMissingFiles $df   # snap root 'X' does not exist -> both required pics missing
Assert-Equal 2 $miss.Count 'Get-PlanMissingFiles flags required missing pictures'

exit (Complete-Tests)
