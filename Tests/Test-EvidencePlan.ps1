#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
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
$valid = Select-ValidCorrelIds @('JIGPF48S', '#VALUE!', '', '  ', 'JIDSF48S')
Assert-Equal 'JIGPF48S|JIDSF48S' ($valid -join '|') 'drops #VALUE!/blank, keeps order'

# -- DF plan (spec 7) --
$df = Build-DfEvidencePlan -SnapRoot 'X' -CorrelOrder @('A','B')
Assert-Equal 'text,picture,blank,text,picture,blank' (Kinds $df) 'DF: text,pic,blank per correl'
$dfPic = FirstOp $df 'DF'
Assert-Equal 'X\DF\A.png' $dfPic.Path 'DF: snap path = snap\DF\<correl>.png'
Assert-True $dfPic.Required 'DF: picture required'

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
