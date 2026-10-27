#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
$root = Split-Path $here -Parent
. (Join-Path $root 'ProjectLabels.ps1')
. (Join-Path $root 'PomBook.ps1')

Reset-Tests 'PomBook'

$L = Get-ProjectLabels
$P = Get-PomBookLabels

# ---- the delivered sample's sheet names (built from code points) ----
Assert-Equal ([int][char]$P.GfixSendBefore[-4]) 0x5909 'before label starts with henkou'
Assert-Equal ([int][char]$P.GfixSendBefore[-2]) 0x524D 'before label ends with mae'
Assert-Equal ([int][char]$P.GfixRecvAfter[-2]) 0x5F8C 'after label ends with go'
Assert-Equal ($P.GfixSendBefore.Substring(0, $L['SheetGfixSend'].Length + 1)) ($L['SheetGfixSend'] + '(') 'ASCII paren after the base name'
Assert-Equal ($P.PomCompare.Length) 14 'GFIX de-ta _ henkou mae go hikaku = 14 chars'

# ---- an old-layout book (YJDSWM39 sample, 10/27) ----
$old = @(
    $L['SheetSoshinData'], $L['SheetSoshinArrow'], $L['SheetGiftSend'], $L['SheetGfixSend'],
    $L['SheetJushinArrow'], $L['SheetGiftRecv'], $L['SheetGfixRecv'], $L['SheetDfCompare'])
$plan = Get-PomBookPlan -SheetNames $old
Assert-Equal $plan.status 'convert' 'old layout converts'
Assert-Equal (@($plan.actions).Count) 5 'two renames, two copies, one new compare'
Assert-Equal $plan.actions[0].op 'rename' 'first: rename send'
Assert-Equal $plan.actions[0].name $P.GfixSendBefore 'send -> (before)'
Assert-Equal $plan.actions[1].op 'copyAfter' 'then copy it'
Assert-Equal $plan.actions[1].sheet $P.GfixSendBefore 'the copy is taken from the renamed sheet'
Assert-Equal $plan.actions[1].name $P.GfixSendAfter 'copy is (after)'
Assert-Equal $plan.actions[4].op 'newCompare' 'new compare last'
Assert-Equal $plan.actions[4].template $L['SheetDfCompare'] 'new compare copies the look of GIFT-vs-GFIX'

$order = Get-PomBookResultOrder -SheetNames $old -Actions $plan.actions
$want = @(
    $L['SheetSoshinData'], $L['SheetSoshinArrow'], $L['SheetGiftSend'],
    $P.GfixSendBefore, $P.GfixSendAfter,
    $L['SheetJushinArrow'], $L['SheetGiftRecv'],
    $P.GfixRecvBefore, $P.GfixRecvAfter,
    $L['SheetDfCompare'], $P.PomCompare)
Assert-Equal ($order -join '|') ($want -join '|') 'GIFT sheets and GIFT-vs-GFIX stay in place; after right behind before'

# ---- already POM (the JJPCRR12 samples) ----
$pom = @($L['SheetSoshinData'], $P.GfixSendBefore, $P.GfixSendAfter, $P.GfixRecvBefore, $P.GfixRecvAfter, $P.PomCompare)
Assert-Equal (Get-PomBookPlan -SheetNames $pom).status 'already' 'POM layout is left alone'
Assert-Equal (@((Get-PomBookPlan -SheetNames $pom).actions).Count) 0 'no actions on a POM book'

# ---- neither layout ----
$odd = @($L['SheetSoshinData'], $L['SheetGfixSend'])
$r = Get-PomBookPlan -SheetNames $odd
Assert-Equal $r.status 'unknown' 'no GFIX recv sheet -> unknown'
Assert-Equal ($r.missing -join ',') $L['SheetGfixRecv'] 'missing names the recv sheet'

# ---- no GIFT-vs-GFIX sheet: the compare sheet is created blank ----
$noCmp = @($L['SheetGfixSend'], $L['SheetGfixRecv'])
$r = Get-PomBookPlan -SheetNames $noCmp
Assert-Equal $r.actions[4].template '' 'no template sheet -> blank compare'

# ---- a compare sheet already there is not created twice ----
$withCmp = @($L['SheetGfixSend'], $L['SheetGfixRecv'], $P.PomCompare)
Assert-Equal (@((Get-PomBookPlan -SheetNames $withCmp).actions).Count) 4 'existing compare is kept'

exit (Complete-Tests)
