# ============================================================
#  EvidencePlan.ps1
#
#  PURE evidence-layout planners -- NO Excel COM. Each builder turns a
#  set of mapping rows into an ordered list of insert operations that
#  encodes the review standard (spec sections 7/8/9). The Excel executor
#  (separate, COM-bound) just walks the plan and inserts; this keeps the
#  fragile "what goes where, in which order" logic unit-testable
#  (Tests\Test-EvidencePlan.ps1) and free of COM.
#
#  Design note (per spec priority): the old ReplaceEvidence looped
#  folder-major (all HM, then all MQ, ...). The review standard is
#  correl-major (one Correl_ID_S = one group). These builders are
#  correl-major. Jenkins / NoGfix are deliberately their own trailing
#  sections, not interleaved per-correl.
#
#  This file is intentionally ASCII-only. The Japanese sheet names and
#  labels are resolved later by the executor via ProjectLabels.ps1, so
#  the plan carries ASCII label KEYS only (e.g. 'GfixLogLabel') and this
#  file stays BOM-safe.
#
#  Plan op shape:
#     @{ Kind='text'|'textbold'|'header'|'picture'|'log'|'blank'
#        Col=2; Text=...; LabelKey=...; Folder=...; Name=...; Path=...
#        Required=$true; Count=1; CorrelIdS=...; JobName=...; ToCode=...
#        Section=... }
# ============================================================

function Get-SnapPath {
    param([string]$SnapRoot, [string]$Folder, [string]$Name)
    return (Join-Path (Join-Path $SnapRoot $Folder) ('{0}.png' -f $Name))
}

# Drop blanks, '#VALUE!' / '#REF!' style errors, and anything that does not
# look like a correl id. Order is preserved.
function Select-ValidCorrelIds {
    param([string[]]$Raw)
    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($item in @($Raw)) {
        if ($null -eq $item) { continue }
        $v = ([string]$item).Trim()
        if ([string]::IsNullOrWhiteSpace($v)) { continue }
        if ($v.StartsWith('#')) { continue }                 # #VALUE! / #REF! / #N/A
        if ($v -notmatch '^[A-Za-z0-9_]{4,}$') { continue }  # not a correl-looking token
        $out.Add($v)
    }
    return $out.ToArray()
}

# ---- op constructors (internal) ----
function New-TextOp     { param($Text,$CorrelIdS='',$JobName='',$Section='') @{ Kind='text';     Col=2; Text=$Text; CorrelIdS=$CorrelIdS; JobName=$JobName; Section=$Section } }
function New-TextBoldOp { param($Text,$CorrelIdS='',$Section='')             @{ Kind='textbold'; Col=2; Text=$Text; CorrelIdS=$CorrelIdS; Section=$Section } }
function New-HeaderOp   { param($LabelKey,$Section='')                       @{ Kind='header';   Col=2; LabelKey=$LabelKey; Section=$Section } }
function New-BlankOp    { param([int]$Count=1)                              @{ Kind='blank';    Count=$Count } }
function New-PicOp {
    param($SnapRoot,$Folder,$Name,[bool]$Required=$true,$CorrelIdS='',$JobName='',$Section='')
    @{ Kind='picture'; Col=2; Folder=$Folder; Name=$Name
       Path=(Get-SnapPath $SnapRoot $Folder $Name); Required=$Required
       CorrelIdS=$CorrelIdS; JobName=$JobName; Section=$Section }
}
function New-LogOp {
    param($CorrelIdS,$ToCode,[bool]$Required=$true)
    @{ Kind='log'; Col=2; CorrelIdS=$CorrelIdS; ToCode=$ToCode; Required=$Required; Section='hm_log' }
}

# ---- default blank-row spacing (config can override) ----
function Get-EvidencePlanSpacing {
    return @{
        AfterGiftExcel = 1   # spec 8.3: 1 blank after excel.png
        AfterGfixExcel = 2   # spec 9.3: 2 blanks after excel.png
        AfterHm      = 1   # spec 8.4 / 9.5
        AfterMq      = 2   # spec 8.4
        AfterLog     = 2   # spec 9.5
        SectionGap   = 2   # spec 8.6 / 9.7 (before Jenkins block)
        AfterJenkins = 2   # spec 8.8 (before NoGfix block)
        AfterPic1    = 1   # generic single blank (Jenkins/NoGfix/DF)
    }
}

# ---- DF plan (spec 7) -- sheet 'GIFT de-ta vs GFIX de-ta' ----
# -CorrelOrder MUST already be the order taken from the 'Soushin data'
# sheet column A (use Select-ValidCorrelIds first). Per correl:
# text, DF png, blank.
function Build-DfEvidencePlan {
    param([string]$SnapRoot, [string[]]$CorrelOrder, [hashtable]$Spacing = $null)
    if ($null -eq $Spacing) { $Spacing = Get-EvidencePlanSpacing }
    $plan = [System.Collections.Generic.List[object]]::new()
    foreach ($cid in @($CorrelOrder)) {
        $plan.Add((New-TextOp -Text $cid -CorrelIdS $cid -Section 'df'))
        $plan.Add((New-PicOp -SnapRoot $SnapRoot -Folder 'DF' -Name $cid -Required $true -CorrelIdS $cid -Section 'df'))
        $plan.Add((New-BlankOp -Count $Spacing.AfterPic1))
    }
    return $plan.ToArray()
}

# ---- GIFT plan (spec 8) -- sheet 'GIFT jushin kekka' ----
# excel.png -> HM/MQ per correl -> Jenkins per correl -> NoGfix per correl.
function Build-GiftEvidencePlan {
    param(
        [string]$SnapRoot,
        [string]$JobName,
        [string[]]$CorrelOrder,
        [hashtable]$Spacing = $null
    )
    if ($null -eq $Spacing) { $Spacing = Get-EvidencePlanSpacing }
    $plan = [System.Collections.Generic.List[object]]::new()

    # 1) Excel snap (named by JOB_NAME), then 1 blank (spec 8.3).
    $plan.Add((New-PicOp -SnapRoot $SnapRoot -Folder 'excel' -Name $JobName -Required $true -JobName $JobName -Section 'excel'))
    $plan.Add((New-BlankOp -Count $Spacing.AfterGiftExcel))

    # 2) HM + MQ, correl-major.
    foreach ($cid in @($CorrelOrder)) {
        $plan.Add((New-TextOp -Text $cid -CorrelIdS $cid -Section 'hm_mq'))
        $plan.Add((New-PicOp -SnapRoot $SnapRoot -Folder 'GIFT_HM' -Name $cid -Required $true -CorrelIdS $cid -Section 'hm_mq'))
        $plan.Add((New-BlankOp -Count $Spacing.AfterHm))
        $plan.Add((New-PicOp -SnapRoot $SnapRoot -Folder 'GIFT_MQ' -Name $cid -Required $true -CorrelIdS $cid -Section 'hm_mq'))
        $plan.Add((New-BlankOp -Count $Spacing.AfterMq))
    }
    $plan.Add((New-BlankOp -Count $Spacing.SectionGap))

    # 3) Jenkins, correl-major.
    foreach ($cid in @($CorrelOrder)) {
        $plan.Add((New-TextOp -Text $cid -CorrelIdS $cid -Section 'jenkins'))
        $plan.Add((New-PicOp -SnapRoot $SnapRoot -Folder 'GIFT_Jenkins' -Name $cid -Required $true -CorrelIdS $cid -Section 'jenkins'))
        $plan.Add((New-BlankOp -Count $Spacing.AfterPic1))
    }
    $plan.Add((New-BlankOp -Count $Spacing.AfterJenkins))

    # 4) NoGfix section. Pictures here are OPTIONAL (GFIX side may have a
    #    file, so there is legitimately no "no-file" snap). Strictness is
    #    decided by the executor, not the plan.
    $plan.Add((New-HeaderOp -LabelKey 'GiftNoGfixHeader' -Section 'nogfix'))
    foreach ($cid in @($CorrelOrder)) {
        $plan.Add((New-TextOp -Text $cid -CorrelIdS $cid -Section 'nogfix'))
        $plan.Add((New-PicOp -SnapRoot $SnapRoot -Folder 'GIFT_noGfixfile' -Name $cid -Required $false -CorrelIdS $cid -Section 'nogfix'))
        $plan.Add((New-BlankOp -Count $Spacing.AfterPic1))
    }
    return $plan.ToArray()
}

# ---- GFIX plan (spec 9) -- sheet 'GFIX jushin kekka' ----
# excel.png -> (HM + bold-log-header + whole log) per correl -> Jenkins per correl.
function Build-GfixEvidencePlan {
    param(
        [string]$SnapRoot,
        [string]$JobName,
        [string[]]$CorrelOrder,
        [string]$ToCode,
        [hashtable]$Spacing = $null
    )
    if ($null -eq $Spacing) { $Spacing = Get-EvidencePlanSpacing }
    $plan = [System.Collections.Generic.List[object]]::new()

    $plan.Add((New-PicOp -SnapRoot $SnapRoot -Folder 'excel' -Name $JobName -Required $true -JobName $JobName -Section 'excel'))
    $plan.Add((New-BlankOp -Count $Spacing.AfterGfixExcel))

    foreach ($cid in @($CorrelOrder)) {
        $plan.Add((New-TextOp -Text $cid -CorrelIdS $cid -Section 'hm_log'))
        $plan.Add((New-PicOp -SnapRoot $SnapRoot -Folder 'GFIX_HM' -Name $cid -Required $true -CorrelIdS $cid -Section 'hm_log'))
        $plan.Add((New-BlankOp -Count $Spacing.AfterHm))
        $plan.Add((New-HeaderOp -LabelKey 'GfixLogLabel' -Section 'hm_log'))   # bold, resolved by executor
        $plan.Add((New-LogOp -CorrelIdS $cid -ToCode $ToCode -Required $true))
        $plan.Add((New-BlankOp -Count $Spacing.AfterLog))
    }
    $plan.Add((New-BlankOp -Count $Spacing.SectionGap))

    foreach ($cid in @($CorrelOrder)) {
        $plan.Add((New-TextOp -Text $cid -CorrelIdS $cid -Section 'jenkins'))
        $plan.Add((New-PicOp -SnapRoot $SnapRoot -Folder 'GFIX_Jenkins' -Name $cid -Required $true -CorrelIdS $cid -Section 'jenkins'))
        $plan.Add((New-BlankOp -Count $Spacing.AfterPic1))
    }
    return $plan.ToArray()
}

# ---- execution-support helpers (FS-touching, used by executor + dry-run) ----

# Returns the list of Required picture ops whose file is missing.
function Get-PlanMissingFiles {
    param([object[]]$Plan)
    $missing = [System.Collections.Generic.List[object]]::new()
    foreach ($op in @($Plan)) {
        if ($op.Kind -eq 'picture' -and $op.Required) {
            if (-not (Test-Path -LiteralPath $op.Path)) {
                $missing.Add([pscustomobject]@{
                    CorrelIdS = $op.CorrelIdS; Folder = $op.Folder; Path = $op.Path
                })
            }
        }
    }
    return $missing.ToArray()
}

# Human-readable dry-run rendering of a plan (for -DryRun and tests).
function Format-EvidencePlan {
    param([object[]]$Plan)
    $lines = [System.Collections.Generic.List[string]]::new()
    $row = 3   # plans always start at row 3 (row 1-2 are the kept title rows)
    foreach ($op in @($Plan)) {
        switch ($op.Kind) {
            'text'     { $lines.Add(('B{0}  TEXT      {1}' -f $row, $op.Text));       $row++ }
            'textbold' { $lines.Add(('B{0}  TEXT*bold {1}' -f $row, $op.Text));       $row++ }
            'header'   { $lines.Add(('B{0}  HEADER    [{1}]' -f $row, $op.LabelKey)); $row++ }
            'picture'  {
                $req = if ($op.Required) { 'req' } else { 'opt' }
                $lines.Add(('B{0}  PIC  ({1}) {2}\{3}.png' -f $row, $req, $op.Folder, $op.Name))
                $row++   # logical advance; real executor advances by image height
            }
            'log'      { $lines.Add(('B{0}  LOG       whole file for {1} (TO={2})' -f $row, $op.CorrelIdS, $op.ToCode)); $row++ }
            'blank'    { $row += [int]$op.Count }
        }
    }
    return $lines.ToArray()
}
