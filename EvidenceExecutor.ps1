# ============================================================
#  EvidenceExecutor.ps1
#
#  Walks an EvidencePlan (built by the pure, unit-tested EvidencePlan.ps1)
#  and performs the actual Excel COM inserts. The "what/where/order" logic
#  lives in the plan; this file only does the mechanical insertion, so the
#  risky ordering rules stay testable without COM.
#
#  Caller must have dot-sourced (in this order):
#     Common.ps1 (optional), ExcelHelpers.ps1, ProjectLabels.ps1, GfixLog.ps1
#  Dot-source only (no param() block).
#
#  Row advance: pictures advance to just past the image bottom (blank rows
#  are explicit 'blank' ops in the plan, so spacing is data-driven). A
#  missing picture advances one row as a placeholder so the layout below
#  does not collapse onto it.
# ============================================================

function Resolve-PlanLabel {
    param([hashtable]$Labels, [string]$LabelKey, [string]$GiftNoGfixOverride)
    if ($LabelKey -eq 'GiftNoGfixHeader' -and -not [string]::IsNullOrWhiteSpace($GiftNoGfixOverride)) {
        return $GiftNoGfixOverride
    }
    if ($null -ne $Labels -and $Labels.ContainsKey($LabelKey)) { return [string]$Labels[$LabelKey] }
    return $LabelKey
}

function Write-CellText {
    param($Ws, [int]$Row, [int]$Col, [string]$Text, [bool]$Bold)
    $cell = $Ws.Cells.Item($Row, $Col)
    $cell.Value2 = $Text
    try {
        $cell.Font.Bold = $Bold
        $cell.Font.ColorIndex = 1
        $cell.Interior.ColorIndex = -4142   # xlColorIndexNone
    } catch {}
}

# Returns a result object:
#   Inserted, LogMatched, MissingRequired[], MissingOptional[], Warnings[], EndRow
# The caller decides whether MissingRequired / MissingOptional block the
# completion bit.
function Invoke-EvidencePlan {
    param(
        $Worksheet,
        [object[]]$Plan,
        [hashtable]$Labels,
        [string]$LogDir = '',
        [int]$StartRow = 3,
        [int]$Col = 2,
        [string]$GiftNoGfixLabelOverride = ''
    )
    $anchor = $StartRow
    $inserted = 0
    $logMatched = 0
    $missingRequired = [System.Collections.Generic.List[object]]::new()
    $missingOptional = [System.Collections.Generic.List[object]]::new()
    $warnings        = [System.Collections.Generic.List[string]]::new()

    foreach ($op in @($Plan)) {
        switch ($op.Kind) {
            'text' {
                Write-CellText $Worksheet $anchor $Col ([string]$op.Text) $false
                $anchor++
            }
            'textbold' {
                Write-CellText $Worksheet $anchor $Col ([string]$op.Text) $true
                $anchor++
            }
            'header' {
                $isBold = ([string]$op.LabelKey -eq 'GfixLogLabel')
                $text   = Resolve-PlanLabel $Labels ([string]$op.LabelKey) $GiftNoGfixLabelOverride
                Write-CellText $Worksheet $anchor $Col $text $isBold
                $anchor++
            }
            'blank' {
                $anchor += [int]$op.Count
            }
            'picture' {
                if (Test-Path -LiteralPath $op.Path) {
                    $insertRow = $anchor
                    $pic = Insert-PictureSendToBack $Worksheet $anchor $Col ([string]$op.Path)
                    Set-ShapeMetadata $pic ([string]$op.Folder) ([string]$op.Name)
                    $anchor = Get-NextAnchorRow $Worksheet $pic 0
                    $inserted++
                    Write-Host ("    [OK]  B{0}  {1}\{2}.png" -f $insertRow, $op.Folder, $op.Name) -ForegroundColor DarkGreen
                } else {
                    $rec = [pscustomobject]@{ CorrelIdS = [string]$op.CorrelIdS; Folder = [string]$op.Folder; Path = [string]$op.Path }
                    if ($op.Required) {
                        $missingRequired.Add($rec)
                        Write-Host ("    [MISS-REQ] {0}\{1}.png" -f $op.Folder, $op.Name) -ForegroundColor Red
                    } else {
                        $missingOptional.Add($rec)
                        Write-Host ("    [INFO] optional {0}\{1}.png absent" -f $op.Folder, $op.Name) -ForegroundColor DarkGray
                    }
                    $anchor++
                }
            }
            'log' {
                $res = Find-GfixLogForCorrel -LogDir $LogDir -ToCode ([string]$op.ToCode) -CorrelIdS ([string]$op.CorrelIdS) -SsCode ([string]$op.SsCode)
                if (-not [string]::IsNullOrWhiteSpace([string]$res.Warning)) { $warnings.Add([string]$res.Warning) }
                if (-not [string]::IsNullOrWhiteSpace([string]$res.Error)) {
                    $missingRequired.Add([pscustomobject]@{ CorrelIdS = [string]$op.CorrelIdS; Folder = 'GFIX_log'; Path = [string]$res.Error })
                    Write-Host ("    [MISS-REQ] gfix log for {0}: {1}" -f $op.CorrelIdS, $res.Error) -ForegroundColor Red
                    $anchor++
                } else {
                    $anchor = Write-LogLines $Worksheet $anchor $Col $res.Chosen.Lines
                    $logMatched++
                    Write-Host ("    [OK]  gfix log for {0}: {1} line(s)" -f $op.CorrelIdS, $res.Chosen.Lines.Count) -ForegroundColor DarkGreen
                }
            }
        }
    }

    return [pscustomobject]@{
        Inserted        = $inserted
        LogMatched      = $logMatched
        MissingRequired = $missingRequired.ToArray()
        MissingOptional = $missingOptional.ToArray()
        Warnings        = $warnings.ToArray()
        EndRow          = $anchor
    }
}
