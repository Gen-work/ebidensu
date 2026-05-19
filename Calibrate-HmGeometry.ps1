# ============================================================
# Calibrate-HmGeometry.ps1
#   HM スナップショットを表示し、4 回クリックして
#   Row1Top / RowHeight / StatusColLeft / StatusColWidth を計算。
#
# 使用例:
#   .\Calibrate-HmGeometry.ps1 -SnapPath "snap\GIFT_HM\JIDSQ48S.png"
#
# 操作:
#   1. Row1 (第1データ行) 内のどこか
#   2. Row2 (第2データ行) 内のどこか (Row1 と縦に違うピクセル)
#   3. 「処理状態」列の LEFT エッジ
#   4. 「処理状態」列の RIGHT エッジ
#   → 結果は clipboard に自動コピー。
# ============================================================
param([Parameter(Mandatory)][string]$SnapPath)
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not (Test-Path -LiteralPath $SnapPath)) { throw "Snap not found: $SnapPath" }

$img = [System.Drawing.Image]::FromFile($SnapPath)

$form = New-Object System.Windows.Forms.Form
$form.Text = "HM Geometry Calibration"
$form.StartPosition = "CenterScreen"
$form.AutoSize = $false
$form.ClientSize = New-Object System.Drawing.Size(($img.Width + 240), ([Math]::Max(420, $img.Height)))

$panel = New-Object System.Windows.Forms.Panel
$panel.AutoScroll = $true
$panel.Location = New-Object System.Drawing.Point(0, 0)
$panel.Size = New-Object System.Drawing.Size($img.Width, $form.ClientSize.Height)
$picBox = New-Object System.Windows.Forms.PictureBox
$picBox.Image = $img
$picBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::AutoSize
$picBox.Location = New-Object System.Drawing.Point(0, 0)
$panel.Controls.Add($picBox)

$script:lblTitle = New-Object System.Windows.Forms.Label
$script:lblTitle.Location = New-Object System.Drawing.Point(($img.Width + 10), 10)
$script:lblTitle.Size = New-Object System.Drawing.Size(220, 80)
$script:lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

$script:lblResult = New-Object System.Windows.Forms.Label
$script:lblResult.Location = New-Object System.Drawing.Point(($img.Width + 10), 100)
$script:lblResult.Size = New-Object System.Drawing.Size(220, 220)
$script:lblResult.Font = New-Object System.Drawing.Font("Consolas", 9)

$script:btnReset = New-Object System.Windows.Forms.Button
$script:btnReset.Text = "Reset"
$script:btnReset.Location = New-Object System.Drawing.Point(($img.Width + 10), 330)
$script:btnReset.Size = New-Object System.Drawing.Size(100, 30)

$script:btnDone = New-Object System.Windows.Forms.Button
$script:btnDone.Text = "Copy and Close"
$script:btnDone.Location = New-Object System.Drawing.Point(($img.Width + 120), 330)
$script:btnDone.Size = New-Object System.Drawing.Size(110, 30)
$script:btnDone.Enabled = $false

$form.Controls.Add($panel)
$form.Controls.Add($script:lblTitle)
$form.Controls.Add($script:lblResult)
$form.Controls.Add($script:btnReset)
$form.Controls.Add($script:btnDone)

$script:clicks   = New-Object System.Collections.ArrayList
$script:prompts  = @(
    "Step 1/4`nClick on Row 1`n(1st data row)",
    "Step 2/4`nClick on Row 2`n(2nd data row)",
    "Step 3/4`nClick LEFT edge`nof Status column",
    "Step 4/4`nClick RIGHT edge`nof Status column",
    "Done.`nClick 'Copy and Close'`nor Reset to retry."
)
$script:resultText = ""
$script:lblTitle.Text = $script:prompts[0]

$refreshUi = {
    $script:lblTitle.Text = $script:prompts[$script:clicks.Count]
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $script:clicks.Count; $i++) {
        $c = $script:clicks[$i]
        [void]$sb.AppendLine(("[{0}] x={1,4} y={2,4}" -f ($i+1), $c.X, $c.Y))
    }
    if ($script:clicks.Count -ge 4) {
        $row1y = $script:clicks[0].Y
        $row2y = $script:clicks[1].Y
        $sL    = $script:clicks[2].X
        $sR    = $script:clicks[3].X
        $rh    = [Math]::Abs($row2y - $row1y)
        $sw    = [Math]::Abs($sR - $sL)
        $top   = [Math]::Min($row1y, $row2y)
        $sLm   = [Math]::Min($sL, $sR)
        if ($rh -eq 0 -or $sw -eq 0) {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("ERROR: same point twice")
        } else {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("Row1Top        = $top")
            [void]$sb.AppendLine("RowHeight      = $rh")
            [void]$sb.AppendLine("StatusColLeft  = $sLm")
            [void]$sb.AppendLine("StatusColWidth = $sw")
            $script:resultText = "Row1Top=$top;RowHeight=$rh;StatusColLeft=$sLm;StatusColWidth=$sw"
            $script:btnDone.Enabled = $true
        }
    }
    $script:lblResult.Text = $sb.ToString()
}

$picBox.Add_MouseClick({
    param($s, $e)
    if ($script:clicks.Count -ge 4) { return }
    [void]$script:clicks.Add(@{X=$e.X; Y=$e.Y})
    & $refreshUi
})
$script:btnReset.Add_Click({
    $script:clicks.Clear()
    $script:btnDone.Enabled = $false
    $script:resultText = ""
    & $refreshUi
})
$script:btnDone.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText($script:resultText)
    [System.Windows.Forms.MessageBox]::Show(
        "Copied to clipboard:`n`n$($script:resultText)",
        "Done") | Out-Null
    $form.Close()
})

[void]$form.ShowDialog()
$img.Dispose()