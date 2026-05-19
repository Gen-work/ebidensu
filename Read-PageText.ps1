# ============================================================
# Read-PageText.ps1
#   フォアグラウンド Edge ページの可視テキストを clipboard 経由取得。
#   frameset ページは呼び出し側で事前に frame_main を click すること。
# ============================================================
param(
    [int]$SelectWaitMs = 400,
    [int]$CopyWaitMs   = 400
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms

[System.Windows.Forms.Clipboard]::Clear()
Start-Sleep -Milliseconds 100

[System.Windows.Forms.SendKeys]::SendWait('^a')
Start-Sleep -Milliseconds $SelectWaitMs

[System.Windows.Forms.SendKeys]::SendWait('^c')
Start-Sleep -Milliseconds $CopyWaitMs

# クリックして選択解除（次操作の妨げにならないように）
[System.Windows.Forms.SendKeys]::SendWait('{ESC}')

[System.Windows.Forms.Clipboard]::GetText()