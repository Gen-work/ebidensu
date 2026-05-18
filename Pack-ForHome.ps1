# Pack-ForHome.ps1
# ---------------------------------------------------------
# 目的：業務依存のコードを除外し、汎用ツール（持ち出し可能）のみを
#       LLMに貼り付けるためのMarkdownフォーマットでクリップボードにコピーする
# ---------------------------------------------------------

# 同期対象となる「汎用ツール」のリスト（業務ロジックを含まないもの）
$TargetFiles = @(
    "VerifyTool.ps1",
    "Mark.ps1",
    "ExcelHelpers.ps1",
    "Probe-Shapes.ps1",
    "ReplaceEvidence.ps1",
    "Common.ps1"
)

# 絶対に除外すべきファイル（設定、データ）の警告用
$ExcludedFiles = @("VerifyConfig.psd1", "verify_session.json")

$output = "本日（$(Get-Date -Format 'yyyy-MM-dd')）のツールチェーンの進捗です。業務データや特定の設定は意図的に除外しています。`n`n"

foreach ($file in $TargetFiles) {
    if (Test-Path $file) {
        $output += "### File: $file`n"
        $output += "```
```powershell`n"
        
        # ファイル内容を読み込む
        $content = Get-Content $file -Raw
        
# 簡易的なサニタイズ（念のため、ハードコードされたパスなどをマスキング）
        $content = $content -replace '\\\\fs-f3170-1\\[^\s"'']+', '[REDACTED_FS_PATH]'
        $content = $content -replace 'KJRV[A-Z0-9]+', '[REDACTED_CORREL_ID]'

        $output += $content
        $output += "`n```````n`n"
    } else {
        Write-Warning "見つかりません (NotFound): $file"
    }
}

$output += "`n`n-- これ以降の指示があるまで、コードを学習・保持するだけで出力は不要です --"

Set-Clipboard -Value $output
Write-Host "[完了] 自宅同期用のコードをクリップボードにコピーしました。AIのWeb UIに貼り付けてください。" -ForegroundColor Green
Write-Host "注意：VerifyConfig.psd1 等の業務ファイルは除外されています。" -ForegroundColor Yellow