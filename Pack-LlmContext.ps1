# Pack-LlmContext.ps1
# ---------------------------------------------------------
# 目的：VerifyToolsHO の全コンテキストを収集し、LLM用にクリップボードにコピーする
# ---------------------------------------------------------

# CLAUDE.md に定義されている主要なファイルマップに基づく
$TargetFiles = @(
    "VerifyTool.ps1", 
    "VerifyConfig.psd1", 
    "ExcelHelpers.ps1", 
    "Mark.ps1", 
    "Probe-Shapes.ps1", 
    "ReplaceEvidence.ps1",
    "Clone.ps1",
    "Validate.ps1",
    "JenkinsSnap.ps1"
)

$DocFiles = @("CLAUDE.md", "CHANGELOG.md", "HANDOFF.md")

$output = "You are working in VerifyTool ? PowerShell automation for Misaki's GIFT->GFIX migration evidence collection.`n"
$output += "Here is the complete current project context for continuing development.`n`n"

# 1. ドキュメントの集約
foreach ($doc in $DocFiles) {
    if (Test-Path $doc) {
        $output += "=== Document: $doc ===`n"
        $output += (Get-Content $doc -Raw -Encoding UTF8) + "`n`n"
    }
}

# 2. 自動変更履歴（存在する場合）
if (Test-Path "auto_changelog.md") {
    $output += "=== Document: auto_changelog.md ===`n"
    $output += (Get-Content "auto_changelog.md" -Raw -Encoding UTF8) + "`n`n"
}

# 3. ソースコードの集約
$output += "=== Source Code ===`n"
foreach ($file in $TargetFiles) {
    if (Test-Path $file) {
        $output += "--- File: $file ---`n"
        $output += "``````powershell`n"
        $output += (Get-Content $file -Raw -Encoding UTF8)
        $output += "`n```````n`n"
    }
}

# クリップボードへコピー
Set-Clipboard -Value $output
Write-Host "[SUCCESS] Project context packed into clipboard. Paste it into your LLM UI." -ForegroundColor Green