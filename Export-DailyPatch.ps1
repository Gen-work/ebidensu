# Export-DailyPatch.ps1
# ---------------------------------------------------------
# 目的：本日のGit差分（パッチ）のみを抽出し、クリップボードにコピーする。
#       LLMに全量コードを食わせるのを防ぐための軽量同期ツール。
# ---------------------------------------------------------

Write-Host "最新のコミットと現在の変更を計算しています..." -ForegroundColor Cyan

# 1. ワークツリーに未コミットの変更があるか確認し、あれば一時的にコミットする
$status = git status --porcelain
if ($status) {
    git add .
    git commit -m "chore: 自宅同期用の一時保存 (WIP)" | Out-Null
    Write-Host "[Info] 未コミットの変更をWIPとして保存しました。" -ForegroundColor Yellow
}

# 2. 直近1回のコミットの完全な差分（パッチ）を取得
# （もし複数回のコミットをまとめたい場合は HEAD~1 を HEAD~3 などに変更）
$patchData = git diff HEAD~1 HEAD

if (-not $patchData) {
    Write-Warning "差分が見つかりません。コードに変更がないようです。"
    exit
}

# 3. LLM用のプロンプトを組み立てる
$output = "以下は本日の作業のGit差分（Patch）です。`n"
$output += "この差分内容を学習し、現在の実装状況を把握してください。コードの全体を出力する必要はありません。`n`n"
$output += "```diff`n"
$output += $patchData
$output += "`n````n"

# クリップボードへコピー
Set-Clipboard -Value $output
Write-Host "[成功] 今日の差分（Patch）をクリップボードにコピーしました！" -ForegroundColor Green
Write-Host "文字数は約 $($output.Length) 文字です。AI Web UIに貼り付けてください。" -ForegroundColor Gray