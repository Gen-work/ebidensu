# Pack-LlmContext.ps1
# ---------------------------------------------------------
# 目的：Gitを利用してプロジェクトの全量ファイルパスと、
#       主要なソースコード（.ps1, .json, .md等）のコンテキストを収集し
#       LLM用にクリップボードへコピーする
# ---------------------------------------------------------

$header = "You are working in VerifyTool - PowerShell automation for Misaki's GIFT->GFIX migration evidence collection.`n"
$header += "Here is the complete current project context for continuing development.`n`n"

# 1. Gitコマンドを利用してファイル一覧を取得（追跡済み + 未追跡ファイル）
$tracked = git ls-files
$untracked = git ls-files --others --exclude-standard
$allFiles = $tracked + $untracked | Sort-Object -Unique

# 内容を読み込む対象の拡張子
$contentExtensions = @('.ps1', '.psd1', '.json', '.md')
# 意図的に除外する不要なドキュメント
$excludeFiles = @('handoff.md', 'auto_changelog.md')

$output = [System.Text.StringBuilder]::new()
[void]$output.AppendLine($header)

# 2. 全量ファイルツリーの出力（相対パスのみ）
[void]$output.AppendLine("=== Project File Structure ===")
foreach ($file in $allFiles) {
    [void]$output.AppendLine($file)
}
[void]$output.AppendLine("==============================`n")

# 3. 対象ソースコードの内容を集約
[void]$output.AppendLine("=== Source Code & Context ===")
foreach ($file in $allFiles) {
    $fileName = Split-Path $file -Leaf
    
    # 廃止されたドキュメントをスキップ
    if ($excludeFiles -contains $fileName.ToLower()) {
        continue
    }

    # 拡張子の判定
    $shouldReadContent = $false
    foreach ($ext in $contentExtensions) {
        if ($file.ToLower().EndsWith($ext)) {
            $shouldReadContent = $true
            break
        }
    }

    if ($shouldReadContent -and (Test-Path $file)) {
        [void]$output.AppendLine("--- File: $file ---")
        
        # LLMが見やすいように言語タグを動的に設定
        $langTag = if ($file.ToLower().EndsWith('.md')) { "markdown" } 
                   elseif ($file.ToLower().EndsWith('.json')) { "json" } 
                   else { "powershell" }

        [void]$output.AppendLine("``````$langTag")
        
        try {
            # エンコーディングの罠（Encoding Trap）を完全に回避するため、
            # PowerShellのGet-Contentではなく.NETクラスを直接呼び出す
            $fullPath = (Resolve-Path $file).ProviderPath
            $content = [System.IO.File]::ReadAllText($fullPath, [System.Text.Encoding]::UTF8)
            [void]$output.AppendLine($content.TrimEnd())
        } catch {
            [void]$output.AppendLine("# Error: Failed to read file content (Encoding or Permission issue)")
        }
        
        [void]$output.AppendLine("```
````n")
    }
}

# クリップボードへ送る
$output.ToString() | Set-Clipboard
Write-Host "LLM context packed and copied to clipboard successfully!" -ForegroundColor Cyan