# Apply-LlmPatch.ps1
# ---------------------------------------------------------
# 目的：クリップボード内のXMLパッチ情報を解析し、ローカルファイルに自動適用する
# ---------------------------------------------------------

$clipboardData = Get-Clipboard -Raw

# パッチブロックを正規表現で抽出
$pattern = '(?s)<patch file="(.*?)">\s*<search>\s*(.*?)\s*</search>\s*<replace>\s*(.*?)\s*</replace>\s*</patch>'
$matches = [regex]::Matches($clipboardData, $pattern)

if ($matches.Count -eq 0) {
    Write-Warning "クリップボードに有効なパッチブロックが見つかりません。"
    exit
}

foreach ($match in $matches) {
    $fileName = $match.Groups[1].Value.Trim()
    $searchString = $match.Groups[2].Value
    $replaceString = $match.Groups[3].Value

    if (-not (Test-Path $fileName)) {
        Write-Warning "ファイルが見つかりません: $fileName"
        continue
    }

    $content = Get-Content $fileName -Raw -Encoding UTF8

    # 厳密な文字列置換（正規表現のエスケープ処理を含む）
    if ($content.Contains($searchString)) {
        $newContent = $content.Replace($searchString, $replaceString)
        Set-Content -Path $fileName -Value $newContent -NoNewline -Encoding UTF8
        Write-Host "[成功] $fileName を更新しました。" -ForegroundColor Green
    } else {
        Write-Host "[失敗] $fileName の中に <search> ブロックと完全に一致するテキストが見つかりません。" -ForegroundColor Red
        Write-Host "--- 検索対象 ---`n$searchString`n----------------" -ForegroundColor Gray
    }
}