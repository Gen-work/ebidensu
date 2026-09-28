#Requires -Version 5.1
# Test-Json.ps1 -- kernel/Json.ps1 (P1-35), the one JSON entry point.
#
# The card's completion criteria: a profile JSON holding Japanese reads back
# unchanged; a value five levels deep round-trips without truncation; a
# written file has no BOM and its Japanese is characters, not \uXXXX.
# Everything else here pins the PS 5.1 traps the file exists to hide.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Json.ps1')

Reset-Tests 'Json'

$jp = [string][char]0x8A3C + [char]0x62E0 + [char]0x30D5 + [char]0x30A1   # evidence + 'fa' in katakana
$utf8 = New-Object System.Text.UTF8Encoding($false)

# ---------------------------------------------------------------- ConvertTo-EbiHashtable
$h = ConvertTo-EbiHashtable ([pscustomobject]@{ a = 1; b = [pscustomobject]@{ c = @(1, 2, [pscustomobject]@{ d = 'x' }) } })
Assert-True ($h -is [hashtable]) 'hashtable: a PSCustomObject becomes a hashtable'
Assert-True ($h['b'] -is [hashtable]) 'hashtable: nested objects too'
Assert-True ($h['b']['c'] -is [array]) 'hashtable: an array stays an array'
Assert-Equal 'x' $h['b']['c'][2]['d'] 'hashtable: an object inside an array becomes a hashtable'
$one = ConvertTo-EbiHashtable @(@{ id = 'a' })
Assert-True ($one -is [array] -and $one.Count -eq 1) 'hashtable: a one-element array is not unrolled'
$none = ConvertTo-EbiHashtable @()
Assert-True ($none -is [array] -and $none.Count -eq 0) 'hashtable: an empty array is an empty array'
Assert-Equal 'plain' (ConvertTo-EbiHashtable 'plain') 'hashtable: a string is untouched'
Assert-True ($null -eq (ConvertTo-EbiHashtable $null)) 'hashtable: null stays null'

# ---------------------------------------------------------------- depth
Assert-Equal 0 (Get-EbiJsonDepth 'x') 'depth: a scalar is 0'
Assert-Equal 0 (Get-EbiJsonDepth 5) 'depth: a number is 0'
Assert-Equal 0 (Get-EbiJsonDepth $null) 'depth: null is 0'
Assert-Equal 1 (Get-EbiJsonDepth @{}) 'depth: an empty map is 1'
Assert-Equal 1 (Get-EbiJsonDepth @(1, 2)) 'depth: a flat list is 1'
Assert-Equal 2 (Get-EbiJsonDepth @{ a = @{ b = 1 } }) 'depth: map in map is 2'
Assert-Equal 3 (Get-EbiJsonDepth @{ a = @( @{ b = 1 } ) }) 'depth: map in list in map is 3'
Assert-Equal 2 (Get-EbiJsonDepth ([pscustomobject]@{ a = @(1) })) 'depth: a PSCustomObject counts as a container'
Assert-Equal 20 (Get-EbiJsonMaxDepth) 'depth: the limit is 20'

function New-Nested { param([int]$Levels, $Leaf) $v = $Leaf; for ($i = 0; $i -lt $Levels; $i++) { $v = @{ ('l' + ($Levels - $i)) = $v } }; return $v }
function Get-Nested { param($Value, [int]$Levels) $v = $Value; for ($i = 1; $i -le $Levels; $i++) { $v = $v[('l' + $i)] }; return $v }

# ---------------------------------------------------------------- ConvertTo-EbiJson
$text = ConvertTo-EbiJson -Value @{ name = $jp; n = 3 }
Assert-True ($text.Contains($jp)) 'to-json: Japanese is written as characters, not \uXXXX'
Assert-True (-not ($text -match '\\u[0-9a-fA-F]{4}')) 'to-json: ... and no \uXXXX escape is left for it'
$text = ConvertTo-EbiJson -Value @{ q = 'a"b'; lt = '<' }
Assert-True ($text.Contains('\"') -or $text.Contains('"')) 'to-json: a quote stays escaped one way or the other (valid JSON)'
Assert-True (((ConvertFrom-EbiJson -Text $text)['value'])['q'] -eq 'a"b') 'to-json: ... and reads back'
Assert-Equal '<' ((ConvertFrom-EbiJson -Text $text)['value'])['lt'] 'to-json: an ASCII character PS 5.1 escapes (<) reads back'
$text = ConvertTo-EbiJson -Value @{ a = 1 } -Compress
Assert-True (-not $text.Contains("`n")) 'to-json: -Compress gives one line'
Assert-Equal '[]' (ConvertTo-EbiJson -Value @() -Compress) 'to-json: a top-level empty array is [] (not "" as a piped @() gives on 5.1)'
Assert-Equal '[1,2]' (ConvertTo-EbiJson -Value @(1, 2) -Compress) 'to-json: a top-level array serializes as an array'
Assert-Equal '{"a":[]}' (ConvertTo-EbiJson -Value @{ a = @() } -Compress) 'to-json: a nested empty array is []'
$deep20 = New-Nested -Levels 20 -Leaf 'bottom'
$text = ConvertTo-EbiJson -Value $deep20 -Compress
Assert-True ($text.Contains('"bottom"')) 'to-json: 20 levels serialize in full'
$deep21 = New-Nested -Levels 21 -Leaf 'bottom'
$threw = ''
try { [void](ConvertTo-EbiJson -Value $deep21) } catch { $threw = $_.Exception.Message }
Assert-True ($threw -like '*21 levels deep*at most 20*') 'to-json: 21 levels is refused loudly, never truncated to a string'
Assert-True (Test-EbiJsonSerializable @{ a = @(1, @{ b = $jp }) }) 'serializable: plain data is'
Assert-True (-not (Test-EbiJsonSerializable $deep21)) 'serializable: too deep is not'
Assert-True (Test-EbiJsonSerializable $null) 'serializable: null is'

# ---------------------------------------------------------------- ConvertFrom-EbiJson
$r = ConvertFrom-EbiJson -Text '{"a": 1, "b": {"c": [1, 2, {"d": "x"}]}, "e": null}'
Assert-True ($r['ok']) 'from-json: an object parses'
Assert-True ($r['value'] -is [hashtable]) 'from-json: ... into a hashtable'
Assert-Equal 'x' $r['value']['b']['c'][2]['d'] 'from-json: nested values are reachable by index'
Assert-True ($r['value'].Contains('e') -and $null -eq $r['value']['e']) 'from-json: null is a present null key'
$r = ConvertFrom-EbiJson -Text '[1, 2, 3]'
Assert-True ($r['ok'] -and $r['value'] -is [array] -and $r['value'].Count -eq 3) 'from-json: a top-level array is one array of three'
$r = ConvertFrom-EbiJson -Text '[[1, 2]]'
Assert-True ($r['value'].Count -eq 1 -and $r['value'][0].Count -eq 2) 'from-json: [[1,2]] is one array holding one array (not unrolled)'
$r = ConvertFrom-EbiJson -Text '[{"id": "a"}]'
Assert-True ($r['value'].Count -eq 1 -and $r['value'][0]['id'] -eq 'a') 'from-json: a one-object array is not unrolled into the object'
$r = ConvertFrom-EbiJson -Text '"just a string"'
Assert-True ($r['ok'] -and $r['value'] -eq 'just a string') 'from-json: a scalar parses'
$r = ConvertFrom-EbiJson -Text '42'
Assert-True ($r['ok'] -and $r['value'] -eq 42) 'from-json: a number parses'
$r = ConvertFrom-EbiJson -Text ''
Assert-True (-not $r['ok'] -and $r['message'] -like '*empty*') 'from-json: blank text is a failure, not null'
$r = ConvertFrom-EbiJson -Text '{ not json'
Assert-True (-not $r['ok'] -and $r['message'] -like 'not valid JSON:*') 'from-json: malformed text is a failure record, not an exception'
$r = ConvertFrom-EbiJson -Text ('{"name": "' + $jp + '"}')
Assert-Equal $jp $r['value']['name'] 'from-json: Japanese in the text reads back'

# ---------------------------------------------------------------- files
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('ebidensu-json-' + [guid]::NewGuid().ToString('N'))
try {
    # the card's criteria: Japanese profile round-trip, 5 levels, no BOM, no \uXXXX
    $profile = @{ pages = @{ transferStatus = @{ title = $jp; crop = @{ left = 6 }; tabs = @('a', 'b') } }; key = @{ columns = @('Correl_ID_S') } }
    $path = Join-Path (Join-Path $tmp 'profiles') 'pages.json'
    $w = Write-EbiJson -Path $path -Value $profile
    Assert-True ($w['ok']) 'write: writes, creating parent directories'
    Assert-True (Test-Path -LiteralPath $path) 'write: the file exists'
    $bytes = [System.IO.File]::ReadAllBytes($path)
    Assert-True (-not ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)) 'write: no UTF-8 BOM'
    $raw = [System.IO.File]::ReadAllText($path, $utf8)
    Assert-True ($raw.Contains($jp)) 'write: the file holds the Japanese characters themselves'
    Assert-True (-not ($raw -match '\\u[0-9a-fA-F]{4}')) 'write: no \uXXXX in the file'
    Assert-True ($raw.EndsWith([Environment]::NewLine)) 'write: one trailing newline'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath (Split-Path $path -Parent) -Filter '*.tmp-*').Count 'write: no temporary file is left behind'
    $r = Read-EbiJson -Path $path
    Assert-True ($r['ok'] -and $r['exists']) 'read: reads back'
    Assert-Equal $jp $r['value']['pages']['transferStatus']['title'] 'read: Japanese survives the file round trip'
    Assert-Equal 6 $r['value']['pages']['transferStatus']['crop']['left'] 'read: a value four levels down is intact'
    Assert-Equal 'b' $r['value']['pages']['transferStatus']['tabs'][1] 'read: an array four levels down is intact'
    Assert-Equal 'Correl_ID_S' $r['value']['key']['columns'][0] 'read: a one-element array is still an array'

    $five = New-Nested -Levels 5 -Leaf @{ leaf = $jp; list = @(1, 2, 3) }
    $p5 = Join-Path $tmp 'five.json'
    [void](Write-EbiJson -Path $p5 -Value $five)
    $back = (Read-EbiJson -Path $p5)['value']
    Assert-Equal $jp (Get-Nested -Value $back -Levels 5)['leaf'] 'round-trip: 5 levels of nesting are not truncated'
    Assert-Equal 3 (Get-Nested -Value $back -Levels 5)['list'].Count 'round-trip: the list at the bottom is a list'
    $p20 = Join-Path $tmp 'twenty.json'
    [void](Write-EbiJson -Path $p20 -Value $deep20)
    Assert-Equal 'bottom' (Get-Nested -Value (Read-EbiJson -Path $p20)['value'] -Levels 20) 'round-trip: 20 levels round-trip'
    $w = Write-EbiJson -Path (Join-Path $tmp 'never.json') -Value $deep21
    Assert-True (-not $w['ok'] -and $w['message'] -like '*21 levels deep*') 'write: a value too deep is refused as a record, and no file is written'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $tmp 'never.json'))) 'write: ... really no file'

    # atomic overwrite: the old content is replaced whole
    $w = Write-EbiJson -Path $path -Value @{ replaced = $true }
    Assert-True ($w['ok']) 'write: overwriting an existing file works'
    Assert-Equal 'True' ([string](Read-EbiJson -Path $path)['value']['replaced']) 'write: ... and the new content is what reads back'
    Assert-Equal 0 @(Get-ChildItem -LiteralPath (Split-Path $path -Parent) -Filter '*.tmp-*').Count 'write: no temporary file after an overwrite either'

    # reading what is not there, or not JSON
    $r = Read-EbiJson -Path (Join-Path $tmp 'absent.json')
    Assert-True (-not $r['ok'] -and -not $r['exists'] -and $r['message'] -like 'file not found:*') 'read: a missing file is ok=false, exists=false'
    $badPath = Join-Path $tmp 'bad.json'
    [System.IO.File]::WriteAllText($badPath, '{ nope', $utf8)
    $r = Read-EbiJson -Path $badPath
    Assert-True (-not $r['ok'] -and $r['exists'] -and $r['message'] -like '*bad.json: not valid JSON:*') 'read: a malformed file is ok=false, exists=true, naming the file'
    $r = Read-EbiJson -Path ''
    Assert-True (-not $r['ok']) 'read: an empty path is a failure record'
    $bomPath = Join-Path $tmp 'bom.json'
    [System.IO.File]::WriteAllText($bomPath, ('{"name":"' + $jp + '"}'), (New-Object System.Text.UTF8Encoding($true)))
    $r = Read-EbiJson -Path $bomPath
    Assert-True ($r['ok'] -and $r['value']['name'] -eq $jp) 'read: a file someone saved WITH a BOM still reads (the decoder skips it)'

    # JSONL
    $jl = Join-Path (Join-Path $tmp 'run') 'ledger.jsonl'
    foreach ($i in 1..3) { $a = Add-EbiJsonLine -Path $jl -Value @{ n = $i; text = $jp; data = @{ lines = @($i) } }; Assert-True ($a['ok']) ('jsonl: line ' + $i + ' appended') }
    $lines = [System.IO.File]::ReadAllLines($jl, $utf8)
    Assert-Equal 3 $lines.Count 'jsonl: one line per value'
    Assert-True ($lines[0].Contains($jp)) 'jsonl: Japanese is readable in the line'
    $r = Read-EbiJsonLines -Path $jl
    Assert-True ($r['ok'] -and $r['value'].Count -eq 3) 'jsonl: three values read back'
    Assert-True ($r['value'][1] -is [hashtable]) 'jsonl: each as a hashtable'
    Assert-Equal 2 $r['value'][1]['n'] 'jsonl: in file order'
    Assert-Equal 3 $r['value'][2]['data']['lines'][0] 'jsonl: nested data intact'
    Assert-Equal 0 @($r['badLines']).Count 'jsonl: no bad lines'
    Assert-True (-not $r['partial']) 'jsonl: not partial'

    [System.IO.File]::AppendAllText($jl, '{"n":', $utf8)   # a writer mid-line
    $r = Read-EbiJsonLines -Path $jl
    Assert-Equal 3 $r['value'].Count 'jsonl: a half-written final line is not a value'
    Assert-True ($r['partial']) 'jsonl: ... and is reported as partial'
    Assert-Equal 0 @($r['badLines']).Count 'jsonl: ... not as a bad line'
    [System.IO.File]::AppendAllText($jl, ('4}' + [Environment]::NewLine + 'garbage' + [Environment]::NewLine + [Environment]::NewLine + '{"n":5}' + [Environment]::NewLine), $utf8)
    $r = Read-EbiJsonLines -Path $jl
    Assert-Equal 5 $r['value'].Count 'jsonl: the completed line and the one after the garbage both count'
    Assert-Equal 5 $r['value'][4]['n'] 'jsonl: ... in order'
    Assert-Equal '5' (@($r['badLines']) -join ',') 'jsonl: a malformed middle line is reported by number, not dropped in silence'
    Assert-True (-not $r['partial']) 'jsonl: a good final line is not partial'
    Assert-True ($r['message'] -like '1 line(s) are not JSON: 5') 'jsonl: the message says which'
    $r = Read-EbiJsonLines -Path (Join-Path $tmp 'absent.jsonl')
    Assert-True (-not $r['ok'] -and -not $r['exists'] -and $r['value'].Count -eq 0) 'jsonl: a missing file is ok=false with an empty value'
} finally {
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

$rc = Complete-Tests
exit $rc
