#Requires -Version 5.1
# ============================================================
#  kernel/Json.ps1
#
#  The ONE place ebi-dance reads and writes JSON (P1-35). Dot-source only
#  (no param() block, ASCII source, no class -- CLAUDE.md conventions).
#
#  Windows PowerShell 5.1 has four JSON traps, and every one of them has
#  either bitten this repository or has a private workaround somewhere in
#  it. This file is where those workarounds live, once:
#
#    1. ConvertFrom-Json hands back PSCustomObjects, not hashtables. Under
#       Set-StrictMode a missing property read with dot syntax throws, and
#       .ContainsKey does not exist. -> everything read here comes back as
#       plain hashtables (index access never throws) and object[] arrays
#       (ConvertTo-EbiHashtable, moved in from the runner spike).
#    2. ConvertTo-Json defaults to -Depth 2 and SILENTLY flattens anything
#       deeper into a string. -> depth is fixed at 20 here, and a value
#       nested deeper than that is refused loudly (an exception naming the
#       depth) rather than truncated: no data written by this file is ever
#       cut short without anyone knowing.
#    3. ConvertTo-Json turns every non-ASCII character into \uXXXX, so the
#       Japanese in a profile or a trace is unreadable in the file -- and
#       on 5.1 also ' < > & (' ...), which 7 writes as themselves. ->
#       the writer turns every escape that is valid unescaped back into
#       its character (ConvertFrom-ConfigJson's ConvertFrom-JsonUnicode
#       Escape, moved in and widened); quotes, backslashes and control
#       characters stay escapes so the text remains valid JSON. And 5.1's
#       indented layout differs from 7's (4-space indent, two spaces after
#       the colon), so ConvertTo-EbiJson lays the text out itself
#       (Format-EbiJsonPretty): the same value is the same bytes on both.
#    4. A file read without an explicit encoding is decoded as ANSI on a JP
#       locale host, which mojibakes UTF-8 Japanese. -> every read here is
#       [IO.File]::ReadAllText with UTF-8, every write is UTF-8 without a
#       BOM (Set-Content -Encoding UTF8 would add one, and would corrupt a
#       JSONL file when used repeatedly).
#
#  Rule R8 (BACKLOG.md iron-rule table; STEP-CONTRACT.md 1.1 and 7):
#  nothing under modules/** or kernel/** calls ConvertFrom-Json /
#  ConvertTo-Json / Get-Content on a .json file directly. The contract
#  checker greps for it.
#
#  Failures are return values, never exceptions, for everything that
#  touches a file or parses text: @{ ok; value; message; ... }. The one
#  exception is the depth guard in ConvertTo-EbiJson (trap 2 above): a
#  value too deep to serialize is a bug in the caller, not data, and a
#  string-returning function has nowhere else to say so. Callers that must
#  never fail (the trace) wrap it.
# ============================================================

function Get-EbiJsonMaxDepth { return 20 }

function ConvertTo-EbiHashtable {
    <#
      ConvertFrom-Json hands back PSCustomObjects; the kernel wants plain
      hashtables (index access never throws under StrictMode, and the step
      contract is written in terms of hashtables). Recursive. Arrays come
      back as object[] built by an explicit loop -- the @() wrap over an
      indexed collection is the shape this repo bans (CLAUDE.md, R4).
    #>
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $Value.Keys) { $h[[string]$k] = ConvertTo-EbiHashtable $Value[$k] }
        return $h
    }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Collections.IList]) {
        $list = New-Object System.Collections.ArrayList
        foreach ($item in $Value) { [void]$list.Add((ConvertTo-EbiHashtable $item)) }
        # The unary comma matters: a one-element array returned bare is
        # unrolled into its element, and a "setup" with a single step call
        # would come back as that call instead of a list of one. Callers
        # assign the result; none wraps it in @().
        return ,$list.ToArray()
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($p in $Value.PSObject.Properties) { $h[$p.Name] = ConvertTo-EbiHashtable $p.Value }
        return $h
    }
    return $Value
}

function Get-EbiJsonDepth {
    <#
      How deep a value nests, counted the way ConvertTo-Json's -Depth counts:
      a scalar is 0, a container is 1 + its deepest child. @{ a = @{ b = 1 } }
      is 2. PSCustomObjects count as containers (the trace hands them in).
    #>
    param($Value)
    if ($null -eq $Value -or $Value -is [string] -or $Value -is [System.ValueType]) { return 0 }
    $max = 0
    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($k in $Value.Keys) { $d = Get-EbiJsonDepth $Value[$k]; if ($d -gt $max) { $max = $d } }
        return ($max + 1)
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        foreach ($item in $Value) { $d = Get-EbiJsonDepth $item; if ($d -gt $max) { $max = $d } }
        return ($max + 1)
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in $Value.PSObject.Properties) { $d = Get-EbiJsonDepth $p.Value; if ($d -gt $max) { $max = $d } }
        return ($max + 1)
    }
    return 0
}

function ConvertFrom-EbiJsonUnicodeEscape {
    # \uXXXX -> the character itself for every code point that is valid
    # unescaped inside a JSON string: >= 0x80 (Japanese), and the printable
    # ASCII PS 5.1 escapes for HTML's sake (' < > & come out as '
    # < > & on 5.1 and as themselves on 7 -- the same value
    # must serialize to the same text on both, or a generated file drifts
    # with the PowerShell that wrote it). Control characters, the quote and
    # the backslash stay escapes so the text remains valid JSON. A literal
    # backslash followed by uXXXX in the DATA is "\\uXXXX" in the text (an
    # odd run of backslashes never precedes the u) and is left alone.
    # Surrogate pairs come out as two chars, which is the pair again.
    param([string]$Json)
    if ([string]::IsNullOrEmpty($Json)) { return $Json }
    $evaluator = {
        param($m)
        $code = [Convert]::ToInt32($m.Groups[2].Value, 16)
        if ($code -lt 0x20 -or $code -eq 0x22 -or $code -eq 0x5C) { return $m.Value }
        return ($m.Groups[1].Value + [string][char]$code)
    }
    return [regex]::Replace($Json, '(?<!\\)((?:\\\\)*)\\u([0-9a-fA-F]{4})', $evaluator)
}

function Format-EbiJsonPretty {
    <#
      PURE. Compressed JSON text -> one indented layout, the same on every
      PowerShell: two-space indent, "key": value with one space, one element
      per line, {} and [] for empty containers (what pwsh 7 prints; 5.1
      prints four-space indents, two spaces after the colon and a blank
      line inside an empty array, so a file written by 5.1 and one written
      by 7 differed on every line). Strings are copied through untouched,
      escapes included; whitespace outside strings is dropped.
    #>
    param([string]$Json, [int]$Indent = 2)
    if ([string]::IsNullOrEmpty($Json)) { return $Json }
    $sb = New-Object System.Text.StringBuilder
    $level = 0; $inStr = $false; $esc = $false
    $n = $Json.Length
    for ($i = 0; $i -lt $n; $i++) {
        $c = [string]$Json[$i]
        if ($inStr) {
            [void]$sb.Append($c)
            if ($esc) { $esc = $false } elseif ($c -eq '\') { $esc = $true } elseif ($c -eq '"') { $inStr = $false }
            continue
        }
        if ($c -eq '"') { $inStr = $true; [void]$sb.Append($c); continue }
        if ($c -eq '{' -or $c -eq '[') {
            $close = if ($c -eq '{') { '}' } else { ']' }
            $j = $i + 1
            while ($j -lt $n -and [char]::IsWhiteSpace($Json[$j])) { $j++ }
            if ($j -lt $n -and [string]$Json[$j] -eq $close) { [void]$sb.Append($c).Append($close); $i = $j; continue }
            $level++
            [void]$sb.Append($c).Append("`n").Append(' ' * ($Indent * $level))
            continue
        }
        if ($c -eq '}' -or $c -eq ']') {
            $level = [Math]::Max(0, $level - 1)
            [void]$sb.Append("`n").Append(' ' * ($Indent * $level)).Append($c)
            continue
        }
        if ($c -eq ',') { [void]$sb.Append(',').Append("`n").Append(' ' * ($Indent * $level)); continue }
        if ($c -eq ':') { [void]$sb.Append(': '); continue }
        if ([char]::IsWhiteSpace($Json[$i])) { continue }
        [void]$sb.Append($c)
    }
    return $sb.ToString()
}

function ConvertTo-EbiJson {
    <#
      Value -> JSON text. Depth fixed at Get-EbiJsonMaxDepth (20); non-ASCII
      and ' < > & readable; -Compress for one-line output (JSONL), else the
      one layout of Format-EbiJsonPretty -- ConvertTo-Json is only ever
      asked for the compressed form here, because its indented form differs
      between 5.1 and 7 and a committed generated file (catalog.json) must
      not change with the PowerShell that regenerated it. Always
      -InputObject, never piped: a piped empty array unrolls to nothing and
      serializes as "" on PS 5.1, which reads back as a string.

      THROWS when the value nests deeper than the limit -- see the header.
    #>
    param($Value, [switch]$Compress)
    $depth = Get-EbiJsonDepth $Value
    $limit = Get-EbiJsonMaxDepth
    if ($depth -gt $limit) {
        throw ('value nests {0} levels deep; ConvertTo-EbiJson serializes at most {1} (deeper would be silently truncated to a string)' -f $depth, $limit)
    }
    $text = ConvertFrom-EbiJsonUnicodeEscape ([string](ConvertTo-Json -InputObject $Value -Depth $limit -Compress))
    if ($Compress.IsPresent) { return $text }
    return (Format-EbiJsonPretty -Json $text)
}

function Test-EbiJsonSerializable {
    # Can this value be written by ConvertTo-EbiJson? False for handles, COM
    # objects, anything ConvertTo-Json rejects, and anything too deep.
    param($Value)
    try { [void](ConvertTo-EbiJson -Value $Value -Compress); return $true }
    catch { return $false }
}

function ConvertFrom-EbiJson {
    <#
      JSON text -> @{ ok; value; message }. value is hashtables / object[] /
      scalars (ConvertTo-EbiHashtable), never a PSCustomObject.

      The text is parsed wrapped as {"v": <text>} so a top-level array comes
      back as one array on every PowerShell version: piped or not, PS 5.1
      unrolls a top-level JSON array into its elements, and once unrolled
      [[1,2]] and [1,2] cannot be told apart. Blank text is not JSON.
    #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @{ ok = $false; value = $null; message = 'empty text is not JSON' }
    }
    try {
        $wrapped = ConvertFrom-Json -InputObject ('{"v":' + $Text + '}') -ErrorAction Stop
        return @{ ok = $true; value = (ConvertTo-EbiHashtable $wrapped.v); message = '' }
    } catch {
        return @{ ok = $false; value = $null; message = ('not valid JSON: ' + $_.Exception.Message) }
    }
}

function Get-EbiJsonEncoding { return (New-Object System.Text.UTF8Encoding($false)) }

function Read-EbiJson {
    <#
      File -> @{ ok; value; message; exists }. A missing file is ok=$false
      with exists=$false, so a caller can tell "no sidecar yet" (P0-R14:
      data=null plus a warning) from "a sidecar that cannot be read". UTF-8;
      a BOM, if present, is skipped by the decoder.
      -AllowCp932: for a file a person edits by hand (<WorkDir>/ebi.local.json):
      Notepad on a Japanese Windows saves "ANSI" = Shift_JIS, and reading
      those bytes as UTF-8 turns every Japanese path into U+FFFD. Bytes that
      are not valid UTF-8 are then read as CP932 instead, and the result
      says so (encoding = 'cp932'; otherwise 'utf8').
    #>
    param([string]$Path, [switch]$AllowCp932)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return @{ ok = $false; value = $null; message = 'no path given'; exists = $false }
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return @{ ok = $false; value = $null; message = ('file not found: ' + $Path); exists = $false }
    }
    $text = ''
    $enc = 'utf8'
    try {
        if ($AllowCp932) {
            $bytes = [System.IO.File]::ReadAllBytes($Path)
            $start = 0
            if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { $start = 3 }
            try {
                $text = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($bytes, $start, $bytes.Length - $start)
            } catch {
                $cp = $null
                try { $cp = [System.Text.Encoding]::GetEncoding(932) } catch {
                    # .NET Core (pwsh 7) ships code pages behind a provider
                    [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
                    $cp = [System.Text.Encoding]::GetEncoding(932)
                }
                $text = $cp.GetString($bytes)
                $enc = 'cp932'
            }
        } else {
            $text = [System.IO.File]::ReadAllText($Path, (Get-EbiJsonEncoding))
        }
    } catch {
        return @{ ok = $false; value = $null; message = ('cannot read {0}: {1}' -f $Path, $_.Exception.Message); exists = $true; encoding = $enc }
    }
    $parsed = ConvertFrom-EbiJson -Text $text
    if (-not $parsed['ok']) {
        return @{ ok = $false; value = $null; message = ('{0}: {1}' -f $Path, $parsed['message']); exists = $true; encoding = $enc }
    }
    return @{ ok = $true; value = $parsed['value']; message = ''; exists = $true; encoding = $enc }
}

function Write-EbiJson {
    <#
      Value -> file, atomically: the text goes to a temporary file in the
      same directory and is then moved over the target (File.Replace when
      the target exists, File.Move when it does not), so a reader never sees
      a half-written file and a crash mid-write leaves the old one intact.
      Parent directories are created. UTF-8, no BOM, one trailing newline.
      Returns @{ ok; message }.
    #>
    param([string]$Path, $Value)
    if ([string]::IsNullOrWhiteSpace($Path)) { return @{ ok = $false; message = 'no path given' } }
    $text = ''
    try { $text = ConvertTo-EbiJson -Value $Value }
    catch { return @{ ok = $false; message = $_.Exception.Message } }

    $tmp = ''
    try {
        $dir = Split-Path -Path $Path -Parent
        if (-not [string]::IsNullOrEmpty($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $tmp = $Path + '.tmp-' + [guid]::NewGuid().ToString('N')
        [System.IO.File]::WriteAllText($tmp, $text + [Environment]::NewLine, (Get-EbiJsonEncoding))
        if (Test-Path -LiteralPath $Path) {
            # [NullString]::Value, not $null: PowerShell turns $null into ''
            # for a [string] argument, and Replace rejects '' as a backup path.
            [System.IO.File]::Replace($tmp, $Path, [NullString]::Value)
        } else {
            [System.IO.File]::Move($tmp, $Path)
        }
        return @{ ok = $true; message = '' }
    } catch {
        if ($tmp -ne '' -and (Test-Path -LiteralPath $tmp)) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        return @{ ok = $false; message = ('cannot write {0}: {1}' -f $Path, $_.Exception.Message) }
    }
}

function Add-EbiJsonLine {
    <#
      Append one value as one compressed line to a JSONL file (trace,
      ledger). Parent directories are created. Returns @{ ok; message }.
      Not atomic across processes -- one writer per file is the contract
      (a run owns its run/<runId>/ directory).
    #>
    param([string]$Path, $Value)
    if ([string]::IsNullOrWhiteSpace($Path)) { return @{ ok = $false; message = 'no path given' } }
    try {
        $line = ConvertTo-EbiJson -Value $Value -Compress
        $dir = Split-Path -Path $Path -Parent
        if (-not [string]::IsNullOrEmpty($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        [System.IO.File]::AppendAllText($Path, $line + [Environment]::NewLine, (Get-EbiJsonEncoding))
        return @{ ok = $true; message = '' }
    } catch {
        return @{ ok = $false; message = ('cannot append to {0}: {1}' -f $Path, $_.Exception.Message) }
    }
}

function Read-EbiJsonLines {
    <#
      JSONL file -> @{ ok; value; badLines; partial; message; exists }.
        value     object[] of parsed lines (hashtables), in file order;
                  blank lines skipped
        badLines  1-based numbers of lines that are not JSON, EXCLUDING a
                  malformed final line
        partial   $true when the final line is malformed -- a reader can
                  see a half-written last line while a writer is appending,
                  and that is not corruption
      ok is $true whenever the file could be read; a bad line in the middle
      is reported in badLines and the parsed lines are still returned, so
      the caller decides whether that is a warning or a failure. Nothing is
      dropped without being counted.
    #>
    param([string]$Path)
    $empty = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return @{ ok = $false; value = $empty.ToArray(); badLines = @(); partial = $false; message = 'no path given'; exists = $false }
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return @{ ok = $false; value = $empty.ToArray(); badLines = @(); partial = $false; message = ('file not found: ' + $Path); exists = $false }
    }
    $lines = $null
    try {
        $lines = [System.IO.File]::ReadAllLines($Path, (Get-EbiJsonEncoding))
    } catch {
        return @{ ok = $false; value = $empty.ToArray(); badLines = @(); partial = $false; message = ('cannot read {0}: {1}' -f $Path, $_.Exception.Message); exists = $true }
    }
    $values = New-Object System.Collections.ArrayList
    $bad    = New-Object System.Collections.ArrayList
    $lastNonBlank = 0
    for ($i = 0; $i -lt $lines.Count; $i++) { if (-not [string]::IsNullOrWhiteSpace($lines[$i])) { $lastNonBlank = $i + 1 } }
    $partial = $false
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parsed = ConvertFrom-EbiJson -Text $line
        if ($parsed['ok']) { [void]$values.Add($parsed['value']); continue }
        if (($i + 1) -eq $lastNonBlank) { $partial = $true } else { [void]$bad.Add($i + 1) }
    }
    $msg = ''
    if ($bad.Count -gt 0) { $msg = ('{0} line(s) are not JSON: {1}' -f $bad.Count, ($bad.ToArray() -join ', ')) }
    return @{ ok = $true; value = $values.ToArray(); badLines = $bad.ToArray(); partial = $partial; message = $msg; exists = $true }
}
