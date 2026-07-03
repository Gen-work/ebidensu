# ============================================================
#  OcrWindows.ps1
#
#  Windows built-in OCR (Windows.Media.Ocr WinRT API) callable from
#  PowerShell 5.1. This is the same engine family behind the Snipping
#  Tool text extraction / PowerToys Text Extractor -- zero installs,
#  Japanese supported when the ja language pack is present.
#
#  Dot-source only (no param() block). Dot-sourcing never throws on a
#  non-Windows host: all WinRT loading is deferred to Initialize-WinOcr
#  and reported via Test-WinOcrAvailable / Get-WinOcrInitError, so the
#  cloud-side parse check and unit-test run stay green.
#
#  Output is projected to plain pscustomobjects (Text + word boxes) so
#  consumers (SendMetadata.ps1) never need WinRT types and stay
#  unit-testable with synthetic fixtures.
# ============================================================

$script:WinOcrReady = $false
$script:WinOcrInitError = ''
$script:WinOcrAsTaskGeneric = $null
$script:WinOcrTextStrategy = ''      # first strategy that produced text
$script:WinOcrNativeReader = $null   # $true/$false after first Add-Type try

# Compiled C# string reader: bypasses the PS WinRT adapter entirely.
# References the in-box metadata under C:\Windows\System32\WinMetadata
# (present on every Windows 10/11; no SDK install needed). Lazy: only
# compiled the first time the cheaper strategies fail.
function Initialize-WinOcrNativeTextReader {
    if ($null -ne $script:WinOcrNativeReader) { return [bool]$script:WinOcrNativeReader }
    $script:WinOcrNativeReader = $false
    try {
        if (([System.Management.Automation.PSTypeName]'VerifyOcr.NativeText').Type) {
            $script:WinOcrNativeReader = $true
            return $true
        }
    } catch {}
    try {
        $winmdDir = Join-Path $env:windir 'System32\WinMetadata'
        $mediaWinmd = Join-Path $winmdDir 'Windows.Media.winmd'
        $foundationWinmd = Join-Path $winmdDir 'Windows.Foundation.winmd'
        if (-not (Test-Path -LiteralPath $mediaWinmd) -or -not (Test-Path -LiteralPath $foundationWinmd)) {
            return $false
        }
        Add-Type -ReferencedAssemblies @($mediaWinmd, $foundationWinmd, 'System.Runtime.WindowsRuntime') -TypeDefinition @'
namespace VerifyOcr {
    public static class NativeText {
        public static string LineText(object line)  { return ((Windows.Media.Ocr.OcrLine)line).Text; }
        public static string WordText(object word)  { return ((Windows.Media.Ocr.OcrWord)word).Text; }
        public static string ResultText(object res) { return ((Windows.Media.Ocr.OcrResult)res).Text; }
    }
}
'@ -ErrorAction Stop
        $script:WinOcrNativeReader = $true
    } catch {
        $script:WinOcrNativeReader = $false
    }
    return [bool]$script:WinOcrNativeReader
}

# Reads a WinRT Text property through layered strategies, because on some
# hosts the PS 5.1 adapter silently returns empty for every HSTRING
# property while collections still enumerate (field-observed: lines=92
# words=489 chars=0 rawChars=0). Kind: 'Line' / 'Word' / 'Result'.
# Records the first strategy that worked in $script:WinOcrTextStrategy.
function Read-WinRtText {
    param($Object, [string]$Kind)
    if ($null -eq $Object) { return '' }
    $v = ''
    try { $v = [string]$Object.Text } catch {}
    if ($v.Length -gt 0) {
        if ([string]::IsNullOrEmpty($script:WinOcrTextStrategy)) { $script:WinOcrTextStrategy = 'adapter' }
        return $v
    }
    try { $v = [string]$Object.psbase.Text } catch {}
    if ($v.Length -gt 0) {
        if ([string]::IsNullOrEmpty($script:WinOcrTextStrategy)) { $script:WinOcrTextStrategy = 'psbase' }
        return $v
    }
    try {
        $pi = $Object.GetType().GetProperty('Text')
        if ($null -ne $pi) { $v = [string]$pi.GetValue($Object, $null) }
    } catch {}
    if ($v.Length -gt 0) {
        if ([string]::IsNullOrEmpty($script:WinOcrTextStrategy)) { $script:WinOcrTextStrategy = 'reflection' }
        return $v
    }
    if (Initialize-WinOcrNativeTextReader) {
        try {
            switch ($Kind) {
                'Line'   { $v = [string][VerifyOcr.NativeText]::LineText($Object) }
                'Word'   { $v = [string][VerifyOcr.NativeText]::WordText($Object) }
                'Result' { $v = [string][VerifyOcr.NativeText]::ResultText($Object) }
            }
        } catch {}
        if ($v.Length -gt 0 -and [string]::IsNullOrEmpty($script:WinOcrTextStrategy)) {
            $script:WinOcrTextStrategy = 'compiled'
        }
    }
    return $v
}

function Get-WinOcrTextStrategy {
    return $script:WinOcrTextStrategy
}

function Initialize-WinOcr {
    if ($script:WinOcrReady) { return $true }
    if (-not [string]::IsNullOrEmpty($script:WinOcrInitError)) { return $false }
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
        # Touching one type per namespace pulls the WinRT projection in.
        $null = [Windows.Media.Ocr.OcrEngine,Windows.Foundation.UniversalApiContract,ContentType=WindowsRuntime]
        $null = [Windows.Globalization.Language,Windows.Globalization,ContentType=WindowsRuntime]
        $null = [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics,ContentType=WindowsRuntime]
        $script:WinOcrAsTaskGeneric = @([System.WindowsRuntimeSystemExtensions].GetMethods() |
            Where-Object {
                $_.Name -eq 'AsTask' -and
                $_.GetParameters().Count -eq 1 -and
                $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
            })[0]
        if ($null -eq $script:WinOcrAsTaskGeneric) { throw 'AsTask(IAsyncOperation) bridge not found' }
        $script:WinOcrReady = $true
        return $true
    } catch {
        $script:WinOcrInitError = $_.Exception.Message
        return $false
    }
}

function Test-WinOcrAvailable {
    return (Initialize-WinOcr)
}

function Get-WinOcrInitError {
    return $script:WinOcrInitError
}

# Blocks on a WinRT IAsyncOperation and returns its result.
function Wait-WinOcrOperation {
    param($Operation, [type]$ResultType)
    $asTask = $script:WinOcrAsTaskGeneric.MakeGenericMethod($ResultType)
    $task = $asTask.Invoke($null, @($Operation))
    [void]$task.Wait(-1)
    return $task.Result
}

function Get-WinOcrLanguageTags {
    if (-not (Initialize-WinOcr)) { return @() }
    $tags = @()
    try {
        foreach ($l in [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages) {
            $tags += [string]$l.LanguageTag
        }
    } catch {}
    return $tags
}

# Returns an OcrEngine for the requested language tag (e.g. 'ja'),
# falling back to the user-profile languages, or $null when none.
function Get-WinOcrEngine {
    param([string]$LanguageTag = '')
    if (-not (Initialize-WinOcr)) { return $null }
    $engine = $null
    if (-not [string]::IsNullOrWhiteSpace($LanguageTag)) {
        try {
            $lang = New-Object Windows.Globalization.Language($LanguageTag.Trim())
            $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
        } catch { $engine = $null }
    }
    if ($null -eq $engine) {
        try { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() } catch { $engine = $null }
    }
    return $engine
}

# Diagnostic sweep for ONE image: tries every installed recognizer
# language (plus the user-profile engine) and reports what each saw,
# alongside the engine's max image dimension and the image's pixel size.
# Pure data out (pscustomobject); callers do the formatting.
function Invoke-WinOcrDiag {
    param([string]$Path)
    if (-not (Initialize-WinOcr)) {
        throw ("Windows OCR not available: {0}" -f $script:WinOcrInitError)
    }
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'image path is empty' }
    if (-not (Test-Path -LiteralPath $Path)) { throw "image not found: $Path" }

    $maxDim = 0
    try { $maxDim = [int][Windows.Media.Ocr.OcrEngine]::MaxImageDimension } catch {}
    $px = ''
    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        $img = [System.Drawing.Image]::FromFile((Resolve-Path -LiteralPath $Path).ProviderPath)
        try { $px = ('{0}x{1}' -f [int]$img.Width, [int]$img.Height) } finally { $img.Dispose() }
    } catch {}

    $attempts = @()
    $tagList = @(Get-WinOcrLanguageTags) + @('')   # '' = user-profile engine
    foreach ($tag in $tagList) {
        $label = if ([string]::IsNullOrWhiteSpace($tag)) { '(user profile)' } else { $tag }
        try {
            $res = Invoke-WinOcrFile -Path $Path -LanguageTag $tag
            $lineCount = @($res.Lines).Count
            $wordCount = 0
            $charCount = 0
            $sample = ''
            foreach ($ln in @($res.Lines)) {
                $wordCount += @($ln.Words).Count
                $t = [string]$ln.Text
                $charCount += $t.Length
                if ([string]::IsNullOrWhiteSpace($sample) -and -not [string]::IsNullOrWhiteSpace($t)) { $sample = $t }
            }
            $rawLen = 0
            try { $rawLen = ([string]$res.RawText).Length } catch {}
            # word-box probe: nonzero X/Width proves struct marshaling works
            # even when string reads fail
            $wordBox = ''
            foreach ($ln in @($res.Lines)) {
                $ws = @($ln.Words)
                if ($ws.Count -gt 0) {
                    $wordBox = ('X={0} Y={1} W={2} H={3}' -f [int]$ws[0].X, [int]$ws[0].Y, [int]$ws[0].Width, [int]$ws[0].Height)
                    break
                }
            }
            $attempts += [pscustomobject]@{
                Language = $label
                Engine   = [string]$res.LanguageTag
                Lines    = [int]$lineCount
                Words    = [int]$wordCount
                Chars    = [int]$charCount
                RawChars = [int]$rawLen
                LineType = [string]$res.LineTypeName
                Strategy = [string]$res.TextStrategy
                WordBox  = $wordBox
                Sample   = $sample
                Error    = ''
            }
        } catch {
            $attempts += [pscustomobject]@{
                Language = $label; Engine = ''; Lines = 0; Words = 0; Chars = 0; RawChars = 0
                LineType = ''; Strategy = ''; WordBox = ''; Sample = ''
                Error    = [string]$_.Exception.Message
            }
        }
    }
    return [pscustomobject]@{
        Path              = $Path
        PixelSize         = $px
        MaxImageDimension = $maxDim
        Attempts          = @($attempts)
    }
}
#   @{ Path; LanguageTag; Text; Lines = @(@{ Text; Words = @(@{ Text; X; Y; Width; Height }) }) }
# Word boxes are in image pixels; SendMetadata.ps1 uses them to rebuild
# the spacing the Japanese recognizer drops.
function Invoke-WinOcrFile {
    param([string]$Path, [string]$LanguageTag = '')
    if (-not (Initialize-WinOcr)) {
        throw ("Windows OCR not available: {0}" -f $script:WinOcrInitError)
    }
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'image path is empty' }
    if (-not (Test-Path -LiteralPath $Path)) { throw "image not found: $Path" }
    $engine = Get-WinOcrEngine $LanguageTag
    if ($null -eq $engine) {
        throw ("no OCR recognizer language available (requested '{0}'; installed: {1})" -f $LanguageTag, ((Get-WinOcrLanguageTags) -join ', '))
    }

    $full = (Resolve-Path -LiteralPath $Path).ProviderPath
    $file = Wait-WinOcrOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($full)) ([Windows.Storage.StorageFile])
    $stream = Wait-WinOcrOperation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder = Wait-WinOcrOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Wait-WinOcrOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        try {
            $ocr = Wait-WinOcrOperation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
            # OcrResult.Text reads the whole text in one call -- keep it as a
            # fallback: in some PS 5.1 WinRT projections enumerating
            # Lines/Words works but their .Text properties silently return
            # null (field-observed: lines=92 words=489 yet every Text empty).
            $rawText = Read-WinRtText $ocr 'Result'
            $lineTypeName = ''
            $charCount = 0
            $lines = @()
            foreach ($ln in $ocr.Lines) {
                if ([string]::IsNullOrEmpty($lineTypeName)) {
                    try { $lineTypeName = $ln.GetType().FullName } catch {}
                }
                $words = @()
                foreach ($w in $ln.Words) {
                    $r = $w.BoundingRect
                    $words += [pscustomobject]@{
                        Text   = Read-WinRtText $w 'Word'
                        X      = [double]$r.X
                        Y      = [double]$r.Y
                        Width  = [double]$r.Width
                        Height = [double]$r.Height
                    }
                }
                $text = Read-WinRtText $ln 'Line'
                $charCount += $text.Length
                $lines += [pscustomobject]@{ Text = $text; Words = @($words) }
            }
            if ($charCount -eq 0 -and -not [string]::IsNullOrWhiteSpace($rawText)) {
                # line/word .Text reads came back empty but the aggregate
                # OcrResult.Text works: rebuild plain lines from it (no word
                # boxes, so spacing rebuild is skipped downstream).
                $lines = @()
                foreach ($t in ($rawText -split "`r?`n")) {
                    $lines += [pscustomobject]@{ Text = [string]$t; Words = @() }
                }
            }
            return [pscustomobject]@{
                Path         = $full
                LanguageTag  = [string]$engine.RecognizerLanguage.LanguageTag
                Lines        = @($lines)
                Text         = (@($lines | ForEach-Object { $_.Text }) -join "`r`n")
                RawText      = $rawText
                LineTypeName = $lineTypeName
                TextStrategy = [string]$script:WinOcrTextStrategy
            }
        } finally {
            try { $bitmap.Dispose() } catch {}
        }
    } finally {
        try { $stream.Dispose() } catch {}
    }
}
