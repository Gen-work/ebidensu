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
    return ,@($tags)
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

# OCRs one image file. Returns:
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

    $full = (Resolve-Path -LiteralPath $Path).Path
    $file = Wait-WinOcrOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($full)) ([Windows.Storage.StorageFile])
    $stream = Wait-WinOcrOperation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder = Wait-WinOcrOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Wait-WinOcrOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        try {
            $ocr = Wait-WinOcrOperation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
            $lines = @()
            foreach ($ln in $ocr.Lines) {
                $words = @()
                foreach ($w in $ln.Words) {
                    $r = $w.BoundingRect
                    $words += [pscustomobject]@{
                        Text   = [string]$w.Text
                        X      = [double]$r.X
                        Y      = [double]$r.Y
                        Width  = [double]$r.Width
                        Height = [double]$r.Height
                    }
                }
                $lines += [pscustomobject]@{ Text = [string]$ln.Text; Words = @($words) }
            }
            return [pscustomobject]@{
                Path        = $full
                LanguageTag = [string]$engine.RecognizerLanguage.LanguageTag
                Lines       = @($lines)
                Text        = (@($lines | ForEach-Object { $_.Text }) -join "`r`n")
            }
        } finally {
            try { $bitmap.Dispose() } catch {}
        }
    } finally {
        try { $stream.Dispose() } catch {}
    }
}
