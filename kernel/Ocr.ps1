#Requires -Version 5.1
# ============================================================
#  kernel/Ocr.ps1
#
#  The ONE Windows OCR binding for ebi-dance steps (Windows.Media.Ocr, the
#  engine behind the Snipping Tool's text extraction; nothing to install).
#  A slim port of the repo-root OcrWindows.ps1 (which the old phases keep
#  using): lazy WinRT load, text read through the adapter / psbase /
#  reflection, the whole result's Text as the fallback.
#
#  Dot-source only (no param() block, ASCII source). Dot-sourcing never
#  throws on Linux: every function that names a WinRT type is a *Core
#  function called only after the Windows check (a PowerShell function
#  that names a type the host cannot load throws on its FIRST call,
#  before running a line -- the same split kernel/Image.ps1 makes for
#  System.Drawing).
#
#  Entry point:
#    Invoke-EbiOcrFile -Path <png> [-Language en-US]
#      -> @{ ok; text; lines = string[]; language; message }
#  ok = $false (never a throw) when OCR is not available on this host,
#  the file is missing, or no recognizer language is installed.
# ============================================================

$script:EbiOcrReady = $false
$script:EbiOcrInitError = ''
$script:EbiOcrAsTask = $null

function Test-EbiOcrHost {
    # PURE-ish: is this a host where Windows OCR can exist at all?
    return ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT)
}

function Initialize-EbiOcrCore {
    if ($script:EbiOcrReady) { return $true }
    if ($script:EbiOcrInitError -ne '') { return $false }
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
        $null = [Windows.Media.Ocr.OcrEngine,Windows.Foundation.UniversalApiContract,ContentType=WindowsRuntime]
        $null = [Windows.Globalization.Language,Windows.Globalization,ContentType=WindowsRuntime]
        $null = [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
        $null = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics,ContentType=WindowsRuntime]
        $script:EbiOcrAsTask = @([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
                $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
        if ($null -eq $script:EbiOcrAsTask) { throw 'AsTask(IAsyncOperation) bridge not found' }
        $script:EbiOcrReady = $true
        return $true
    } catch {
        $script:EbiOcrInitError = $_.Exception.Message
        return $false
    }
}

function Wait-EbiOcrOperation {
    param($Operation, [type]$ResultType)
    $t = $script:EbiOcrAsTask.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    [void]$t.Wait(-1)
    return $t.Result
}

function Read-EbiOcrText {
    # A WinRT Text property: on some PS 5.1 hosts the adapter returns ''
    # for every HSTRING while collections still enumerate (OcrWindows.ps1).
    param($Object)
    if ($null -eq $Object) { return '' }
    $v = ''
    try { $v = [string]$Object.Text } catch { }
    if ($v.Length -gt 0) { return $v }
    try { $v = [string]$Object.psbase.Text } catch { }
    if ($v.Length -gt 0) { return $v }
    try { $pi = $Object.GetType().GetProperty('Text'); if ($null -ne $pi) { $v = [string]$pi.GetValue($Object, $null) } } catch { }
    return $v
}

function Invoke-EbiOcrFileCore {
    param([string]$Path, [string]$Language)
    $engine = $null
    foreach ($tag in @($Language, 'en-US', 'en')) {
        if ([string]::IsNullOrWhiteSpace($tag)) { continue }
        try { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage((New-Object Windows.Globalization.Language($tag))) } catch { $engine = $null }
        if ($null -ne $engine) { break }
    }
    if ($null -eq $engine) { try { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() } catch { $engine = $null } }
    if ($null -eq $engine) { return @{ ok = $false; text = ''; lines = @(); language = ''; message = 'no OCR recognizer language is installed' } }
    $full = (Resolve-Path -LiteralPath $Path).ProviderPath
    $file = Wait-EbiOcrOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($full)) ([Windows.Storage.StorageFile])
    $stream = Wait-EbiOcrOperation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    try {
        $decoder = Wait-EbiOcrOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Wait-EbiOcrOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        try {
            $res = Wait-EbiOcrOperation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
            $lines = New-Object System.Collections.ArrayList
            $boxes = New-Object System.Collections.ArrayList
            foreach ($ln in $res.Lines) {
                $t = Read-EbiOcrText $ln
                [void]$lines.Add($t)
                # where the line sits (top of its words), for callers that need rows
                $y = -1.0; $hh = 0.0
                try { foreach ($w in $ln.Words) { $r = $w.BoundingRect; if ($y -lt 0 -or [double]$r.Y -lt $y) { $y = [double]$r.Y }; if ([double]$r.Height -gt $hh) { $hh = [double]$r.Height } } } catch { }
                [void]$boxes.Add(@{ text = $t; y = $y; h = $hh })
            }
            $joined = (@($lines) -join "`n")
            if ($joined.Trim() -eq '') {
                $raw = Read-EbiOcrText $res
                $lines = New-Object System.Collections.ArrayList
                foreach ($t in ($raw -split "`r?`n")) { [void]$lines.Add([string]$t) }
                $joined = $raw
            }
            return @{ ok = $true; text = $joined; lines = $lines.ToArray(); boxes = $boxes.ToArray(); language = [string]$engine.RecognizerLanguage.LanguageTag; message = '' }
        } finally { try { $bitmap.Dispose() } catch { } }
    } finally { try { $stream.Dispose() } catch { } }
}

function Invoke-EbiOcrFile {
    <#
      OCR one image file. Returns @{ ok; text; lines; boxes (@{ text; y; h } per line, y = -1 when unknown); language; message };
      never throws. -Language: a BCP-47 tag tried first (default en-US:
      job names and file names are ASCII, and the en recognizer reads
      O/0, I/1 better than ja); then en, then the user-profile languages.
    #>
    param([string]$Path, [string]$Language = 'en-US')
    if (-not (Test-EbiOcrHost)) { return @{ ok = $false; text = ''; lines = @(); language = ''; message = 'Windows OCR is only available on Windows' } }
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{ ok = $false; text = ''; lines = @(); language = ''; message = ('no image at ' + $Path) } }
    if (-not (Initialize-EbiOcrCore)) { return @{ ok = $false; text = ''; lines = @(); language = ''; message = ('Windows OCR not available: ' + $script:EbiOcrInitError) } }
    try { return (Invoke-EbiOcrFileCore -Path $Path -Language $Language) }
    catch { return @{ ok = $false; text = ''; lines = @(); language = ''; message = $_.Exception.Message } }
}
