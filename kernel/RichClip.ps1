#Requires -Version 5.1
# ============================================================
#  kernel/RichClip.ps1
#
#  PURE builders for a "rich" clipboard: text and pictures in ONE paste,
#  the way copying a mixed selection out of Word lands in a chat box.
#  Three representations go on the clipboard together (Native.ps1
#  Set-EbiClipboardRich puts them there); the target picks the best it
#  understands:
#    HTML Format  CF_HTML, pictures inline as data: URIs
#    Rich Text    RTF, pictures as \pngblip
#    Unicode text the text alone (a target that takes no pictures still
#                 gets the words)
#  Dot-source only (no param(), ASCII source). No Windows type is named
#  here, so all of it runs (and is unit-tested) on Linux.
#
#    Get-EbiPngInfo      PNG bytes -> @{ ok; width; height }
#    ConvertTo-EbiHtmlText  HTML-escape one line
#    New-EbiShareHtml    lines + pictures -> an HTML fragment
#    New-EbiCfHtml       fragment -> the CF_HTML string (byte offsets in
#                        UTF-8, as the format demands)
#    ConvertTo-EbiRtfText   one line -> RTF with \uN? escapes
#    New-EbiShareRtf     lines + pictures -> an RTF document
# ============================================================

function Get-EbiPngInfo {
    # PURE. Width / height from the IHDR chunk; ok=$false when not a PNG.
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -lt 24) { return @{ ok = $false; width = 0; height = 0 } }
    $sig = @(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)
    for ($i = 0; $i -lt 8; $i++) { if ([int]$Bytes[$i] -ne $sig[$i]) { return @{ ok = $false; width = 0; height = 0 } } }
    $w = ([int]$Bytes[16] -shl 24) -bor ([int]$Bytes[17] -shl 16) -bor ([int]$Bytes[18] -shl 8) -bor [int]$Bytes[19]
    $h = ([int]$Bytes[20] -shl 24) -bor ([int]$Bytes[21] -shl 16) -bor ([int]$Bytes[22] -shl 8) -bor [int]$Bytes[23]
    return @{ ok = $true; width = $w; height = $h }
}

function ConvertTo-EbiHtmlText {
    # PURE. Escape & < > " for HTML text.
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function New-EbiShareHtml {
    <#
      PURE. An HTML fragment: each text line a <div>, then each picture an
      <img> with a data: URI, one per line, in the order given. Pictures
      are @{ bytes; width; height } (width/height in CSS px; 0 = natural).
    #>
    param([string[]]$Lines = @(), [object[]]$Pictures = @())
    $sb = New-Object System.Text.StringBuilder
    foreach ($l in @($Lines)) {
        if ([string]::IsNullOrEmpty([string]$l)) { [void]$sb.Append('<div><br></div>') }
        else { [void]$sb.Append('<div>' + (ConvertTo-EbiHtmlText -Text ([string]$l)) + '</div>') }
    }
    foreach ($p in @($Pictures)) {
        if ($null -eq $p) { continue }
        $b64 = [Convert]::ToBase64String([byte[]]$p['bytes'])
        $size = ''
        if ([int]$p['width'] -gt 0 -and [int]$p['height'] -gt 0) { $size = (' width="' + [int]$p['width'] + '" height="' + [int]$p['height'] + '"') }
        [void]$sb.Append('<div><img src="data:image/png;base64,' + $b64 + '"' + $size + '></div>')
    }
    return $sb.ToString()
}

function New-EbiCfHtml {
    <#
      PURE. Wrap an HTML fragment in the CF_HTML header. StartHTML /
      EndHTML / StartFragment / EndFragment are BYTE offsets into the UTF-8
      encoding of the whole string -- character offsets break the moment
      the text holds Japanese. Offsets are written zero-padded to 10 digits
      so the header length does not depend on the numbers.
    #>
    param([string]$Fragment)
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $headerTemplate = "Version:0.9`r`nStartHTML:{0:D10}`r`nEndHTML:{1:D10}`r`nStartFragment:{2:D10}`r`nEndFragment:{3:D10}`r`n"
    $pre = '<html><head><meta charset="utf-8"></head><body><!--StartFragment-->'
    $post = '<!--EndFragment--></body></html>'
    $headerLen = $utf8.GetByteCount(($headerTemplate -f 0, 0, 0, 0))
    $startHtml = $headerLen
    $startFrag = $startHtml + $utf8.GetByteCount($pre)
    $endFrag = $startFrag + $utf8.GetByteCount([string]$Fragment)
    $endHtml = $endFrag + $utf8.GetByteCount($post)
    return (($headerTemplate -f $startHtml, $endHtml, $startFrag, $endFrag) + $pre + $Fragment + $post)
}

function ConvertTo-EbiRtfText {
    # PURE. One line of text as RTF: \ { } escaped, anything past ASCII as
    # \uN? with N the signed 16-bit code unit (RTF's rule).
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        $c = [int]$ch
        if ($c -eq 0x5C) { [void]$sb.Append('\\') }
        elseif ($c -eq 0x7B) { [void]$sb.Append('\{') }
        elseif ($c -eq 0x7D) { [void]$sb.Append('\}') }
        elseif ($c -ge 0x20 -and $c -lt 0x7F) { [void]$sb.Append($ch) }
        else {
            $n = if ($c -gt 32767) { $c - 65536 } else { $c }
            [void]$sb.Append('\u' + $n + '?')
        }
    }
    return $sb.ToString()
}

function New-EbiShareRtf {
    <#
      PURE. An RTF document: the lines as paragraphs, then each picture as
      {\pict\pngblip ...} in its own paragraph. picwgoal / pichgoal are in
      twips at 96 dpi (15 twips per pixel).
    #>
    param([string[]]$Lines = @(), [object[]]$Pictures = @())
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('{\rtf1\ansi\ansicpg932\deff0{\fonttbl{\f0\fnil Meiryo UI;}}\f0\fs20 ')
    foreach ($l in @($Lines)) { [void]$sb.Append((ConvertTo-EbiRtfText -Text ([string]$l)) + '\par ') }
    foreach ($p in @($Pictures)) {
        if ($null -eq $p) { continue }
        $bytes = [byte[]]$p['bytes']
        $w = [int]$p['width']; $h = [int]$p['height']
        $hex = New-Object System.Text.StringBuilder ($bytes.Length * 2)
        foreach ($b in $bytes) { [void]$hex.Append(([int]$b).ToString('x2')) }
        [void]$sb.Append('{\pict\pngblip\picw' + $w + '\pich' + $h + '\picwgoal' + ($w * 15) + '\pichgoal' + ($h * 15) + ' ' + $hex.ToString() + '}\par ')
    }
    [void]$sb.Append('}')
    return $sb.ToString()
}
