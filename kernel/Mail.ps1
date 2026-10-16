#Requires -Version 5.1
# ============================================================
#  kernel/Mail.ps1
#
#  Mail templates for human.draft_mail (10/15: the review-request mail,
#  one per deliverable). A template is a sent mail with ONE word in it --
#  the token, e.g. the job name of the mail it was copied from -- that is
#  replaced by the current item's value in the subject, the text body and
#  the HTML body. Recipients come with it.
#
#  Template files this reads (PURE, no Outlook):
#    .json      { token, subject, body, html, recipients: [ { name,
#               address, type } ] }   (type: 1 To, 2 CC, 3 BCC)
#    .bat/.cmd  the same JSON, base64, after a line "#==TEMPLATE_JSON=="
#               (the operator's sendj4mail.bat carries one)
#  A .msg / .oft (a mail saved from Outlook) is read by the step through
#  Outlook itself (CreateItemFromTemplate); its token comes from the
#  workflow.
#
#  Dot-source only (no param() block, ASCII source). JSON goes through
#  kernel/Json.ps1 (R8).
# ============================================================

. (Join-Path $PSScriptRoot 'Json.ps1')

function Get-EbiMailTemplateKind {
    # PURE. 'json' | 'embedded' | 'outlook' | '' by extension.
    param([string]$Path)
    switch -Regex ([System.IO.Path]::GetExtension([string]$Path).ToLowerInvariant()) {
        '^\.json$'        { return 'json' }
        '^\.(bat|cmd)$'   { return 'embedded' }
        '^\.(msg|oft)$'   { return 'outlook' }
        default           { return '' }
    }
}

function ConvertTo-EbiMailTemplate {
    # PURE. A parsed JSON value -> @{ ok; template; message } with every
    # field a string and recipients a list of @{ name; address; type }.
    param($Value)
    if (-not ($Value -is [System.Collections.IDictionary])) { return @{ ok = $false; template = $null; message = 'the template is not a JSON object' } }
    $t = @{ token = ''; subject = ''; body = ''; html = ''; recipients = @() }
    foreach ($k in @('token', 'subject', 'body', 'html')) { if ($Value.Contains($k) -and $null -ne $Value[$k]) { $t[$k] = [string]$Value[$k] } }
    $list = New-Object System.Collections.ArrayList
    foreach ($r in @($Value['recipients'])) {
        if (-not ($r -is [System.Collections.IDictionary])) { continue }
        $addr = ([string]$r['address']).Trim()
        if ($addr -eq '') { return @{ ok = $false; template = $null; message = 'a recipient has no address' } }
        $type = 1; if ($r.Contains('type') -and $null -ne $r['type']) { try { $type = [int]$r['type'] } catch { $type = 1 } }
        [void]$list.Add(@{ name = [string]$r['name']; address = $addr; type = $type })
    }
    $t['recipients'] = $list.ToArray()
    if ($t['subject'] -eq '' -and $t['html'] -eq '' -and $t['body'] -eq '') { return @{ ok = $false; template = $null; message = 'the template has no subject and no body' } }
    return @{ ok = $true; template = $t; message = '' }
}

function Read-EbiMailTemplate {
    <#
      A .json or .bat/.cmd template file -> @{ ok; template; kind; message }.
      kind 'outlook' (a .msg / .oft) comes back ok = $false with kind set:
      the step reads those through Outlook.
    #>
    param([string]$Path)
    $kind = Get-EbiMailTemplateKind -Path $Path
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{ ok = $false; template = $null; kind = $kind; message = ('no template at ' + $Path) } }
    if ($kind -eq 'json') {
        $r = Read-EbiJson -Path $Path
        if (-not $r['ok']) {
            # 10/16: the operator pasted sendj4mail.bat's #==TEMPLATE_JSON==
            # body (base64) into review-mail.json -> "invalid JSON primitive:
            # eyJ0b2tl...". That body IS the template: read it as such.
            $raw = ''
            try { $raw = ([System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding($false)))) -replace '\s', '' } catch { }
            $raw = $raw.TrimStart([char]0xFEFF)
            if ($raw -match '^[A-Za-z0-9+/]+={0,2}$') {
                $json = ''
                try { $json = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($raw)) } catch { $json = '' }
                if ($json -ne '') {
                    $r2 = ConvertFrom-EbiJson -Text $json
                    if ($r2['ok']) {
                        $c = ConvertTo-EbiMailTemplate -Value $r2['value']
                        return @{ ok = $c['ok']; template = $c['template']; kind = $kind; message = $c['message'] }
                    }
                }
                return @{ ok = $false; template = $null; kind = $kind; message = ($Path + ' holds base64, not JSON, and it does not decode to a template: point paths.mailTemplate at sendj4mail.bat itself') }
            }
            return @{ ok = $false; template = $null; kind = $kind; message = $r['message'] }
        }
        $c = ConvertTo-EbiMailTemplate -Value $r['value']
        return @{ ok = $c['ok']; template = $c['template']; kind = $kind; message = $c['message'] }
    }
    if ($kind -eq 'embedded') {
        $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding($false)))
        $parts = [regex]::Split($text, '(?m)^#==TEMPLATE_JSON==\r?$', 2)
        if ($parts.Count -ne 2) { return @{ ok = $false; template = $null; kind = $kind; message = 'no "#==TEMPLATE_JSON==" line in the file' } }
        $json = ''
        try { $json = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String(($parts[1] -replace '\s', ''))) } catch { return @{ ok = $false; template = $null; kind = $kind; message = ('the embedded template is not base64: ' + $_.Exception.Message) } }
        $r = ConvertFrom-EbiJson -Text $json
        if (-not $r['ok']) { return @{ ok = $false; template = $null; kind = $kind; message = $r['message'] } }
        $c = ConvertTo-EbiMailTemplate -Value $r['value']
        return @{ ok = $c['ok']; template = $c['template']; kind = $kind; message = $c['message'] }
    }
    if ($kind -eq 'outlook') { return @{ ok = $false; template = $null; kind = $kind; message = 'a .msg / .oft is read through Outlook' } }
    return @{ ok = $false; template = $null; kind = ''; message = ('not a template file (.json, .bat, .cmd, .msg, .oft): ' + $Path) }
}

function Expand-EbiMailTemplate {
    <#
      PURE. Replace the token by the value in subject / body / html.
      Returns @{ ok; subject; body; html; message }. ok = $false when the
      token is empty or appears in none of them (the wrong template, or
      the token mistyped: a mail that would go out unchanged).
      -TrimBefore drops spaces / &nbsp; between "_" or ">" and the token
      (the 10/15 template's file name read "_ MJDSWM39.xlsx": the token
      sat in its own <span> after a space).
    #>
    param([hashtable]$Template, [string]$Token = '', [string]$Value, [switch]$TrimBefore)
    $tok = if ($Token -ne '') { $Token } else { [string]$Template['token'] }
    if ($tok -eq '') { return @{ ok = $false; subject = ''; body = ''; html = ''; message = 'no token: the template names none and the workflow gave none' } }
    $hits = 0
    $out = @{}
    foreach ($k in @('subject', 'body', 'html')) {
        $s = [string]$Template[$k]
        if ($s.Contains($tok)) { $hits++ }
        if ($TrimBefore) { $s = [regex]::Replace($s, '(?<=[_>])(?:[ \t]|&nbsp;)+(?=' + [regex]::Escape($tok) + ')', '') }
        $out[$k] = $s.Replace($tok, $Value)
    }
    if ($hits -eq 0) { return @{ ok = $false; subject = ''; body = ''; html = ''; message = ('the token "' + $tok + '" is in neither the subject nor the body') } }
    return @{ ok = $true; subject = $out['subject']; body = $out['body']; html = $out['html']; message = '' }
}
