# modules/human/human.draft_mail.ps1
# One Outlook DRAFT from a template (kernel/Mail.ps1), saved and opened for
# the operator to check and send. It NEVER sends: the person presses Send
# in Outlook, then answers the gate that follows (10/15: the J4 review
# request, one mail per deliverable -- what sendj4mail.bat did by hand).
# Needs Classic Outlook (the new Outlook has no automation).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Mail.ps1')

$Manifest = @{
  id         = 'human.draft_mail'
  group      = 'human'
  summary    = 'Make an Outlook draft from a template (token -> value), save and open it; never sends'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $false
  inputs     = @{
    template   = @{ type='path';   required=$true; desc='.json, .bat/.cmd (a #==TEMPLATE_JSON== section) or .msg/.oft; relative paths resolve under the work dir' }
    value      = @{ type='string'; required=$true; desc='what the token becomes, e.g. {{item.Excel_NAME}}' }
    token      = @{ type='string'; default=''; desc='the word in the template to replace; empty = the template''s own token (a .msg/.oft has none: give it)' }
    to         = @{ type='list';   default=@(); desc='addresses that replace the template''s To; empty = keep the template''s' }
    cc         = @{ type='list';   default=@(); desc='addresses that replace the template''s CC; empty = keep the template''s' }
    trimBefore = @{ type='bool';   default=$true; desc='drop spaces between "_" or ">" and the token (a template whose file name read "_ JOB.xlsx")' }
    display    = @{ type='bool';   default=$true; desc='open the draft window for the operator' }
  }
  outputs    = @{
    subject = @{ type='string'; desc='the draft''s subject' }
    to      = @{ type='list';   desc='To addresses' }
    cc      = @{ type='list';   desc='CC addresses' }
    saved   = @{ type='bool';   desc='the draft is in the Drafts folder' }
  }
  failures   = @(
    @{ id = 'file_not_found';      transient = $false }
    @{ id = 'template_invalid';    transient = $false }
    @{ id = 'outlook_unavailable'; transient = $false }
    @{ id = 'draft_failed';        transient = $true  }
  )
  example    = @{ use = 'human.draft_mail'; with = @{ template = '{{profile.paths.mailTemplate}}'; value = '{{item.Excel_NAME}}' } }
  notes      = 'Not idempotent: every run makes one more draft (a rerun after a failure leaves the earlier draft in Drafts -- delete it there). The template file stays on the PC (addresses and names are not committed); profiles/gfix-recv/README.md says where.'
}

function HumanDraftMail-Recipients {
    # PURE. The template's recipients with To / CC replaced when given.
    param($Recipients, $To, $Cc)
    $to2 = @(@($To) | Where-Object { $null -ne $_ -and ([string]$_).Trim() -ne '' } | ForEach-Object { ([string]$_).Trim() })
    $cc2 = @(@($Cc) | Where-Object { $null -ne $_ -and ([string]$_).Trim() -ne '' } | ForEach-Object { ([string]$_).Trim() })
    $out = New-Object System.Collections.ArrayList
    foreach ($r in @($Recipients)) {
        if ($null -eq $r) { continue }
        $t = [int]$r['type']
        if ($t -eq 1 -and $to2.Count -gt 0) { continue }
        if ($t -eq 2 -and $cc2.Count -gt 0) { continue }
        [void]$out.Add(@{ address = [string]$r['address']; type = $t })
    }
    foreach ($a in $to2) { [void]$out.Add(@{ address = $a; type = 1 }) }
    foreach ($a in $cc2) { [void]$out.Add(@{ address = $a; type = 2 }) }
    return $out.ToArray()
}

function HumanDraftMail-ReadOutlookTemplate {
    # A .msg / .oft through Outlook -> the same template shape. Windows only.
    param($Outlook, [string]$Path)
    $item = $Outlook.CreateItemFromTemplate($Path)
    try {
        $rec = New-Object System.Collections.ArrayList
        foreach ($r in $item.Recipients) { [void]$rec.Add(@{ name = [string]$r.Name; address = [string]$r.Address; type = [int]$r.Type }) }
        return @{ token = ''; subject = [string]$item.Subject; body = [string]$item.Body; html = [string]$item.HTMLBody; recipients = $rec.ToArray() }
    } finally { try { $item.Close(1) } catch { } }
}

function Invoke-Step {
    param($In, $Ctx)
    $path = Resolve-EbiWorkPath -PathValue ([string]$In['template']) -WorkDir ([string]$Ctx['WorkDir'])
    $none = @{ ok = $true; subject = ''; to = @(); cc = @(); saved = $false }
    $kind = Get-EbiMailTemplateKind -Path $path
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        if ($Ctx['DryRun']) { $Ctx.Log.Info(('would make a draft from {0} (not there yet)' -f $path)); return $none }
        return @{ ok = $false; failure = 'file_not_found'; message = ('no mail template at ' + $path + ' (set paths.mailTemplate in ebi.local.json)'); subject = ''; to = @(); cc = @(); saved = $false }
    }
    $tpl = $null
    if ($kind -ne 'outlook') {
        $r = Read-EbiMailTemplate -Path $path
        if (-not $r['ok']) { return @{ ok = $false; failure = 'template_invalid'; message = $r['message']; subject = ''; to = @(); cc = @(); saved = $false } }
        $tpl = $r['template']
    }
    if ($Ctx['DryRun'] -and $null -ne $tpl) {
        $x = Expand-EbiMailTemplate -Template $tpl -Token ([string]$In['token']) -Value ([string]$In['value']) -TrimBefore:([bool]$In['trimBefore'])
        $rcp = @(HumanDraftMail-Recipients -Recipients $tpl['recipients'] -To $In['to'] -Cc $In['cc'])
        $Ctx.Log.Info(('would make a draft: {0}' -f $x['subject']))
        return @{ ok = $true; subject = [string]$x['subject']; to = @($rcp | Where-Object { $_['type'] -eq 1 } | ForEach-Object { $_['address'] }); cc = @($rcp | Where-Object { $_['type'] -eq 2 } | ForEach-Object { $_['address'] }); saved = $false }
    }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would make a draft from {0}' -f $path)); return $none }
    $ol = $null
    try { $ol = New-Object -ComObject Outlook.Application } catch { return @{ ok = $false; failure = 'outlook_unavailable'; message = 'Classic Outlook could not be started (the new Outlook cannot be automated): install / switch to Classic Outlook'; subject = ''; to = @(); cc = @(); saved = $false } }
    if ($null -eq $tpl) {
        try { $tpl = HumanDraftMail-ReadOutlookTemplate -Outlook $ol -Path $path } catch { return @{ ok = $false; failure = 'template_invalid'; message = ('Outlook could not read ' + $path + ': ' + $_.Exception.Message); subject = ''; to = @(); cc = @(); saved = $false } }
    }
    $x = Expand-EbiMailTemplate -Template $tpl -Token ([string]$In['token']) -Value ([string]$In['value']) -TrimBefore:([bool]$In['trimBefore'])
    if (-not $x['ok']) { return @{ ok = $false; failure = 'template_invalid'; message = $x['message']; subject = ''; to = @(); cc = @(); saved = $false } }
    $rcp = @(HumanDraftMail-Recipients -Recipients $tpl['recipients'] -To $In['to'] -Cc $In['cc'])
    if (@($rcp | Where-Object { $_['type'] -eq 1 }).Count -eq 0) { return @{ ok = $false; failure = 'template_invalid'; message = 'the draft would have no To address'; subject = ''; to = @(); cc = @(); saved = $false } }
    $mail = $null; $saved = $false
    try {
        $mail = $ol.CreateItem(0)
        foreach ($r in $rcp) { $a = $mail.Recipients.Add([string]$r['address']); $a.Type = [int]$r['type'] }
        [void]$mail.Recipients.ResolveAll()
        $mail.Subject = [string]$x['subject']
        if ([string]$x['html'] -ne '') { $mail.BodyFormat = 2; $mail.HTMLBody = [string]$x['html'] } else { $mail.Body = [string]$x['body'] }
        $mail.Save(); $saved = $true                       # Drafts; Send is the operator's
        if ([bool]$In['display']) { $mail.Display($false) }
    } catch {
        return @{ ok = $false; failure = 'draft_failed'; message = $_.Exception.Message; subject = [string]$x['subject']; to = @(); cc = @(); saved = $saved }
    }
    $Ctx.Log.Info(('draft saved{0}: {1} (nothing sent)' -f $(if ([bool]$In['display']) { ' and opened' } else { '' }), $x['subject']))
    return @{ ok = $true; subject = [string]$x['subject']; to = @($rcp | Where-Object { $_['type'] -eq 1 } | ForEach-Object { $_['address'] }); cc = @($rcp | Where-Object { $_['type'] -eq 2 } | ForEach-Object { $_['address'] }); saved = $saved }
}
