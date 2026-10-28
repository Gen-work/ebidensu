# modules/verify/verify.book_review.ps1
# The pre-delivery self-review of one evidence workbook, read straight from
# the .xlsx (no Excel, no Python): the non-OCR checks of
# tools/review/check_book.py, ported 10/28 because the office PC has no
# Python. NG lines come back as `alert` for the gate; the two OCR checks
# (K-50 GIFT list box, K-53 DF file names) stay in the cloud review.
# kernel/BookReview.ps1 holds the checks; this step reads the book and
# turns the findings into a verdict.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')       # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\BookReview.ps1')   # + kernel\Xlsx.ps1

$Manifest = @{
  id         = 'verify.book_review'
  group      = 'verify'
  summary    = 'Review an evidence .xlsx without Excel: view, fonts, logs, yellow, labels, job logs, boxes'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path    = @{ type='path';   required=$true; desc='the evidence workbook (.xlsx); read with a shared handle, never written' }
    sheets  = @{ type='map';    required=$true; desc='sheet names: sendData, giftResult, gfixResult, compare (profile.layout.evidence.sheets)' }
    font    = @{ type='string'; required=$true; desc='the one font every cell with a value must use (profile.layout.evidence.font.name)' }
    jobLog  = @{ type='string'; required=$true; desc='the GFIX job-log label on the GFIX receive sheet (profile.layout.evidence.labels.jobLog)' }
    maxAlert = @{ type='int';   default=8; desc='how many NG lines go to alert (the rest are counted)' }
  }
  outputs    = @{
    code     = @{ type='string'; desc='ok (no NG; WARN lines may be there) | ng | unknown (dry run, no book)' }
    ng       = @{ type='int';    desc='how many NG findings' }
    warn     = @{ type='int';    desc='how many WARN findings' }
    findings = @{ type='list';   desc='@{ level = NG|WARN; where; message; ref } in check order' }
    alert    = @{ type='list';   desc='the NG lines ("where: message (ref)"), at most maxAlert, then one "... and N more"' }
    reason   = @{ type='string'; desc='one line for the gate' }
    path     = @{ type='path' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $false }
    @{ id = 'read_failed';    transient = $true }
    @{ id = 'input_invalid';  transient = $false }
  )
  example    = @{ use = 'verify.book_review'; with = @{ path = '{{steps.book.out.path}}'; sheets = '{{profile.layout.evidence.sheets}}'; font = '{{profile.layout.evidence.font.name}}'; jobLog = '{{profile.layout.evidence.labels.jobLog}}' } }
  notes      = 'Read-only: the .xlsx is opened as a zip with a shared read handle, so a book open in Excel can be reviewed (what Excel has not saved yet is not seen). Pictures are not read: the OCR checks are tools/review/check_book.py''s. A WARN never makes the code ng.'
}

function VerifyBookReview-Line {
    # PURE. A finding -> one line.
    param($Finding)
    return ([string]$Finding['where'] + ': ' + [string]$Finding['message'] + $(if ([string]$Finding['ref'] -ne '') { ' (' + [string]$Finding['ref'] + ')' } else { '' }))
}

function VerifyBookReview-Verdict {
    # PURE. Findings -> @{ code; ng; warn; alert; reason }.
    param($Findings, [int]$MaxAlert, [string]$Name)
    $ngs = @($Findings | Where-Object { $_['level'] -eq 'NG' })
    $warns = @($Findings | Where-Object { $_['level'] -ne 'NG' })
    $alert = New-Object System.Collections.ArrayList
    foreach ($f in ($ngs | Select-Object -First ([math]::Max(0, $MaxAlert)))) { [void]$alert.Add((VerifyBookReview-Line -Finding $f)) }
    if ($ngs.Count -gt $MaxAlert) { [void]$alert.Add(('... and ' + ($ngs.Count - $MaxAlert) + ' more NG')) }
    $reason = $Name + ': ' + $(if ($ngs.Count -eq 0) { 'no NG' } else { [string]$ngs.Count + ' NG' }) + $(if ($warns.Count -gt 0) { ', ' + $warns.Count + ' WARN (' + ((@($warns | Select-Object -First 3 | ForEach-Object { VerifyBookReview-Line -Finding $_ })) -join ' / ') + ')' } else { '' }) + ' -- the OCR checks (GIFT list box, DF file names) are the cloud review''s'
    return @{ code = $(if ($ngs.Count -gt 0) { 'ng' } else { 'ok' }); ng = $ngs.Count; warn = $warns.Count; alert = $alert.ToArray(); reason = $reason }
}

function Invoke-Step {
    param($In, $Ctx)
    $src = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    $empty = @{ code = 'unknown'; ng = 0; warn = 0; findings = @(); alert = @(); reason = ''; path = $src }
    if ($Ctx['DryRun'] -and -not [System.IO.File]::Exists($src)) {
        $Ctx.Log.Info(('would review ' + $src + ' (dry run: no such book here)'))
        return @{ ok = $true } + $empty
    }
    $sh = $In['sheets']
    $names = @{ font = [string]$In['font']; jobLog = [string]$In['jobLog'] }
    foreach ($pair in @(@('send', 'sendData'), @('giftRecv', 'giftResult'), @('gfixRecv', 'gfixResult'), @('compare', 'compare'))) {
        if (-not ($sh -is [System.Collections.IDictionary]) -or -not $sh.Contains($pair[1]) -or [string]::IsNullOrWhiteSpace([string]$sh[$pair[1]])) {
            return @{ ok = $false; failure = 'input_invalid'; message = ('sheets has no "' + $pair[1] + '" (profile.layout.evidence.sheets)') } + $empty
        }
        $names[$pair[0]] = [string]$sh[$pair[1]]
    }
    $book = Read-EbiXlsx -Path $src
    if (-not $book['ok']) { return @{ ok = $false; failure = [string]$book['failure']; message = [string]$book['message'] } + $empty }
    try {
        $model = New-EbiBookModel -Book $book -Names $names
        $findings = @(Invoke-EbiBookReview -Model $model -Names $names)
    } catch {
        return @{ ok = $false; failure = 'read_failed'; message = ('could not review ' + $src + ': ' + $_.Exception.Message) } + $empty
    }
    $v = VerifyBookReview-Verdict -Findings $findings -MaxAlert ([int]$In['maxAlert']) -Name ([System.IO.Path]::GetFileName($src))
    foreach ($f in $findings) { if ($f['level'] -eq 'NG') { $Ctx.Log.Warn((VerifyBookReview-Line -Finding $f)) } else { $Ctx.Log.Info((VerifyBookReview-Line -Finding $f)) } }
    return @{ ok = $true; code = $v['code']; ng = $v['ng']; warn = $v['warn']; findings = $findings; alert = @($v['alert']); reason = $v['reason']; path = $src }
}
