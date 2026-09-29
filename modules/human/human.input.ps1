# modules/human/human.input.ps1
# Ask the operator for one value, once per run, with a default and a check
# (P2-07). Ported from Resolve-ExpectedTime.ps1's batch prompt: Enter takes
# the default, r takes "recent" (now minus an hour), anything else is parsed
# in the declared kind. kind=timeWindow is the P0-R6 run.timeWindow: the
# answer is written into $Ctx.Run (the runner persists run/<runId>/run.json
# at once) and returned in outputs. When run.timeWindow is ALREADY set --
# the runner restored it on -Resume, or the CLI gave -TimeWindow -- the
# step keeps it and does not ask (kept=true): a resumed run never asks
# again, and never swaps the window its finished rows were judged by.
# persistTo names a worklist column to fill on rows whose cell is BLANK
# (never overwriting an operator's own value); the table is flushed.

. (Join-Path $PSScriptRoot '..\..\kernel\Gate.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Parse.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Table.ps1')

$Manifest = @{
  id         = 'human.input'
  group      = 'human'
  summary    = 'Ask the operator for a value (text, time, or the run time window) with a default'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    question  = @{ type='string'; required=$true }
    kind      = @{ type='string'; default='text'; enum=@('text', 'time', 'timeWindow') }
    default   = @{ type='string'; default=''; desc='what Enter means; for timeWindow "from..to", empty = the last hour' }
    persistTo = @{ type='string'; default=''; desc='worklist column to fill where blank (needs worklist)' }
    worklist  = @{ type='session'; sessionKind='worklist'; desc='only with persistTo' }
    format    = @{ type='string'; default='yyyy/MM/dd H:mm:ss'; desc='how a time answer is stored in the column' }
  }
  outputs    = @{
    value      = @{ type='string'; desc='the answer as text (a time in `format`; a window as "from..to")' }
    timeWindow = @{ type='map';    desc='{ from; to } ISO when kind is timeWindow, else null' }
    filled     = @{ type='int';    desc='rows whose blank cell was filled' }
    auto       = @{ type='bool';   desc='nobody could answer (dry run / no console): the default was taken' }
    kept       = @{ type='bool';   desc='run.timeWindow was already set (-TimeWindow or a resume): returned as is, nobody asked' }
  }
  failures   = @(
    @{ id = 'operator_quit'; transient = $false }
    @{ id = 'input_invalid'; transient = $false }
    @{ id = 'write_failed';  transient = $true  }
  )
  example    = @{ use = 'human.input'; with = @{ question = 'Batch run window for today?'; kind = 'timeWindow'; default = '' } }
  notes      = 'Runs in setup, once per run. The ledger does not replay setup, so on a resume the step runs again -- but run.json carries the window, the runner restores it into run.timeWindow first, and a kind=timeWindow question then returns it with kept=true instead of asking (the same when the CLI gave -TimeWindow). Under DryRun or without a console the default is taken and reported with auto=true.'
}

function HumanInput-Parse {
    <#
      PURE. Typed answer + kind -> @{ ok; value (text); timeWindow; message }.
        text        anything
        time        one time -> formatted with Format
        timeWindow  "from..to" (times as ConvertTo-EbiDateTime reads them; a
                    bare H:mm is today) -> ISO from / to
    #>
    param([string]$Answer, [string]$Kind, [string]$Format, [datetime]$Now = (Get-Date))
    $a = if ($null -eq $Answer) { '' } else { $Answer.Trim() }
    switch ($Kind) {
        'text' { return @{ ok = $true; value = $a; timeWindow = $null; message = '' } }
        'time' {
            $t = ConvertTo-EbiDateTime -Text $a -Date $Now.Date
            if (-not $t['ok']) { return @{ ok = $false; value = $a; timeWindow = $null; message = ('"' + $a + '" is not a time (yyyy/MM/dd H:mm:ss or H:mm)') } }
            return @{ ok = $true; value = $t['value'].ToString($Format); timeWindow = $null; message = '' }
        }
        'timeWindow' {
            $parts = @($a -split '\.\.')
            if ($parts.Count -ne 2) { return @{ ok = $false; value = $a; timeWindow = $null; message = ('"' + $a + '" is not "from..to"') } }
            $f = ConvertTo-EbiDateTime -Text $parts[0] -Date $Now.Date; $t = ConvertTo-EbiDateTime -Text $parts[1] -Date $Now.Date
            if (-not $f['ok'] -or -not $t['ok']) { return @{ ok = $false; value = $a; timeWindow = $null; message = ('cannot read "' + $(if (-not $f['ok']) { $parts[0] } else { $parts[1] }) + '" as a time') } }
            if ($t['value'] -lt $f['value']) { return @{ ok = $false; value = $a; timeWindow = $null; message = '"to" is before "from"' } }
            $w = @{ from = $f['value'].ToString('yyyy-MM-ddTHH:mm:ss'); to = $t['value'].ToString('yyyy-MM-ddTHH:mm:ss') }
            return @{ ok = $true; value = ($w['from'] + '..' + $w['to']); timeWindow = $w; message = '' }
        }
    }
    return @{ ok = $false; value = $a; timeWindow = $null; message = ('unknown kind ' + $Kind) }
}

function HumanInput-Default {
    # PURE. The text Enter stands for: the given default, else for a
    # timeWindow the last hour, for a time now minus an hour (the old
    # "recent"), for text ''.
    param([string]$Default, [string]$Kind, [datetime]$Now = (Get-Date))
    if (-not [string]::IsNullOrWhiteSpace($Default)) { return $Default }
    switch ($Kind) {
        'timeWindow' { return ($Now.AddHours(-1).ToString('yyyy/MM/dd H:mm:ss') + '..' + $Now.ToString('yyyy/MM/dd H:mm:ss')) }
        'time'       { return $Now.AddHours(-1).ToString('yyyy/MM/dd H:mm:ss') }
    }
    return ''
}

function Invoke-Step {
    param($In, $Ctx)
    $kind = [string]$In['kind']; $format = [string]$In['format']
    $now = Get-Date
    $default = HumanInput-Default -Default ([string]$In['default']) -Kind $kind -Now $now
    $hasDefault = -not [string]::IsNullOrWhiteSpace($default) -or $kind -eq 'text'
    $answer = ''; $auto = $false; $kept = $false
    if ($kind -eq 'timeWindow' -and $null -ne $Ctx['Run'] -and ($Ctx['Run'] -is [System.Collections.IDictionary]) -and $Ctx['Run'].Contains('timeWindow') -and ($Ctx['Run']['timeWindow'] -is [System.Collections.IDictionary]) -and $Ctx['Run']['timeWindow'].Count -gt 0) {
        # run.timeWindow is already set: -TimeWindow on the CLI, or a -Resume that
        # restored it from run.json. The question was answered once; asking again
        # would overwrite the window the finished rows were judged by (and with no
        # console it would silently become "the last hour"). Keep it, say so.
        $w = $Ctx['Run']['timeWindow']
        $p = @{ ok = $true; value = ([string]$w['from'] + '..' + [string]$w['to']); timeWindow = $w; message = '' }
        $kept = $true
        $Ctx.Log.Info('run.timeWindow already set (resume or -TimeWindow): kept ' + $p['value'] + ', not asked')
    }
    $tries = 0
    while (-not $kept) {
        $tries++
        $r = Show-EbiGate -Title ('INPUT ' + $kind) -What @([string]$In['question'], $(if ($default -ne '') { 'Enter = ' + $default } else { '' })) -Next @('type the value' + $(if ($kind -eq 'timeWindow') { ' as from..to (e.g. 9:00..12:00 for today)' } elseif ($kind -eq 'time') { ' (yyyy/MM/dd H:mm:ss or H:mm)' } else { '' }), 'q: cancel the whole run') -Actions @(@{ key = 'i'; label = 'value (type it, Enter = default)' }, @{ key = 'q'; label = 'quit' }) -Default 'i' -Auto 'i' -DryRun ([bool]$Ctx['DryRun']) -Reader $(if ($null -ne $Ctx['Reader']) { $Ctx['Reader'] } else { $null }) -Raw
        if ($r['auto']) { $auto = $true; $answer = $default; break }
        if ([string]$r['action'] -eq 'q') { return @{ ok = $false; failure = 'operator_quit'; message = 'operator answered q at the input'; value = ''; timeWindow = $null; filled = 0; auto = $false; kept = $false } }
        $answer = if ([string]$r['note'] -ne '') { [string]$r['note'] } else { $default }
        $p = HumanInput-Parse -Answer $answer -Kind $kind -Format $format -Now $now
        if ($p['ok']) { break }
        Write-Host ('  ' + $p['message']) -ForegroundColor DarkYellow
        if ($tries -ge 5) { return @{ ok = $false; failure = 'input_invalid'; message = $p['message']; value = $answer; timeWindow = $null; filled = 0; auto = $false; kept = $false } }
    }
    if (-not $kept) {
        $p = HumanInput-Parse -Answer $answer -Kind $kind -Format $format -Now $now
        if (-not $p['ok']) { return @{ ok = $false; failure = 'input_invalid'; message = $p['message']; value = $answer; timeWindow = $null; filled = 0; auto = $auto; kept = $false } }
        if ($auto) { $Ctx.Log.Info(('nobody to ask (dry run or no console): took ' + $(if ($p['value'] -ne '') { $p['value'] } else { '(empty)' }))) }
    }
    if ($kind -eq 'timeWindow' -and $null -ne $Ctx['Run'] -and ($Ctx['Run'] -is [System.Collections.IDictionary])) { $Ctx['Run']['timeWindow'] = $p['timeWindow'] }
    $filled = 0
    $col = [string]$In['persistTo']
    if ($col -ne '' -and $In.Contains('worklist') -and $null -ne $In['worklist']) {
        $wl = $In['worklist']
        if (-not (@($wl['columns']) -contains $col)) { $wl['columns'] = @(@($wl['columns']) + @($col)) }
        foreach ($row in @($wl['rows'])) { if ($null -eq $row) { continue }; $cur = if ($row.Contains($col) -and $null -ne $row[$col]) { ([string]$row[$col]).Trim() } else { '' }; if ($cur -eq '') { $row[$col] = $p['value']; $filled++ } }
        if ($filled -gt 0 -and -not $Ctx['DryRun'] -and -not [string]::IsNullOrWhiteSpace([string]$wl['path'])) {
            $s = Save-EbiWorklist -Worklist $wl
            if (-not $s['ok']) { return @{ ok = $false; failure = 'write_failed'; message = $s['message']; value = $p['value']; timeWindow = $p['timeWindow']; filled = $filled; auto = $auto; kept = $kept } }
        }
        if ($Ctx['DryRun'] -and $filled -gt 0) { $Ctx.Log.Info(('would fill ' + $filled + ' blank ' + $col + ' cell(s)')) }
    }
    return @{ ok = $true; value = $p['value']; timeWindow = $p['timeWindow']; filled = $filled; auto = $auto; kept = $kept }
}
