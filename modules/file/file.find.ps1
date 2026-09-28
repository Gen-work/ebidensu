# modules/file/file.find.ps1
# Find the file for a key (or a glob) in a folder, tolerating what the old
# tool learned the hard way (P1-22): full-width ASCII in the file name
# (WorkbookResolver.ps1's FullWidthFilenameResolver), a key saved under its
# stamped batch-run name ABC123.260824.10515511.dat while the worklist says
# ABC123 (MappingStore.ps1 Resolve-CorrelFilePath), and the reverse.
#
# Normalization is kernel/Key.ps1's (P1-27): the stem is handed to
# Get-EbiKeyMatchTier under the profile's confirmedRules; this file has no
# comparison of its own. Several hits with expect=one is
# 'ambiguous' and the candidates come back in the P0-R4 standard shape
# (PROFILE-SCHEMA.md 6.6 c) for human.choose to render -- the step does
# not pick.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')

$Manifest = @{
  id         = 'file.find'
  group      = 'file'
  summary    = 'Find the file(s) for a key or glob; full-width and stamped names tolerated'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    dir     = @{ type='path';   required=$true; desc='folder to look in; relative paths resolve under the work dir' }
    key     = @{ type='string'; default=''; desc='the item key the file is named after (stem); empty = glob only' }
    ext     = @{ type='string'; default=''; desc='extension the file must have (e.g. dat, .png); empty = any' }
    glob    = @{ type='string'; default='*'; desc='wildcard filter applied first' }
    recurse = @{ type='bool';   default=$false }
    expect  = @{ type='string'; default='one'; enum=@('one','any'); desc='one: several hits are ambiguous; any: return them all' }
  }
  outputs    = @{
    path       = @{ type='path';   desc='the hit (the only one, or the newest with expect=any)' }
    files      = @{ type='list';   desc='every hit, newest first' }
    found      = @{ type='int';    desc='how many files matched' }
    matchedBy  = @{ type='string'; desc='exact | stripped | fullwidth | case | glob (kernel/Key.ps1 tiers)' }
    candidates = @{ type='map';    desc='P0-R4 candidate shape when ambiguous, else null' }
  }
  failures   = @(
    @{ id = 'dir_not_found';  transient = $false }
    @{ id = 'file_not_found'; transient = $true  }
    @{ id = 'ambiguous';      transient = $false }
  )
  example    = @{ use = 'file.find'; with = @{ dir = '{{profile.paths.downloads}}'; key = '{{item.keySafe}}'; ext = 'dat' } }
  notes      = 'file_not_found is transient on purpose: a download that has not landed yet is the usual cause, and a retry after a wait is the right first move. Match order is kernel/Key.ps1''s: exact stem, suffix/prefix rules stripped (the batch stamp), full-width folded, case folded; the first tier with hits wins.'
}

function FileFind-Stem {
    param([string]$Name)
    return [System.IO.Path]::GetFileNameWithoutExtension($Name)
}

function FileFind-Tier {
    # PURE. How a file name's stem relates to the key, through kernel/Key.ps1
    # (Get-EbiKeyMatchTier: exact > stripped > fullwidth > case under the
    # profile's rules); '' when it does not; 'glob' when there is no key.
    param([string]$Name, [string]$Key, $Rules = $null)
    if ([string]::IsNullOrEmpty($Key)) { return 'glob' }
    return (Get-EbiKeyMatchTier -Value (FileFind-Stem -Name $Name) -Key $Key -Rules $Rules)
}

function FileFind-Rank {
    # PURE. Names + key -> @{ tier; names } for the best tier that hit
    # (exact > stripped > fullwidth > case > glob); @{ tier=''; names=@() } if none.
    param($Names, [string]$Key, [string]$Ext = '', $Rules = $null)
    $byTier = @{}
    foreach ($tn in @(Get-EbiKeyTierOrder) + @('glob')) { $byTier[$tn] = New-Object System.Collections.ArrayList }
    $e = if ([string]::IsNullOrWhiteSpace($Ext)) { '' } else { '.' + $Ext.TrimStart('.') }
    foreach ($n in @($Names)) {
        $name = [string]$n
        if ($e -ne '' -and -not $name.EndsWith($e, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $t = FileFind-Tier -Name $name -Key $Key -Rules $Rules
        if ($t -ne '') { [void]$byTier[$t].Add($name) }
    }
    foreach ($t in @(Get-EbiKeyTierOrder) + @('glob')) {
        if ($byTier[$t].Count -gt 0) { return @{ tier = $t; names = $byTier[$t].ToArray() } }
    }
    return @{ tier = ''; names = @() }
}

function FileFind-Candidates {
    # PURE. File records (@{ name; path; modifiedAt; size }) newest first ->
    # the P0-R4 candidate shape.
    param($Files, [string]$Dir, [string]$Tier)
    $list = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($f in @($Files)) {
        $i++
        [void]$list.Add(@{ id = ('c' + $i); candidate = [string]$f['name']; evidence = @{ source = $Dir; modifiedAt = [string]$f['modifiedAt']; size = [string]$f['size']; matchedBy = $Tier } })
    }
    $sugg = $null
    if ($list.Count -gt 0) { $sugg = @{ id = 'c1'; reason = 'newest by modified time' } }
    return @{ candidates = $list.ToArray(); suggestion = $sugg; doubts = ('' + $list.Count + ' files match; a rerun of one transfer leaves several, and only the run window tells which is this run''s') }
}

function Invoke-Step {
    param($In, $Ctx)
    $dir = Resolve-EbiWorkPath -PathValue ([string]$In['dir']) -WorkDir ([string]$Ctx['WorkDir'])
    $key = [string]$In['key']; $ext = [string]$In['ext']; $glob = [string]$In['glob']; $expect = [string]$In['expect']
    if ([string]::IsNullOrWhiteSpace($glob)) { $glob = '*' }
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
        if ($Ctx['DryRun']) {
            # A dry run walks the plan on a machine that may have none of the
            # folders; the lookup itself is read-only, so it runs when it can.
            $Ctx.Log.Warn(('would fail: no folder at {0}' -f $dir))
            return @{ ok = $true; path = ''; files = @(); found = 0; matchedBy = ''; candidates = $null; warnings = @( @{ code = 'would_fail'; message = ('no folder at ' + $dir); data = @{} } ) }
        }
        return @{ ok = $false; failure = 'dir_not_found'; message = $dir; path = ''; files = @(); found = 0; matchedBy = ''; candidates = $null }
    }
    $items = @(Get-ChildItem -LiteralPath $dir -Filter $glob -File -Recurse:([bool]$In['recurse']) -ErrorAction SilentlyContinue)
    $byName = @{}
    foreach ($it in $items) { $byName[$it.Name] = $it }
    $rank = FileFind-Rank -Names @($byName.Keys) -Key $key -Ext $ext -Rules (Get-EbiKeyRules -Profile $Ctx['Profile'])
    if ($rank['tier'] -eq '') {
        return @{ ok = $false; failure = 'file_not_found'; message = ('nothing for key "' + $key + '"' + $(if ($ext -ne '') { ' (.' + $ext.TrimStart('.') + ')' } else { '' }) + ' under ' + $dir + ' (' + $items.Count + ' file(s) seen)'); path = ''; files = @(); found = 0; matchedBy = ''; candidates = $null }
    }
    $hits = New-Object System.Collections.ArrayList
    foreach ($it in @(@($rank['names']) | ForEach-Object { $byName[$_] } | Sort-Object -Property @{ Expression = 'LastWriteTimeUtc'; Descending = $true }, @{ Expression = 'Name' })) {
        [void]$hits.Add(@{ name = $it.Name; path = $it.FullName; modifiedAt = $it.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'); size = [int64]$it.Length })
    }
    $paths = @(@($hits.ToArray()) | ForEach-Object { $_['path'] })
    if ($hits.Count -gt 1 -and $expect -ne 'any') {
        $cand = FileFind-Candidates -Files $hits.ToArray() -Dir $dir -Tier $rank['tier']
        return @{ ok = $false; failure = 'ambiguous'; message = ('' + $hits.Count + ' files match key "' + $key + '" (' + $rank['tier'] + ')'); path = ''; files = $paths; found = $hits.Count; matchedBy = $rank['tier']; candidates = $cand }
    }
    $warnings = @()
    if ($rank['tier'] -eq 'fullwidth') { $warnings = @( @{ code = 'full_width_name'; message = 'matched only after folding full-width characters'; data = @{ name = $hits[0]['name'] } } ) }
    return @{ ok = $true; path = $paths[0]; files = $paths; found = $hits.Count; matchedBy = $rank['tier']; candidates = $null; warnings = $warnings }
}
