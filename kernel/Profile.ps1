#Requires -Version 5.1
# ============================================================
#  kernel/Profile.ps1
#
#  Loading profiles/<name>/ (P1-10; the profile CONTENT is P2-01).
#  Dot-source only (no param() block, ASCII source, no class).
#
#  A profile is a directory of JSON files (PROFILE-SCHEMA.md 1):
#      vocabulary.json pages.json grammar.json rules.json worklist.json
#      layout.json calibration.json window.json paths.json
#  Each file becomes the same-named top-level key of the profile
#  hashtable; a file that is not there is simply absent (and listed in
#  'missing'), a file that is not JSON is an error. <WorkDir>/ebi.local.json
#  (PROFILE-SCHEMA.md 0: the per-machine / per-job overlay) is deep-merged
#  on top -- nested maps merge key by key, anything else is replaced whole,
#  the same rule ConfigOverlay.ps1 has used for years.
#  Templates inside the profile ({{run.operator}} in worklist.file) are
#  NOT expanded here: Context.ps1 does that on the way out (4.3).
# ============================================================

. (Join-Path $PSScriptRoot 'Json.ps1')

function Get-EbiProfileFiles {
    return @('vocabulary', 'pages', 'grammar', 'rules', 'worklist', 'layout', 'calibration', 'window', 'paths')
}

function Get-EbiDefaultProfilesRoot {
    return (Join-Path (Split-Path $PSScriptRoot -Parent) 'profiles')
}

function Merge-EbiHashtable {
    # PURE. A new hashtable: Base with Overlay merged in (nested
    # dictionaries merge, everything else is replaced by the overlay).
    param($Base, $Overlay)
    $out = @{}
    if ($Base -is [System.Collections.IDictionary]) { foreach ($k in $Base.Keys) { $out[[string]$k] = $Base[$k] } }
    if (-not ($Overlay -is [System.Collections.IDictionary])) { return $out }
    foreach ($k in $Overlay.Keys) {
        $key = [string]$k
        if ($out.Contains($key) -and ($out[$key] -is [System.Collections.IDictionary]) -and ($Overlay[$k] -is [System.Collections.IDictionary])) {
            $out[$key] = Merge-EbiHashtable -Base $out[$key] -Overlay $Overlay[$k]
        } else {
            $out[$key] = $Overlay[$k]
        }
    }
    return $out
}

function Resolve-EbiProfileDir {
    # A profile name -> profiles/<name>; a path that is a directory -> itself.
    # '' or 'none' -> ''.
    param([string]$NameOrPath, [string]$ProfilesRoot = '')
    if ([string]::IsNullOrWhiteSpace($NameOrPath) -or $NameOrPath -eq 'none') { return '' }
    if (Test-Path -LiteralPath $NameOrPath -PathType Container) { return (Resolve-Path -LiteralPath $NameOrPath).ProviderPath }
    if ([string]::IsNullOrWhiteSpace($ProfilesRoot)) { $ProfilesRoot = Get-EbiDefaultProfilesRoot }
    return (Join-Path $ProfilesRoot $NameOrPath)
}

function Read-EbiProfile {
    <#
      profiles/<name>/ (+ <WorkDir>/ebi.local.json) -> @{ ok; value; name;
      dir; missing; overlay; message }. value is the profile hashtable
      (one key per file present, plus 'name'). ok=$false when the
      directory is missing or a file is not valid JSON; a missing file is
      not an error (a capture-only profile has no layout.json).
    #>
    param([string]$Dir, [string]$WorkDir = '')
    $name = if ([string]::IsNullOrWhiteSpace($Dir)) { '' } else { Split-Path -Leaf $Dir }
    $result = @{ ok = $false; value = @{}; name = $name; dir = $Dir; missing = @(); overlay = $false; message = '' }
    if ([string]::IsNullOrWhiteSpace($Dir)) { $result['message'] = 'no profile directory given'; return $result }
    if (-not (Test-Path -LiteralPath $Dir -PathType Container)) { $result['message'] = ('profile directory not found: ' + $Dir); return $result }
    $profile = @{ name = $name }
    $missing = New-Object System.Collections.ArrayList
    foreach ($f in (Get-EbiProfileFiles)) {
        $path = Join-Path $Dir ($f + '.json')
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { [void]$missing.Add($f); continue }
        $r = Read-EbiJson -Path $path
        if (-not $r['ok']) { $result['message'] = $r['message']; $result['missing'] = $missing.ToArray(); return $result }
        $profile[$f] = $r['value']
    }
    if (-not [string]::IsNullOrWhiteSpace($WorkDir)) {
        $local = Join-Path $WorkDir 'ebi.local.json'
        if (Test-Path -LiteralPath $local -PathType Leaf) {
            # hand-edited: tolerate Notepad's "ANSI" (Shift_JIS) save, but say so
            $r = Read-EbiJson -Path $local -AllowCp932
            if (-not $r['ok']) { $result['message'] = $r['message']; $result['missing'] = $missing.ToArray(); return $result }
            $result['localEncoding'] = $r['encoding']
            if ($r['encoding'] -eq 'cp932') { Write-Host ('  [warn ] ' + $local + ' is not UTF-8; read it as Shift_JIS (CP932). Save it as UTF-8 to silence this.') -ForegroundColor Yellow }
            if ($r['value'] -is [System.Collections.IDictionary]) { $profile = Merge-EbiHashtable -Base $profile -Overlay $r['value']; $result['overlay'] = $true }
        }
    }
    $result['ok'] = $true
    $result['value'] = $profile
    $result['missing'] = $missing.ToArray()
    return $result
}
