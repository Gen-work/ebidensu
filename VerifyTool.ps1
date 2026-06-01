# ============================================================
# VerifyTool.ps1 - main entry for evidence verification workflow
#
# Daily usage:
#   .\VerifyTool.ps1
#   .\VerifyTool.ps1 -Help
#   .\VerifyTool.ps1 -Phase Status
#   .\VerifyTool.ps1 -Phase Clone
#   .\VerifyTool.ps1 -Phase ReplaceGift
#   .\VerifyTool.ps1 -Phase ReviewEvidence
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner   = '',
    [string]$Phase   = 'Menu',

    [int]$WindowWidth  = 0,
    [int]$WindowHeight = 0,
    [int]$CropPx       = -1,

    [string[]]$TargetIds = @(),
    [string]$EvidenceDir = '',
    [string]$CursorCell  = '',

    # Clone / Replace
    [string]$CloneSourceDir = '',
    [string]$J4BaseDir      = '',
    [string[]]$BizCodes     = @(),

    # DfSnap override (takes precedence over VerifyConfig.psd1 -> Df.ExePath)
    [string]$DfExePath = '',

    # ProbeShapes
    [string]$ProbeFile  = '',
    [string]$ProbeSheet = '',

    # CheckSheet
    [string]$CheckSheetPath = '',

    [switch]$Force,
    [switch]$Interactive,
    [switch]$NoResize,
    [switch]$RefreshUrls,
    [switch]$DryRun,
    [switch]$Help,

    [string]$ConfigPath = ''
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

function Read-Choice([string]$Prompt, [string]$Default = '') {
    if ([string]::IsNullOrWhiteSpace($Default)) { return (Read-Host $Prompt) }
    $v = Read-Host ("{0} [{1}]" -f $Prompt, $Default)
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v
}

function To-BoolText([bool]$v) { if ($v) { return 'ON' } else { return 'off' } }

function Load-VerifyConfig([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { $Path = Join-Path $PSScriptRoot 'VerifyConfig.psd1' }
    if (-not (Test-Path -LiteralPath $Path)) { throw "Config not found: $Path" }
    return Import-PowerShellDataFile -LiteralPath $Path
}

function Show-VerifyHelp([hashtable]$Config) {
    Write-Host ''
    Write-Host '===== VerifyTool Help =====' -ForegroundColor Green
    Write-Host ''
    Write-Host 'Daily entry:'
    Write-Host '  .\VerifyTool.ps1'
    Write-Host ''
    Write-Host 'Useful commands:'
    Write-Host '  .\VerifyTool.ps1 -Help'
    Write-Host '  .\VerifyTool.ps1 -Phase Status'
    Write-Host '  .\VerifyTool.ps1 -Phase Validate'
    Write-Host '  .\VerifyTool.ps1 -Phase Mapping -Force'
    Write-Host '  .\VerifyTool.ps1 -Phase ExcelSnap'
    Write-Host '  .\VerifyTool.ps1 -Phase GiftHmSnap -TargetIds JIGPL48S'
    Write-Host '  .\VerifyTool.ps1 -Phase GiftMqSnap -TargetIds JIGPL48S,JIDSL48S'
    Write-Host '  .\VerifyTool.ps1 -Phase GiftJenkins -RefreshUrls'
    Write-Host '  .\VerifyTool.ps1 -Phase GiftJenkinsNoFile'
    Write-Host '  .\VerifyTool.ps1 -Phase GfixHmSnap'
    Write-Host '  .\VerifyTool.ps1 -Phase GfixJenkins'
    Write-Host '  .\VerifyTool.ps1 -Phase GfixLog'
    Write-Host '  .\VerifyTool.ps1 -Phase DfSnap -DfExePath "C:\tools\df.exe"'
    Write-Host '  .\VerifyTool.ps1 -Phase MarkGfixLog          # standalone re-highlight utility (folded into MarkGfix)'
    Write-Host '  .\VerifyTool.ps1 -Phase Clone -CloneSourceDir <ext_path>'
    Write-Host '  .\VerifyTool.ps1 -Phase Align -J4BaseDir <j4_path>'
    Write-Host '  .\VerifyTool.ps1 -Phase ReplaceGift'
    Write-Host '  .\VerifyTool.ps1 -Phase ReplaceGfix -TargetIds JIGPL48S'
    Write-Host '  .\VerifyTool.ps1 -Phase ReplaceDf'
    Write-Host '  .\VerifyTool.ps1 -Phase MarkGift'
    Write-Host '  .\VerifyTool.ps1 -Phase MarkGift -TargetIds KJRVWD64 -Force'
    Write-Host '  .\VerifyTool.ps1 -Phase MarkGfix             # red rect + GFIX log highlight (one pass)'
    Write-Host '  .\VerifyTool.ps1 -Phase MarkDf'
    Write-Host '  .\VerifyTool.ps1 -Phase ProbeShapes -ProbeFile <evidence.xlsx>'
    Write-Host '  .\VerifyTool.ps1 -Phase RepairMapping'
    Write-Host '  .\VerifyTool.ps1 -Phase ReviewGift'
    Write-Host '  .\VerifyTool.ps1 -Phase ReviewGfix'
    Write-Host '  .\VerifyTool.ps1 -Phase ReviewDf'
    Write-Host '  .\VerifyTool.ps1 -Phase ReviewEvidence'
    Write-Host '  .\VerifyTool.ps1 -Phase ReviewEvidence -CursorCell A1'
    Write-Host '  .\VerifyTool.ps1 -Phase Comments            # list recorded review comments'
    Write-Host '  .\VerifyTool.ps1 -Phase CheckSheet          # fill the review check sheet (temp-preview, then commit)'
    Write-Host '  .\VerifyTool.ps1 -Phase CheckSheet -CheckSheetPath "\\srv\...\check.xlsx"'
    Write-Host '  .\VerifyTool.ps1 -Phase DeliverMail         # one Outlook draft per Excel; you click Send + Enter'
    Write-Host '  .\VerifyTool.ps1 -Phase DeliverMail -TargetIds SJRVWD64'
    Write-Host ''
    Write-Host 'Common options:'
    Write-Host '  -WorkDir <path>       Work folder. If omitted, last used path is remembered.'
    Write-Host ('  -Owner {0}             mapping_{0}.csv owner suffix. Default comes from config.' -f ([char]0x53B3))
    Write-Host '  -TargetIds A,B        Limit run by Correl_ID_S / Correl_ID_M / JOB_NAME / Excel_NAME.'
    Write-Host '  -CloneSourceDir <p>   External path for Clone (existing evidence files per bizcode).'
    Write-Host '  -J4BaseDir <p>        J4 baseline root for Align. If omitted, config/session/CloneSourceDir is used.'
    Write-Host '  -BizCodes A,B         Override bizcode candidate list for Clone.'
    Write-Host '  -Force                Re-run rows whose flag/bit is already set.'
    Write-Host '  -Interactive          Ask before each row in supported stages.'
    Write-Host '  -WindowWidth 1050 -WindowHeight 761 -CropPx 6'
    Write-Host '  -NoResize             Do not resize Edge window.'
    Write-Host '  -RefreshUrls          Re-capture Jenkins folder URLs.'
    Write-Host '  -DryRun               Print arguments instead of invoking child scripts.'
    Write-Host ''
    Write-Host 'Phases:'
    foreach ($p in $Config.PhaseOrder) {
        $field = if ([string]::IsNullOrWhiteSpace($p.Field)) { '-' } else { $p.Field }
        $bv = if ($p.ContainsKey('BitValue')) { (' bit=' + [string]$p.BitValue) } else { '' }
        Write-Host ("  {0,-20} {1,-12}{2} {3} [{4}]" -f $p.Key, $field, $bv, $p.Label, $p.Status)
    }
    Write-Host ''
    Write-Host 'isReplaced/isMarked/isReviewed bitmask: 1=GIFT, 2=GFIX, 4=DF. 7 = all done.'
    Write-Host ''
    Write-Host 'Aliases kept for old commands:'
    Write-Host '  Excel -> ExcelSnap, HmGift -> GiftHmSnap, MqGift -> GiftMqSnap'
    Write-Host '  JenkinsGift -> GiftJenkins, NoGfix -> GiftJenkinsNoFile'
    Write-Host '  HmGfix -> GfixHmSnap, JenkinsGfix -> GfixJenkins, Review -> ReviewEvidence'
    Write-Host '  MkExcel/RenameExcel -> Clone'
    Write-Host '  Replace/ReplaceEvidence -> ReplaceGift (default)'
    Write-Host '  Rgift/Rgfix/Rdf -> ReplaceGift/ReplaceGfix/ReplaceDf'
    Write-Host ''
}

function Resolve-Phase([hashtable]$Config, [string]$RawPhase) {
    if ([string]::IsNullOrWhiteSpace($RawPhase)) { return 'Menu' }
    foreach ($k in $Config.Aliases.Keys) {
        if ($k -ieq $RawPhase) { return [string]$Config.Aliases[$k] }
    }
    foreach ($p in $Config.PhaseOrder) {
        if ($p.Key -ieq $RawPhase) { return [string]$p.Key }
    }
    throw "Unknown phase: $RawPhase. Use -Help to see available phases."
}

function Resolve-ToolPath([hashtable]$Config, [string]$ScriptKey) {
    $name = $Config.Scripts[$ScriptKey]
    if ([string]::IsNullOrWhiteSpace($name)) { throw "Script key not configured: $ScriptKey" }
    $p = Join-Path $PSScriptRoot $name
    if (-not (Test-Path -LiteralPath $p)) { throw "Script not found: $p" }
    return (Resolve-Path -LiteralPath $p).Path
}

function Load-Session([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return @{} }
    try {
        $obj = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        $h = @{}
        foreach ($p in $obj.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    } catch { return @{} }
}

function Save-Session([string]$Path, [hashtable]$Session) {
    try { $Session | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8 } catch {}
}


function Resolve-AlignJ4BaseDir([hashtable]$Config, [hashtable]$State) {
    foreach ($candidate in @(
        [string]$State.J4BaseDir,
        [string]$Config.Align.J4BaseDir,
        [string]$State.CloneSourceDir
    )) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }
    return ''
}

function Get-MappingPath([hashtable]$Config, [string]$Dir, [string]$OwnerName) {
    $pat = [string]$Config.Paths.MappingPattern
    if ([string]::IsNullOrWhiteSpace($pat)) { $pat = 'mapping_{0}.csv' }
    return (Join-Path $Dir ($pat -f $OwnerName))
}

function Load-MappingSafe([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    return @(Import-Csv -LiteralPath $Path -Encoding UTF8)
}

function Ensure-PhaseColumns([hashtable]$Config, [string]$MappingPath, [switch]$Quiet) {
    <#
    Reads mapping CSV, collects every non-empty `Field` from $Config.PhaseOrder,
    and adds any missing column to every row with default value '0'. Writes
    the mapping back only if at least one column was added. Existing data is
    never modified. Returns the count of columns added.
    #>
    if (-not (Test-Path -LiteralPath $MappingPath)) { return 0 }
    $rows = @(Import-Csv -LiteralPath $MappingPath -Encoding UTF8)
    if ($rows.Count -eq 0) { return 0 }

    $fields = @()
    foreach ($p in $Config.PhaseOrder) {
        $f = [string]$p.Field
        if (-not [string]::IsNullOrWhiteSpace($f)) { $fields += $f }
    }
    $fields = @($fields | Select-Object -Unique)

    $first = $rows | Select-Object -First 1
    $existing = @($first.PSObject.Properties.Name)
    $missing = @($fields | Where-Object { $existing -notcontains $_ })
    if ($missing.Count -eq 0) { return 0 }

    foreach ($r in $rows) {
        foreach ($f in $missing) {
            if (-not ($r.PSObject.Properties.Name -contains $f)) {
                $r | Add-Member -NotePropertyName $f -NotePropertyValue '0' -Force
            }
        }
    }
    $rows | Export-Csv -LiteralPath $MappingPath -Encoding UTF8 -NoTypeInformation -Force
    if (-not $Quiet.IsPresent) {
        Write-Host ("  [auto-repair] added {0} missing column(s) to mapping: {1}" -f $missing.Count, ($missing -join ', ')) -ForegroundColor DarkYellow
    }
    return $missing.Count
}

function Get-FieldStats([array]$Rows, [string]$Field, [int]$BitValue = 0) {
    $total = $Rows.Count
    if ($total -eq 0 -or [string]::IsNullOrWhiteSpace($Field)) {
        return @{ Total=$total; Done=0; Pending=$total; Missing=$true }
    }
    $first = $Rows | Select-Object -First 1
    if ($null -eq $first -or -not ($first.PSObject.Properties.Name -contains $Field)) {
        return @{ Total=$total; Done=0; Pending=$total; Missing=$true }
    }
    if ($BitValue -gt 0) {
        # Bitmask check: (value -band BitValue) == BitValue
        $done = @($Rows | Where-Object {
            $v = 0
            try { $v = [int]$_.$Field } catch { $v = 0 }
            ($v -band $BitValue) -eq $BitValue
        }).Count
    } else {
        $done = @($Rows | Where-Object { $_.$Field -eq '1' }).Count
    }
    return @{ Total=$total; Done=$done; Pending=($total - $done); Missing=$false }
}

function Get-PhaseBit($Phase) {
    if ($null -eq $Phase) { return 0 }
    if ($Phase.ContainsKey('BitValue')) {
        try { return [int]$Phase.BitValue } catch { return 0 }
    }
    return 0
}

function Show-Status([hashtable]$Config, [array]$Rows) {
    Write-Host ''
    Write-Host '===== Current mapping status =====' -ForegroundColor Cyan
    if ($Rows.Count -eq 0) {
        Write-Host '  mapping csv not found yet, or it has no rows.' -ForegroundColor Yellow
        return
    }
    Write-Host ("  Rows: {0}" -f $Rows.Count)
    foreach ($p in $Config.PhaseOrder) {
        if ([string]::IsNullOrWhiteSpace($p.Field)) { continue }
        $bv = Get-PhaseBit $p
        $s  = Get-FieldStats $Rows $p.Field $bv
        $suffix = if ($p.Status -eq 'planned') { ' (planned)' } elseif ($p.Status -eq 'legacy') { ' (legacy)' } else { '' }
        $tag = if ($bv -gt 0) { ("{0} bit={1}" -f $p.Field, $bv) } else { $p.Field }
        if ($s.Missing) {
            Write-Host ("  {0,-22} : column missing ({1}){2}" -f $p.Key, $tag, $suffix) -ForegroundColor DarkYellow
        } else {
            Write-Host ("  {0,-22} : {1}/{2} done, {3} pending  [{4}]{5}" -f $p.Key, $s.Done, $s.Total, $s.Pending, $tag, $suffix)
        }
    }

    foreach ($field in @('TO_code','FROM_code','isMultiAppl')) {
        $first = $Rows | Select-Object -First 1
        if ($first -and ($first.PSObject.Properties.Name -contains $field)) {
            $groups = @($Rows | Group-Object $field | Sort-Object Name)
            if ($groups.Count -gt 0) {
                Write-Host ''
                Write-Host ("  {0} groups:" -f $field) -ForegroundColor DarkGray
                foreach ($g in $groups) { Write-Host ("    {0,-10} {1}" -f $g.Name, $g.Count) -ForegroundColor DarkGray }
            }
        }
    }
}

function Get-RecommendPhase([hashtable]$Config, [array]$Rows) {
    if ($Rows.Count -eq 0) { return 'Mapping' }
    foreach ($p in $Config.PhaseOrder) {
        if ($p.Status -eq 'planned') { continue }
        if ([string]::IsNullOrWhiteSpace($p.Field)) { continue }
        $bv = Get-PhaseBit $p
        $s  = Get-FieldStats $Rows $p.Field $bv
        if (-not $s.Missing -and $s.Pending -gt 0) { return $p.Key }
    }
    return 'Status'
}

function Show-PhaseNotes([string]$PhaseKey) {
    $lines = switch -Regex ($PhaseKey) {
        '^Align$' { @(
            '  Phase params:',
            '    f=Force        -> -Apply  : sync DIFF sheets (work <- J4 values). EXPERIMENTAL - run on a copy first.',
            '    h=HostTypes    -> which FROM_sys / TO_sys column values count as Host (mainframe).',
            '                      e.g. enter: HOST   (check your mapping FROM_sys/TO_sys column for the actual literal)',
            '                      HostToOpen  = 3 receive sheets only',
            '                      OpenToOpen / OpenToHost = GIFT+GFIX send sheets + 3 receive sheets',
            '    j=J4BaseDir    -> root folder of J4 baseline workbooks (searched recursively)',
            '    t=TargetIds    -> limit to specific Excel_NAME / Correl_ID / JOB_NAME'
        ) }
        '^Clone$' { @(
            '  Phase params:',
            '    d=SourceDir    -> external folder containing per-bizcode evidence files to copy',
            '    b=BizCodes     -> override bizcode list (default: derived from mapping TO_code/FROM_code)',
            '    t=TargetIds    -> limit rows'
        ) }
        '^Replace(Gift|Gfix|Df)$' { @(
            '  Phase params:',
            '    t=TargetIds    -> limit to specific Correl_ID_S / JOB_NAME / Excel_NAME',
            '    f=Force        -> re-run rows already marked done'
        ) }
        '^MarkGfix$' { @(
            '  Phase params:',
            '    t=TargetIds    -> limit rows',
            '    f=Force        -> overwrite existing marks',
            '  NOTE: the GFIX log yellow-highlight is done in this same pass (folded in; no separate phase).'
        ) }
        '^Mark(Gift|Df)$' { @(
            '  Phase params:',
            '    t=TargetIds    -> limit rows',
            '    f=Force        -> overwrite existing marks'
        ) }
        '^MarkGfixLog$' { @(
            '  Phase params:',
            '    t=TargetIds    -> limit rows',
            '    f=Force        -> re-highlight already-done rows'
        ) }
        '^Review(Gift|Gfix|Df|Evidence)$' { @(
            '  Phase params:',
            '    a=CursorCell   -> cell to activate when workbook opens (default: A3)',
            '    t=TargetIds    -> limit rows',
            '    f=Force        -> re-open already-reviewed workbooks',
            '  NOTE: ReviewGift/Gfix/Df open the matching sheet up front. At the per-workbook',
            '        prompt, append  -m "comment"  (works with Enter / s / q) to record a note;',
            '        prior notes are shown on open. List all notes with the Comments phase.'
        ) }
        '^(Gift|Gfix)Jenkins$|^GiftJenkinsNoFile$' { @(
            '  Phase params:',
            '    t=TargetIds    -> limit rows (Correl_ID_S / JOB_NAME)',
            '    i=Interactive  -> pause before each row for manual confirmation',
            '    f=Force        -> re-capture already-done rows',
            '    n=NoResize     -> do not resize Edge window',
            '    w=Window       -> Edge window size',
            '    c=CropPx       -> crop captured PNG edges',
            '    r=RefreshUrls  -> re-fetch Jenkins folder URLs (use when URLs changed)'
        ) }
        '^GiftMqSnap$|^(Gift|Gfix)HmSnap$' { @(
            '  Phase params:',
            '    t=TargetIds    -> limit rows (Correl_ID_S / JOB_NAME)',
            '    i=Interactive  -> pause before each row for manual confirmation',
            '    f=Force        -> re-capture already-done rows',
            '    n=NoResize     -> do not resize Edge window',
            '    w=Window       -> Edge window size',
            '    c=CropPx       -> crop captured PNG edges'
        ) }
        '^GfixLogDownload$' { @(
            '  Phase params:',
            '    t=TargetIds    -> limit rows',
            '    f=Force        -> re-download already-done rows',
            '  NOTE: GoAnywhere rows-per-page must be set to 100 manually before running.'
        ) }
        '^DfSnap$' { @(
            '  Phase params:',
            '    t=TargetIds    -> limit rows',
            '    f=Force        -> re-capture already-done rows',
            '    e=DfExePath    -> override df.exe path for this run',
            '  NOTE: set Df.ExePath in VerifyConfig.psd1 to skip the path prompt.'
        ) }
        '^Validate$' { @(
            '  Phase params:',
            '    t=TargetIds    -> limit rows  (read-only, no mapping changes)'
        ) }
        '^ProbeShapes$' { @(
            '  Phase params:',
            '    p=ProbeFile    -> Excel file to inspect',
            '    s=ProbeSheet   -> sheet name to inspect'
        ) }
        '^Crop$' { @(
            '  Phase params:',
            '    c=CropPx       -> crop existing PNG edges',
            '    f=Force        -> re-crop directories already marked cropped'
        ) }
        '^CheckSheet$' { @(
            '  Phase params:',
            '    k=CheckSheetPath -> review check sheet .xlsx (prompts if config path is missing)',
            '    t=TargetIds      -> limit to specific Excel_NAME / Correl_ID / JOB_NAME',
            '    f=Force          -> add a row even if the Excel is already listed',
            '  NOTE: edits are previewed in a TEMP copy first; you press Enter to commit.',
            '        If the real check sheet changes during the preview, the write is held.'
        ) }
        '^DeliverMail$' { @(
            '  Phase params:',
            '    t=TargetIds    -> limit to specific Excel_NAME / Correl_ID / JOB_NAME',
            '    f=Force        -> re-draft rows already marked delivered',
            '  NOTE: one Outlook DRAFT per Excel (never auto-sent). You click Send, then',
            '        press Enter to mark isDelivered. Append  -m "comment"  to record a note.'
        ) }
        default { @() }
    }
    foreach ($l in $lines) { Write-Host $l -ForegroundColor DarkGray }
}

function Get-PhaseOptionKeys([string]$PhaseKey) {
    switch -Regex ($PhaseKey) {
        '^Align$' { return @('f','apply','h','j','t') }
        '^Clone$' { return @('d','b','t','f') }
        '^Replace(Gift|Gfix|Df)$' { return @('t','f') }
        '^Mark(Gift|Gfix|Df)$' { return @('t','f') }
        '^MarkGfixLog$' { return @('t','f') }
        '^Review(Gift|Gfix|Df|Evidence)$' { return @('a','t','f') }
        '^(Gift|Gfix)Jenkins$|^GiftJenkinsNoFile$' { return @('t','i','f','n','w','c','r') }
        '^GiftMqSnap$|^(Gift|Gfix)HmSnap$' { return @('t','i','f','n','w','c') }
        '^GfixLogDownload$' { return @('t','f') }
        '^DfSnap$' { return @('t','f','e') }
        '^Validate$' { return @('t') }
        '^ProbeShapes$' { return @('p','s') }
        '^Crop$' { return @('c','f') }
        '^CheckSheet$' { return @('t','f','k') }
        '^DeliverMail$' { return @('t','f') }
        '^Mapping$|^ExcelSnap$' { return @('f') }
        default { return @() }
    }
}

function Test-PhaseOption([string[]]$Allowed, [string]$Key) {
    return (@($Allowed) -contains $Key)
}

function Write-UnusedOption([string]$PhaseKey, [string]$Key) {
    Write-Host ("  option '{0}' is not used by phase {1}" -f $Key, $PhaseKey) -ForegroundColor Yellow
}

function Ask-RunOptions([hashtable]$State, [string]$PhaseKey = '') {
    $allowed = @(Get-PhaseOptionKeys $PhaseKey)

    Write-Host ("Selected phase : {0}" -f $PhaseKey) -ForegroundColor Cyan
    Show-PhaseNotes $PhaseKey
    Write-Host ''
    Write-Host 'Options for this run:' -ForegroundColor Cyan

    if (Test-PhaseOption $allowed 'apply') { Write-Host ("  Apply(Align)   : {0}" -f (To-BoolText $State.Force)) }
    elseif (Test-PhaseOption $allowed 'f') { Write-Host ("  Force          : {0}" -f (To-BoolText $State.Force)) }
    if (Test-PhaseOption $allowed 'i') { Write-Host ("  Interactive    : {0}" -f (To-BoolText $State.Interactive)) }
    if (Test-PhaseOption $allowed 'n') { Write-Host ("  NoResize       : {0}" -f (To-BoolText $State.NoResize)) }
    if (Test-PhaseOption $allowed 'r') { Write-Host ("  RefreshUrls    : {0}" -f (To-BoolText $State.RefreshUrls)) }
    if (Test-PhaseOption $allowed 'w') { Write-Host ("  Window         : {0}x{1}" -f $State.WindowWidth, $State.WindowHeight) }
    if (Test-PhaseOption $allowed 'c') { Write-Host ("  CropPx         : {0}" -f $State.CropPx) }
    if (Test-PhaseOption $allowed 'a') { Write-Host ("  CursorCell     : {0}" -f $State.CursorCell) }
    if (Test-PhaseOption $allowed 't') { Write-Host ("  TargetIds      : {0}" -f $(if ($State.TargetIds.Count -gt 0) { $State.TargetIds -join ', ' } else { '(all)' })) }
    if (Test-PhaseOption $allowed 'd') { Write-Host ("  CloneSourceDir : {0}" -f $State.CloneSourceDir) }
    if (Test-PhaseOption $allowed 'j') { Write-Host ("  J4BaseDir      : {0}" -f $State.J4BaseDir) }
    if (Test-PhaseOption $allowed 'b') { Write-Host ("  BizCodes       : {0}" -f $(if ($State.BizCodes.Count -gt 0) { $State.BizCodes -join ', ' } else { '(auto)' })) }
    if (Test-PhaseOption $allowed 'h') { Write-Host ("  HostSystemTypes: {0}" -f $(if ($State.HostSystemTypes.Count -gt 0) { $State.HostSystemTypes -join ', ' } else { '(config/auto)' })) }
    if (Test-PhaseOption $allowed 'e') { Write-Host ("  DfExePath      : {0}" -f $State.DfExePath) }
    if (Test-PhaseOption $allowed 'p') { Write-Host ("  ProbeFile      : {0}" -f $State.ProbeFile) }
    if (Test-PhaseOption $allowed 's') { Write-Host ("  ProbeSheet     : {0}" -f $State.ProbeSheet) }
    if (Test-PhaseOption $allowed 'k') { Write-Host ("  CheckSheetPath : {0}" -f $(if ([string]::IsNullOrWhiteSpace($State.CheckSheetPath)) { '(config/prompt)' } else { $State.CheckSheetPath })) }

    Write-Host ''
    if ($allowed.Count -eq 0) {
        Write-Host '  This phase has no interactive options. Enter=continue' -ForegroundColor DarkGray
    } else {
        $help = @()
        if (Test-PhaseOption $allowed 'f') {
            if (Test-PhaseOption $allowed 'apply') { $help += 'f/-apply=Apply mode' }
            else { $help += 'f=Force' }
        }
        if (Test-PhaseOption $allowed 'i') { $help += 'i=Interactive' }
        if (Test-PhaseOption $allowed 'n') { $help += 'n=NoResize' }
        if (Test-PhaseOption $allowed 'r') { $help += 'r=RefreshUrls' }
        if (Test-PhaseOption $allowed 'w') { $help += 'w=window size' }
        if (Test-PhaseOption $allowed 'c') { $help += 'c=crop px' }
        if (Test-PhaseOption $allowed 't') { $help += 't=target IDs' }
        if (Test-PhaseOption $allowed 'a') { $help += 'a=review cursor cell' }
        if (Test-PhaseOption $allowed 'd') { $help += 'd=Clone SourceDir' }
        if (Test-PhaseOption $allowed 'j') { $help += 'j=J4 BaseDir' }
        if (Test-PhaseOption $allowed 'b') { $help += 'b=BizCodes' }
        if (Test-PhaseOption $allowed 'h') { $help += 'h=HostSystemTypes' }
        if (Test-PhaseOption $allowed 'e') { $help += 'e=DfExePath' }
        if (Test-PhaseOption $allowed 'p') { $help += 'p=ProbeFile' }
        if (Test-PhaseOption $allowed 's') { $help += 's=ProbeSheet' }
        if (Test-PhaseOption $allowed 'k') { $help += 'k=CheckSheet path' }
        Write-Host ("  {0}, Enter=continue" -f ($help -join ', '))
    }

    while ($true) {
        $x = Read-Host 'option'
        if ([string]::IsNullOrWhiteSpace($x)) { break }
        switch -Regex ($x.Trim().ToLower()) {
            '^f$' {
                if (-not (Test-PhaseOption $allowed 'f')) { Write-UnusedOption $PhaseKey 'f' }
                else {
                    $State.Force = -not $State.Force
                    $label = if (Test-PhaseOption $allowed 'apply') { 'Apply(Align)' } else { 'Force' }
                    Write-Host ("  {0,-15}: {1}" -f $label, (To-BoolText $State.Force)) -ForegroundColor DarkGray
                }
            }
            '^-?apply$' {
                if (-not (Test-PhaseOption $allowed 'apply')) { Write-UnusedOption $PhaseKey 'apply' }
                else { $State.Force = $true; Write-Host '  Apply=ON' -ForegroundColor DarkGray }
            }
            '^i$' {
                if (-not (Test-PhaseOption $allowed 'i')) { Write-UnusedOption $PhaseKey 'i' }
                else { $State.Interactive = -not $State.Interactive; Write-Host ("  Interactive    : {0}" -f (To-BoolText $State.Interactive)) -ForegroundColor DarkGray }
            }
            '^n$' {
                if (-not (Test-PhaseOption $allowed 'n')) { Write-UnusedOption $PhaseKey 'n' }
                else { $State.NoResize = -not $State.NoResize; Write-Host ("  NoResize       : {0}" -f (To-BoolText $State.NoResize)) -ForegroundColor DarkGray }
            }
            '^r$' {
                if (-not (Test-PhaseOption $allowed 'r')) { Write-UnusedOption $PhaseKey 'r' }
                else { $State.RefreshUrls = -not $State.RefreshUrls; Write-Host ("  RefreshUrls    : {0}" -f (To-BoolText $State.RefreshUrls)) -ForegroundColor DarkGray }
            }
            '^w$' {
                if (-not (Test-PhaseOption $allowed 'w')) { Write-UnusedOption $PhaseKey 'w' }
                else {
                    $v = Read-Choice 'Window size, e.g. 1050x761' ("{0}x{1}" -f $State.WindowWidth, $State.WindowHeight)
                    if ($v -match '^\s*(\d+)\s*[xX, ]\s*(\d+)\s*$') {
                        $State.WindowWidth  = [int]$Matches[1]
                        $State.WindowHeight = [int]$Matches[2]
                    } else { Write-Host '  invalid size' -ForegroundColor Yellow }
                }
            }
            '^c$' {
                if (-not (Test-PhaseOption $allowed 'c')) { Write-UnusedOption $PhaseKey 'c' }
                else {
                    $v = Read-Choice 'CropPx' ([string]$State.CropPx)
                    if ($v -match '^\d+$') { $State.CropPx = [int]$v }
                    else { Write-Host '  invalid crop px' -ForegroundColor Yellow }
                }
            }
            '^a$' {
                if (-not (Test-PhaseOption $allowed 'a')) { Write-UnusedOption $PhaseKey 'a' }
                else { $State.CursorCell = Read-Choice 'Review cursor cell' $State.CursorCell }
            }
            '^t$' {
                if (-not (Test-PhaseOption $allowed 't')) { Write-UnusedOption $PhaseKey 't' }
                else {
                    $v = Read-Choice 'Target IDs, comma-separated. Empty = all' ($State.TargetIds -join ',')
                    if ([string]::IsNullOrWhiteSpace($v)) { $State.TargetIds = @() }
                    else { $State.TargetIds = @($v -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
                }
            }
            '^d$' {
                if (-not (Test-PhaseOption $allowed 'd')) { Write-UnusedOption $PhaseKey 'd' }
                else { $State.CloneSourceDir = Read-Choice 'Clone SourceDir (external)' $State.CloneSourceDir }
            }
            '^j$' {
                if (-not (Test-PhaseOption $allowed 'j')) { Write-UnusedOption $PhaseKey 'j' }
                else { $State.J4BaseDir = Read-Choice 'J4 BaseDir (Align baseline)' $State.J4BaseDir }
            }
            '^b$' {
                if (-not (Test-PhaseOption $allowed 'b')) { Write-UnusedOption $PhaseKey 'b' }
                else {
                    $v = Read-Choice 'BizCodes, comma-separated. Empty = use TO_code/FROM_code' ($State.BizCodes -join ',')
                    if ([string]::IsNullOrWhiteSpace($v)) { $State.BizCodes = @() }
                    else { $State.BizCodes = @($v -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
                }
            }
            '^h$' {
                if (-not (Test-PhaseOption $allowed 'h')) { Write-UnusedOption $PhaseKey 'h' }
                else {
                    $v = Read-Choice 'HostSystemTypes, comma-separated (e.g. HOST,MF). Empty = auto-detect' ($State.HostSystemTypes -join ',')
                    if ([string]::IsNullOrWhiteSpace($v)) { $State.HostSystemTypes = @() }
                    else { $State.HostSystemTypes = @($v -split ',' | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ }) }
                }
            }
            '^e$' {
                if (-not (Test-PhaseOption $allowed 'e')) { Write-UnusedOption $PhaseKey 'e' }
                else { $State.DfExePath = Read-Choice 'DfExePath' $State.DfExePath }
            }
            '^p$' {
                if (-not (Test-PhaseOption $allowed 'p')) { Write-UnusedOption $PhaseKey 'p' }
                else { $State.ProbeFile = Read-Choice 'Probe Excel file path' $State.ProbeFile }
            }
            '^s$' {
                if (-not (Test-PhaseOption $allowed 's')) { Write-UnusedOption $PhaseKey 's' }
                else { $State.ProbeSheet = Read-Choice 'Probe sheet name. Empty = all/first' $State.ProbeSheet }
            }
            '^k$' {
                if (-not (Test-PhaseOption $allowed 'k')) { Write-UnusedOption $PhaseKey 'k' }
                else { $State.CheckSheetPath = Read-Choice 'Review check sheet .xlsx path. Empty = use config' $State.CheckSheetPath }
            }
            default { Write-Host '  unknown option' -ForegroundColor Yellow }
        }
    }
}

function Show-PlannedPhase([string]$PhaseKey) {
    switch ($PhaseKey) {
        default {
            Write-Host ("  Phase '{0}' is not yet implemented." -f $PhaseKey) -ForegroundColor DarkGray
        }
    }
}

function Invoke-ToolPhase([string]$PhaseKey, [hashtable]$Config, [hashtable]$State) {
    $common = Resolve-ToolPath $Config 'Common'
    $base = @{ WorkDir = $State.WorkDir; Owner = $State.Owner }

    if ($PhaseKey -eq 'Mapping') {
        $p = Resolve-ToolPath $Config 'GenerateMapping'
        $args = $base.Clone()
        if ($State.Force) { $args['Force'] = $true }
        Write-Host ("[RUN] {0}" -f (Split-Path $p -Leaf)) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'ExcelSnap') {
        $p = Resolve-ToolPath $Config 'Excel'
        $args = $base.Clone()
        if ($State.Force) { $args['Force'] = $true }
        Write-Host ("[RUN] {0}" -f (Split-Path $p -Leaf)) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'GiftHmSnap' -or $PhaseKey -eq 'GfixHmSnap') {
        $p = Resolve-ToolPath $Config 'Hm'
        $stage = if ($PhaseKey -eq 'GiftHmSnap') { 'GIFT' } else { 'GFIX' }
        $args = $base.Clone()
        $args['Stage'] = $stage
        $args['CropPx'] = $State.CropPx
        $args['WindowWidth'] = $State.WindowWidth
        $args['WindowHeight'] = $State.WindowHeight
        $args['ActionWaitMs'] = $Config.Timing.ActionWaitMs
        $args['ResultWaitSec'] = $Config.Timing.ResultWaitSec
        $args['TabsToCorrelid'] = $Config.Hm.TabsToCorrelid
        $args['TabsBackFromSearch'] = $Config.Hm.TabsBackFromSearch
        $args['TabsBackToInput'] = $Config.Hm.TabsBackToInput
        $args['CommonScript'] = $common
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        if ($State.Interactive) { $args['Interactive'] = $true }
        if ($State.NoResize) { $args['NoResize'] = $true }
        Write-Host ("[RUN] HmSnap {0}" -f $stage) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'GiftMqSnap') {
        $p = Resolve-ToolPath $Config 'Mq'
        $args = $base.Clone()
        $args['CropPx'] = $State.CropPx
        $args['WindowWidth'] = $State.WindowWidth
        $args['WindowHeight'] = $State.WindowHeight
        $args['ActionWaitMs'] = $Config.Timing.ActionWaitMs
        $args['ResultWaitSec'] = $Config.Timing.ResultWaitSec
        $args['TabsToInquiry'] = $Config.Mq.TabsToInquiry
        $args['TabsToCorrelid'] = $Config.Mq.TabsToCorrelid
        $args['CommonScript'] = $common
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        if ($State.Interactive) { $args['Interactive'] = $true }
        if ($State.NoResize) { $args['NoResize'] = $true }
        Write-Host '[RUN] MqSnap GIFT' -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'GiftJenkins' -or $PhaseKey -eq 'GfixJenkins' -or $PhaseKey -eq 'GiftJenkinsNoFile') {
        $p = Resolve-ToolPath $Config 'Jenkins'
        $mode = switch ($PhaseKey) {
            'GiftJenkins'       { 'GiftRecv' }
            'GfixJenkins'       { 'GfixRecv' }
            'GiftJenkinsNoFile' { 'NoGfix' }
        }
        $args = $base.Clone()
        $args['Mode'] = $mode
        $args['CropPx'] = $State.CropPx
        $args['WindowWidth'] = $State.WindowWidth
        $args['WindowHeight'] = $State.WindowHeight
        $args['ActionWaitMs'] = $Config.Timing.ActionWaitMs
        $args['ResultWaitMs'] = $Config.Timing.ResultWaitMs
        $args['CommonScript'] = $common
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        if ($State.Interactive) { $args['Interactive'] = $true }
        if ($State.NoResize) { $args['NoResize'] = $true }
        if ($State.RefreshUrls) { $args['RefreshUrls'] = $true }
        Write-Host ("[RUN] JenkinsSnap {0}" -f $mode) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'Crop') {
        $p = Resolve-ToolPath $Config 'Crop'
        $dir = Join-Path $State.WorkDir $Config.Paths.SnapDir
        $args = @{ Dir = $dir; CropPx = $State.CropPx; Recurse = $true }
        if ($State.Force) { $args['Force'] = $true }
        Write-Host ("[RUN] Crop-Snap {0}" -f $dir) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'ReviewGift' -or $PhaseKey -eq 'ReviewGfix' -or $PhaseKey -eq 'ReviewDf' -or $PhaseKey -eq 'ReviewEvidence') {
        $p  = Resolve-ToolPath $Config 'Review'
        $eh = Resolve-ToolPath $Config 'ExcelHelpers'
        $bit = switch ($PhaseKey) {
            'ReviewGift'     { 1 }
            'ReviewGfix'     { 2 }
            'ReviewDf'       { 4 }
            'ReviewEvidence' { 7 }
        }
        $args = $base.Clone()
        $args['EvidenceDir'] = $State.EvidenceDir
        $args['CursorCell'] = $State.CursorCell
        $args['ReviewField'] = $Config.Review.Field
        $args['ReviewBit'] = $bit
        $args['ExcelHelpersScript'] = $eh
        if ($Config.Review.SaveWaitMs) { $args['SaveWaitMs'] = [int]$Config.Review.SaveWaitMs }
        if ($Config.Review.Maximize) { $args['Maximize'] = $true }
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        if ($State.DryRun) { $args['DryRun'] = $true }
        Write-Host ("[RUN] {0} bit={1}" -f (Split-Path $p -Leaf), $bit) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'Clone') {
        $p = Resolve-ToolPath $Config 'Clone'
        $args = $base.Clone()
        if (-not [string]::IsNullOrWhiteSpace($State.CloneSourceDir)) { $args['SourceDir'] = $State.CloneSourceDir }
        if ($State.BizCodes.Count -gt 0) { $args['BizCodes'] = $State.BizCodes }
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        Write-Host '[RUN] Clone' -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'ReplaceGift' -or $PhaseKey -eq 'ReplaceGfix' -or $PhaseKey -eq 'ReplaceDf') {
        $p  = Resolve-ToolPath $Config 'Replace'
        $eh = Resolve-ToolPath $Config 'ExcelHelpers'
        $mode = switch ($PhaseKey) {
            'ReplaceGift' { 'Gift' }
            'ReplaceGfix' { 'Gfix' }
            'ReplaceDf'   { 'Df' }
        }
        $args = $base.Clone()
        $args['Mode'] = $mode
        $args['CommonScript'] = $common
        $args['ExcelHelpersScript'] = $eh
        if ($Config.Replace) {
            if ($Config.Replace.BlankRowsBetween) { $args['BlankRowsBetween'] = [int]$Config.Replace.BlankRowsBetween }
            if (-not [string]::IsNullOrWhiteSpace([string]$Config.Replace.GiftNoGfixLabel)) {
                $args['GiftNoGfixLabel'] = [string]$Config.Replace.GiftNoGfixLabel
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$Config.Replace.GfixLogLabel)) {
                $args['GfixLogLabel'] = [string]$Config.Replace.GfixLogLabel
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$Config.Replace.GfixLogTodoText)) {
                $args['GfixLogTodoText'] = [string]$Config.Replace.GfixLogTodoText
            }
        }
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        Write-Host ("[RUN] Replace -Mode {0}" -f $mode) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'GfixLogDownload') {
        $p = Resolve-ToolPath $Config 'GfixLogDownload'
        $args = $base.Clone()
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        Write-Host '[RUN] GfixLogDownload' -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'DfSnap') {
        $p = Resolve-ToolPath $Config 'DfSnap'
        $args = $base.Clone()
        # CLI -DfExePath overrides config
        if (-not [string]::IsNullOrWhiteSpace($State.DfExePath)) {
            $args['DfExePath'] = $State.DfExePath
        }
        if ($Config.Df) {
            if (-not $args.ContainsKey('DfExePath') -and -not [string]::IsNullOrWhiteSpace([string]$Config.Df.ExePath)) {
                $args['DfExePath'] = [string]$Config.Df.ExePath
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$Config.Df.GiftDataDir)) { $args['GiftDataDir']   = [string]$Config.Df.GiftDataDir }
            if (-not [string]::IsNullOrWhiteSpace([string]$Config.Df.GfixDataDir)) { $args['GfixDataDir']   = [string]$Config.Df.GfixDataDir }
            if (-not [string]::IsNullOrWhiteSpace([string]$Config.Df.FilePattern)) { $args['FilePattern']   = [string]$Config.Df.FilePattern }
            if ($Config.Df.LoadWaitSec)  { $args['LoadWaitSec']  = [int]$Config.Df.LoadWaitSec }
            if (-not [string]::IsNullOrWhiteSpace([string]$Config.Df.CaptureMode)) { $args['CaptureMode']   = [string]$Config.Df.CaptureMode }
            foreach ($k in @('RegionX','RegionY','RegionWidth','RegionHeight','CropLeft','CropTop','CropRight','CropBottom')) {
                if ($Config.Df.ContainsKey($k) -and $null -ne $Config.Df[$k]) { $args[$k] = [int]$Config.Df[$k] }
            }
        }
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force)  { $args['Force']  = $true }
        if ($State.DryRun) { $args['DryRun'] = $true }
        Write-Host '[RUN] DfSnap' -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'Align') {
        $p  = Resolve-ToolPath $Config 'Align'
        $eh = Resolve-ToolPath $Config 'ExcelHelpers'
        $args = $base.Clone()
        $args['ExcelHelpersScript'] = $eh
        if ($Config.Align) {
            $j4BaseDir = Resolve-AlignJ4BaseDir $Config $State
            if (-not [string]::IsNullOrWhiteSpace($j4BaseDir)) { $args['J4BaseDir'] = $j4BaseDir }
            $hostTypes = @($State.HostSystemTypes)
            if ($hostTypes.Count -eq 0 -and @($Config.Align.HostSystemTypes).Count -gt 0) {
                $hostTypes = [string[]]$Config.Align.HostSystemTypes
            }
            if ($hostTypes.Count -gt 0) { $args['HostSystemTypes'] = $hostTypes }
        }
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        # Align defaults to a read-only DryRun report; -Force opts into -Apply.
        if ($State.Force) { $args['Apply'] = $true }
        Write-Host '[RUN] Align' -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'WatchProgress') {
        $p = Resolve-ToolPath $Config 'WatchProgress'
        $args = $base.Clone()
        Write-Host '[RUN] Watch-MappingProgress' -ForegroundColor Green
        if ($State.DryRun) { $args['Once'] = $true }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'MarkGfixLog') {
        $p  = Resolve-ToolPath $Config 'MarkGfixLog'
        $eh = Resolve-ToolPath $Config 'ExcelHelpers'
        $args = $base.Clone()
        $args['ExcelHelpersScript'] = $eh
        if ($Config.GfixLog) {
            if (-not [string]::IsNullOrWhiteSpace([string]$Config.GfixLog.LogAnchor))       { $args['LogAnchor']         = [string]$Config.GfixLog.LogAnchor }
            if (-not [string]::IsNullOrWhiteSpace([string]$Config.GfixLog.CommandPattern))  { $args['CommandPattern']    = [string]$Config.GfixLog.CommandPattern }
            if ($Config.GfixLog.HighlightColor)    { $args['HighlightColor']    = [long]$Config.GfixLog.HighlightColor }
            if ($Config.GfixLog.HighlightColStart) { $args['HighlightColStart'] = [int]$Config.GfixLog.HighlightColStart }
            if ($Config.GfixLog.HighlightColEnd)   { $args['HighlightColEnd']   = [int]$Config.GfixLog.HighlightColEnd }
        }
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force)  { $args['Force']  = $true }
        if ($State.DryRun) { $args['DryRun'] = $true }
        Write-Host '[RUN] MarkGfixLog' -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'MarkGift' -or $PhaseKey -eq 'MarkGfix' -or $PhaseKey -eq 'MarkDf') {
        $p  = Resolve-ToolPath $Config 'Mark'
        $eh = Resolve-ToolPath $Config 'ExcelHelpers'
        $mode = switch ($PhaseKey) {
            'MarkGift' { 'Gift' }
            'MarkGfix' { 'Gfix' }
            'MarkDf'   { 'Df' }
        }
        $args = $base.Clone()
        $args['Mode'] = $mode
        $args['CommonScript'] = $common
        $args['ExcelHelpersScript'] = $eh
        if ($Config.Mark) {
            if ($Config.Mark.Boxes)      { $args['BoxesConfig'] = $Config.Mark.Boxes }
            if ($Config.Mark.NamePrefix) { $args['NamePrefix']  = [string]$Config.Mark.NamePrefix }
            if ($Config.Mark.LineWeight) { $args['LineWeight']  = [double]$Config.Mark.LineWeight }
        }
        # GFIX log yellow-highlight is folded into MarkGfix (no separate phase).
        # Pass the GfixLog settings through for -Mode Gfix.
        if ($mode -eq 'Gfix' -and $Config.GfixLog) {
            if (-not [string]::IsNullOrWhiteSpace([string]$Config.GfixLog.LogAnchor))      { $args['GfixLogAnchor']         = [string]$Config.GfixLog.LogAnchor }
            if (-not [string]::IsNullOrWhiteSpace([string]$Config.GfixLog.CommandPattern)) { $args['GfixLogCommandPattern']  = [string]$Config.GfixLog.CommandPattern }
            if ($Config.GfixLog.HighlightColor)    { $args['GfixLogHighlightColor'] = [long]$Config.GfixLog.HighlightColor }
            if ($Config.GfixLog.HighlightColStart) { $args['GfixLogColStart']       = [int]$Config.GfixLog.HighlightColStart }
            if ($Config.GfixLog.HighlightColEnd)   { $args['GfixLogColEnd']         = [int]$Config.GfixLog.HighlightColEnd }
        }
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        Write-Host ("[RUN] Mark -Mode {0}" -f $mode) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'CheckSheet') {
        $p  = Resolve-ToolPath $Config 'FillCheckSheet'
        $eh = Resolve-ToolPath $Config 'ExcelHelpers'
        $args = $base.Clone()
        $args['ExcelHelpersScript'] = $eh
        if ($Config.CheckSheet) {
            $cs = $Config.CheckSheet
            $csPath = $State.CheckSheetPath
            if ([string]::IsNullOrWhiteSpace($csPath)) { $csPath = [string]$cs.Path }
            if (-not [string]::IsNullOrWhiteSpace($csPath))               { $args['CheckSheetPath'] = $csPath }
            if (-not [string]::IsNullOrWhiteSpace([string]$cs.SheetName)) { $args['SheetName']      = [string]$cs.SheetName }
            if (-not [string]::IsNullOrWhiteSpace([string]$cs.Language))  { $args['Language']       = [string]$cs.Language }
            if (-not [string]::IsNullOrWhiteSpace([string]$cs.Phase))     { $args['Phase']          = [string]$cs.Phase }
            if (-not [string]::IsNullOrWhiteSpace([string]$cs.DateFormat)){ $args['DateFormat']     = [string]$cs.DateFormat }
            foreach ($k in @('ColNo','ColDate','ColLang','ColResourceId','ColPhase','ColTarget','ColOwner','ColViewer')) {
                if ($cs.ContainsKey($k) -and $null -ne $cs[$k]) { $args[$k] = [int]$cs[$k] }
            }
        }
        if ($Config.Reviewer -and -not [string]::IsNullOrWhiteSpace([string]$Config.Reviewer.ShortName)) {
            $args['Viewer'] = [string]$Config.Reviewer.ShortName
        }
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        Write-Host '[RUN] FillCheckSheet' -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'DeliverMail') {
        $p = Resolve-ToolPath $Config 'DeliverMail'
        $args = $base.Clone()
        if ($Config.Mail) {
            $m = $Config.Mail
            if (-not [string]::IsNullOrWhiteSpace([string]$m.From))             { $args['From']             = [string]$m.From }
            if (-not [string]::IsNullOrWhiteSpace([string]$m.Phase))            { $args['Phase']            = [string]$m.Phase }
            if (-not [string]::IsNullOrWhiteSpace([string]$m.SubjectTemplate))  { $args['SubjectTemplate']  = [string]$m.SubjectTemplate }
            if ($m.BodyLines)                                                  { $args['BodyLines']        = [string[]]$m.BodyLines }
            if (-not [string]::IsNullOrWhiteSpace([string]$m.EvidenceFolder))   { $args['EvidenceFolder']   = [string]$m.EvidenceFolder }
            if (-not [string]::IsNullOrWhiteSpace([string]$m.CheckSheetFolder)) { $args['CheckSheetFolder'] = [string]$m.CheckSheetFolder }
            if (-not [string]::IsNullOrWhiteSpace([string]$m.CheckSheetFile))   { $args['CheckSheetFile']   = [string]$m.CheckSheetFile }
        }
        if ($Config.Reviewer) {
            $rv = $Config.Reviewer
            if (-not [string]::IsNullOrWhiteSpace([string]$rv.Address))     { $args['ReviewerAddress'] = [string]$rv.Address }
            if (-not [string]::IsNullOrWhiteSpace([string]$rv.DisplayName)) { $args['ReviewerDisplay'] = [string]$rv.DisplayName }
            if (-not [string]::IsNullOrWhiteSpace([string]$rv.ShortName))   { $args['ReviewerShort']   = [string]$rv.ShortName }
        }
        $args['EvidenceDir'] = $State.EvidenceDir
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        Write-Host '[RUN] DeliverMail' -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'RepairMapping') {
        $mp = Get-MappingPath $Config $State.WorkDir $State.Owner
        Write-Host '[RUN] RepairMapping' -ForegroundColor Green
        Write-Host ("  Mapping: {0}" -f $mp)
        $added = Ensure-PhaseColumns $Config $mp
        if ($added -eq 0) { Write-Host '  no missing columns. mapping unchanged.' -ForegroundColor Green }
        return
    }

    if ($PhaseKey -eq 'ProbeShapes') {
        $p  = Resolve-ToolPath $Config 'Probe'
        $eh = Resolve-ToolPath $Config 'ExcelHelpers'
        $f  = $State.ProbeFile
        if ([string]::IsNullOrWhiteSpace($f)) {
            $f = Read-Choice 'Path to Excel file to probe' ''
        }
        if ([string]::IsNullOrWhiteSpace($f)) {
            Write-Host '[INFO] no file provided, cancelled.' -ForegroundColor Yellow
            return
        }
        $args = @{ File = $f; ExcelHelpersScript = $eh }
        if (-not [string]::IsNullOrWhiteSpace($State.ProbeSheet)) { $args['Sheet'] = $State.ProbeSheet }
        Write-Host ("[RUN] Probe-Shapes  {0}" -f $f) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'Validate') {
        $p = Resolve-ToolPath $Config 'Validate'
        $args = $base.Clone()
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        Write-Host '[RUN] Validate' -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'Comments') {
        $mp = Get-MappingPath $Config $State.WorkDir $State.Owner
        Write-Host '[RUN] Review comments' -ForegroundColor Green
        Write-Host ("  Mapping: {0}" -f $mp)
        if (-not (Test-Path -LiteralPath $mp)) {
            Write-Host ("  mapping not found: {0}" -f $mp) -ForegroundColor Yellow; return
        }
        $rows  = @(Import-Csv -LiteralPath $mp -Encoding UTF8)
        $first = $rows | Select-Object -First 1
        if ($null -eq $first -or -not ($first.PSObject.Properties.Name -contains 'ReviewComment')) {
            Write-Host '  no ReviewComment column yet (no comments recorded).' -ForegroundColor DarkGray; return
        }
        $withComments = @($rows |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.ReviewComment) } |
            Group-Object Excel_NAME | Sort-Object Name)
        if ($withComments.Count -eq 0) {
            Write-Host '  (no comments recorded)' -ForegroundColor DarkGray; return
        }
        Write-Host ''
        Write-Host ("===== Review comments ({0} workbook(s)) =====" -f $withComments.Count) -ForegroundColor Cyan
        foreach ($g in $withComments) {
            $c = [string]($g.Group | Select-Object -First 1).ReviewComment
            Write-Host ("  {0,-28} {1}" -f $g.Name, $c)
        }
        return
    }

    if ($PhaseKey -eq 'Status') { return }
    throw "Unknown phase: $PhaseKey"
}

$Config = Load-VerifyConfig $ConfigPath
$ResolvedPhase = Resolve-Phase $Config $Phase
if ($Help.IsPresent -or $ResolvedPhase -eq 'Help') { Show-VerifyHelp $Config; return }

$sessionPath = Join-Path $PSScriptRoot 'verify_session.json'
$session = Load-Session $sessionPath

if ([string]::IsNullOrWhiteSpace($Owner)) {
    if ($session.ContainsKey('Owner') -and -not [string]::IsNullOrWhiteSpace([string]$session['Owner'])) { $Owner = [string]$session['Owner'] }
    else { $Owner = [string]$Config.DefaultOwner }
}

if ([string]::IsNullOrWhiteSpace($WorkDir)) {
    $candidate = ''
    if ($session.ContainsKey('WorkDir')) { $candidate = [string]$session['WorkDir'] }
    if ([string]::IsNullOrWhiteSpace($candidate)) { $candidate = [string]$Config.DefaultWorkDir }
    if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
        $WorkDir = Read-Choice 'WorkDir path' $candidate
    } else {
        $WorkDir = Read-Host 'WorkDir path'
    }
}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { throw 'WorkDir is empty.' }
if (-not (Test-Path -LiteralPath $WorkDir)) { throw "WorkDir not found: $WorkDir" }

if ($WindowWidth -le 0)  { $WindowWidth  = [int]$Config.Window.Width }
if ($WindowHeight -le 0) { $WindowHeight = [int]$Config.Window.Height }
if ($CropPx -lt 0)       { $CropPx       = [int]$Config.Window.CropPx }
if ([string]::IsNullOrWhiteSpace($CursorCell)) { $CursorCell = [string]$Config.Review.CursorCell }
if ([string]::IsNullOrWhiteSpace($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir ([string]$Config.Review.EvidenceDir) }
elseif (-not [System.IO.Path]::IsPathRooted($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir $EvidenceDir }

if ([string]::IsNullOrWhiteSpace($CloneSourceDir) -and $session.ContainsKey('CloneSourceDir')) {
    $CloneSourceDir = [string]$session['CloneSourceDir']
}
if ([string]::IsNullOrWhiteSpace($J4BaseDir) -and $session.ContainsKey('J4BaseDir')) {
    $J4BaseDir = [string]$session['J4BaseDir']
}
if ([string]::IsNullOrWhiteSpace($J4BaseDir) -and -not [string]::IsNullOrWhiteSpace($CloneSourceDir)) {
    $J4BaseDir = $CloneSourceDir
}
if ([string]::IsNullOrWhiteSpace($CheckSheetPath) -and $session.ContainsKey('CheckSheetPath')) {
    $CheckSheetPath = [string]$session['CheckSheetPath']
}

$flatTargets = @()
foreach ($rawId in @($TargetIds)) {
    if ($null -eq $rawId) { continue }
    foreach ($part in ($rawId.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $flatTargets += $v }
    }
}
$TargetIds = @($flatTargets | Select-Object -Unique)

$flatBiz = @()
foreach ($raw in @($BizCodes)) {
    if ($null -eq $raw) { continue }
    foreach ($part in ($raw.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $flatBiz += $v }
    }
}
$BizCodes = @($flatBiz | Select-Object -Unique)

$state = @{
    WorkDir         = $WorkDir
    Owner           = $Owner
    WindowWidth     = $WindowWidth
    WindowHeight    = $WindowHeight
    CropPx          = $CropPx
    EvidenceDir     = $EvidenceDir
    CursorCell      = $CursorCell
    TargetIds       = $TargetIds
    CloneSourceDir  = $CloneSourceDir
    J4BaseDir       = $J4BaseDir
    BizCodes        = $BizCodes
    HostSystemTypes = @()
    ProbeFile       = $ProbeFile
    ProbeSheet      = $ProbeSheet
    DfExePath       = $DfExePath
    CheckSheetPath  = $CheckSheetPath
    Force           = [bool]$Force.IsPresent
    Interactive     = [bool]$Interactive.IsPresent
    NoResize        = [bool]$NoResize.IsPresent
    RefreshUrls     = [bool]$RefreshUrls.IsPresent
    DryRun          = [bool]$DryRun.IsPresent
}

$session['WorkDir'] = $WorkDir
$session['Owner'] = $Owner
$session['WindowWidth'] = $WindowWidth
$session['WindowHeight'] = $WindowHeight
$session['CropPx'] = $CropPx
$session['EvidenceDir'] = $EvidenceDir
$session['CursorCell'] = $CursorCell
$session['CloneSourceDir'] = $CloneSourceDir
$session['J4BaseDir'] = $J4BaseDir
$session['CheckSheetPath'] = $CheckSheetPath
Save-Session $sessionPath $session

$mappingPath = Get-MappingPath $Config $WorkDir $Owner

# Auto-repair: ensure every PhaseOrder field has a column in the mapping.
# Safe — never modifies existing data, only adds missing columns with '0'.
if (Test-Path -LiteralPath $mappingPath) {
    Ensure-PhaseColumns $Config $mappingPath | Out-Null
}

$mappingRows = Load-MappingSafe $mappingPath

Write-Host ''
Write-Host '===== VerifyTool =====' -ForegroundColor Green
Write-Host ("  WorkDir        : {0}" -f $WorkDir)
Write-Host ("  Owner          : {0}" -f $Owner)
Write-Host ("  Mapping        : {0}" -f $mappingPath)
Write-Host ("  EvidenceDir    : {0}" -f $EvidenceDir)
Write-Host ("  Window         : {0}x{1}, CropPx={2}" -f $WindowWidth, $WindowHeight, $CropPx)
if (-not [string]::IsNullOrWhiteSpace($CloneSourceDir)) {
    Write-Host ("  CloneSourceDir : {0}" -f $CloneSourceDir)
}
if (-not [string]::IsNullOrWhiteSpace($J4BaseDir)) {
    Write-Host ("  J4BaseDir      : {0}" -f $J4BaseDir)
}
Write-Host ("  Phase          : {0}" -f $ResolvedPhase)

if ($ResolvedPhase -ne 'Menu') {
    if ($ResolvedPhase -eq 'Status') { Show-Status $Config $mappingRows; return }
    Invoke-ToolPhase $ResolvedPhase $Config $state
    return
}

while ($true) {
    $mappingRows = Load-MappingSafe $mappingPath
    Show-Status $Config $mappingRows
    $rec = Get-RecommendPhase $Config $mappingRows

    Write-Host ''
    Write-Host ("Recommended next: {0}" -f $rec) -ForegroundColor Green
    Write-Host ''
    Write-Host 'Choose phase:' -ForegroundColor Cyan
    $idx = 1
    $menu = @{}
    foreach ($p in $Config.PhaseOrder) {
        $status = if ($p.Status -eq 'planned') { ' planned' } elseif ($p.Status -eq 'legacy') { ' legacy' } else { '' }
        $bv = Get-PhaseBit $p
        $bvtxt = if ($bv -gt 0) { (' bit={0}' -f $bv) } else { '' }
        Write-Host ("  {0,2}  {1,-20} {2}{3}{4}" -f $idx, $p.Key, $p.Label, $bvtxt, $status)
        $menu[[string]$idx] = [string]$p.Key
        $idx++
    }
    Write-Host '   s  Status only'
    Write-Host '   h  Help'
    Write-Host '   q  Quit'

    $ans = Read-Choice 'phase' $rec
    $lower = $ans.Trim().ToLower()
    if ($lower -eq 'q') { break }
    if ($lower -eq 'h') { Show-VerifyHelp $Config; continue }
    if ($lower -eq 's') { continue }

    if ($menu.ContainsKey($ans.Trim())) { $key = $menu[$ans.Trim()] }
    else { $key = Resolve-Phase $Config $ans.Trim() }

    Ask-RunOptions $state $key
    Invoke-ToolPhase $key $Config $state

    $session['WorkDir'] = $state.WorkDir
    $session['Owner'] = $state.Owner
    $session['WindowWidth'] = $state.WindowWidth
    $session['WindowHeight'] = $state.WindowHeight
    $session['CropPx'] = $state.CropPx
    $session['EvidenceDir'] = $state.EvidenceDir
    $session['CursorCell'] = $state.CursorCell
    $session['CloneSourceDir'] = $state.CloneSourceDir
    $session['J4BaseDir'] = $state.J4BaseDir
    $session['CheckSheetPath'] = $state.CheckSheetPath
    Save-Session $sessionPath $session

    Write-Host ''
    Write-Host 'Back to VerifyTool menu. Enter to refresh / q to quit : ' -ForegroundColor Magenta -NoNewline
    $again = Read-Host
    if ($again -eq 'q') { break }
}