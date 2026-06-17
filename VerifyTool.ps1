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

    # Mapping
    [string]$FromBizCode = '',
    [int]$WbsStartRow = 0,
    [int]$WbsEndRow = 0,
    [string[]]$CorrelIdsM = @(),
    [string[]]$JobNames = @(),
    [string[]]$ExcelNames = @(),

    [int]$WindowWidth  = 0,
    [int]$WindowHeight = 0,
    [int]$CropPx       = -1,

    [string[]]$TargetIds = @(),
    [string]$EvidenceDir = '',
    [string]$CursorCell  = '',

    # Clone / Replace
    [string]$CloneSourceDir = '',
    [string]$J4BaseDir      = '',
    [string]$ExcelPrefix    = '',
    [string[]]$BizCodes     = @(),

    # DfSnap override (takes precedence over VerifyConfig.psd1 -> Df.ExePath)
    [string]$DfExePath = '',

    # ProbeShapes
    [string]$ProbeFile  = '',
    [string]$ProbeSheet = '',

    # CheckSheet
    [string]$CheckSheetPath = '',

    [switch]$Force,
    [switch]$Ocr,
    [switch]$Interactive,
    [switch]$NoResize,
    [switch]$RefreshUrls,
    [switch]$DryRun,
    [switch]$MoveData,
    [switch]$AllowTempMapping,
    [switch]$Add,
    [switch]$Help,

    [string]$ConfigPath = ''
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

# Per-work-folder JSON config overlay helpers (pure; unit-tested). No param().
. (Join-Path $PSScriptRoot 'ConfigOverlay.ps1')

function Read-Choice([string]$Prompt, [string]$Default = '') {
    if ([string]::IsNullOrWhiteSpace($Default)) { return (Read-Host $Prompt) }
    $v = Read-Host ("{0} [{1}]" -f $Prompt, $Default)
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v
}

function ConvertTo-TargetIdSelection([object]$RawValue) {
    $items = @()
    foreach ($raw in @($RawValue)) {
        if ($null -eq $raw) { continue }
        foreach ($part in ($raw.ToString() -split ',')) {
            $v = $part.Trim()
            if ([string]::IsNullOrWhiteSpace($v)) { continue }
            if ($v -ieq 'all') { return @() }
            $items += $v
        }
    }
    return @($items | Select-Object -Unique)
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
    Write-Host '  .\VerifyTool.ps1 -Phase Mapping -Owner 0602 -Force'
    Write-Host '  .\VerifyTool.ps1 -Phase Mapping -Owner 0602 -CorrelIdsM JIDSC02M,JIDSC03M -AllowTempMapping -Force'
    Write-Host '  .\VerifyTool.ps1 -Phase Mapping -Owner 0602 -Add -JobNames CJODJDEU,CJODJDB5   # grow map, keep progress'
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
    Write-Host '  .\VerifyTool.ps1 -Phase SendVsGift          # gather GIFT metadata + SEND/GIFT review (OCR on by default)'
    Write-Host '  .\OcrTool.ps1 -Path <png|dir|wildcard>      # standalone OCR tool (also -Workbook <xlsx>)'
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
    Write-Host '  .\VerifyTool.ps1 -Phase DeliverFiles       # copy evidence Excel + DATA to J4'
    Write-Host '  .\VerifyTool.ps1 -Phase DeliverFiles -MoveData  # move DATA files'
    Write-Host '  .\VerifyTool.ps1 -Phase InitConfig          # write/update per-folder verify_config.json'
    Write-Host '  .\VerifyTool.ps1 -Phase InitConfig -Interactive # grouped config editor (peek/edit/delete/save)'
    Write-Host ''
    Write-Host 'Common options:'
    Write-Host '  -WorkDir <path>       Work folder. If omitted, last used path is remembered.'
    Write-Host '  -Owner <Owner>        mapping_<Owner>.csv owner suffix. No personal default is configured.'
    Write-Host '  -TargetIds A,B        Limit run by Correl_ID_S / Correl_ID_M / JOB_NAME / Excel_NAME.'
    Write-Host '  -CloneSourceDir <p>   External path for Clone (existing evidence files per bizcode).'
    Write-Host '  -J4BaseDir <p>        J4 baseline root for Align. If omitted, config/CloneSourceDir/session is used.'
    Write-Host '  -ExcelPrefix <text>   Evidence workbook filename prefix before _<Excel_NAME>.'
    Write-Host '  -BizCodes A,B         Override bizcode candidate list for Clone.'
    Write-Host '  -Force                Re-run rows whose flag/bit is already set.'
    Write-Host '  Mapping options      owner/-Owner, from/-FromBizCode, range, cm/-CorrelIdsM, jobs/-JobNames, temp.'
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
    Write-Host 'Per-work-folder config overlay:'
    Write-Host '  <WorkDir>\verify_config.json overrides VerifyConfig.psd1 (JSON wins; CLI still wins).'
    Write-Host '  Create or refresh it with:  .\VerifyTool.ps1 -Phase InitConfig'
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

function Import-ConfigOverlay([hashtable]$Config, [string]$WorkDir) {
    # Find the per-work-folder overlay (Paths.OverlayName, default
    # verify_config.json) and deep-merge it over $Config in place. JSON wins.
    # Returns @{ Loaded=<bool>; Path=<string>; Overlay=<hashtable> }.
    $name = [string]$Config.Paths.OverlayName
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'verify_config.json' }
    $path = Join-Path $WorkDir $name
    $result = @{ Loaded = $false; Path = $path; Overlay = @{} }
    if (-not (Test-Path -LiteralPath $path)) { return $result }
    try {
        $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return $result }
        $overlay = ConvertFrom-ConfigJson $raw
        if (($overlay -is [hashtable]) -and ($overlay.Count -gt 0)) {
            Merge-ConfigHashtable $Config $overlay | Out-Null
            $result.Loaded = $true
            $result.Overlay = $overlay
        }
    } catch {
        Write-Host ("[WARN] config overlay not loaded ({0}): {1}" -f $path, $_.Exception.Message) -ForegroundColor Yellow
    }
    return $result
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
        '^Mapping$' { @(
            '  Phase params:',
            '    owner=Owner        -> set mapping owner suffix, e.g. 0602 creates mapping_0602.csv',
            '    from=FromBizCode   -> filter GFIX from biz code',
            '    range              -> WBS row range, e.g. 1275-2250',
            '    cm=Correl_ID_M     -> temp mode: comma-separated M IDs, finds JOB in GFIX',
            '    jobs=JOB_NAME      -> temp mode: comma-separated JOB names',
            '    temp               -> if WBS has no matching rows, ask to create temp mapping from GFIX',
            '    f=Force            -> overwrite mapping_<owner>.csv if it already exists'
        ) }
        '^Align$' { @(
            '  Phase params:',
            '    diff=DiffMode  -> -DiffMode : report only (do NOT replace sheets). Default is force-replace.',
            '    h=HostTypes    -> which FROM_sys / TO_sys column values count as Host (mainframe).',
            '                      e.g. enter: HOST   (check your mapping FROM_sys/TO_sys column for the actual literal)',
            '                      HostToOpen  = send data + GIFT/GFIX send result sheets',
            '                      OpenToOpen / OpenToHost = GIFT+GFIX send sheets + 3 receive sheets',
            '    j=J4BaseDir    -> root folder of J4 baseline workbooks (searched recursively)',
            '    t=TargetIds    -> limit to specific Excel_NAME / Correl_ID / JOB_NAME'
        ) }
        '^InitConfig$' { @(
            '  Writes verify_config.json in the work folder: a JSON overlay that',
            '  overrides VerifyConfig.psd1 for THIS folder only (owner, window,',
            '  mail format, mark boxes, expected-time defaults, ...).',
            '  Edit the file, then re-run any phase. f=Force regenerates (keeps a .bak).'
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
        '^SendVsGift$' { @(
            '  Phase params:',
            '    a=CursorCell   -> fallback cursor cell (default: A3)',
            '    t=TargetIds    -> limit rows',
            '    f=Force        -> re-open rows already marked SendVsGift=1',
            '    o=Ocr          -> toggle OCR on/off (default: on; ok->auto 1, ng->manual prompt)',
            '  Scans DATA\GIFT, writes data\gift_metadata.csv, ensures the SendVsGift',
            '  mapping column, then opens each pending workbook ONCE (rows grouped per',
            '  Excel). The cursor jumps to each Correl_ID_S label in column A of the',
            '  send-data sheet; after every console answer Excel is refocused.',
            '  Enter=mark 1, n=mark 2 (NG), s=skip, q=quit.',
            '  With OCR on: each correl section (pictures between its column-A label',
            '  and the next) is exported + OCR-compared against gift_metadata;',
            '  verdict ok -> auto-mark 1, ng -> auto-mark 2 (listed at the end),',
            '  unknown -> manual prompt. NG rows (=2) stay pending.'
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
        '^DeliverFiles$' { @(
            '  Phase params:',
            '    t=TargetIds    -> limit to specific Excel_NAME / Correl_ID / JOB_NAME',
            '    f=Force        -> re-copy files already marked delivered',
            '    mv=MoveData    -> Move DATA files (delete source). Evidence Excel is always Copied.'
        ) }
        '^InitConfig$' { @(
            '  Phase params:',
            '    i=Interactive  -> grouped editor: peek, edit, delete, save with confirmation',
            '    f=Force        -> keep for compatibility; InitConfig now updates/refreshes existing JSON by default',
            '  NOTE: this phase writes all editable config keys into verify_config.json,',
            '        preserving loaded per-folder values and adding new defaults when the tool changes.'
        ) }
        default { @() }
    }
    foreach ($l in $lines) { Write-Host $l -ForegroundColor DarkGray }
}

function Get-PhaseOptionKeys([string]$PhaseKey) {
    switch -Regex ($PhaseKey) {
        '^Align$' { return @('diff','h','j','t') }
        '^Clone$' { return @('d','b','t','f') }
        '^Replace(Gift|Gfix|Df)$' { return @('t','f') }
        '^Mark(Gift|Gfix|Df)$' { return @('t','f') }
        '^MarkGfixLog$' { return @('t','f') }
        '^SendVsGift$' { return @('a','t','f','o') }
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
        '^DeliverFiles$' { return @('t','f','mv') }
        '^Mapping$' { return @('f','owner','from','range','cm','jobs','ex','temp','add') }
        '^ExcelSnap$' { return @('f') }
        '^InitConfig$' { return @('i','f') }
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

    if (Test-PhaseOption $allowed 'diff') { Write-Host ("  DiffMode(Align): {0}" -f (To-BoolText $State.DiffMode)) }
    elseif (Test-PhaseOption $allowed 'f') { Write-Host ("  Force          : {0}" -f (To-BoolText $State.Force)) }
    if (Test-PhaseOption $allowed 'owner') { Write-Host ("  Owner          : {0}" -f $State.Owner) }
    if (Test-PhaseOption $allowed 'from') { Write-Host ("  FromBizCode    : {0}" -f $(if ([string]::IsNullOrWhiteSpace($State.FromBizCode)) { '(none)' } else { $State.FromBizCode })) }
    if (Test-PhaseOption $allowed 'range') { Write-Host ("  WBS range      : {0}" -f $(if ($State.WbsStartRow -gt 0 -and $State.WbsEndRow -gt 0) { "$($State.WbsStartRow)-$($State.WbsEndRow)" } else { '(full)' })) }
    if (Test-PhaseOption $allowed 'cm') { Write-Host ("  CorrelIdsM     : {0}" -f $(if ($State.CorrelIdsM.Count -gt 0) { $State.CorrelIdsM -join ', ' } else { '(none)' })) }
    if (Test-PhaseOption $allowed 'jobs') { Write-Host ("  JobNames       : {0}" -f $(if ($State.JobNames.Count -gt 0) { $State.JobNames -join ', ' } else { '(none)' })) }
    if (Test-PhaseOption $allowed 'ex') { Write-Host ("  ExcelNames     : {0}" -f $(if ($State.ExcelNames.Count -gt 0) { $State.ExcelNames -join ', ' } else { '(none)' })) }
    if (Test-PhaseOption $allowed 'temp') { Write-Host ("  TempMapping    : {0}" -f (To-BoolText $State.AllowTempMapping)) }
    if (Test-PhaseOption $allowed 'add') { Write-Host ("  Add (merge)    : {0}" -f (To-BoolText $State.AddRows)) }
    if (Test-PhaseOption $allowed 'i') { Write-Host ("  Interactive    : {0}" -f (To-BoolText $State.Interactive)) }
    if (Test-PhaseOption $allowed 'n') { Write-Host ("  NoResize       : {0}" -f (To-BoolText $State.NoResize)) }
    if (Test-PhaseOption $allowed 'r') { Write-Host ("  RefreshUrls    : {0}" -f (To-BoolText $State.RefreshUrls)) }
    if (Test-PhaseOption $allowed 'w') { Write-Host ("  Window         : {0}x{1}" -f $State.WindowWidth, $State.WindowHeight) }
    if (Test-PhaseOption $allowed 'c') { Write-Host ("  CropPx         : {0}" -f $State.CropPx) }
    if (Test-PhaseOption $allowed 'a') { Write-Host ("  CursorCell     : {0}" -f $State.CursorCell) }
    if (Test-PhaseOption $allowed 'o') { Write-Host ("  Ocr            : {0}" -f (To-BoolText $State.Ocr)) }
    if (Test-PhaseOption $allowed 't') { Write-Host ("  TargetIds      : {0}" -f $(if ($State.TargetIds.Count -gt 0) { $State.TargetIds -join ', ' } else { '(all)' })) }
    if (Test-PhaseOption $allowed 'd') { Write-Host ("  CloneSourceDir : {0}" -f $State.CloneSourceDir) }
    if (Test-PhaseOption $allowed 'j') { Write-Host ("  J4BaseDir      : {0}" -f $State.J4BaseDir) }
    if (Test-PhaseOption $allowed 'b') { Write-Host ("  BizCodes       : {0}" -f $(if ($State.BizCodes.Count -gt 0) { $State.BizCodes -join ', ' } else { '(auto)' })) }
    if (Test-PhaseOption $allowed 'h') { Write-Host ("  HostSystemTypes: {0}" -f $(if ($State.HostSystemTypes.Count -gt 0) { $State.HostSystemTypes -join ', ' } else { '(config/auto)' })) }
    if (Test-PhaseOption $allowed 'e') { Write-Host ("  DfExePath      : {0}" -f $State.DfExePath) }
    if (Test-PhaseOption $allowed 'p') { Write-Host ("  ProbeFile      : {0}" -f $State.ProbeFile) }
    if (Test-PhaseOption $allowed 's') { Write-Host ("  ProbeSheet     : {0}" -f $State.ProbeSheet) }
    if (Test-PhaseOption $allowed 'k') { Write-Host ("  CheckSheetPath : {0}" -f $(if ([string]::IsNullOrWhiteSpace($State.CheckSheetPath)) { '(config/prompt)' } else { $State.CheckSheetPath })) }
    if (Test-PhaseOption $allowed 'mv') { Write-Host ("  MoveData       : {0}" -f (To-BoolText $State.MoveData)) }

    Write-Host ''
    if ($allowed.Count -eq 0) {
        Write-Host '  This phase has no interactive options. Enter=continue' -ForegroundColor DarkGray
    } else {
        $help = @()
        if (Test-PhaseOption $allowed 'diff') { $help += 'diff=DiffMode (toggle report-only)' }
        elseif (Test-PhaseOption $allowed 'f') { $help += 'f=Force' }
        if (Test-PhaseOption $allowed 'i') { $help += 'i=Interactive' }
        if (Test-PhaseOption $allowed 'n') { $help += 'n=NoResize' }
        if (Test-PhaseOption $allowed 'r') { $help += 'r=RefreshUrls' }
        if (Test-PhaseOption $allowed 'w') { $help += 'w=window size' }
        if (Test-PhaseOption $allowed 'c') { $help += 'c=crop px' }
        if (Test-PhaseOption $allowed 't') { $help += 't=target IDs' }
        if (Test-PhaseOption $allowed 'a') { $help += 'a=review cursor cell' }
        if (Test-PhaseOption $allowed 'o') { $help += 'o=Ocr toggle' }
        if (Test-PhaseOption $allowed 'd') { $help += 'd=Clone SourceDir' }
        if (Test-PhaseOption $allowed 'j') { $help += 'j=J4 BaseDir' }
        if (Test-PhaseOption $allowed 'b') { $help += 'b=BizCodes' }
        if (Test-PhaseOption $allowed 'h') { $help += 'h=HostSystemTypes' }
        if (Test-PhaseOption $allowed 'e') { $help += 'e=DfExePath' }
        if (Test-PhaseOption $allowed 'p') { $help += 'p=ProbeFile' }
        if (Test-PhaseOption $allowed 's') { $help += 's=ProbeSheet' }
        if (Test-PhaseOption $allowed 'k') { $help += 'k=CheckSheet path' }
        if (Test-PhaseOption $allowed 'mv') { $help += 'mv=MoveData' }
        if (Test-PhaseOption $allowed 'owner') { $help += 'owner=Owner' }
        if (Test-PhaseOption $allowed 'from') { $help += 'from=FromBizCode' }
        if (Test-PhaseOption $allowed 'range') { $help += 'r/range=WBS rows' }
        if (Test-PhaseOption $allowed 'cm') { $help += 'cm=Correl_ID_M list' }
        if (Test-PhaseOption $allowed 'jobs') { $help += 'jobs=JOB_NAME list' }
        if (Test-PhaseOption $allowed 'ex') { $help += 'ex=Excel_NAME list' }
        if (Test-PhaseOption $allowed 'temp') { $help += 'temp=AllowTempMapping' }
        if (Test-PhaseOption $allowed 'add') { $help += 'add=incremental merge' }
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
                if (Test-PhaseOption $allowed 'range') {
                    $defaultRange = if ($State.WbsStartRow -gt 0 -and $State.WbsEndRow -gt 0) { "{0}-{1}" -f $State.WbsStartRow, $State.WbsEndRow } else { '' }
                    $v = Read-Choice 'WBS row range, e.g. 1275-2250. Empty = full WBS scan' $defaultRange
                    if ([string]::IsNullOrWhiteSpace($v)) { $State.WbsStartRow = 0; $State.WbsEndRow = 0 }
                    elseif ($v -match '^\s*(\d+)\s*[-,~ ]\s*(\d+)\s*$') { $State.WbsStartRow = [int]$Matches[1]; $State.WbsEndRow = [int]$Matches[2] }
                    else { Write-Host '  invalid range' -ForegroundColor Yellow }
                    Write-Host ("  WBS range      : {0}" -f $(if ($State.WbsStartRow -gt 0 -and $State.WbsEndRow -gt 0) { "{0}-{1}" -f $State.WbsStartRow, $State.WbsEndRow } else { '(full)' })) -ForegroundColor DarkGray
                } elseif (-not (Test-PhaseOption $allowed 'r')) { Write-UnusedOption $PhaseKey 'r' }
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
                    $currentTargets = if ($State.TargetIds.Count -gt 0) { $State.TargetIds -join ',' } else { 'all' }
                    $v = Read-Choice 'Target IDs, comma-separated. Enter=keep, all=all' $currentTargets
                    $State.TargetIds = @(ConvertTo-TargetIdSelection $v)
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

            '^-?o(cr)?$' {
                # 'o'/'ocr'/'-ocr' -> OCR toggle where the phase supports it;
                # plain 'o' falls through to the owner alias below otherwise.
                if (Test-PhaseOption $allowed 'o') {
                    $State.Ocr = -not $State.Ocr
                    Write-Host ("  Ocr            : {0}" -f (To-BoolText $State.Ocr)) -ForegroundColor DarkGray
                    break
                }
                if ($x.Trim().ToLower() -ne 'o') { Write-UnusedOption $PhaseKey 'ocr'; break }
            }
            '^-?owner$|^o$' {
                if (-not (Test-PhaseOption $allowed 'owner')) { Write-UnusedOption $PhaseKey 'owner' }
                else {
                    $State.Owner = Read-Choice 'Mapping owner suffix (example: 0602)' $State.Owner
                    Write-Host ("  Owner          : {0}" -f $State.Owner) -ForegroundColor DarkGray
                }
            }
            '^-?from(bizcode)?$' {
                if (-not (Test-PhaseOption $allowed 'from')) { Write-UnusedOption $PhaseKey 'from' }
                else {
                    $State.FromBizCode = Read-Choice 'FromBizCode. Empty = no GFIX from-code filter' $State.FromBizCode
                    Write-Host ("  FromBizCode    : {0}" -f $(if ([string]::IsNullOrWhiteSpace($State.FromBizCode)) { '(none)' } else { $State.FromBizCode })) -ForegroundColor DarkGray
                }
            }
            '^-?range$' {
                if (-not (Test-PhaseOption $allowed 'range')) { Write-UnusedOption $PhaseKey 'range' }
                else {
                    $defaultRange = if ($State.WbsStartRow -gt 0 -and $State.WbsEndRow -gt 0) { "{0}-{1}" -f $State.WbsStartRow, $State.WbsEndRow } else { '' }
                    $v = Read-Choice 'WBS row range, e.g. 1275-2250. Empty = full WBS scan' $defaultRange
                    if ([string]::IsNullOrWhiteSpace($v)) { $State.WbsStartRow = 0; $State.WbsEndRow = 0 }
                    elseif ($v -match '^\s*(\d+)\s*[-,~ ]\s*(\d+)\s*$') { $State.WbsStartRow = [int]$Matches[1]; $State.WbsEndRow = [int]$Matches[2] }
                    else { Write-Host '  invalid range' -ForegroundColor Yellow }
                    Write-Host ("  WBS range      : {0}" -f $(if ($State.WbsStartRow -gt 0 -and $State.WbsEndRow -gt 0) { "{0}-{1}" -f $State.WbsStartRow, $State.WbsEndRow } else { '(full)' })) -ForegroundColor DarkGray
                }
            }
            '^-?cm$|^-?correl(idsm|idm)?$' {
                if (-not (Test-PhaseOption $allowed 'cm')) { Write-UnusedOption $PhaseKey 'cm' }
                else {
                    $v = Read-Choice 'Correl_ID_M list, comma-separated (example: JIDSC02M,JIDSC03M). Empty = none' ($State.CorrelIdsM -join ',')
                    if ([string]::IsNullOrWhiteSpace($v)) { $State.CorrelIdsM = @() }
                    else { $State.CorrelIdsM = @($v -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique) }
                }
            }
            '^-?jobs?$|^-?jobnames?$' {
                if (-not (Test-PhaseOption $allowed 'jobs')) { Write-UnusedOption $PhaseKey 'jobs' }
                else {
                    $v = Read-Choice 'JOB_NAME list, comma-separated (example: CJODJDEI,CJODJDB7). Empty = none' ($State.JobNames -join ',')
                    if ([string]::IsNullOrWhiteSpace($v)) { $State.JobNames = @() }
                    else { $State.JobNames = @($v -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique) }
                }
            }
            '^-?ex$|^-?excelnames?$' {
                if (-not (Test-PhaseOption $allowed 'ex')) { Write-UnusedOption $PhaseKey 'ex' }
                else {
                    $v = Read-Choice 'Excel_NAME list, comma-separated (example: LJRVWD64,LJRVWD65). Empty = none' ($State.ExcelNames -join ',')
                    if ([string]::IsNullOrWhiteSpace($v)) { $State.ExcelNames = @() }
                    else { $State.ExcelNames = @($v -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique) }
                }
            }
            '^-?temp$' {
                if (-not (Test-PhaseOption $allowed 'temp')) { Write-UnusedOption $PhaseKey 'temp' }
                else {
                    $State.AllowTempMapping = -not $State.AllowTempMapping
                    Write-Host ("  TempMapping    : {0}" -f (To-BoolText $State.AllowTempMapping)) -ForegroundColor DarkGray
                }
            }
            '^-?add$' {
                if (-not (Test-PhaseOption $allowed 'add')) { Write-UnusedOption $PhaseKey 'add' }
                else {
                    $State.AddRows = -not $State.AddRows
                    Write-Host ("  Add (merge)    : {0}" -f (To-BoolText $State.AddRows)) -ForegroundColor DarkGray
                }
            }
            '^-?diff$' {
                if (-not (Test-PhaseOption $allowed 'diff')) { Write-UnusedOption $PhaseKey 'diff' }
                else {
                    $State.DiffMode = -not $State.DiffMode
                    Write-Host ("  DiffMode(Align): {0}" -f (To-BoolText $State.DiffMode)) -ForegroundColor DarkGray
                }
            }
            '^-?mv$' {
                if (-not (Test-PhaseOption $allowed 'mv')) { Write-UnusedOption $PhaseKey 'mv' }
                else {
                    $State.MoveData = -not $State.MoveData
                    Write-Host ("  MoveData       : {0}" -f (To-BoolText $State.MoveData)) -ForegroundColor DarkGray
                }
            }
            default { Write-Host '  unknown option' -ForegroundColor Yellow }
        }
    }
}


function Copy-ConfigObject([object]$Value) {
    if ($null -eq $Value) { return $null }
    return (ConvertTo-ConfigHashtable ($Value | ConvertTo-Json -Depth 30 | ConvertFrom-Json))
}

function Get-ConfigPathParts([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return @() }
    return @($Path -split '\.' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Test-IntegerText([string]$Text) { return ($Text -match '^\d+$') }

function Get-ConfigValueByPath([object]$Root, [string]$Path) {
    $cur = $Root
    foreach ($part in (Get-ConfigPathParts $Path)) {
        if ($cur -is [System.Collections.IDictionary]) {
            if (-not $cur.ContainsKey($part)) { return $null }
            $cur = $cur[$part]
        } elseif (($cur -is [System.Collections.IList]) -and (Test-IntegerText $part)) {
            $idx = [int]$part
            if ($idx -lt 0 -or $idx -ge $cur.Count) { return $null }
            $cur = $cur[$idx]
        } else { return $null }
    }
    return $cur
}

function Set-ConfigValueByPath([hashtable]$Root, [string]$Path, [object]$Value) {
    $parts = @(Get-ConfigPathParts $Path)
    if ($parts.Count -eq 0) { throw 'Path is required.' }
    $cur = $Root
    for ($i = 0; $i -lt ($parts.Count - 1); $i++) {
        $part = $parts[$i]
        $nextPart = $parts[$i + 1]
        if ($cur -is [System.Collections.IDictionary]) {
            if (-not $cur.ContainsKey($part) -or $null -eq $cur[$part]) {
                if (Test-IntegerText $nextPart) { $cur[$part] = @() } else { $cur[$part] = @{} }
            }
            $cur = $cur[$part]
        } elseif (($cur -is [System.Collections.IList]) -and (Test-IntegerText $part)) {
            $idx = [int]$part
            if ($idx -lt 0 -or $idx -ge $cur.Count) { throw "Array index out of range: $part" }
            $cur = $cur[$idx]
        } else { throw "Cannot traverse path at: $part" }
    }
    $leaf = $parts[$parts.Count - 1]
    if ($cur -is [System.Collections.IDictionary]) { $cur[$leaf] = $Value; return }
    if (($cur -is [System.Collections.IList]) -and (Test-IntegerText $leaf)) {
        $idx = [int]$leaf
        if ($idx -lt 0 -or $idx -ge $cur.Count) { throw "Array index out of range: $leaf" }
        $cur[$idx] = $Value
        return
    }
    throw "Cannot set path: $Path"
}

function Remove-ConfigValueByPath([hashtable]$Root, [string]$Path) {
    $parts = @(Get-ConfigPathParts $Path)
    if ($parts.Count -eq 0) { throw 'Path is required.' }
    $cur = $Root
    $parent = $null
    $parentKey = $null
    for ($i = 0; $i -lt ($parts.Count - 1); $i++) {
        $part = $parts[$i]
        $parent = $cur
        $parentKey = $part
        if ($cur -is [System.Collections.IDictionary]) {
            if (-not $cur.ContainsKey($part)) { return $false }
            $cur = $cur[$part]
        } elseif (($cur -is [System.Collections.IList]) -and (Test-IntegerText $part)) {
            $idx = [int]$part
            if ($idx -lt 0 -or $idx -ge $cur.Count) { return $false }
            $cur = $cur[$idx]
        } else { return $false }
    }
    $leaf = $parts[$parts.Count - 1]
    if ($cur -is [System.Collections.IDictionary]) {
        if (-not $cur.ContainsKey($leaf)) { return $false }
        $cur.Remove($leaf)
        return $true
    }
    if (($cur -is [System.Collections.IList]) -and (Test-IntegerText $leaf)) {
        $idx = [int]$leaf
        if ($idx -lt 0 -or $idx -ge $cur.Count) { return $false }
        $newList = @()
        for ($i = 0; $i -lt $cur.Count; $i++) {
            if ($i -ne $idx) { $newList += ,$cur[$i] }
        }
        if ($null -eq $parent) { throw 'Cannot delete the root array.' }
        if ($parent -is [System.Collections.IDictionary]) { $parent[$parentKey] = $newList; return $true }
        if (($parent -is [System.Collections.IList]) -and (Test-IntegerText $parentKey)) { $parent[[int]$parentKey] = $newList; return $true }
        return $false
    }
    return $false
}

function ConvertFrom-ConfigEditorValue([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    try {
        $parsed = $Text | ConvertFrom-Json
        return (ConvertTo-ConfigHashtable $parsed)
    } catch {
        return $Text
    }
}

function Show-ConfigEditorGroup([hashtable]$Data, [hashtable]$Group) {
    Write-Host ''
    Write-Host ("[{0}] {1}" -f $Group.Key, $Group.Label) -ForegroundColor Cyan
    foreach ($path in @($Group.Paths)) {
        if ($path -eq '*') {
            Write-Host (Get-ConfigOverlayJson $Data)
            return
        }
        $value = Get-ConfigValueByPath $Data $path
        if ($null -eq $value) { continue }
        Write-Host ("--- {0} ---" -f $path) -ForegroundColor DarkGray
        Write-Host (Get-ConfigOverlayJson $value)
    }
}

function Invoke-ConfigOverlayEditor([hashtable]$Data, [string]$DestPath) {
    $groups = @(Get-ConfigOverlayGroups)
    Write-Host ''
    Write-Host '===== InitConfig editor =====' -ForegroundColor Green
    Write-Host 'View by group, edit any JSON path, delete paths, then save with confirmation.' -ForegroundColor DarkGray
    Write-Host 'Path examples: Window.Width, Mail.BodyLines, Mark.Boxes.GIFT_HM.0.OffsetX, PhaseOrder.0.Label' -ForegroundColor DarkGray
    Write-Host 'Value input accepts JSON (true, 123, [..], {..}, "text") or raw text.' -ForegroundColor DarkGray
    Write-Host ("Target: {0}" -f $DestPath) -ForegroundColor DarkGray

    while ($true) {
        Write-Host ''
        Write-Host 'Groups:' -ForegroundColor Cyan
        $menu = @{}
        $idx = 1
        foreach ($g in $groups) {
            Write-Host ("  {0,2}  {1,-7} {2}" -f $idx, $g.Key, $g.Label)
            $menu[[string]$idx] = $g
            $menu[[string]($g.Key)] = $g
            $idx++
        }
        Write-Host '   v  Peek path'
        Write-Host '   e  Edit path'
        Write-Host '   d  Delete path'
        Write-Host '   s  Save'
        Write-Host '   q  Quit without saving'
        $ans = (Read-Host 'config').Trim().ToLower()
        if ([string]::IsNullOrWhiteSpace($ans)) { continue }
        if ($ans -eq 'q') { return $null }
        if ($ans -eq 's') {
            $confirm = (Read-Choice 'Save changes? type YES to write' 'no')
            if ($confirm -ceq 'YES') { return $Data }
            Write-Host '  save cancelled; still in editor.' -ForegroundColor Yellow
            continue
        }
        if ($ans -eq 'v') {
            $path = Read-Choice 'JSON path to peek (empty = all)' ''
            if ([string]::IsNullOrWhiteSpace($path)) { Write-Host (Get-ConfigOverlayJson $Data) }
            else {
                $value = Get-ConfigValueByPath $Data $path
                if ($null -eq $value) { Write-Host '  (not found)' -ForegroundColor Yellow }
                else { Write-Host (Get-ConfigOverlayJson $value) }
            }
            continue
        }
        if ($ans -eq 'e') {
            $path = Read-Choice 'JSON path to edit' ''
            if ([string]::IsNullOrWhiteSpace($path)) { Write-Host '  path is required' -ForegroundColor Yellow; continue }
            $old = Get-ConfigValueByPath $Data $path
            Write-Host 'Current value:' -ForegroundColor DarkGray
            if ($null -eq $old) { Write-Host '  (new path)' -ForegroundColor DarkGray } else { Write-Host (Get-ConfigOverlayJson $old) }
            $raw = Read-Choice 'New value as JSON or raw text. Empty = empty string' ''
            $newValue = ConvertFrom-ConfigEditorValue $raw
            try {
                Set-ConfigValueByPath $Data $path $newValue
                Write-Host '  updated in memory (choose s to save).' -ForegroundColor Green
            } catch { Write-Host ("  update failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }
            continue
        }
        if ($ans -eq 'd') {
            $path = Read-Choice 'JSON path to delete' ''
            if ([string]::IsNullOrWhiteSpace($path)) { Write-Host '  path is required' -ForegroundColor Yellow; continue }
            $old = Get-ConfigValueByPath $Data $path
            if ($null -eq $old) { Write-Host '  (not found)' -ForegroundColor Yellow; continue }
            Write-Host 'Deleting value:' -ForegroundColor DarkGray
            Write-Host (Get-ConfigOverlayJson $old)
            $confirm = Read-Choice 'Delete this path? type DELETE' 'no'
            if ($confirm -ceq 'DELETE') {
                if (Remove-ConfigValueByPath $Data $path) { Write-Host '  deleted in memory (choose s to save).' -ForegroundColor Green }
                else { Write-Host '  delete failed' -ForegroundColor Yellow }
            }
            continue
        }
        if ($menu.ContainsKey($ans)) { Show-ConfigEditorGroup $Data $menu[$ans]; continue }
        Write-Host '  unknown editor command' -ForegroundColor Yellow
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
        if (-not [string]::IsNullOrWhiteSpace($State.FromBizCode)) { $args['FromBizCode'] = $State.FromBizCode }
        if ($State.WbsStartRow -gt 0) { $args['WbsStartRow'] = $State.WbsStartRow }
        if ($State.WbsEndRow -gt 0) { $args['WbsEndRow'] = $State.WbsEndRow }
        if ($State.CorrelIdsM.Count -gt 0) { $args['CorrelIdsM'] = $State.CorrelIdsM }
        if ($State.JobNames.Count -gt 0) { $args['JobNames'] = $State.JobNames }
        if ($State.ExcelNames.Count -gt 0) { $args['ExcelNames'] = $State.ExcelNames }
        if ($State.AllowTempMapping) { $args['AllowTempMapping'] = $true }
        if ($State.AddRows) { $args['Add'] = $true }
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
        # SnapVerify F2 (MQ NG detection) + Expected_Time window config.
        if ($Config.ContainsKey('SnapVerify')) {
            $sv = $Config.SnapVerify
            $args['SnapEnabled']     = [bool]$sv.Enabled
            $args['ToleranceMinutes'] = [int]$sv.ToleranceMinutes
            $args['SaveText']        = [bool]$sv.SaveText
            $args['PollTimeoutSec']  = [int]$sv.PollTimeoutSec
            $args['PollIntervalMs']  = [int]$sv.PollIntervalMs
        }
        if ($Config.ContainsKey('ExpectedTime')) {
            $args['TimeColumn'] = [string]$Config.ExpectedTime.TimeColumn
            $args['TimeFormat'] = [string]$Config.ExpectedTime.TimeFormat
        }
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
        # SnapVerify F3 (Jenkins file NG detection) + Expected_Time window config.
        if ($Config.ContainsKey('SnapVerify')) {
            $sv = $Config.SnapVerify
            $args['SnapEnabled']      = [bool]$sv.Enabled
            $args['ToleranceMinutes'] = [int]$sv.ToleranceMinutes
            $args['SaveText']         = [bool]$sv.SaveText
            $args['PollTimeoutSec']   = [int]$sv.PollTimeoutSec
            $args['PollIntervalMs']   = [int]$sv.PollIntervalMs
        }
        if ($Config.ContainsKey('ExpectedTime')) {
            $args['TimeColumn'] = [string]$Config.ExpectedTime.TimeColumn
            $args['TimeFormat'] = [string]$Config.ExpectedTime.TimeFormat
        }
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

    if ($PhaseKey -eq 'SendVsGift') {
        $p  = Resolve-ToolPath $Config 'SendVsGift'
        $eh = Resolve-ToolPath $Config 'ExcelHelpers'
        $args = $base.Clone()
        if (-not [string]::IsNullOrWhiteSpace($State.ExcelPrefix)) { $args['ExcelPrefix'] = $State.ExcelPrefix }
        $args['EvidenceDir'] = $State.EvidenceDir
        $args['CursorCell'] = $State.CursorCell
        $args['ExcelHelpersScript'] = $eh
        if ($Config.Review.SaveWaitMs) { $args['SaveWaitMs'] = [int]$Config.Review.SaveWaitMs }
        if ($Config.Review.Maximize) { $args['Maximize'] = $true }
        if ($State.Ocr) { $args['Ocr'] = $true }
        if ($Config.ContainsKey('SendVsGift') -and $null -ne $Config.SendVsGift) {
            $svg = $Config.SendVsGift
            if ($svg.ContainsKey('AutoMark') -and -not $svg.AutoMark) { $args['NoAutoMark'] = $true }
            if ($svg.ContainsKey('OcrLanguage') -and -not [string]::IsNullOrWhiteSpace([string]$svg.OcrLanguage)) { $args['OcrLanguage'] = [string]$svg.OcrLanguage }
            if ($svg.ContainsKey('SendSheetName') -and -not [string]::IsNullOrWhiteSpace([string]$svg.SendSheetName)) { $args['SendSheetName'] = [string]$svg.SendSheetName }
            if ($svg.ContainsKey('ZeroBytePattern') -and -not [string]::IsNullOrWhiteSpace([string]$svg.ZeroBytePattern)) { $args['ZeroBytePattern'] = [string]$svg.ZeroBytePattern }
            if ($svg.ContainsKey('ZeroTemplate') -and -not [string]::IsNullOrWhiteSpace([string]$svg.ZeroTemplate)) { $args['ZeroTemplate'] = [string]$svg.ZeroTemplate }
        }
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        if ($State.DryRun) { $args['DryRun'] = $true }
        Write-Host '[RUN] SendVsGift' -ForegroundColor Green
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
        if (-not [string]::IsNullOrWhiteSpace($State.ExcelPrefix)) { $args['ExcelPrefix'] = $State.ExcelPrefix }
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
        if (-not [string]::IsNullOrWhiteSpace($State.ExcelPrefix)) { $args['ExcelPrefix'] = $State.ExcelPrefix }
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
        if (-not [string]::IsNullOrWhiteSpace($State.ExcelPrefix)) { $args['ExcelPrefix'] = $State.ExcelPrefix }
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
        if (-not [string]::IsNullOrWhiteSpace($State.ExcelPrefix)) { $args['ExcelPrefix'] = $State.ExcelPrefix }
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
        # Align default is force-replace; -DiffMode switches to report-only.
        if ($State.DiffMode) { $args['DiffMode'] = $true }
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
        if (-not [string]::IsNullOrWhiteSpace($State.ExcelPrefix)) { $args['ExcelPrefix'] = $State.ExcelPrefix }
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
        if (-not [string]::IsNullOrWhiteSpace($State.ExcelPrefix)) { $args['ExcelPrefix'] = $State.ExcelPrefix }
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
        if (-not [string]::IsNullOrWhiteSpace($State.ExcelPrefix)) { $args['ExcelPrefix'] = $State.ExcelPrefix }
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
        if (-not [string]::IsNullOrWhiteSpace($State.ExcelPrefix)) { $args['ExcelPrefix'] = $State.ExcelPrefix }
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

    if ($PhaseKey -eq 'DeliverFiles') {
        $p = Resolve-ToolPath $Config 'DeliverFiles'
        $args = $base.Clone()
        if (-not [string]::IsNullOrWhiteSpace($State.ExcelPrefix)) { $args['ExcelPrefix'] = $State.ExcelPrefix }
        $j4Ev = ''
        if ($Config.DeliverFiles) {
            $df = $Config.DeliverFiles
            if (-not [string]::IsNullOrWhiteSpace([string]$df.J4EvidenceDir)) { $j4Ev = [string]$df.J4EvidenceDir; $args['J4EvidenceDir'] = $j4Ev }
            if (-not [string]::IsNullOrWhiteSpace([string]$df.J4GfixDataDir)) { $args['J4GfixDataDir'] = [string]$df.J4GfixDataDir }
            if (-not [string]::IsNullOrWhiteSpace([string]$df.J4GiftDataDir)) { $args['J4GiftDataDir'] = [string]$df.J4GiftDataDir }
            if ($df.MoveData -or $State.MoveData)  { $args['MoveData'] = $true }
        }
        if ([string]::IsNullOrWhiteSpace($j4Ev)) {
            # Fall back to Mail.EvidenceFolder
            if ($Config.Mail -and -not [string]::IsNullOrWhiteSpace([string]$Config.Mail.EvidenceFolder)) {
                $args['J4EvidenceDir'] = [string]$Config.Mail.EvidenceFolder
            }
        }
        if ($State.TargetIds.Count -gt 0) { $args['TargetIds'] = $State.TargetIds }
        if ($State.Force) { $args['Force'] = $true }
        $args['EvidenceDir'] = $State.EvidenceDir
        Write-Host '[RUN] DeliverFiles' -ForegroundColor Green
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
        if (-not [string]::IsNullOrWhiteSpace($State.ExcelPrefix)) { $args['ExcelPrefix'] = $State.ExcelPrefix }
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

    if ($PhaseKey -eq 'InitConfig') {
        $overlayName = [string]$Config.Paths.OverlayName
        if ([string]::IsNullOrWhiteSpace($overlayName)) { $overlayName = 'verify_config.json' }
        $dest = Join-Path $State.WorkDir $overlayName
        Write-Host '[RUN] InitConfig' -ForegroundColor Green
        Write-Host ("  Overlay target : {0}" -f $dest)
        $exists = Test-Path -LiteralPath $dest
        $snap = New-ConfigOverlaySnapshot $Config
        if ($State.Interactive) {
            $editable = Copy-ConfigObject $snap
            $edited = Invoke-ConfigOverlayEditor $editable $dest
            if ($null -eq $edited) {
                Write-Host '  [CANCEL] no config changes were written.' -ForegroundColor Yellow
                return
            }
            $snap = $edited
        }
        if ($State.DryRun) {
            Write-Host '  [dry-run] would write this overlay snapshot:' -ForegroundColor DarkGray
            Write-Host (Get-ConfigOverlayJson $snap)
            return
        }
        $json = Get-ConfigOverlayJson $snap
        if ($exists) {
            $bak = ('{0}.bak.{1}' -f $dest, (Get-Date -Format 'yyyyMMdd_HHmmss'))
            Copy-Item -LiteralPath $dest -Destination $bak -Force
            Write-Host ("  [backup] {0}" -f $bak) -ForegroundColor DarkGray
        }
        [System.IO.File]::WriteAllText($dest, $json, (New-Object System.Text.UTF8Encoding($false)))
        $readmePath = Join-Path $State.WorkDir 'verify_config.README.txt'
        $readmeText = Get-ConfigOverlayReadmeText $overlayName
        [System.IO.File]::WriteAllText($readmePath, $readmeText, (New-Object System.Text.UTF8Encoding($false)))
        Write-Host '  [OK] wrote/updated work-folder config overlay (UTF-8, no BOM).' -ForegroundColor Green
        Write-Host ("  [OK] wrote config field guide: {0}" -f $readmePath) -ForegroundColor Green
        Write-Host '       Edit values, then re-run any phase. JSON overrides VerifyConfig.psd1.' -ForegroundColor DarkGray
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

# WorkDir must be resolved first: the per-work-folder config overlay lives under it.
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

# Per-work-folder JSON overlay: deep-merged over VerifyConfig.psd1 (JSON wins,
# CLI args still win over JSON). Lets each work folder fully customize owner,
# window size, mark boxes, mail format, etc. without editing the .psd1.
$overlayInfo = Import-ConfigOverlay $Config $WorkDir
$overlay     = $overlayInfo.Overlay

# Owner: CLI > overlay (DefaultOwner) > last session > psd1 default.
if ([string]::IsNullOrWhiteSpace($Owner)) {
    if ($overlay.ContainsKey('DefaultOwner') -and -not [string]::IsNullOrWhiteSpace([string]$overlay['DefaultOwner'])) {
        $Owner = [string]$overlay['DefaultOwner']
    } elseif ($session.ContainsKey('Owner') -and -not [string]::IsNullOrWhiteSpace([string]$session['Owner'])) {
        $Owner = [string]$session['Owner']
    } else {
        $Owner = [string]$Config.DefaultOwner
    }
}

if ($WindowWidth -le 0)  { $WindowWidth  = [int]$Config.Window.Width }
if ($WindowHeight -le 0) { $WindowHeight = [int]$Config.Window.Height }
if ($CropPx -lt 0)       { $CropPx       = [int]$Config.Window.CropPx }
if ([string]::IsNullOrWhiteSpace($CursorCell)) { $CursorCell = [string]$Config.Review.CursorCell }
if ([string]::IsNullOrWhiteSpace($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir ([string]$Config.Review.EvidenceDir) }
elseif (-not [System.IO.Path]::IsPathRooted($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir $EvidenceDir }

# CloneSourceDir: CLI > overlay (Clone.SourceDir) > last session.
if ([string]::IsNullOrWhiteSpace($CloneSourceDir) -and $Config.ContainsKey('Clone') -and -not [string]::IsNullOrWhiteSpace([string]$Config.Clone.SourceDir)) {
    $CloneSourceDir = [string]$Config.Clone.SourceDir
}
if ([string]::IsNullOrWhiteSpace($CloneSourceDir) -and $session.ContainsKey('CloneSourceDir')) {
    $CloneSourceDir = [string]$session['CloneSourceDir']
}

# J4BaseDir: CLI > work-folder config Align.J4BaseDir > CloneSourceDir fallback > last session.
if ([string]::IsNullOrWhiteSpace($J4BaseDir) -and $Config.Align -and -not [string]::IsNullOrWhiteSpace([string]$Config.Align.J4BaseDir)) {
    $J4BaseDir = [string]$Config.Align.J4BaseDir
}
if ([string]::IsNullOrWhiteSpace($J4BaseDir) -and -not [string]::IsNullOrWhiteSpace($CloneSourceDir)) {
    $J4BaseDir = $CloneSourceDir
}
if ([string]::IsNullOrWhiteSpace($J4BaseDir) -and $session.ContainsKey('J4BaseDir')) {
    $J4BaseDir = [string]$session['J4BaseDir']
}

# CheckSheetPath: CLI > work-folder config CheckSheet.Path > last session.
if ([string]::IsNullOrWhiteSpace($CheckSheetPath) -and $Config.CheckSheet -and -not [string]::IsNullOrWhiteSpace([string]$Config.CheckSheet.Path)) {
    $CheckSheetPath = [string]$Config.CheckSheet.Path
}
if ([string]::IsNullOrWhiteSpace($CheckSheetPath) -and $session.ContainsKey('CheckSheetPath')) {
    $CheckSheetPath = [string]$session['CheckSheetPath']
}

# Project-level evidence workbook prefix: CLI > work-folder config Workbook.ExcelPrefix.
if ([string]::IsNullOrWhiteSpace($ExcelPrefix) -and $Config.Workbook -and -not [string]::IsNullOrWhiteSpace([string]$Config.Workbook.ExcelPrefix)) {
    $ExcelPrefix = [string]$Config.Workbook.ExcelPrefix
}

$TargetIds = @(ConvertTo-TargetIdSelection $TargetIds)

$flatBiz = @()
foreach ($raw in @($BizCodes)) {
    if ($null -eq $raw) { continue }
    foreach ($part in ($raw.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $flatBiz += $v }
    }
}
$BizCodes = @($flatBiz | Select-Object -Unique)

$flatCorrelIdsM = @()
foreach ($raw in @($CorrelIdsM)) {
    if ($null -eq $raw) { continue }
    foreach ($part in ($raw.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $flatCorrelIdsM += $v }
    }
}
$CorrelIdsM = @($flatCorrelIdsM | Select-Object -Unique)

$flatJobNames = @()
foreach ($raw in @($JobNames)) {
    if ($null -eq $raw) { continue }
    foreach ($part in ($raw.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $flatJobNames += $v }
    }
}
$JobNames = @($flatJobNames | Select-Object -Unique)

$flatExcelNames = @()
foreach ($raw in @($ExcelNames)) {
    if ($null -eq $raw) { continue }
    foreach ($part in ($raw.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $flatExcelNames += $v }
    }
}
$ExcelNames = @($flatExcelNames | Select-Object -Unique)

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
    ExcelPrefix     = $ExcelPrefix
    BizCodes        = $BizCodes
    FromBizCode    = $FromBizCode
    WbsStartRow    = $WbsStartRow
    WbsEndRow      = $WbsEndRow
    CorrelIdsM     = $CorrelIdsM
    JobNames       = $JobNames
    ExcelNames     = $ExcelNames
    AllowTempMapping = [bool]$AllowTempMapping.IsPresent
    AddRows        = [bool]$Add.IsPresent
    HostSystemTypes = @()
    ProbeFile       = $ProbeFile
    ProbeSheet      = $ProbeSheet
    DfExePath       = $DfExePath
    CheckSheetPath  = $CheckSheetPath
    Force           = [bool]$Force.IsPresent
    # OCR for SendVsGift: on by default; 'o' menu option toggles off/on per run.
    Ocr             = $true
    Interactive     = [bool]$Interactive.IsPresent
    NoResize        = ([bool]$NoResize.IsPresent -or ($Config.Window -and [bool]$Config.Window.NoResize))
    RefreshUrls     = [bool]$RefreshUrls.IsPresent
    DryRun          = [bool]$DryRun.IsPresent
    DiffMode        = $false
    MoveData        = [bool]$MoveData.IsPresent
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
# Safe - never modifies existing data, only adds missing columns with '0'.
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
if ($overlayInfo.Loaded) { Write-Host ("  Config overlay : {0}" -f $overlayInfo.Path) -ForegroundColor DarkGray }
Write-Host ("  Window         : {0}x{1}, CropPx={2}" -f $WindowWidth, $WindowHeight, $CropPx)
if (-not [string]::IsNullOrWhiteSpace($CloneSourceDir)) {
    Write-Host ("  CloneSourceDir : {0}" -f $CloneSourceDir)
}
if (-not [string]::IsNullOrWhiteSpace($J4BaseDir)) {
    Write-Host ("  J4BaseDir      : {0}" -f $J4BaseDir)
}
if (-not [string]::IsNullOrWhiteSpace($ExcelPrefix)) {
    Write-Host ("  ExcelPrefix    : {0}" -f $ExcelPrefix)
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
    $mappingPath = Get-MappingPath $Config $state.WorkDir $state.Owner

    Write-Host ''
    Write-Host 'Back to VerifyTool menu. Enter to refresh / q to quit : ' -ForegroundColor Magenta -NoNewline
    $again = Read-Host
    if ($again -eq 'q') { break }
}