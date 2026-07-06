# ============================================================
#  SnapLocalize.ps1   (M5 / F5 wiring glue -- NOT a pure library)
#
#  Turns a snap verdict into a <correl>.loc.json sidecar by combining:
#    - the PURE geometry/builder functions in SnapVerify.ps1
#      (Get-MatchedRowIndex / Get-RowPixelRect / Get-JenkinsHighlightRect /
#       New-SnapLocRect / Save-SnapLocSidecar), and
#    - image I/O (System.Drawing) + the Jenkins highlight scan
#      (Find-ActiveHighlightRow.ps1).
#  Because it touches GDI+ it lives OUTSIDE SnapVerify.ps1 (which stays pure /
#  unit-tested). No param() block -> safe to dot-source per CLAUDE.md.
#
#  Dot-source AFTER SnapVerify.ps1 and Find-ActiveHighlightRow.ps1, and after
#  Add-Type System.Drawing. Write-SnapLocalize NEVER throws into the caller:
#  any failure (disabled, uncalibrated geometry, no matched row, no highlight,
#  GDI error) returns $null so the screenshot workflow is never blocked.
#
#  Also holds Write-MarkTemplateHits: an unrelated-but-similar snap-time
#  helper that runs Mark.ps1's 'Template' image-match boxes right after a
#  screenshot is captured and persists any hits to <correl>.tplhit.json, so
#  Mark.ps1 can reuse that anchor later instead of re-scanning the archived
#  PNG. Same never-blocks-the-caller contract as Write-SnapLocalize.
# ============================================================

# ---------------------------------------------------------------------------
# Get-PngSize -- final (cropped) PNG pixel dimensions for the sidecar, so the
# Mark phase can scale pixel -> point via Shape.Width / imageWidth.
# ---------------------------------------------------------------------------
function Get-PngSize {
    param([Parameter(Mandatory)][string]$Path)
    $img = [System.Drawing.Image]::FromFile($Path)
    try { return @{ Width = [int]$img.Width; Height = [int]$img.Height } }
    finally { $img.Dispose() }
}

# ---------------------------------------------------------------------------
# Write-SnapLocalize
#   Compute + persist <correl>.loc.json for one snapped row. Returns the
#   sidecar path on success, else $null (silent no-op on any problem).
#
#   Parameters:
#     Page          'Hm' | 'Mq' | 'Jenkins'
#     Localize      Config.SnapVerify.Localize hashtable
#     SnapDir       folder containing the PNG (sidecar written here)
#     Correl        Correl_ID_S (sidecar basename + HM/MQ row filter)
#     PngPath       full path to the saved (already cropped) PNG
#     Rows          parsed rows (HM/MQ); ignored for Jenkins
#     Expected      [datetime] | $null  (verdict time window)
#     ToleranceMin  int
#     CropLeft/CropTop  px Invoke-CropPng trimmed (HM/MQ pixel offset)
# ---------------------------------------------------------------------------
function Write-SnapLocalize {
    param(
        [string]$Page,
        $Localize,
        [string]$SnapDir,
        [string]$Correl,
        [string]$PngPath,
        [object[]]$Rows    = @(),
        [object]$Expected  = $null,
        [int]$ToleranceMin = 30,
        [int]$CropLeft     = 0,
        [int]$CropTop      = 0
    )

    if ($null -eq $Localize) { return $null }
    if (-not [bool]$Localize['Enabled']) { return $null }

    try {
        if ([string]::IsNullOrWhiteSpace($PngPath) -or -not (Test-Path -LiteralPath $PngPath)) { return $null }
        $size     = Get-PngSize -Path $PngPath
        $rect     = $null
        $source   = ''
        $rowIndex = 0

        if ($Page -eq 'Jenkins') {
            if (-not [bool]$Localize['Jenkins']) { return $null }
            $band = Find-ActiveHighlightRow -ImagePath $PngPath `
                -ActiveR ([int]$Localize['JenkinsActiveR']) `
                -ActiveG ([int]$Localize['JenkinsActiveG']) `
                -ActiveB ([int]$Localize['JenkinsActiveB']) `
                -Tolerance ([int]$Localize['JenkinsTolerance']) `
                -MinPixelsPerRow ([int]$Localize['JenkinsMinPixelsPerRow'])
            if ($null -eq $band) { return $null }   # no active match found
            $rect = Get-JenkinsHighlightRect -Top ([int]$band.Top) -Bottom ([int]$band.Bottom) `
                -ColLeft ([int]$Localize['JenkinsColLeft']) -ColWidth ([int]$Localize['JenkinsColWidth']) `
                -ImageWidth ([int]$size.Width) -Pad ([int]$Localize['JenkinsPad'])
            $source = 'jenkins-highlight'
        }
        elseif ($Page -eq 'Hm' -or $Page -eq 'Mq') {
            $rowHeight = [int]$Localize["${Page}RowHeight"]
            $colWidth  = [int]$Localize["${Page}ColWidth"]
            if ($rowHeight -le 0 -or $colWidth -le 0) { return $null }   # geometry not calibrated
            $row1Top   = [int]$Localize["${Page}Row1Top"]
            $colLeft   = [int]$Localize["${Page}ColLeft"]
            $dateProp  = if ($Page -eq 'Hm') { 'StartTime' } else { 'RecvDate' }
            $rowIndex  = Get-MatchedRowIndex -Rows $Rows -CorrelId $Correl `
                -DateProperty $dateProp -Expected $Expected -ToleranceMin $ToleranceMin
            if ($rowIndex -lt 1) { return $null }   # no matched data row
            $rect = Get-RowPixelRect -RowIndex $rowIndex -Row1Top $row1Top -RowHeight $rowHeight `
                -ColLeft $colLeft -ColWidth $colWidth -CropLeft $CropLeft -CropTop $CropTop
            $source = "$($Page.ToLower())-geometry"
        }
        else { return $null }

        if ($null -eq $rect) { return $null }

        $loc = New-SnapLocRect -CorrelId $Correl -X $rect.x -Y $rect.y -W $rect.w -H $rect.h `
            -Source $source -RowIndex $rowIndex -ImageWidth ([int]$size.Width) -ImageHeight ([int]$size.Height)
        $sidecar = Join-Path $SnapDir ("{0}.loc.json" -f $Correl)
        return (Save-SnapLocSidecar -Loc $loc -Path $sidecar)
    } catch {
        # Localisation is best-effort -- never break snapping. Swallow + skip.
        return $null
    }
}

# ---------------------------------------------------------------------------
# Resolve-SnapTemplatePath -- snap-time counterpart of Mark.ps1's
# Resolve-MarkTemplatePath (duplicated rather than shared: Mark.ps1 has a
# param() block, so per CLAUDE.md's dot-source rule it can only ever be
# invoked via '&', never dot-sourced, and its copy is not reachable here).
# ---------------------------------------------------------------------------
function Resolve-SnapTemplatePath {
    param([string]$Template, [string]$TemplateDir)
    if ([string]::IsNullOrWhiteSpace($Template)) { return $null }
    if (Test-Path -LiteralPath $Template) { return (Resolve-Path -LiteralPath $Template).ProviderPath }
    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($TemplateDir)) { $candidates += (Join-Path $TemplateDir $Template) }
    $candidates += (Join-Path (Join-Path $PSScriptRoot 'mark_templates') $Template)
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return (Resolve-Path -LiteralPath $c).ProviderPath }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Write-MarkTemplateHits
#   Snap-time counterpart of Mark.ps1's image-match boxes: runs the same
#   Locate-ByImage match for every Mark.Boxes[<folder>] entry that carries a
#   'Template' key, against the screenshot JUST captured (while the page is
#   known-good), and persists any hits to <correl>.tplhit.json next to the
#   PNG. Mark.ps1 reads this sidecar first (Get-MarkTemplateHitFromSidecar)
#   instead of re-running the match against the archived PNG later, and
#   silently falls back to a live match when the sidecar is absent, stale,
#   or has no entry for a given box -- so a missing/failed sidecar never
#   blocks Mark. This function itself never throws into the caller: any
#   failure (no boxes configured, Locate-ByImage missing, no match found for
#   any box) returns $null so the snap workflow is never blocked.
#
#   Parameters:
#     SnapDir      folder containing the PNG (sidecar written here)
#     Correl       Correl_ID_S (sidecar basename)
#     PngPath      full path to the saved (already cropped) PNG
#     Boxes        Mark.Boxes[<folder>] array of box hashtables
#     TemplateDir  Mark.TemplateDir (falls back to <repo>\mark_templates)
#     Tolerance    default LockBits tolerance (a box's own 'Tolerance' wins)
#     LocateScript full path to Locate-ByImage.ps1
# ---------------------------------------------------------------------------
function Write-MarkTemplateHits {
    param(
        [string]$SnapDir,
        [string]$Correl,
        [string]$PngPath,
        [object[]]$Boxes      = @(),
        [string]$TemplateDir  = '',
        [int]$Tolerance       = 15,
        [string]$LocateScript = ''
    )

    if ([string]::IsNullOrWhiteSpace($PngPath) -or -not (Test-Path -LiteralPath $PngPath)) { return $null }
    if ($null -eq $Boxes -or @($Boxes).Count -eq 0) { return $null }
    if ([string]::IsNullOrWhiteSpace($LocateScript) -or -not (Test-Path -LiteralPath $LocateScript)) { return $null }

    try {
        $size = Get-PngSize -Path $PngPath
        $hits = New-Object System.Collections.Generic.List[object]
        $idx = 0
        foreach ($b in @($Boxes)) {
            if ($b -is [hashtable] -and $b.ContainsKey('Template') -and -not [string]::IsNullOrWhiteSpace([string]$b.Template)) {
                $tplName = [string]$b.Template
                $tplPath = Resolve-SnapTemplatePath -Template $tplName -TemplateDir $TemplateDir
                if ($null -ne $tplPath) {
                    $tol = $Tolerance
                    if ($b.ContainsKey('Tolerance')) { try { $tol = [int]$b.Tolerance } catch {} }
                    try {
                        $hit = & $LocateScript -SourcePath $PngPath -TemplatePath $tplPath -Tolerance $tol -Quiet
                        if ($null -ne $hit) {
                            $hits.Add([ordered]@{
                                Index    = $idx
                                Template = $tplName
                                X        = [int]$hit.X
                                Y        = [int]$hit.Y
                                Width    = [int]$hit.Width
                                Height   = [int]$hit.Height
                            })
                        }
                    } catch {}
                }
            }
            $idx++
        }
        if ($hits.Count -eq 0) { return $null }

        $payload = [ordered]@{
            Correl       = $Correl
            SourceWidth  = [int]$size.Width
            SourceHeight = [int]$size.Height
            CapturedAt   = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
            Boxes        = $hits
        }
        Ensure-Dir $SnapDir
        $sidecar = Join-Path $SnapDir ("{0}.tplhit.json" -f $Correl)
        $enc = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($sidecar, ($payload | ConvertTo-Json -Depth 6), $enc)
        return $sidecar
    } catch {
        return $null
    }
}
