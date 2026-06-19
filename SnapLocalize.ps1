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
