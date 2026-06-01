@{
    # VerifyTool.ps1 remembers the last WorkDir in verify_session.json.
    DefaultWorkDir = ''
    DefaultOwner   = '厳'

    Paths = @{
        MappingPattern  = 'mapping_{0}.csv'
        EvidenceDir     = 'evidence'
        SnapDir         = 'snap'
        FileDir         = 'DATA'
        ExcelWorkbook   = 'wipGFIX一覧.xlsx'
        TemplatePrefix  = 'template_'    # template_<bizcode>.xlsx in WorkDir
        TemplateGeneric = 'template.xlsx'
    }

    Window = @{
        Width    = 1050
        Height   = 761
        CropPx   = 6
        NoResize = $false
    }

    Timing = @{
        ActionWaitMs  = 500
        ResultWaitSec = 2
        ResultWaitMs  = 500
    }

    ExcelSnap = @{
        WorkbookName  = 'wipGFIX一覧.xlsx'
        FilterCell    = 'O5'
        SnapStartCell = 'B4'
        SnapEndColumn = 'O'
        SnapFolder    = 'excel'
    }

    Hm = @{
        TabsToCorrelid     = 1
        TabsBackFromSearch = 1
        TabsBackToInput    = 4
    }

    Mq = @{
        TabsToInquiry  = 1
        TabsToCorrelid = 4
    }

    Review = @{
        EvidenceDir = 'evidence'
        CursorCell  = 'A3'
        Field       = 'isReviewed'
        Maximize    = $true
        SaveWaitMs  = 5000
    }

    # Replace phase configuration
    Replace = @{
        StartRow         = 3
        ColAnchor        = 2           # column B
        BlankRowsBetween = 1
        ClearEndColumn   = 20
        # Labels for ReplaceEvidence.ps1. Empty -> tool uses [char] defaults.
        # Edit here when final wording is confirmed.
        GiftNoGfixLabel  = ''
        GfixLogLabel     = 'GFIX受信log'
        GfixLogTodoText  = '<<TODO: GFIX 受信 log>>'
    }

    # Mark phase configuration (マークフェーズ設定)
    Mark = @{
        NamePrefix = 'verifyMark_'   # すべてのマーク用図形はこのプレフィックスで開始（クリーンアップ用）
        LineWeight = 1.5
        # 各ソースフォルダごとの設定：画像の左上隅を基準とした相対的な赤枠（矩形）の描画リスト。
        # 幅 (Width) / 高さ (Height) の単位はポイント（1pt = 1/72インチ）。
        # 空のリスト = 該当フォルダにはマークを描画しない。
        # 調整方法：手動でマークした検証用ワークブックに対して Probe-Shapes.ps1 を実行し、
        # AutoShape とその親となる Picture のオフセット差分を読み取る。
        Boxes = @{
            'excel'           = @()
            'GIFT_HM'         = @( @{ OffsetX = 384.8; OffsetY = 181.6; Width = 62.2; Height = 16.5 } )
            'GIFT_MQ'         = @( @{ OffsetX = 155.2; OffsetY = 170.2; Width = 528.8; Height = 63 } )
            'GIFT_Jenkins'    = @( @{ OffsetX = 303;   OffsetY = 297;   Width = 288.8; Height = 18.8 } )
            'GIFT_noGfixfile' = @()
            'GFIX_HM'         = @( @{ OffsetX = 384.8; OffsetY = 181.6; Width = 62.2; Height = 16.5 } ) # GIFT_HMと同じと仮定
            'GFIX_Jenkins'    = @( @{ OffsetX = 303;   OffsetY = 297;   Width = 288.8; Height = 18.8 } ) # GIFT_Jenkinsと同じと仮定
            'DF'              = @( @{ CellCols = 'AW:BC'; RowsFromBottom = 2 } )
        }
    }

    Scripts = @{
        GenerateMapping = 'Generate-HostOpenMapping.ps1'
        Excel           = 'ExcelSnap.ps1'
        Hm              = 'HmSnap.ps1'
        Mq              = 'MqSnap.ps1'
        Jenkins         = 'JenkinsSnap.ps1'
        Crop            = 'Crop-Snap.ps1'
        Review          = 'ReviewEvidence.ps1'
        Common          = 'Common.ps1'
        ExcelHelpers    = 'ExcelHelpers.ps1'
        Clone           = 'Clone.ps1'
        Replace         = 'ReplaceEvidence.ps1'
        Validate        = 'Validate.ps1'
        Mark            = 'Mark.ps1'
        MarkGfixLog     = 'MarkGfixLog.ps1'
        Probe           = 'Probe-Shapes.ps1'
        GfixLogDownload = 'GfixLogDownload.ps1'
        DfSnap          = 'DfSnap.ps1'
        Align           = 'Align.ps1'
        WatchProgress   = 'Watch-MappingProgress.ps1'
    }

    # DF snap / mark configuration
    Df = @{
        # Path to df.exe comparison tool. User must set for their environment.
        ExePath      = ''
        # Seconds to wait after the df.exe window appears before capturing.
        LoadWaitSec  = 8
        # 'region' (recommended; df.exe window handle is flaky), 'window', 'fullscreen'.
        CaptureMode  = 'region'
        # Fixed capture rectangle for 'region' mode (target screen ~1980x1020).
        RegionX      = 120
        RegionY      = 280
        RegionWidth  = 1250
        RegionHeight = 657
        # Per-direction crop in px (window shadow is asymmetric).
        CropLeft     = 0
        CropTop      = 0
        CropRight    = 0
        CropBottom   = 0
        # Base directories containing data files to compare.
        # Defaults to <WorkDir>\DATA\GIFT and <WorkDir>\DATA\GFIX when empty.
        GiftDataDir  = ''
        GfixDataDir  = ''
        # Wildcard pattern for file lookup. {0} = Correl_ID_S.
        FilePattern  = '{0}*'
    }

    # Align / Precheck configuration (compare work evidence vs J4 baseline)
    Align = @{
        # Root folder holding the J4 baseline workbooks (searched recursively
        # by Excel_NAME). Example:
        #   \\Fs-f3170-1\...\REQ-000xxxxx_GIFT...\40.J4\07.GPCS
        J4BaseDir       = ''
        # FROM_sys / TO_sys literal values that count as "Host" (mainframe).
        # Until set, migration type is Unknown and Align falls back to the
        # Host->Open (3 receive sheets) scope with a warning.
        HostSystemTypes = @()
    }

    # GFIX log cell-mark configuration
    GfixLog = @{
        # Exact text in column B that marks the start of each log region.
        LogAnchor        = ''       # empty -> MarkGfixLog.ps1 uses [char]0x25BC + 'GFIX' + ...
        # Regex matched against B column to find the row to highlight.
        CommandPattern   = "Command:\s*'/appl/[A-Za-z0-9]+/shell/"
        # OLE color for highlight. Yellow RGB(255,255,0) = 65535.
        HighlightColor   = 65535
        # Column range to highlight (B=2, AY=51).
        HighlightColStart = 2
        HighlightColEnd   = 51
    }

    # Phase entries: Field + optional BitValue.
    # If BitValue > 0, the field is read as a bitmask and "done" means
    # (value -band BitValue) -eq BitValue.
    PhaseOrder = @(
        @{ Key='Mapping';            Field='';                     Label='mapping 生成 / 更新';               Status='implemented' }
        @{ Key='ExcelSnap';          Field='Excel_snap';           Label='Excel 証跡';                        Status='legacy' }
        @{ Key='GiftHmSnap';         Field='GIFT_HM_snap';         Label='GIFT HM 証跡';                      Status='implemented' }
        @{ Key='GiftMqSnap';         Field='GIFT_MQ_snap';         Label='GIFT MQ 証跡';                      Status='implemented' }
        @{ Key='GiftJenkins';        Field='GIFT_Jenkins_snap';    Label='GIFT Jenkins 証跡 + DL';            Status='implemented' }
        @{ Key='GiftJenkinsNoFile';  Field='GIFT_noGfixfile_snap'; Label='GIFT Jenkins no-GFIX 証跡';         Status='implemented' }
        @{ Key='GfixHmSnap';         Field='GFIX_HM_snap';         Label='GFIX HM 証跡';                      Status='implemented' }
        @{ Key='GfixJenkins';        Field='GFIX_Jenkins_snap';    Label='GFIX Jenkins 証跡 + DL';            Status='implemented' }
        @{ Key='GfixLogDownload';    Field='GFIX_log';             Label='GFIX LOG download';                 Status='implemented' }
        @{ Key='DfSnap';             Field='DF_snap';              Label='DF 証跡 (df.exe 截图)';             Status='implemented' }
        @{ Key='Clone';              Field='';                     Label='証跡 Excel 複製 (mkexcel)';         Status='implemented' }
        @{ Key='Align';              Field='';                     Label='J4 基準 sheet 比較/同期 (precheck)'; Status='implemented' }
        @{ Key='ReplaceGift';        Field='isReplaced'; BitValue=1; Label='GIFT 証跡置換';                   Status='implemented' }
        @{ Key='ReplaceGfix';        Field='isReplaced'; BitValue=2; Label='GFIX 証跡置換';                   Status='implemented' }
        @{ Key='ReplaceDf';          Field='isReplaced'; BitValue=4; Label='DF 証跡置換';                     Status='implemented' }
        @{ Key='MarkGift';           Field='isMarked';   BitValue=1; Label='GIFT 赤枠 mark';                  Status='implemented' }
        @{ Key='MarkGfix';           Field='isMarked';   BitValue=2; Label='GFIX 赤枠 mark';                  Status='implemented' }
        @{ Key='MarkDf';             Field='isMarked';   BitValue=4; Label='DF 赤枠 mark';                    Status='implemented' }
        @{ Key='ReviewGift';         Field='isReviewed'; BitValue=1; Label='GIFT 目視 review';                Status='implemented' }
        @{ Key='ReviewGfix';         Field='isReviewed'; BitValue=2; Label='GFIX 目視 review';                Status='implemented' }
        @{ Key='ReviewDf';           Field='isReviewed'; BitValue=4; Label='DF 目視 review';                  Status='implemented' }
        @{ Key='ReviewEvidence';     Field='isReviewed'; BitValue=7; Label='全体 目視 review + 保存';          Status='implemented' }
        @{ Key='Comments';           Field='';                     Label='review コメント 一覧 (read-only)';   Status='implemented' }
        @{ Key='Validate';           Field='';                     Label='就緒状態 診断 (read-only)';         Status='implemented' }
        @{ Key='RepairMapping';      Field='';                     Label='mapping 列補完 (auto on startup)';  Status='implemented' }
        @{ Key='ProbeShapes';        Field='';                     Label='Excel shape 座標 / AltText 一覧';    Status='implemented' }
        @{ Key='Crop';               Field='';                     Label='既存 PNG 一括 crop';                Status='implemented' }
        @{ Key='WatchProgress';      Field='';                     Label='進捗モニタ (read-only, 非ロック)';   Status='implemented' }
    )

    Aliases = @{
        Menu              = 'Menu'
        Help              = 'Help'
        Status            = 'Status'
        Mapping           = 'Mapping'
        Excel             = 'ExcelSnap'
        ExcelSnap         = 'ExcelSnap'
        HmGift            = 'GiftHmSnap'
        GiftHm            = 'GiftHmSnap'
        GiftHmSnap        = 'GiftHmSnap'
        MqGift            = 'GiftMqSnap'
        GiftMq            = 'GiftMqSnap'
        GiftMqSnap        = 'GiftMqSnap'
        JenkinsGift       = 'GiftJenkins'
        GiftJenkins       = 'GiftJenkins'
        NoGfix            = 'GiftJenkinsNoFile'
        GiftNoFile        = 'GiftJenkinsNoFile'
        GiftJenkinsNoFile = 'GiftJenkinsNoFile'
        HmGfix            = 'GfixHmSnap'
        GfixHm            = 'GfixHmSnap'
        GfixHmSnap        = 'GfixHmSnap'
        JenkinsGfix       = 'GfixJenkins'
        GfixJenkins       = 'GfixJenkins'
        GfixLog           = 'GfixLogDownload'
        GfixLogDownload   = 'GfixLogDownload'
        Df                = 'DfSnap'
        DfSnap            = 'DfSnap'
        # MarkGfixLog is folded into MarkGfix; these keep the standalone
        # re-highlight utility reachable by name.
        MarkGfixLog       = 'MarkGfixLog'
        Mgl               = 'MarkGfixLog'
        GfixLogMark       = 'MarkGfixLog'

        # New: Clone + Replace per mode
        Clone             = 'Clone'
        MkExcel           = 'Clone'
        RenameExcel       = 'Clone'

        Align             = 'Align'
        Precheck          = 'Align'
        AlignCompare      = 'Align'

        WatchProgress     = 'WatchProgress'
        Watch             = 'WatchProgress'
        Progress          = 'WatchProgress'

        Replace           = 'ReplaceGift'    # bare "Replace" -> GIFT
        ReplaceEvidence   = 'ReplaceGift'    # legacy planned name -> GIFT
        ReplaceGift       = 'ReplaceGift'
        Rgift             = 'ReplaceGift'
        ReplaceGfix       = 'ReplaceGfix'
        Rgfix             = 'ReplaceGfix'
        ReplaceDf         = 'ReplaceDf'
        Rdf               = 'ReplaceDf'

        Comments          = 'Comments'
        Comment           = 'Comments'

        Review            = 'ReviewEvidence'
        ReviewEvidence    = 'ReviewEvidence'
        ReviewAll         = 'ReviewEvidence'
        ReviewGift        = 'ReviewGift'
        Rvgift            = 'ReviewGift'
        ReviewGfix        = 'ReviewGfix'
        Rvgfix            = 'ReviewGfix'
        ReviewDf          = 'ReviewDf'
        Rvdf              = 'ReviewDf'

        Mark              = 'MarkGift'
        MarkGift          = 'MarkGift'
        Mgift             = 'MarkGift'
        MarkGfix          = 'MarkGfix'
        Mgfix             = 'MarkGfix'
        MarkDf            = 'MarkDf'
        Mdf               = 'MarkDf'

        Validate          = 'Validate'
        Check             = 'Validate'
        Diagnose          = 'Validate'

        RepairMapping     = 'RepairMapping'
        EnsureCols        = 'RepairMapping'
        Repair            = 'RepairMapping'

        ProbeShapes       = 'ProbeShapes'
        Probe             = 'ProbeShapes'

        Crop              = 'Crop'
    }
}
