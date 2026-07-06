@{
    # VerifyTool.ps1 remembers the last WorkDir in verify_session.json.
    DefaultWorkDir = ''
    DefaultOwner   = ''

    Paths = @{
        MappingPattern  = 'mapping_{0}.csv'
        EvidenceDir     = 'evidence'
        SnapDir         = 'snap'
        FileDir         = 'DATA'
        ExcelWorkbook   = 'wipGFIX一覧.xlsx'
        TemplatePrefix  = 'template_'    # template_<bizcode>.xlsx in WorkDir
        TemplateGeneric = 'template.xlsx'
        # Per-work-folder JSON overlay file (lives in WorkDir). Deep-merged over
        # this .psd1 at startup; JSON wins. Generate it with -Phase InitConfig.
        OverlayName     = 'verify_config.json'
    }

    # Evidence/J4 workbook naming. ExcelPrefix is a project-level prefix before
    # _<Excel_NAME>, e.g. J4 review title (REQ-000xxxxx_GIFT project).
    # Legacy mapping rows with Excel_Prefix still override this for compatibility.
    Workbook = @{
        ExcelPrefix = ''
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

    # SendVsGift phase, Stage 2 OCR compare (off by default; Stage 1 manual
    # flow is unchanged). Overridable per work folder via verify_config.json.
    #   Ocr             -> $true = export send-sheet pictures + Windows OCR + auto compare
    #                      (also togglable per run: 'o' at the option prompt, or -Ocr)
    #   AutoMark        -> $true (default) = OCR verdict ok marks SendVsGift=1,
    #                      ng marks SendVsGift=2 (NG rows stay pending and are
    #                      listed at the end); $false = verdict is advisory only
    #   OcrLanguage     -> recognizer language tag passed to Windows.Media.Ocr
    #   SendSheetName   -> blank = the ProjectLabels send-data sheet name
    #   ZeroBytePattern -> regex override for the 0-byte screenshot pattern
    SendVsGift = @{
        Ocr             = $false
        AutoMark        = $true
        OcrLanguage     = 'ja'
        SendSheetName   = ''
        ZeroBytePattern = ''
    }

    # Clone phase: external source folder of existing evidence files (per bizcode).
    # Blank -> falls back to the -CloneSourceDir CLI arg / last session value.
    Clone = @{
        SourceDir = ''
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
        # Font forced on every pasted GFIX receive-log line (ReplaceGfix).
        # 'MS Gothic' is the ASCII-typeable alias Windows/Excel resolve to
        # the Japanese fixed-width font "MS ゴシック". Blank -> leave the
        # workbook's default font untouched.
        GfixLogFontName  = 'MS Gothic'
        # Font size forced on every pasted GFIX receive-log line (ReplaceGfix).
        # 0 -> leave the workbook's default size untouched. The same font +
        # size also drive MarkGfix's auto-width highlight measurement.
        GfixLogFontSize  = 11
    }

    # Mark phase configuration (マークフェーズ設定)
    Mark = @{
        NamePrefix = 'verifyMark_'   # すべてのマーク用図形はこのプレフィックスで開始（クリーンアップ用）
        LineWeight = 1.5
        # 各ソースフォルダごとの設定：画像の左上隅を基準とした相対的な赤枠（矩形）の描画リスト。
        # 幅 (Width) / 高さ (Height) の単位はポイント（1pt = 1/72インチ）。
        # 空のリスト = 該当フォルダにはマークを描画しない。
        #
        # 調整方法（固定オフセット方式）：手動でマークした検証用ワークブックに対して
        # Probe-Shapes.ps1 を実行し、AutoShape とその親となる Picture のオフセット差分を読み取る。
        #
        # 画像認識方式（任意）：各エントリに Template = '<ファイル名>.png' を追加すると、
        # そのフォルダ/correl の元スクリーンショット（<WorkDir>\snap\<folder>\<correl>.png）に対して
        # テンプレートマッチング（Locate-ByImage.ps1）を先に試み、一致した領域に赤枠を描画する
        # （見つからない場合は下の OffsetX/OffsetY/Width/Height にフォールバックするので安全）。
        #   Template  : テンプレート画像ファイル名 (TemplateDir または <repo>\mark_templates を検索)
        #               絶対/相対パスも可。
        #   Tolerance : このボックスだけの許容誤差上書き (既定 ImageMatch.Tolerance)
        #   PadX/PadY : 一致した領域の左右/上下に加える余白 (px, 既定 0)
        #
        # StampImage方式（任意、Templateとの組み合わせ専用）：矩形の代わりに、
        # Templateが一致した箇所（スケール済みシート座標）に画像そのものを
        # 等倍で貼り付ける。フォールバックのOffsetX/OffsetYは無い -- 一致しない
        # 場合は「対象が見つからなかった」とみなして何も貼らない（例：
        # GIFT_noGfixfile で過去分データが実際に存在するときだけ検出用テンプレート
        # がヒットし、その位置に already_exists.png を貼る／ヒットしなければ
        # 何もしない＝通常の「該当ファイルなし」ケース）。
        #   StampImage : 貼り付ける画像ファイル名 (TemplateDir または
        #                <repo>\mark_templates を検索)
        Boxes = @{
            'excel'           = @()
            'GIFT_HM'         = @( @{ OffsetX = 395.3; OffsetY = 189.2; Width = 62.2; Height = 16.5 } )
            'GIFT_MQ'         = @( @{ OffsetX = 167.9; OffsetY = 176.9; Width = 528.8; Height = 63 } )
            'GIFT_Jenkins'    = @( @{ OffsetX = 301.5; OffsetY = 282.0; Width = 288.8; Height = 18.8 } )
            # Image-recognition-only stamp: when the NoGfixHit.png pattern (a
            # crop of what a past-data "hit" looks like on the Jenkins list
            # page) is found on the source snap PNG, already_exists.png is
            # inserted at that location; no match = no stamp (see StampImage
            # note above). Both files must be added to mark_templates/ first.
            'GIFT_noGfixfile' = @( @{ Template = 'NoGfixHit.png'; StampImage = 'already_exists.png' } )
            'GFIX_HM'         = @( @{ OffsetX = 395.3; OffsetY = 189.2; Width = 62.2; Height = 16.5 } )
            'GFIX_Jenkins'    = @( @{ OffsetX = 301.5; OffsetY = 282.0; Width = 288.8; Height = 18.8 } )
            'DF'              = @( @{ CellCols = 'AW:BC'; RowsFromBottom = 2 } )
        }
        # Folder searched for a Box's 'Template' filename when it is not
        # already an absolute/relative path that resolves on its own.
        # Blank -> <repo>\mark_templates (always tried as a second fallback
        # even when this is set).
        TemplateDir = ''
        ImageMatch = @{
            # LockBits color tolerance per channel (0 = exact, higher = looser).
            # Per-box 'Tolerance' overrides this.
            Tolerance = 15
        }
        # Optional stamp images attached to a 'verifyNote' annotation instead
        # of a per-picture Boxes entry -- keyed by the note's Folder value
        # (the first '|' field of the AltText payload EvidenceExecutor.ps1
        # stamps via Set-ShapeMetadata 'verifyNote'), so future note kinds can
        # be added here later without touching Mark.ps1. Only 'GIFT_noGfixfile'
        # (the F4/M6 past-data annotation) is wired today: -Mode Gift reuses
        # the pixel rect already carried in that payload (from the snap-time
        # <correl>.loc.json / .note.json sidecars) to find the highlighted
        # row, then inserts Image at (that row + RowOffset, Column).
        #   Image     : template filename (TemplateDir or <repo>\mark_templates)
        #   Column    : column letter the stamp is inserted at
        #   RowOffset : rows added to the highlighted row (0 = same row; may be negative)
        NoteStamps = @{
            'GIFT_noGfixfile' = @{ Image = 'already_exists.png'; Column = 'AF'; RowOffset = 0 }
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
        ConfigOverlay   = 'ConfigOverlay.ps1'
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
        DeliverMail     = 'DeliverMail.ps1'
        DeliverFiles    = 'DeliverFiles.ps1'
        BackupJ4        = 'BackupJ4.ps1'
        FillCheckSheet  = 'FillCheckSheet.ps1'
        SendVsGift      = 'SendVsGift.ps1'
        SnapVerify      = 'SnapVerify.ps1'
        SnapLocalize    = 'SnapLocalize.ps1'
    }

    # The ONE J4 evidence folder (canonical). Read by DeliverFiles, BackupJ4,
    # the DeliverMail body path, and the review phases' j4 option. REQUIRED
    # by those phases (no personal default is committed here) -- set it per
    # work folder in verify_config.json via -Phase InitConfig. The legacy
    # duplicates DeliverFiles.J4EvidenceDir / Mail.EvidenceFolder still win
    # when non-empty (old configs keep working) but are no longer generated.
    J4EvidenceDir = ''

    # Reviewer mail address (canonical; Outlook draft To). The legacy
    # duplicate Reviewer.Address still wins when non-empty but is no longer
    # generated by InitConfig.
    Address = ''

    # Reviewer (To / 確認者). The single "viewer" param: used as the mail
    # recipient, the body greeting, and the check-sheet 確認者 column.
    # Address here is legacy -- prefer the top-level Address field above.
    Reviewer = @{
        DisplayName = ''
        Address     = ''
        ShortName   = ''
    }

    # DeliverMail phase: one Outlook draft per Excel_NAME. Operator clicks Send.
    Mail = @{
        From  = ''
        # フェーズ token in the subject. 対象物 = Excel_NAME (short code).
        Phase = 'JRV→IDS,IGP_J4'
        # Subject = 【GIFT廃止対応】<Phase>レビュー依頼(<対象物>)  ({0}=Phase, {1}=Excel_NAME)
        SubjectTemplate = '【GIFT廃止対応】{0}レビュー依頼({1})'
        # UNC folder + filename shown in the body. LEGACY duplicate of the
        # top-level J4EvidenceDir -- wins when non-empty, but set the
        # top-level field instead (InitConfig no longer generates this one).
        EvidenceFolder   = ''
        CheckSheetFolder = ''
        CheckSheetFile   = 'レビューチェックシート_REQ-000xxxxx_GIFT廃止対応_OPEN.xlsx'
        # Body lines, joined with CRLF. Placeholders:
        #   {0}=Reviewer.ShortName {1}=Owner {2}=EvidenceFolder
        #   {3}=evidence filename (per Excel) {4}=CheckSheetFolder {5}=CheckSheetFile
        BodyLines = @(
            '{0}さん',
            '',
            'お疲れ様です。{1}です。',
            '',
            '標記件、下記レビューをお願い致します。',
            '{2}',
            '{3}',
            '',
            '・レビューチェックシート',
            '{4}',
            '{5}',
            '',
            '以上、よろしくお願いいたします。'
        )
    }

    # DeliverFiles phase: replaces the delivery-scope sheets (GIFT/GFIX recv
    # result + GIFT-vs-GFIX data compare -- Align.ps1's Get-AlignRecvSheets
    # set) in the corresponding J4 evidence workbook with the matching work
    # sheets, in place (other J4 sheets untouched); first delivery for an
    # Excel_NAME (no existing J4 workbook yet) copies the whole file instead.
    # Also copies DATA\GFIX and DATA\GIFT files to J4. Source files are never
    # deleted -- this phase only copies / in-place sheet-replaces.
    DeliverFiles = @{
        # Root J4 folder for evidence Excel files. LEGACY duplicate of the
        # top-level J4EvidenceDir -- wins when non-empty, but set the
        # top-level field instead (InitConfig no longer generates this one).
        J4EvidenceDir = ''
        # Destination for DATA\GFIX files.
        # Defaults to J4EvidenceDir + '\DATA\GFIX' if blank.
        J4GfixDataDir = ''
        # Destination for DATA\GIFT files.
        # Defaults to J4EvidenceDir + '\DATA\GIFT' if blank.
        J4GiftDataDir = ''
        # true = back up the whole J4 file this phase is about to modify
        # (into J4EvidenceDir\_bak) before replacing its sheets. false = edit
        # directly. For a LOCAL rollback copy instead, use the BackupJ4 phase.
        Backup = $false
        # Local folder the BackupJ4 phase copies J4 files into.
        # Defaults to <WorkDir>\bk if blank.
        BackupLocalDir = ''
    }

    # CheckSheet phase: append one row per Excel to the shared review check
    # sheet, sheet "Check Sheet_J4". Edited via a temp copy first (double-check).
    CheckSheet = @{
        # Actual file to edit (half-width katakana name). Replace REQ-000xxxxx.
        Path      = ''
        SheetName = 'Check Sheet_J4'
        Language  = 'JAVA'          # col C  (COBOL/JAVA)
        Phase     = 'J4内部ﾚﾋﾞｭｰ'    # col E  (確認ﾌｪｰｽﾞ)
        # 1-indexed columns: A No. / B 記入日 / C COBOL/JAVA / D ﾘｿｰｽID /
        # E 確認ﾌｪｰｽﾞ / F レビュー対象 / G 担当 / H 確認者 / (I 完了希望日, J~ blank)
        ColNo = 1; ColDate = 2; ColLang = 3; ColResourceId = 4
        ColPhase = 5; ColTarget = 6; ColOwner = 7; ColViewer = 8
        # Fallback date format if there is no prior row to copy from.
        DateFormat = 'yyyy/m/d'
    }

    # DF snap / mark configuration
    Df = @{
        # Path to df.exe comparison tool.
        #   - Leave empty to be PROMPTED on the first run (the prompt offers
        #     DefaultExePath; press Enter to accept). The answer is remembered
        #     in verify_session.json so later runs skip the prompt.
        #   - Set this to a real path to lock it and never be prompted.
        # CLI -DfExePath / verify_config.json overlay still win over this.
        ExePath        = ''
        # Suggested df.exe path the first-run prompt pre-fills (Enter accepts).
        DefaultExePath = 'C:\tools\DF\DF.exe'
        # Seconds to wait after the df.exe window appears before capturing.
        LoadWaitSec  = 8
        # 'window' (default; auto-fits the df.exe window, falls back to
        # 'region' if the window handle/rect looks bad), 'region', 'fullscreen'.
        CaptureMode  = 'window'
        # Fixed capture rectangle for 'region' mode (target screen ~1980x1020)
        # -- also used as the fallback rectangle when 'window' mode's handle
        # or rect comes back invalid.
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
        #   \\server\share\...\project\40.J4\07.GPCS
        J4BaseDir       = ''
        # FROM_sys / TO_sys literal values that count as "Host" (mainframe).
        # Until set, migration type is Unknown and Align falls back to
        # DefaultMigrationType below (Host->Open send sheets) with a warning.
        HostSystemTypes = @()
        # Scope used when the migration type cannot be classified (HostSystemTypes
        # empty). 'HostToOpen' = delete + copy the send sheets (Soushin data /
        # GIFT / GFIX send result) from J4 -- the project's normal case.
        DefaultMigrationType = 'HostToOpen'
        # A J4 send-result sheet (GIFT/GFIX send result) with this many text rows
        # or fewer is treated as "not prepared" (blank template) and its replace
        # is skipped. The send-data sheet uses a picture-present check instead.
        MinSendResultRows = 3
    }

    # GFIX log cell-mark configuration
    GfixLog = @{
        # Exact text in column B that marks the start of each log region.
        LogAnchor        = ''       # empty -> MarkGfixLog.ps1 uses [char]0x25BC + 'GFIX' + ...
        # Regex matched against B column to find the row to highlight.
        CommandPattern   = "Command:\s*'/appl/[A-Za-z0-9]+/shell/"
        # OLE color for highlight. Yellow RGB(255,255,0) = 65535.
        HighlightColor   = 65535
        # Column range to highlight (B=2, AY=51). When AutoHighlightWidth is
        # on (default), this is also the CAP: the auto-computed width never
        # exceeds it, and the idempotent re-run clear still scans this full
        # range regardless of what width was actually applied last time.
        HighlightColStart = 2
        HighlightColEnd   = 51
        # true (default) = size the highlight to the target row's ACTUAL
        # pasted text width (measured with its real cell font) instead of
        # always filling out to HighlightColStart..HighlightColEnd -- a long
        # Command: line no longer risks running past a short fixed range, and
        # a short one no longer leaves an oversized highlight band.
        # false = legacy fixed-width behavior (always HighlightColStart..HighlightColEnd).
        AutoHighlightWidth = $true
        # Extra columns of padding added after the measured text width.
        HighlightPadCols   = 1
    }

    # SnapVerify: instant NG detection for HM / MQ / Jenkins snap phases.
    # All keys overridable per work folder via verify_config.json overlay.
    SnapVerify = @{
        Enabled           = $true   # master switch; $false = pure screenshot (no detection)
        TimeCheck         = $false  # $false = existence/abend checks only (no run-time
                                    # window prompt or comparison). The time window is
                                    # mostly nice-to-have, so it is OFF by default; set
                                    # $true to be prompted for a run time + tolerance.
        ToleranceMinutes  = 30      # default +-minutes when TimeCheck is on
        SaveText          = $true   # save Ctrl+A text as <correl>.txt alongside PNG
        PollTimeoutSec    = 10
        PollIntervalMs    = 500
        NoGfixNoteColumn  = 'AZ'    # F4: column for past-data annotation

        # M5/F5 pixel localisation: write <correl>.loc.json beside each PNG so
        # the Mark phase can red-box the exact data row the verdict judged.
        # OFF by default -- HM/MQ geometry must be measured for the office-PC
        # window size first (Calibrate-HmGeometry.ps1). Jenkins needs no
        # geometry (it reuses the orange Ctrl+F highlight). Enable per work
        # folder once the *Row1Top / *RowHeight / *ColLeft / *ColWidth are set.
        Localize = @{
            Enabled         = $false  # master switch for sidecar localisation
            Jenkins         = $true   # Jenkins leg (orange highlight; no geometry)
            CropLeft        = 0       # pixels Invoke-CropPng trims (match CropPx)
            CropTop         = 0
            # HM status-column geometry (pre-crop px; 0 = not calibrated -> skip)
            HmRow1Top       = 0
            HmRowHeight     = 0
            HmColLeft       = 0
            HmColWidth      = 0
            # MQ recv-date column (MQ records span 2 lines; RowHeight covers both)
            MqRow1Top       = 0
            MqRowHeight     = 0
            MqColLeft       = 0
            MqColWidth      = 0
            # Jenkins highlight box horizontal extent (0 width -> full PNG width)
            JenkinsColLeft  = 0
            JenkinsColWidth = 0
            JenkinsPad      = 2
            # Find-ActiveHighlightRow tuning (orange FF9632 active match)
            JenkinsActiveR        = 255
            JenkinsActiveG        = 150
            JenkinsActiveB        = 50
            JenkinsTolerance      = 25
            JenkinsMinPixelsPerRow= 30
        }
    }

    # Expected_Time helper (Resolve-ExpectedTime.ps1). The time VALUES live per
    # row in the mapping CSV (not JSON, since they are per-correl); these are the
    # centralized defaults for that helper. Override per work folder in the JSON.
    ExpectedTime = @{
        TimeColumn    = 'Expected_Time'
        IdColumn      = 'Correl_ID_S'
        LookbackHours = 1.0
        TimeFormat    = 'yyyy/MM/dd HH:mm:ss'
    }

    # Phase entries: Field + optional BitValue.
    # If BitValue > 0, the field is read as a bitmask and "done" means
    # (value -band BitValue) -eq BitValue.
    PhaseOrder = @(
        @{ Key='InitConfig';         Field='';                     Label='work-folder config JSON (verify_config.json)'; Status='implemented' }
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
        @{ Key='SendVsGift';         Field='SendVsGift';            Label='SEND data vs GIFT data metadata review'; Status='implemented' }
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
        @{ Key='CheckSheet';         Field='';                     Label='レビューチェックシート 記入';        Status='implemented' }
        @{ Key='BackupJ4';           Field='';                     Label='J4 ﾌｧｲﾙ ﾛｰｶﾙ backup (bk)';        Status='implemented' }
        @{ Key='DeliverFiles';       Field='isFilesDelivered';     Label='J4 ﾌｧｲﾙ転送 (証跡Excel sheet + DATA)'; Status='implemented' }
        @{ Key='DeliverMail';        Field='isDelivered';          Label='レビュー依頼メール 送付';            Status='implemented' }
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

        CheckSheet        = 'CheckSheet'
        FillCheckSheet    = 'CheckSheet'
        RvCheck           = 'CheckSheet'

        DeliverMail       = 'DeliverMail'
        Mail              = 'DeliverMail'
        SendMail          = 'DeliverMail'
        Deliver           = 'DeliverMail'

        DeliverFiles      = 'DeliverFiles'
        FilesDeliver      = 'DeliverFiles'
        CopyJ4            = 'DeliverFiles'
        MoveJ4            = 'DeliverFiles'

        BackupJ4          = 'BackupJ4'
        Bk                = 'BackupJ4'
        BkJ4              = 'BackupJ4'

        SendVsGift        = 'SendVsGift'
        Svgift            = 'SendVsGift'
        GiftMeta          = 'SendVsGift'

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

        InitConfig        = 'InitConfig'
        Config            = 'InitConfig'
        MakeConfig        = 'InitConfig'
        EditConfig        = 'InitConfig'
    }
}
