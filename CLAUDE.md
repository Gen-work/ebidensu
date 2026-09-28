# CLAUDE.md — VerifyTool Project Context

Read this file first when opening the project from an IDE or LLM session.

## Project purpose

**VerifyTool** automates GIFT→GFIX migration evidence collection at Honda Japan.
Operator: project user. Environment: Windows 10/11 + PowerShell 5.1 + Excel 2019.

The tool captures screenshots from HM / MQ / Jenkins, inserts them into evidence Excel
workbooks, draws red rectangles on the relevant cells, and tracks completion state in
a CSV mapping file.

## Repository

Remote: `gen-work/ebidensu`
Branch convention: `claude/<slug>`
Local clone: configure per environment (do not commit personal paths).

Versioning: use `MAJOR.MINOR.PATCH`; release headings in `CHANGELOG.md` carry the current `vX.Y.Z`. See `docs/Versioning.md` for bump rules and release automation guidance.

## File map

```
VerifyTool.ps1          main entry, menu, phase router, status display
VerifyConfig.psd1       project config (paths, scripts, PhaseOrder, Aliases, Mark.Boxes)
verify_session.json     last settings (WorkDir, Owner, WindowSize, CursorCell, CloneSourceDir)
verify_config.json      OPTIONAL per-work-folder JSON overlay (lives in WorkDir): deep-
                        merged over VerifyConfig.psd1 (JSON wins; CLI still wins).
                        Generate via -Phase InitConfig. Customizes owner / window /
                        Mark.Boxes / Mail / Reviewer / Df / ExpectedTime / etc.

ExcelHelpers.ps1        dot-source lib: Excel COM, bitmask, shape metadata helpers (no param())

  -- ebi-dance refactor tree (docs/ebi-dance/) : the target layout the repo is
     being moved into, one card at a time. Empty dirs carry a README.md that
     says what belongs there --
modules/                capability-oriented ebi-dance steps, one .ps1 per step
                        with an inline $Manifest, grouped by capability
                        (browser/screen/file/excel/table/verify/human/progress/
                        flow). Files here that are NOT steps are pre-conversion
                        libraries, listed and exempted on every test run.
                        First three real steps (P0-08, ported from Common.ps1):
                        human/human.prepare.ps1, browser/browser.ensure.ps1
                        (provides 'window'), screen/screen.capture_window.ps1
                        (consumes it via a type='session' input).
legacy/                 retired implementations, kept only while they still have
                        a backlog to clear. Not in the catalog.
kernel/                 runner internals. Trace.ps1 (append-only run trace,
                        unit-tested via Tests\Test-Trace.ps1), Registry.ps1
                        (P1-02: step discovery under modules/ -- Find-EbiStep
                        Files / Get-EbiStepCatalog, a manifest-only scan for
                        lint/help/docs -- and loading for running: Import-
                        EbiStep MUST be dot-sourced (`. Import-EbiStep ...`)
                        so the step's helpers land in the caller's scope; it
                        captures Invoke-Step into the registry table at once
                        and removes the bare name. Also the "with" -> $In
                        pipeline: 'as' lifted out, Test-EbiStepInputs checks
                        required / type / enum / default / unknown parameters
                        against the manifest (every problem names the
                        parameter; failure 'input_invalid', or 'contract_
                        violation' when the manifest is wrong), then session
                        names -> instances; and Test-EbiStepReturn (3.1).
                        Unit-tested via Tests\Test-Registry.ps1) and Runner.ps1
                        (P0-07 spike of the workflow runner: Invoke-EbiWorkflow
                        reads a workflow JSON, loads each step through the
                        registry, runs setup then teardown in a finally, keeps
                        $Ctx.Session and the STEP-CONTRACT 3.4 point 7 resource
                        channel -- reserved return key 'resource', 'with.as'
                        registration, session-input name -> instance
                        replacement -- enforces the 3.1 return contract and
                        traces every step; refuses source/each/templates/
                        when/onError up front with 'unsupported_in_spike'.
                        Unit-tested via Tests\Test-Runner.ps1). Context.ps1
                        (P1-01: pure {{}} template evaluation -- scopes
                        vars/profile/page/run/item/steps, whole-value type
                        preservation, \{\{ escape, one-pass evaluation of
                        profile/page subtrees, failures as records naming
                        the unresolved segment; not yet wired into Runner,
                        P1-03 does that; Tests\Test-Context.ps1) and Key.ps1
                        (P1-27 seed: full-width folding, item.key display
                        form, item.keySafe file-safe form -- the ONE place
                        key normalization lives).
workflows/              JSON workflows (the artifact a human or Agent writes).
                        spike.capture_window.json is the P0-08 end-to-end
                        spike (prepare -> ensure -> capture one window).
ebi.ps1                 ebi-dance CLI entry: run / dryrun / help so far
                        (P1-07..P1-10 add lint / explain / doctor and the
                        real run options). Has param(): call via -File or &,
                        never dot-source.
profiles/               per-project data: page bindings, decision rules, schemas

  -- shared dot-source libraries (no param(); ASCII source; no BOM) --
MappingStore.ps1        single source of truth for mapping_<Owner>.csv: read/filter/
                        atomic-write. Import-Mapping, Export-MappingAtomic,
                        Ensure-MappingColumns, ConvertTo-TargetIdList, Test-TargetRow,
                        Get-PendingRows, Set-MappingBit. ALL scripts use this.
EvidencePlan.ps1        pure correl-major Replace plan builders (Build-Gift/Gfix/Df
                        EvidencePlan) encoding the review order. No Excel. Unit-tested.
EvidenceExecutor.ps1    walks an EvidencePlan and performs the Excel inserts.
ProjectLabels.ps1       Japanese sheet/label names from [char] (keeps consumers ASCII /
                        codepage-agnostic). Get-AlignSendSheets / Get-AlignRecvSheets.
ProgressLog.ps1         append-only status\progress.jsonl events (UTF-8 no BOM).
AlignCompare.ps1        pure sheet-compare + migration-type logic. Unit-tested.
ConfigOverlay.ps1       pure per-work-folder JSON overlay: deep-merge + JSON<->hashtable
                        + InitConfig snapshot/generator helpers. Unit-tested.
SendMetadata.ps1        pure SEND-side OCR-line parsing + send-vs-gift compare for
                        SendVsGift Stage 2 (word-box spacing rebuild, 0-byte rules:
                        used-CYLINDERS-0 / begin+end-on-one-image, row-label record
                        extraction, 80% prefix-similarity record compare,
                        Compare-SendGiftEvidence ok/ng/unknown verdict). Unit-tested.
OcrWindows.ps1          Windows built-in OCR (Windows.Media.Ocr WinRT) from PS 5.1;
                        same engine family as Snipping Tool text extraction. Safe to
                        dot-source anywhere (lazy init; Test-WinOcrAvailable).
EvidenceImageExport.ps1 Excel COM: export embedded sheet pictures to PNG via temp
                        ChartObject (skips verifyMark_* shapes; flattens Ctrl+G groups
                        to child pictures; optional Top range filter for one correl
                        section; clipboard clobbered).
SnapLocalize.ps1        M5/F5 wiring glue (NOT pure: System.Drawing + the
                        Find-ActiveHighlightRow scan). Write-SnapLocalize turns a
                        verdict into a snap\<folder>\<correl>.loc.json sidecar via
                        the pure SnapVerify geometry; swallows all errors (never
                        blocks snapping). Dot-sourced by Hm/Mq/JenkinsSnap when
                        SnapVerify.Localize.Enabled. No param() = safe to dot-source.
WorkbookResolver.ps1    dot-source helper: evidence/J4 workbook filename resolution
                        (prefix + Excel_NAME stem) plus reusable full-width
                        ASCII filename fallback (`FullWidthFilenameResolver`).
                        Unit-tested.
ProcessTimeParse.ps1    pure HM processing start/end/duration helpers for the
                        ProcessTime phase: Get-ProcessDurationText (HH:mm:ss,
                        not clamped to 24h), ConvertTo-ProcessTimeNormalizedLine
                        (repairs OCR-injected spaces INSIDE time tokens:
                        '10 :58 :20' / '00 :00 : 0 1' -> HH:mm:ss),
                        ConvertFrom-ProcessTimeOcrLines (anchors on normalized
                        datetime tokens per OCR'd row instead of column
                        position; fuzzy '...shuuryo' status match; emits
                        PageDuration / CorrelSeen / Partial for cross-checks),
                        Select-ProcessTimeRow + Get-ProcessTimeRowRank (full >
                        partial, correl-seen > unseen, newest among equals),
                        Get-ProcessTimeOcrMissNote (why a read yielded nothing),
                        Get-NewestProcessTimeRow (newest-by-StartTime;
                        -MinimumTimeOfDay default 09:00),
                        Resolve-ProcessTimeRowPlan (-Stage + sidecar-exists +
                        ProcessTime_Inserted bitmask bits + -Force -> per-row
                        NeedsOcr/NeedsWrite) +
                        Get-ProcessTimeMigratedInsertedValue (legacy plain
                        '1' -> bitmask '3'), Get-ProcessTimeOutputTag /
                        Get-ProcessTimeOutputFileName / Resolve-ProcessTime
                        OutputDir (config-driven output tag classification,
                        not hardcoded to JDL/JRV; -DeriveFromName derives an
                        unlisted tag from Excel_NAME chars 2-4) and
                        Get-ProcessTimeCheckSummaryLine (end-of-run
                        manual-check summary line).
                        ConvertTo-ProcessTimeDateTimeValue / ConvertTo-
                        ProcessTimeDurationValue parse the
                        sidecar's formatted stamps back into real Excel
                        date/time serials for the output workbook's value
                        cells + check-formula columns.
                        ConvertTo-ProcessTimeCorrelKey
                        also strips OCR-inserted whitespace before folding;
                        Select-ProcessTimeRow's -MinimumTimeOfDay
                        (default 09:00) drops HM history rows; Get-ProcessTime
                        RecordCount falls back to the count immediately before
                        the result diamond when no datestamp anchors it (JDL).
                        No Excel/OCR. Unit-tested (Tests\Test-ProcessTimeParse.ps1).
ProcessTimeCheck.ps1    pure ProcessTime output-workbook audit ("check") column
                        module (dot-source, no param(), no COM): Get-ProcessTime
                        CheckColumnSpec returns the data-driven spec for the
                        columns appended after A..H -- I 処理時間(検算) (=E-D),
                        J チェック (T/F compare of written vs re-derived
                        duration), K 件数チェック (record-count check; first
                        version flags a blank/zero/non-numeric count) -- and
                        New-ProcessTimeCheckFormula fills a formula template's
                        {0} with a row number. With -CountReference
                        (Resolve-ProcessTimeCountReference expands {Tag}/
                        {Month} in the reference workbook/sheet names;
                        New-ProcessTimeExternalRange builds the
                        'dir\[Book.xlsx]Sheet'!$O$1:$O$20000 qualifier;
                        New-ProcessTimeCountLookupFormula emits the operator's
                        INDEX/MATCH) and the spec grows to four columns --
                        K 件数(参照) pulls the EXPECTED count out of the
                        project's monthly workbook and L 件数チェック compares
                        it against the OCR-read count. The layout is
                        ALWAYS I/J/K/L -- with no reference configured K holds
                        an inert TEXT placeholder carrying <DIR>/<BOOK>/
                        <SHEET> tokens (New-ProcessTimeCountPlaceholderFormula,
                        CountReference.PlaceholderWhenUnset) -- and L compares
                        the count against K when K has a value, else against
                        the SAME correl's other side (GIFT vs GFIX) via
                        Get-ProcessTimeCountPairMap + the template's {1}.
                        Equal counts read OK, INCLUDING 0 vs 0. ProcessTime.ps1's COM-side
                        Set-ProcessTimeCheckColumns walks the spec to write the
                        headers/formulas/number-formats uniformly after the
                        data rows. Japanese headers via [char]. Unit-tested
                        (Tests\Test-ProcessTimeCheck.ps1).

  -- modules/verify/ : the same pure libs, moved out of the repo root by
     P0-04. Still plain dot-source libraries, NOT ebi-dance steps yet --
GfixLog.ps1             pure GFIX receive-log matcher (SS_CODE=Substring(4,1); newest
                        wins; whole-file lines). No Excel. Unit-tested.
GfixJobList.ps1         pure parser for the GoAnywhere completed-jobs LIST page text
                        (Ctrl+A/Ctrl+C capture): ConvertFrom-GfixJobListText (tab-
                        delimited rows keyed by JobNo, data rows identified by a
                        numeric JobNo regex -- no Japanese literals needed) +
                        Get-GfixJobListRowsForIf (filter by normalized IF_NO,
                        receive-side only by default). No COM. Unit-tested. Lets
                        GfixLogDownload fetch every job matching a needed IF_NO
                        (job numbers are unique; IF_NO/project-name text is not).
ScreenRegion.ps1        pure screen-region clamp math + Resolve-DirectionalCrop
                        (four-side snap crop resolution: CropPx + per-side +
                        per-folder overrides). Unit-tested. Dot-sourced by
                        VerifyTool.ps1.
SnapVerify.ps1          pure snap-phase NG detection + localisation library
                        (no COM, no SendKeys). ASCII source (Japanese via [char]).
                        ConvertFrom-HmPageText / Test-HmAbend (F1),
                        ConvertFrom-MqPageText / Test-MqRecord (F2),
                        ConvertFrom-JenkinsListText / Test-JenkinsFile (F3/F4),
                        Get-JenkinsSearchTerm (the EXACT listed file name
                        Ctrl+F should look for -- resolves the intended entry
                        via Select-JenkinsFileCandidate so a correl with
                        several listed reruns highlights the newest row
                        instead of whichever the page listed first; '' when
                        nothing matches, so the caller keeps its base-id
                        fallback),
                        Get-SnapPageKind (A3 sentinel), Resolve-SnapRunTime (2.2),
                        and M5/F5 pixel localisation: Get-MatchedRowIndex /
                        Get-RowPixelRect / Get-JenkinsHighlightRect /
                        New-SnapLocRect / Save-SnapLocSidecar.
                        Unit-tested via Tests\Test-SnapVerify.ps1.
OwnerFilter.ps1         pure WBS owner-cell matching (Test-OwnerMatch: exact /
                        owner<-other / other->owner; reverse dir = not owned)
                        + Select-JobsByOwner (filter explicit -Add JOB_NAMEs by
                        WBS owner; jobs absent from WBS kept as temp). No Excel.
                        Unit-tested (Tests\Test-OwnerFilter.ps1).
GiftMqProcessTime.ps1   pure library behind the standalone GiftMqProcessTime
                        .ps1 driver: GIFT MQ LIST-page Ctrl+A text
                        -> records (ConvertFrom-GiftMqListText; two-line
                        records, Send date = job START), Detail-page text ->
                        fields (ConvertFrom-GiftMqDetailText;
                        Get-GiftMqDetailEndTime = INSERTDATETIME = END,
                        Test-GiftMqDetailMatchesRecord = the per-click
                        safety net), the operator's mapping.xlsx schedule
                        (Get-GiftMqMappingColumns / ConvertTo-GiftMq
                        Schedules / ConvertTo-GiftMqScheduledTime -- Excel
                        float-artefact times rounded), schedule-window +
                        day-order matching (Resolve-GiftMqJobMatches:
                        ok / ambiguous / notime / none), Teams chat text ->
                        counts (ConvertFrom-GiftMqTeamsText, W-name -> J-job
                        via Get-GiftMqJobFromExcelName), output-sheet row
                        plan (Get-GiftMqOutputPlan: update blank cells /
                        append pair; n-th match <-> n-th row). No COM.
                        Unit-tested (Tests\Test-GiftMqProcessTime.ps1).

  -- legacy/ : retired by P0-05. These exist only to clear the backlog of
     OCR-only old HM snapshots; outside the step catalog, no new workflow
     may depend on them. Retirement conditions: legacy/README.md --
OldSnapVerify.ps1       pure old-snap 9->3 hand-verification helpers (dot-source,
                        no param(), no COM): Resolve-OldSnapImagePath (build the
                        snap\<GIFT|GFIX>_HM\<correl>.png path via
                        [IO.Path]::Combine so a rooted Windows path is
                        CI-testable), Test-OldSnapDurationArithmetic (dur == end
                        - start, the J-column mirror), Repair-ProcessTimeStart
                        FromStamp (adopt the clean 14-digit datestamp HH:mm on a
                        pure 3<->9 swap), Get-OldSnapVerifyVerdict (the
                        conservative Txt/OcrOk/NeedsCheck/NoSnap decision) +
                        Get-OldSnapVerifyLabel / Get-OldSnapVerifyColumnSpec (the
                        検証 column), the D1 FALLBACK image
                        helpers: Resolve-OldSnapExportImageDir (the per-correl
                        snap\ProcessTime\<correl> export folder) and
                        Select-OldSnapFallbackImageName (rank the pictures this
                        phase exported OUT of the evidence workbook -- section >
                        below-label > whole-sheet > above-label, lowest index
                        first; *_pre.png OCR derivatives never linked) so a
                        correl with no standalone snap PNG still gets a
                        clickable image, plus Resolve-OldSnapPromotionMarker
                        Path ('<snap>.promoted.json', marking a snap
                        PNG that was copied out of the evidence workbook
                        rather than captured, so the pixel check keeps
                        excluding its re-scaled geometry on later runs).
                        Japanese via [char]. Unit-tested
                        (Tests\Test-OldSnapVerify.ps1).
PixelDigitMatch.ps1     pure D2 per-digit 3/9 image scorer (dot-source, no
                        param(), no COM/GDI): grayscale ink -> binarize -> trim
                        bbox -> average-pool to a normalized grid -> normalized
                        cross-correlation (ConvertTo-DigitInk, Get-DigitInkBBox,
                        ConvertTo-DigitNormalizedGrid, Get-DigitNcc, Get-Digit
                        Similarity, Compare-DigitCandidate, Get-DigitPixelVerdict,
                        Merge-DigitPixelVerdicts). The PS port of the Phase-0
                        GO-proven mock-page/pixeldiff.mjs metric. Unit-tested
                        (Tests\Test-PixelDigitMatch.ps1).
OldSnapPixelVerify.ps1  PARKED with the rest of the D2 image-check line
                        (docs/Parked-Ideas.md); kept, off by default -- do
                        not run two competing D2 paths.
                        NON-pure GDI+ glue for D2 (dot-source, no param();
                        static-checked only). New-DigitTemplateGray (render MS
                        Gothic 3/9), Get-BitmapGrayRegion (crop a snap digit
                        box), Get-OldSnapDigitVerdict, Resolve-OldSnapTimeDigit
                        Rects (per-digit rects from calibrated cell geometry --
                        the one office-PC-calibrated input; empty until then ->
                        conservative '') and Get-OldSnapRowPixelVerdict. Drives
                        PixelDigitMatch. Every entry point swallows errors ->
                        '' so an image check never blocks the write. Off by
                        default (OldSnapVerify.PixelDiff.Enabled).
TimeDigitVerify.ps1     pure 3<->9 digit-risk analysis for OCR'd HM times
                        (dot-source, no param(), no COM/OCR/IO, ASCII).
                        The project's answer to the ja recognizer reading
                        MS Gothic '9' as '3': act only where the answer is
                        FORCED, flag everything else, never rewrite a
                        plausible reading.
                        Repair-ImpossibleTimeDigit -- fix a digit only when
                        exactly one 3<->9 substitution brings its field back
                        into range ('10:93:20' -> '10:33:20'; '2026/19/03'
                        untouched + Invalid).
                        Resolve-ProcessTimeDurationConflict -- treat start,
                        end and the page's own printed processing-time column
                        as three readings of one fact; a disagreement that
                        EXACTLY ONE substitution reconciles is repaired by
                        arithmetic ('repaired'), several or none leaves the
                        values as read ('ambiguous'/'conflict') and flags the
                        row. This replaced the old silent "kept derived"
                        override that shipped wrong durations.
                        Get-TimeDigitRisk -- 'none'/'suspect'/'invalid'
                        classification, never rewrites.
                        Get-ProcessTimeDigitFormatRule +
                        New-ProcessTimeDigitFormatFormula -- the conditional
                        format that reddens every start/end/duration cell
                        whose SECONDS digit is a 3 or 9 (applied by
                        ProcessTime.ps1's Set-ProcessTimeDigitFormat; gated
                        on ProcessTime.EmitDigitFormat).
                        Also Get-TimeDigitFieldSpec / Get-TimeDigitSwapVariants
                        / ConvertTo-TimeDigitDurationSeconds /
                        Format-TimeDigitDuration. Unit-tested
                        (Tests\Test-TimeDigitVerify.ps1).


Clone.ps1               Phase Clone
Align.ps1               Phase Align/Precheck: compare work evidence vs J4 baseline
ReplaceEvidence.ps1     Phase ReplaceGift / ReplaceGfix / ReplaceDf (plan-driven)
ProcessTime.ps1         Phase ProcessTime: extracts each correl's HM batch
                        processing start/end time (GIFT + GFIX) and derives
                        the duration -- archived snap\GIFT_HM|GFIX_HM\
                        <correl>.txt first (ConvertFrom-HmPageText), else OCR
                        of the HM screenshot already inserted into the
                        evidence workbook: content-validated candidate
                        pictures (section, below-label, above-label; relaxed
                        candidates must show the correl id in their OCR text)
                        each read via Invoke-WinOcrFile +
                        ConvertFrom-ProcessTimeOcrLines / Select-ProcessTimeRow.
                        Writes one row per GIFT/GFIX side per correl to
                        <label>(<Tag>).xlsx evidence workbooks under
                        ProcessTime.OutputDirectory (classified per
                        ProcessTime.OutputTags, default JDL/JRV but
                        extendable, e.g. JDS; a row matching no
                        tag goes to UnclassifiedTag instead of aborting the
                        run; OutputMode 'Single' writes one untagged
                        workbook instead; OutputDirectoryByTag routes a tag
                        to its own destination directory). Run after
                        ReplaceGift/ReplaceGfix. Sets ProcessTime_Inserted,
                        a bitmask (bit 1 = OCR'd, bit 2 =
                        written; a legacy plain '1' is migrated to '3').
                        -Stage Ocr|Write|Both runs the
                        extract-and-cache-to-sidecar step and the
                        write-the-output-workbooks step independently; each
                        correl's OCR result (both sides, combined) is cached
                        at snap\ProcessTime\<correl>\result.json so a
                        Write-only rerun opens no evidence workbook at all.
                        Prints an end-of-run "needs manual check" summary
                        listing every correl whose GIFT and/or GFIX side was
                        not matched.
Mark.ps1                Phase MarkGift / MarkGfix / MarkDf. Each Mark.Boxes
                        entry may add a 'Template' key to try image-recognition
                        placement (Locate-ByImage.ps1 LockBits match against
                        the source snap PNG) before falling back to the fixed
                        OffsetX/OffsetY box -- see mark_templates/README.txt.
                        A box may also add BaseRow/RowHeight (GIFT_MQ) to
                        shift OffsetY for correls whose page shows a
                        different record count than the calibrated baseline
                        row; the target row/count come from a snap-time
                        <correl>.mqrow.json sidecar, else a re-parse of
                        <correl>.txt, else English OCR of <correl>.png.
                        -Mode Gfix also highlights the GFIX log Command: row,
                        auto-sized to the row's actual text width (GfixLog.
                        AutoHighlightWidth, capped at HighlightColEnd).
mark_templates/         reference images for Mark.ps1's optional image-match
                        box placement (see mark_templates/README.txt). Ships
                        empty; populated per project on an office PC.
ReviewEvidence.ps1      Phase ReviewGift / ReviewGfix / ReviewDf / ReviewEvidence
SendVsGift.ps1          Phase SendVsGift: GIFT file metadata vs SEND evidence review.
                        Rows grouped per workbook (opened once); cursor jumps to each
                        Correl_ID_S label in column A of the send sheet; Excel is
                        refocused after every console answer. Enter=1, n=2(NG), s, q.
                        Stage 2 -Ocr exports each correl's section pictures, OCRs and
                        auto-marks ok->1 / ng->2 / unknown->prompt (docs/SendVsGift.md).
OcrTool.ps1             standalone Windows-OCR CLI over OcrWindows/SendMetadata/
                        EvidenceImageExport: images, dirs, wildcards or -Workbook
                        picture export; -Json output; -ListLanguages. Reusable by
                        future features (has param(): call via &, never dot-source).
FillCheckSheet.ps1      Phase CheckSheet: append a row per Excel to the shared
                        review check sheet (Check Sheet_J4) via a temp-copy
                        preview, then commit only if the original is unchanged.
DeliverMail.ps1         Phase DeliverMail: one Outlook *draft* per Excel_NAME
                        (CreateItem+Display, never auto-sent); operator clicks
                        Send then Enter -> sets isDelivered. ASCII source;
                        subject/body/reviewer come from config (Mail/Reviewer).
DeliverFiles.ps1        Phase DeliverFiles: replaces the 3 delivery-scope
                        sheets (GIFT/GFIX recv result + GIFT-vs-GFIX data
                        compare -- Align.ps1's Get-AlignRecvSheets set) in
                        the corresponding J4 workbook with the matching
                        work sheets, in place (other J4 sheets untouched);
                        first delivery for an Excel_NAME copies the whole
                        file instead. Also copies DATA\GFIX/GIFT. Never
                        deletes source files. Sets isFilesDelivered.
BackupJ4.ps1            Phase BackupJ4 ("bk"): read-only against J4 --
                        copies each targeted Excel_NAME's current J4
                        workbook into a local, timestamped backup folder
                        (default <WorkDir>\bk). Run before DeliverFiles to
                        keep a local rollback point (DeliverFiles now edits
                        J4 workbooks in place instead of always overwriting
                        the whole file).
Validate.ps1            Phase Validate (read-only diagnostic)
Watch-MappingProgress.ps1  read-only progress monitor (does NOT lock mapping)
Check-Encoding.ps1      read-only encoding policy checker + label self-test
Tests/                  Run-Tests.ps1 (parse-checks every .ps1 in the tree, then
                        runs Tests\**\Test-*.ps1) + Test-*.ps1.
                        Test-Docs.ps1 + DocsCheck.ps1 + docs-checks.json make the
                        ebi-dance docs self-checking: retired spellings, card-count
                        agreement across BACKLOG/Plan/README, cross-file section
                        refs, and workflow-example integrity.
                        StepContract.ps1 + Test-StepContract.ps1 are the step
                        contract checker (P0-06): every step file under modules/
                        is checked against spec/STEP-CONTRACT.md section 7 --
                        dot-sourceable, no param(), id == file name, no
                        required+default, non-empty failures each with a boolean
                        transient, example params declared, JSON-serializable
                        outputs, sessionKind on session inputs, at most one
                        provides, releases backed by a session input, resource
                        kinds declared in the spec's mustRelease table, resource
                        steps idempotent, prefixed helper names, ASCII source.
                        The mustRelease kind table is PARSED OUT of the spec, so
                        adding a kind stays a one-file edit.
docs/Parked-Ideas.md    designed-then-deliberately-shelved work, with what
                        it would take to resume. NOT on the TODO list.
                        Currently holds the whole D2 image-check line
                        (mock-page + PixelDigitMatch + OldSnapPixelVerify)
                        and the never-wired Repair-ProcessTimeStartFromStamp.
docs/ProcessTime-OldSnap-MockMatch-Plan.md
                        PARKED (see docs/Parked-Ideas.md). Designed
                        replacement for the D2 image check: render the
                        reference row from mock-page in Edge on the office PC
                        (same engine/font/CSS as the snap) and whole-field
                        template-match it via Locate-ByImage, instead of GDI+
                        per-digit templates. Never built -- both D2
                        generations need an office-PC calibration session
                        that never happened, and the deterministic checks in
                        TimeDigitVerify.ps1 cover the day-to-day 3/9 problem
                        without one.
docs/ProcessTime-OcrBenchmark-Plan.md
                        Phase 2/3 design: OCR benchmark test set + 3/9 fix
                        (real-snap + synthetic-HTML ground truth, runner,
                        datestamp hh:mm cross-correction, preprocessing tuning,
                        optional bounding-box second-pass). Office-PC-run;
                        CI proves pure logic only.

JenkinsSnap.ps1         Phase GiftJenkins / GfixJenkins / GiftJenkinsNoFile
HmSnap.ps1              Phase GiftHmSnap / GfixHmSnap. MappingStore + ProgressLog
                        + SnapVerify F1 detection (page-text poll, page-kind
                        sentinel, HM abend verdict ok=1/ng=2/ask, newest-wins
                        within the time window, batch Expected_Time prompt).
                        Per-TO_code appl grouping (one HM page per appl).
MqSnap.ps1              Phase GiftMqSnap. MappingStore + ProgressLog + SnapVerify
                        F2 detection (page-text poll, page-kind sentinel, MQ
                        record verdict ok=1/ng=2, batch Expected_Time prompt).
                        Also writes <correl>.mqrow.json (the verdict's target
                        row index + record count, via Get-MatchedRowIndex)
                        so Mark.ps1 can shift the GIFT_MQ red box when a
                        correl shows other than the usual 2 records.
ExcelSnap.ps1           Phase ExcelSnap                 (legacy, kept as-is)
Common.ps1              shared WinAPI/screenshot/SendKeys helpers (dot-sourceable)
Generate-HostOpenMapping.ps1  generates mapping CSV from wipGFIX一覧.xlsx.
                        -Add merges new selectors (JobNames / CorrelIdsM /
                        ExcelNames / WBS range) into an existing mapping,
                        keeping every existing row + its progress.

Calibrate-HmGeometry.ps1  4-click WinForms calibration for HM geometry offsets
Find-Abend.ps1          template-match for HM status cell
Find-ActiveHighlightRow.ps1  detects Edge Ctrl+F active-match (orange) row
Locate-ByImage.ps1      C#-compiled LockBits template matcher; called by
                        Mark.ps1 (see the Mark.ps1 entry above) for optional
                        image-recognition box placement
Pack-LlmContext.ps1     packs project context to clipboard for LLM ingestion
Apply-LlmPatch.ps1      applies XML / git-unified-diff patches from clipboard
Export-DailyPatch.ps1   extracts today's git diff to clipboard
Parse-GiftMq.ps1        parses GIFT/MQ transfer status page text
GiftMqProcessTime.ps1   STANDALONE (not a phase, has param(): call via -File /
                        &) daily 処理時間(<Tag>).xlsx filler from the GIFT MQ
                        page (docs/GiftMqProcessTime.md): reads the
                        operator's own mapping.xlsx (JOB / owner / GIFT run
                        date / GIFT TIME), captures the MQ LIST page once
                        (Send date = start), keyboard-navigates each row's
                        Detail (Ctrl+F row text -> Esc -> Tab -> Enter, or
                        Ctrl+F title -> Tab N; every page reached is verified
                        against its record before INSERTDATETIME = end is
                        trusted; Back button / Alt+Left to return), takes
                        counts from a pasted Teams text file, and updates /
                        appends GIFT rows in the output sheet (real Excel
                        date/time cells, =E-D kept, '<n>件' text). Prints the
                        end/count cells still blank so the operator types
                        only those. COM + SendKeys glue, static-checked only.
Parse-JenkinsList.ps1   parses Jenkins file list page text (standalone; with
                        -CorrelId resolves the NEWEST matching entry, not the
                        first listed)
JenkinsDownload.ps1     Jenkins receive-file download glue: Select-Jenkins
                        DownloadFiles (-PreferNewest, default on: several
                        entries for one correl are reruns of one transfer, so
                        only the newest is fetched; the JOB_NAME fallback is
                        NOT narrowed) + Sort-JenkinsFilesNewestFirst /
                        Get-JenkinsFileTime, Invoke-JenkinsFileDownload
                        (reports the passed-over entries as Superseded).
                        Unit-tested via Tests\Test-JenkinsDownload.ps1.
Probe-Shapes.ps1        lists all shapes in an evidence workbook (calibration aid)
Probe-SheetFormat.ps1   read-only cell-FORMAT probe (calibration aid):
                        dumps a workbook's / one sheet's column widths, row
                        heights and distinct format signatures (NumberFormat,
                        font, colors as raw BGR Longs, alignment, borders)
                        with sample addresses; optional -Json report. Use to
                        match generated output (e.g. ProcessTime) to a
                        delivery template. Has param() -> call via &.
Read-ClipboardJson.ps1  polls clipboard for JSON from bookmarklet
Read-PageText.ps1       captures visible text from foreground Edge page via clipboard
Resolve-ExpectedTime.ps1  interactive Expected_Time column helper
Sample-HighlightColor.ps1  samples a single pixel RGB for highlight calibration

CLAUDE.md               this file
README.md               user-facing documentation
CHANGELOG.md            iteration log
```

## Conventions

### Dot-source safety rule

Only files with **no** `param()` block are ever dot-sourced. In the repo root:
`ExcelHelpers.ps1`, `MappingStore.ps1`, `EvidencePlan.ps1`, `EvidenceExecutor.ps1`,
`ProjectLabels.ps1`, `ProgressLog.ps1`, `AlignCompare.ps1`, `ConfigOverlay.ps1`,
`Common.ps1`, `WorkbookResolver.ps1`, `SendMetadata.ps1`, `OcrWindows.ps1`,
`EvidenceImageExport.ps1`, `SnapLocalize.ps1`, `Find-ActiveHighlightRow.ps1`,
`ProcessTimeParse.ps1`, `ProcessTimeCheck.ps1`. In `modules/verify/`:
`GfixLog.ps1`, `GfixJobList.ps1`, `ScreenRegion.ps1`, `SnapVerify.ps1`,
`OwnerFilter.ps1`, `GiftMqProcessTime.ps1`. In `legacy/`: `OldSnapVerify.ps1`, `PixelDigitMatch.ps1`,
`OldSnapPixelVerify.ps1`, `TimeDigitVerify.ps1`. In `kernel/`: `Trace.ps1`,
`Registry.ps1`, `Runner.ps1`, `Context.ps1`, `Key.ps1`. Every `modules/**/<group>.<verb>.ps1` step file is dot-sourced by
the runner too (STEP-CONTRACT: no `param()`, helpers prefixed with the step id).
In `Tests/`: `_TestCommon.ps1`, `DocsCheck.ps1`, `StepContract.ps1`.
All phase scripts have `param()` and are called via `& $path @args`.

The dot-source **path** moved with the file -- `. (Join-Path $PSScriptRoot
'modules/verify/SnapVerify.ps1')` -- and so did `VerifyConfig.psd1`'s
`Scripts.SnapVerify` entry, which `VerifyTool.ps1` resolves with the same
`Join-Path $PSScriptRoot`.

The critical pattern before any dot-source:
```powershell
$forceFlag = [bool]$Force.IsPresent   # capture switch BEFORE dot-sourcing
. $cfg.Scripts.ExcelHelpers            # dot-source (no param() = safe)
```

Never dot-source a script that has a `param()` block — it will overwrite the caller's
switch parameters with `$false`.

### Full-width filename fallback

`WorkbookResolver.ps1` exposes a reusable `FullWidthFilenameResolver` class and
wrapper functions for filename misses caused by full-width ASCII characters
(e.g. `０` instead of `0`). Use `Resolve-FullWidthFileName` after a normal exact
lookup fails when any file type needs the same tolerance:

```powershell
$path = Resolve-FullWidthFileName -Dir $dir -Name 'report0.txt' -Filter '*.txt' `
    -ItemKind 'file' -FullWidthFallback Prompt
```

Workbook callers should continue to use `Find-WorkbookByExcelName`; it preserves
exact and wildcard matching first, then delegates to the generic resolver for
full-width fallback with `Prompt` / `Accept` / `Reject` policy. Interactive tools
should keep the default `Prompt`; tests and non-interactive batch flows should
pass `Accept` or `Reject` explicitly.

### Encoding table

| File type | Encoding | BOM |
|-----------|----------|-----|
| .ps1 | UTF-8 | **no** (keep source ASCII; build Japanese via `[char]`) |
| .psd1 | UTF-8 | **yes** if it holds raw Japanese (Import-PowerShellDataFile can't use `[char]`); no if pure-ASCII |
| .json / .jsonl | UTF-8 | no |
| .csv (mapping) | UTF-8 | yes (BOM; Excel needs it for Japanese) |
| .md | UTF-8 | no |

New/rewritten `.ps1` must be ASCII-only: any Japanese used at runtime comes from
`ProjectLabels.ps1` ([char] code points). This makes the file work on **any**
Windows codepage. Raw Japanese in a no-BOM `.ps1` mojibakes on a JP-locale host
(this silently broke owner-matching in Generate-HostOpenMapping). Older files
still carry raw Japanese / a BOM; migrate them when touched. `Check-Encoding.ps1`
enforces this policy; `Apply-LlmPatch.ps1` preserves original BOM state.

### Bitmask fields

Three integer CSV columns track multi-mode completion:

| Field | bit 1 (1) | bit 2 (2) | bit 4 (4) | all done |
|-------|-----------|-----------|-----------|---------|
| isReplaced | GIFT replace | GFIX replace | DF replace | 7 |
| isMarked | GIFT mark | GFIX mark | DF mark | 7 |
| isReviewed | GIFT review | GFIX review | DF review | 7 |

Test: `($value -band $bit) -eq $bit` (use `Test-BitDone` / `Set-MappingBit` from
MappingStore). The GFIX-log yellow highlight is now part of MarkGfix (bit 2 of
`isMarked`); the old standalone `isGfixLogMarked` column was removed. Replace marks
a mode's bit only when ALL its required pieces inserted; which correl/step/file
failed is recorded in `status\progress.jsonl`, not in extra columns.

A free-text `ReviewComment` column (per Excel_NAME group) holds review notes
captured via the `-m "comment"` option at the Review prompt; list them with the
`Comments` phase.

`SendVsGift` is a plain value column (NOT a bitmask): `0`/empty = pending,
`1` = OK, `2` = NG (OCR auto-compare or the operator's `n` answer flagged a
disagreement). `2` still counts as pending: it is re-offered on the next
SendVsGift run and listed in the end-of-run NG summary.

`isDelivered` is a plain `0/1` flag (NOT a bitmask): set to `1` per Excel_NAME
when the operator confirms the DeliverMail draft was sent. `DeliverComment` is
its free-text note column (captured with `-m "comment"` at the DeliverMail
prompt). Both are defaulted by MappingStore; `isDelivered` is a `PhaseOrder`
field so it is auto-added on startup and shown in Status. The `CheckSheet` phase
writes only to the external review check sheet workbook — it does not touch the
mapping.

`GIFT_ProcessTime` / `GFIX_ProcessTime` are plain, informational per-side
value columns (NOT a bitmask, NOT this phase's `Get-PendingRows` field):
`0` not yet attempted, `1` start/end extracted, `2` not found. The
`ProcessTime` phase gates its own reprocessing on `ProcessTime_Inserted`,
which **is** a bitmask (matching the `isReplaced`/`isMarked`/
`isReviewed` convention above): bit 1 (1) = this correl's OCR result has
been extracted and cached (per-correl sidecar under
`snap\ProcessTime\<correl>\result.json`); bit 2 (2) = the row has been
written into an output workbook; `3` = both done. A pre-v2.15.0 mapping's
plain `1` (the old "written" flag) is migrated to `3` once, on load
(`Get-ProcessTimeMigratedInsertedValue`, `ProcessTimeParse.ps1`) — a legacy
write could only ever happen after OCR succeeded, so it is never migrated to
just bit 2. `-Force` redoes whichever stage(s) `-Stage` selects regardless of
its bit. Output is written per configurable tag (`ProcessTime.OutputTags`,
default `JDL`/`JRV` but not limited to them — e.g. add `JDS`) into
`<ProcessTime label>(<Tag>).xlsx`; a result row matching no configured tag is
routed to `ProcessTime.UnclassifiedTag` (default `Other`) instead of
aborting the whole write, and `ProcessTime.OutputMode = 'Single'` writes
every row into one untagged workbook instead. `ProcessTime.
OutputDirectoryByTag` routes a tag to its own destination directory.

`PhaseOrder` in `VerifyConfig.psd1` has a `BitValue` key for each bitmask phase.

### Excel COM rules

- Always `$xl.Visible = $true` first, then `$xl.DisplayAlerts = $false`.
- Release COM objects in reverse order: `[Runtime.InteropServices.Marshal]::ReleaseComObject($ws)` etc.
- Use `$xl.Quit()` only when you opened a fresh Excel instance.
- `ExcelHelpers.ps1` functions: `Open-ExcelWorkbook`, `Save-ExcelWorkbook`, `Close-ExcelWorkbook`,
  `Set-MappingBit`, `Get-MappingValue`, `Find-ShapeByAltText`, `Add-VerifyMarkRect`.

### Shape metadata

Red rectangle mark shapes are named `verifyMark_<folder>_<idx>` and have AltText
`verifyMark|<folder>|<idx>|<correl>`.

`Probe-Shapes.ps1` reads these to help calibrate `Mark.Boxes` offsets in `VerifyConfig.psd1`.

### Switch flag pattern

```powershell
param(
    [switch]$Force,
    ...
)
$forceFlag = [bool]$Force.IsPresent
. $cfg.Scripts.ExcelHelpers
# use $forceFlag from here on, NOT $Force
```

### Config overlay groups must track VerifyConfig.psd1

`ConfigOverlay.ps1`'s `Get-ConfigOverlayGroups` is a second, hand-maintained
index of `VerifyConfig.psd1`'s top-level sections (used by the `-Phase
InitConfig -Interactive` grouped field walker and by
`Get-ConfigOverlayReadmeText`). `New-ConfigOverlaySnapshot`/
`Update-ConfigOverlayData` read `VerifyConfig.psd1` generically (every
top-level key is captured and schema-repaired automatically), so a new
top-level section written into `.psd1` is silently correct at the JSON/repair
layer but invisible in the grouped editor and README until someone also adds
it to a NAMED group in `Get-ConfigOverlayGroups` -- it stays reachable only
via the catch-all `all` group, unlabeled. This actually happened:
`SnapVerify` (added v2.9.4, a major feature spanning six changelog entries)
had no named group until this was caught and fixed. **Whenever a phase gains
a new top-level `VerifyConfig.psd1` config section, add it to the most
relevant group in `Get-ConfigOverlayGroups` (and mention it under "Common
fields" in `Get-ConfigOverlayReadmeText`) in the same change.**
`Tests\Test-ConfigOverlay.ps1` has a schema-drift guard that fails the build
when a snapshot field is reachable only via `all` -- run
`Tests\Run-Tests.ps1` after any `VerifyConfig.psd1` structural change. That
same test file also repairs a reduced copy of the REAL `VerifyConfig.psd1`
defaults (not just hand-built fixtures) to confirm `-Phase InitConfig`
repair never drops an operator value and never throws against the actual
production config shape.

## Current state

The current version and per-release history live in `CHANGELOG.md` (newest
entry at the top); the phase list lives in `docs/Operations.md`. This file
does not record either, so it cannot go stale against them.

Pure (COM-free) libs are unit-tested via `Tests\Run-Tests.ps1`; COM/Edge phases
are validated by static analysis only (no PowerShell/Excel in the cloud build
env) and need a Windows + Excel 2019 run to confirm end to end.

To run the tests on Windows: `powershell -File Tests\Run-Tests.ps1` (parse-checks
every .ps1 + runs the unit tests). Encoding check: `powershell -File Check-Encoding.ps1`.

## Known issues / open points

- **Not yet run on Windows/Excel**: this refactor was authored in a Linux cloud
  env without PowerShell or Excel. Run `Tests\Run-Tests.ps1`, then smoke-test
  ReplaceGift/Gfix/Df and DfSnap on a copy before trusting them in production.
- **Align full branching** needs two domain facts: which FROM_sys/TO_sys literals
  mean "Host" (set `Align.HostSystemTypes` in VerifyConfig), and confirmation of
  the per-migration-type sheet sets in `AlignCompare.ps1`. Until then Align uses
  the send-result sheets (send[2], send[3]) and warns.
- **Align recv sheets are never synced** — recv sheets hold operator-captured evidence;
  only the host-team-managed send sheets are fetched from J4.
- **Align -Apply** syncs values + formats (Range.Copy) and is experimental.
- JenkinsSnap.ps1 matches the known-good repo logic (real Common.ps1 helpers); the
  earlier `Get-EdgeHwnd`/`Capture-Window` phantom-function risk is resolved.

## TODOs

Open work lives in [`docs/TODO.md`](docs/TODO.md).

## Cross-environment workflow

```
Office PC  →  Pack-LlmContext.ps1  →  clipboard
clipboard  →  paste to Claude/Cursor
Claude     →  XML patch or git diff
patch      →  Apply-LlmPatch.ps1   →  local files
today diff →  Export-DailyPatch.ps1 → clipboard → git push from home
```

`Apply-LlmPatch.ps1` accepts:
1. XML patch: `<patch><file name="..."><search>...</search><replace>...</replace></file></patch>`
2. git unified diff: standard `--- a/file` / `+++ b/file` / `@@ ... @@` format
3. Markdown fences around either format are stripped automatically.

## Session config (verify_session.json)

Machine/operator state ONLY -- see `docs/Configuration.md` for the full
layering rule (psd1 = shipped defaults, work-folder `verify_config.json` =
everything project-scoped, session = machine ephemera). Project-scoped
first-run prompt answers (e.g. `CheckSheet.Path`) persist to the work
folder's `verify_config.json` since v2.10.7, not here.

Remembered between runs:
- `WorkDir` — last work folder path
- `Owner` — mapping owner suffix (no personal default)
- `WindowWidth`, `WindowHeight`, `CropPx` — screenshot window size
- `CursorCell` — review cursor cell (default: A3)
- `CloneSourceDir` — external path for Clone phase
- `EvidenceDir` — evidence output folder (default: `<WorkDir>\evidence`)
- `DfExePath` — df.exe path; remembered after the first DfSnap run so the
  prompt fires only once (seeded from `Df.DefaultExePath`)
