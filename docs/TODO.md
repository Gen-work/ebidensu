# TODOs

Open work for VerifyTool, moved out of `CLAUDE.md` so the agent context
file stays small. Designed-then-shelved work lives in
[`Parked-Ideas.md`](Parked-Ideas.md) instead; release history lives in
[`../CHANGELOG.md`](../CHANGELOG.md).

- **ProcessTime: ja-OCR digit 9<->3 -- DETERMINISTIC FIX SHIPPED (v2.21.0),
  measurement still open.** The confusion itself is now handled without
  guessing, by `TimeDigitVerify.ps1`: (a) a digit that makes its field
  illegal is forced back when exactly one 3<->9 substitution can do it
  (`Repair-ImpossibleTimeDigit` -- '10:93:20' can only have been
  '10:33:20'); (b) start, end and the page's own printed processing-time
  column are treated as three readings of one fact, so a disagreement that
  exactly one substitution reconciles is repaired by arithmetic
  (`Resolve-ProcessTimeDurationConflict`) and anything ambiguous is left
  exactly as read and flagged; (c) every remaining ambiguous 3/9 -- one
  sitting where the opposite digit would also be legal -- is marked red in
  the output workbook by conditional formatting for a human glance. **No
  heuristic ever rewrites a plausible reading**; that was the v2.20.0-era
  bug (a misread end second turned a printed 00:00:01 into a derived
  00:00:07 silently). Still open: there is no way to MEASURE 3/9 accuracy,
  so preprocessing (`ConvertTo-ProcessTimeOcrImage`, v2.15.3) cannot be
  tuned -- see the benchmark item below.

- **ProcessTime: OCR benchmark harness** (Phase 2,
  `docs/ProcessTime-OcrBenchmark-Plan.md`) -- new office-PC-only
  `Export-OcrBenchmarkTruth.ps1` (reverse-export truth manifest from a
  confirmed `処理時間(*).xlsx`), `Build-OcrBenchmarkImages.ps1` (Edge render of
  a faithful HM template + crop reusing `ScreenRegion.ps1`), and
  `Test-OcrAccuracy.ps1` (`-Sweep`/`-Json`). The synthetic HTML template needs
  operator-supplied CSS + office-PC visual calibration (real captured snaps
  stay the primary ground truth; synthetic pages only cover 3/9 combinations
  absent from real data). Windows/Edge/OCR paths are static-checked only; pure
  `Compare-OcrDigits` / `Get-OcrBenchmarkScore` are CI-unit-tested. NOTE the
  plan's `Repair-ProcessTimeStartFromStamp` step is parked, not pending --
  it exists, was never wired, and is superseded (`docs/Parked-Ideas.md`).

- **NEXT: ReplaceGfix duplicate-candidate confirmation** — since v2.9.18,
  `GfixLogDownload` deliberately downloads *every* GoAnywhere job matching a
  needed IF_NO (not just one), because duplicate-IF_NO rows are common and
  content matching (`Find-GfixLogForCorrel`) is what actually decides which
  correl a log belongs to. This means `log\` can now legitimately hold more
  than one candidate log for a single correl (e.g. genuine retries of the
  same job) more often than before. **Partly addressed in v2.21.0**:
  `Find-GfixLogForCorrel` now decides run identity from the log's own
  `Command:` line (`Select-GfixLogCandidate`), so the plain and
  batch-stamped spellings of ONE download collapse into a single run
  instead of warning, and a batch-stamped `Correl_ID_S` selects its own run
  outright. What is left is the genuine case: two real reruns of the same
  job, where it still silently picks the newest and only prints a
  `[WARN] N different receive runs matched; chose newest (...)` — it never
  stops for operator confirmation. Planned next step: when `ReplaceGfix` (or `GfixLogDownload`'s
  finalize step) hits a multi-candidate `Warning`, show the operator each
  candidate's file name + parsed timestamp and require an explicit pick
  (Enter = accept newest, or choose another) before the log is pasted into
  the evidence workbook / before `GFIX_log`/`isReplaced` is marked done,
  instead of trusting "newest wins" silently. Needs: (1) deciding where the
  prompt belongs (`GfixLogDownload` finalize vs `ReplaceGfix`'s log op in
  `EvidenceExecutor.ps1` — the latter runs later and closer to when the log
  is actually inserted, so may be the more meaningful place to ask), (2) a
  non-interactive fallback (keep "newest wins" under `-NonInteractive`, same
  as the rest of this codebase's interactive/non-interactive split).

- **Mark: image-recognition placement for the red rectangle -- WIRING DONE**
  (v2.9.23), calibration still open. `Mark.Boxes` entries can now add a
  `Template` key (filename resolved against `Mark.TemplateDir`, then
  `mark_templates/`); when present, Mark.ps1 calls the existing
  `Locate-ByImage.ps1` (LockBits template match) against the original snap
  PNG (`<WorkDir>\snap\<folder>\<correl>.png`, the same file
  ReplaceEvidence pasted) instead of trusting a fixed offset, scales the hit
  from source-PNG pixels to the inserted picture's on-sheet point size, and
  falls back to the configured `OffsetX/OffsetY/Width/Height` box whenever
  there is no Template, the file is missing, or no match is found -- so this
  degrades gracefully and never blocks Mark. Per-box `Tolerance`/`PadX`/`PadY`
  overrides; console lines are tagged `[MARK-IMG]` (matched) vs `[MARK]`
  (fixed offset fallback) so a run makes it obvious which path was used.
  Still needs: real reference template PNGs per mark target (a small,
  visually distinctive crop of the target field -- see
  `mark_templates/README.txt` for the how-to) captured from real evidence,
  and an office-PC/Excel session to calibrate and confirm the pixel->point
  scaling -- no Windows/Excel in this dev environment, so `mark_templates/`
  ships empty and this stays fixed-offset-only until templates are added.

- **GiftJenkinsNoFile: callout bubble on the past-data mark** — SnapVerify M6
  (v2.9.11) already detects an unexpected *old* file in the no-GFIX-expected
  case (`Test-JenkinsFile -ExpectExists:$false`), draws the red box on the
  file's timestamp field, and stamps `過去分データー` (`ProjectLabels.NoGfixPastData`)
  into `SnapVerify.NoGfixNoteColumn` (default `AZ`). Requested follow-up: add
  an actual callout/comment-bubble shape next to the mark (not just the AZ
  column text) so the "this is old/past data" note is visible directly on the
  evidence picture itself. Needs a design decision on the shape to use (Excel
  `msoShapeCallout` via COM ~= `Shapes.AddCallout`, sized/positioned relative
  to the existing `verifyNote` AltText rect) plus an office-PC/Excel session
  to confirm placement -- no Windows/Excel in this dev environment.

- **Edge activation robustness DONE** (v2.9.18) — `Common.ps1`'s
  `Activate-EdgeWindow` (used by `Switch-ToEdge`, which `GfixLogDownload` /
  `MqSnap` / `HmSnap` all call after the operator presses Enter) used to
  activate Edge purely via `$Shell.AppActivate("Microsoft Edge")`, a
  title-substring match, and silently discarded its success/failure return
  value. `JenkinsSnap.ps1` had already independently fixed this exact
  flakiness for itself with a process-name-based lookup
  (`Get-EdgeMainWindowHandle` / `Activate-JenkinsEdgeWindow`, msedge.exe by
  process rather than window title), but that fix never made it into the
  shared `Common.ps1` helper the other phases use. Promoted the
  process-handle-first / title-match-fallback approach into
  `Common.ps1.Activate-EdgeWindow` (title match is now only a fallback, and a
  real `[WARN]` is printed when both paths fail instead of silently
  "activating" whatever window already happened to be foreground);
  `JenkinsSnap.ps1`'s duplicate local copy was removed in favor of the shared
  one. Static-checked only (no Windows/Edge in this dev environment) --
  confirm on an office PC that `Switch-ToEdge` reliably reaches GoAnywhere
  again.

- **SnapVerify M1–M5 done** — M1: `SnapVerify.ps1` pure library +
  `Tests/Test-SnapVerify.ps1` unit tests + `SnapVerify` config section in
  `VerifyConfig.psd1`. M2: `MqSnap.ps1` migrated to MappingStore/ProgressLog and
  wired to F2 (page-text poll, page-kind sentinel, MQ verdict ok=1/ng=2, batch
  `Expected_Time` prompt); two new pure helpers (`ConvertTo-ExpectedDateTime`,
  `Set-EmptyRunTimeCells`) are unit-tested. M3: `JenkinsSnap.ps1` wired
  to F3 (GiftRecv/GfixRecv NG=2 + summary, batch time prompt, sentinel,
  `Test-JenkinsSnapDone`); NoGfix stays pure-screenshot until M6. **M4 done** --
  `HmSnap.ps1` migrated to MappingStore/ProgressLog and wired to F1
  (page-text poll, page-kind sentinel, `Test-HmAbend` verdict ok=1/ng=2/ask with
  newest-wins in the time window, batch `Expected_Time` prompt, local
  `Test-HmSnapDone`); per-`TO_code` appl grouping preserved; VerifyTool dispatch
  passes SnapVerify+ExpectedTime config (mirrors MqSnap). **M5 done** -- F5 pixel
  localisation: pure `Get-MatchedRowIndex` / `Get-RowPixelRect` /
  `Get-JenkinsHighlightRect` / `New-SnapLocRect` / `Save-SnapLocSidecar` in
  `SnapVerify.ps1` (unit-tested) produce a `snap\<folder>\<correl>.loc.json` rect
  for the verdict's row; non-pure glue `SnapLocalize.ps1` (`Write-SnapLocalize`,
  System.Drawing + `Find-ActiveHighlightRow` scan) is dot-sourced by the three
  snap scripts and writes the sidecar after each verdict when
  `SnapVerify.Localize.Enabled` (default `$false`; HM/MQ geometry must be
  calibrated first, Jenkins uses the orange highlight). **M6 done** (v2.9.11) --
  NoGfix annotation: `GiftJenkinsNoFile` detects an unexpected file
  (`Test-JenkinsFile -ExpectExists:$false`), writes `<correl>.note.json` when
  `Localize.Enabled`, ReplaceEvidence stamps `verifyNote` AltText, MarkGift
  pixel->point scales + draws the red box and writes `過去分データー`
  (`ProjectLabels.NoGfixPastData`) to `SnapVerify.NoGfixNoteColumn` (default `AZ`).
  v2.9.12 field fixes: `TimeCheck` default-off, time-only run-time input, Edge
  refocus after prompts, NoGfix poll `-RequireTerm $false`, stale-note cleanup.
  M3/M4 copied MqSnap's `Test-MqSnapDone` pattern (done == exactly '1')
  so NG='2' rows stay pending -- `Get-PendingRows`/`Test-SnapDone` treat any
  non-'0' value as done and would hide NG rows. Design + open questions (only Q5,
  Rtncd/Rsncd semantics, is non-blocking) live in `docs/SnapVerify-Plan.md`. The
  M5 COM/GDI+ wiring is static-checked only; confirm on an office PC + calibrate
  `SnapVerify.Localize.*Row1Top/*RowHeight/*ColLeft/*ColWidth` before trusting it.

- **Generate-HostOpenMapping `-Add` + owner filter compose DONE** (v2.9.13) —
  explicit `-Add` selectors (`JOB_NAME` / `Correl_ID_M` / `Excel_NAME`) used to
  bypass the WBS owner-match scan, so jobs were added regardless of owner. They
  are now looked up in the WBS (col A) via the new `Build-WbsJobOwnerMap` and
  filtered through pure `Select-JobsByOwner` (`OwnerFilter.ps1`): a job whose WBS
  owner cell (col P) belongs to another operator is dropped (warned); a job
  absent from the WBS is kept as a temp/not-yet-listed job and reported. The
  WBS-range `-Add` path already owner-filtered (Step C) and is unchanged. Pure
  logic unit-tested in `Tests\Test-OwnerFilter.ps1`; the COM scan needs an
  office-PC run to confirm.

- **GfixLogDownload: auto-set GoAnywhere max rows to 100**
  Currently requires manual setup (default GoAnywhere list shows 20 rows — not enough for
  busy BIZ codes). Future: use SendKeys / UI automation to set the rows-per-page dropdown
  to 100 automatically after `Switch-ToEdge`, before the per-row search loop.

- **DfSnap: DfExePath configurable + first-run prompt DONE** — `Df.ExePath`
  (empty by default) holds a locked path that skips the prompt entirely; the new
  `Df.DefaultExePath` (default `C:\tools\DF\DF.exe`) is the suggestion the
  first-run prompt pre-fills (Enter accepts). Resolution is CLI `-DfExePath` >
  `verify_session.json` > `Df.ExePath` > prompt(`Df.DefaultExePath`). VerifyTool
  prompts once on the first DfSnap run, remembers the answer in
  `verify_session.json` (`DfExePath`), and passes both values to `DfSnap.ps1`
  (which keeps its own default-pre-filled prompt for standalone use). To lock a
  path and never be prompted, set `Df.ExePath` directly. COM/Excel parts are
  static-checked only; confirm the prompt + persistence on an office PC.

- **DfSnap region calibration** — default capture is `region` (x=120,y=280,w=1250,h=657
  for ~1980x1020). Tune `Df.RegionX/Y/Width/Height` and per-direction
  `Df.CropLeft/Top/Right/Bottom` (the window shadow is asymmetric). A pixel-color
  auto-detect of the window edge is a future option (no vision in a PS script).

- **GfixLogDownload max-rows** — still relies on manual "rows=100" setup.
  - **SS_CODE override DONE**: ReplaceGfix now reads an optional `SS_CODE` mapping
    column and threads it through the plan (`Build-GfixEvidencePlan -CorrelToSs`
    -> log op `SsCode` -> `Find-GfixLogForCorrel`). When the column is present and
    non-empty it wins; otherwise `GfixLog.ps1` infers SS from `Correl_ID_S`
    (5th char, or `J` for `J<biz>LxxS` jobs) exactly as before. Add an `SS_CODE`
    column to the mapping to take effect.
