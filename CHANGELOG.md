# Changelog

Tracks iterations across Misaki's browser (work) ↔ IDE (home) workflow.
Bump the date heading whenever a new bundle is delivered.


## 2026-06-18 - MqSnap focus regression fix (v2.9.7)

### Fixed
- **`MqSnap.ps1` no longer clicks the console/ISE window instead of the MQ page.**
  The v3 rewrite (v2.9.4) called `Switch-ToEdge` + `Click-PageBody` at the top of
  every per-row attempt. `Switch-ToEdge` does an `Alt+Tab`, which toggles relative
  to the *current* foreground window. After the previous row's screenshot Edge is
  already foreground, so the `Alt+Tab` flipped to the previously-used window (the
  PowerShell ISE / console), and `Click-PageBody` (which clicks
  `GetForegroundWindow()`) then clicked that window's border instead of the MQ
  page. Restored the known-good pattern used by the legacy MqSnap and the
  still-current `HmSnap.ps1`:
  - `Switch-ToEdge` runs **once before the loop** and **only inside the
    interactive branch** (right after `Bring-ShellToFront`, where the console is
    foreground so `Alt+Tab` correctly lands on Edge).
  - Per-row refocus is `Reset-FocusToBody` only (`Activate-EdgeWindow` ->
    AppActivate **by title** + `Click-PageBody`) -- no blind `Alt+Tab`, so focus
    never flips away from Edge between rows.
  Detection / screenshot behavior is unchanged.


## 2026-06-18 - SendVsGift OCR-dropout tolerance + clean-read preference (v2.9.6)

### Changed
- **`Compare-SendRecordCheck` gains an OCR-dropout tier** (`SendMetadata.ps1`).
  The JP recognizer drops runs of characters from long ASCII record strings
  (field-observed: ~12 digits lost off the first record), which symmetric
  edit-distance scored as a hard `mismatch`. After the existing exact /
  prefix-similarity / compact checks fail, a new tier asks whether the shorter
  side is materially shorter (length ratio `<= 0.85`) AND almost entirely an
  in-order subsequence of the other (`LCS / shorter >= 0.9`): if so the chars
  were dropped, not changed, so it scores `fuzzy` instead of `mismatch`. A
  survivor too short to judge (`< 6` chars) scores `unknown`, never a false
  `mismatch`. Genuine comparable-length conflicts still score `mismatch`.
  Thresholds are function params (defaults `0.85 / 6 / 0.9`).
- **`Find-SendRecordByRowNumber` prefers the fullest read** of a row. Each
  image is OCR'd with both `ja` and `en-US` and the lines are merged, so one
  row label can carry a garbled/short `ja` record and a clean `en-US` one. It
  now returns the LONGEST record after the label (the most complete read)
  instead of the first, so a clean en-US read is no longer shadowed by a
  dropped ja read.
- **Verdict rule made explicit: a matching row count is authoritative.** When
  the max row label matches the gift's `MaxRowNumber` the verdict stays `ok`
  even if a record disagrees -- record text is too OCR-noisy to auto-flag NG.
  The disagreement is still surfaced in the per-field `Checks` for the
  operator. The `Test-SendMetadata.ps1` "first record disagrees" case now
  expects `ok` (was `ng`) to match this rule and asserts the check still
  reports `mismatch`.

### Added
- **`Get-SendLcsLength`** pure helper (longest-common-subsequence length) in
  `SendMetadata.ps1`, backing the OCR-dropout tier. Unit-tested.
- New `Tests/Test-SendMetadata.ps1` cases: LCS length; clean-read-wins record
  selection; dropout -> fuzzy; comparable conflict -> mismatch; heavy dropout
  -> unknown.

### Note
- Pure-function changes only; validated by static analysis (no PowerShell /
  Excel in the cloud build env). Run `Tests\Run-Tests.ps1` on Windows to
  confirm green. The OCR-dropout thresholds can be lifted into the `SendVsGift`
  config block in a follow-up if on-site tuning is wanted.


## 2026-06-17 - SnapVerify M3: JenkinsSnap instant NG detection (v2.9.5)

### Added
- **`JenkinsSnap.ps1` is wired to SnapVerify F3** (plan `docs/SnapVerify-Plan.md` M3),
  for the `GiftRecv` / `GfixRecv` modes. After the Ctrl+F search and screenshot it
  now polls the page text (A2), classifies the page with `Get-SnapPageKind`
  (sentinel A3), archives the Ctrl+A text as `snap\<folder>\<correl>.txt` (A1),
  then runs `ConvertFrom-JenkinsListText` + `Test-JenkinsFile` to decide:
  - `<field>_snap = 1` when the file is in the list and within the
    `Expected_Time` +- tolerance window (or no time check).
  - `<field>_snap = 2` (NG) when the file is missing from the list or its
    timestamp is outside the window. NG stays **pending** (re-offered next run)
    and is listed in an end-of-run NG summary.
  The same polled page text feeds the existing receive-file download, so the
  page is read once per row.
- **Batch run-time prompt** (`Resolve-SnapRunTime`) and **page-kind sentinel**
  (`r=retry / s=skip / q=quit`, max 3 retries) mirror the M2 MqSnap wiring.

### Changed
- **`JenkinsSnap.ps1` migrated off `Get-PendingRows`** to a local "done == exactly
  '1'" rule (`Test-JenkinsSnapDone`) so NG='2' rows stay pending and are not
  hidden. `Ensure-MappingColumns` now also defaults the `Expected_Time` column.
  `VerifyTool.ps1` threads the `SnapVerify` + `ExpectedTime` config into the
  GiftJenkins / GfixJenkins / GiftJenkinsNoFile dispatch.
- **NoGfix mode keeps the pure-screenshot path** (`<field>_snap = 1`) even when
  SnapVerify is enabled; its F4 detection lands in M6. `SnapVerify.Enabled=$false`
  reverts all Jenkins modes to pure screenshot.
- The pure F3 library functions (`ConvertFrom-JenkinsListText`, `Test-JenkinsFile`)
  and their unit tests already shipped in M1; M3 is wiring only, no library change.


## 2026-06-17 - SnapVerify M2: MqSnap instant NG detection + MappingStore migration (v2.9.4)

### Added
- **`MqSnap.ps1` is wired to SnapVerify F2** (plan `docs/SnapVerify-Plan.md` M2).
  After the inquiry search it now polls the page text (A2), classifies the page
  with `Get-SnapPageKind` (sentinel A3), archives the Ctrl+A text as
  `snap\GIFT_MQ\<correl>.txt` (A1), screenshots as before, then runs
  `ConvertFrom-MqPageText` + `Test-MqRecord` to decide the verdict:
  - `GIFT_MQ_snap = 1` when the record is found, in the time window, and
    Rtncd/Rsncd are zero.
  - `GIFT_MQ_snap = 2` (NG) on "No Data!", no matching Correl_ID, RecvDate outside
    the window, or non-zero Rtncd/Rsncd. NG stays **pending** (re-offered next run)
    and is listed in an end-of-run NG summary.
- **Batch run-time prompt** (`Resolve-SnapRunTime`): one question at start
  (`[Enter]`=now / `yyyy/MM/dd HH:mm:ss` / `n`=no time, plus tolerance). Empty
  `Expected_Time` cells on the pending rows are filled and persisted; existing
  values are kept (plan 2.2).
- **Page-kind sentinel**: an off-page text (OuterFrame / Empty / Unknown) stops
  and asks the operator `r=retry / s=skip / q=quit` (max 3 retries), logging a
  `warn` event with the page kind + a 200-char preview.
- **Two pure helpers in `SnapVerify.ps1`**, unit-tested in `Tests/Test-SnapVerify.ps1`:
  - `ConvertTo-ExpectedDateTime` — per-row `Expected_Time` cell -> `[datetime]`
    or `$null` (empty/unparseable = no time window; never throws).
  - `Set-EmptyRunTimeCells` — batch-fill empty time cells, keep existing values,
    return the count filled.

### Changed
- **`MqSnap.ps1` migrated off bare `Import-Csv`/`Export-Csv`** to MappingStore
  (`Import-Mapping` / `Ensure-MappingColumns` / `Export-MappingAtomic`, atomic
  writes that never clobber non-target rows) and ProgressLog (`status\progress.jsonl`
  events). Source is now ASCII-only per the encoding policy. `VerifyTool.ps1`
  threads the `SnapVerify` + `ExpectedTime` config into the GiftMqSnap dispatch.
- Pending filter uses a local "done == exactly '1'" rule (`Test-MqSnapDone`), not
  `Get-PendingRows`, so NG='2' rows are not hidden. `SnapVerify.Enabled=$false`
  reverts MqSnap to pure screenshot (legacy behavior).

### Docs / cleanup
- Removed the stale `TODO-2026-05-29.md` (its actionable items -- JenkinsSnap
  phantom-functions, JIGPMB1S log move, DfSnap df.exe path, GoAnywhere 100-rows --
  are resolved or already captured in CLAUDE.md TODOs) and the leftover
  `Align.ps1.bak.20260601_153221` backup. Refreshed CLAUDE.md "Current state"
  (was still on v2.8.1) and the SnapVerify milestone status.

### Not yet wired
M3 (JenkinsSnap NG=2) / M4 (HmSnap) / M5 (pixel localisation) / M6 (NoGfix) pending.
M3/M4 should copy MqSnap's `Test-MqSnapDone` (done == '1') so NG='2' stays pending.

---

## 2026-06-16 - ReplaceGfix: thread an optional SS_CODE mapping column (v2.9.3)

### Added
- **SS_CODE override is now wired end to end** (closes the GfixLog SS_CODE TODO).
  `GfixLog.ps1` always supported an SS override, but the plan never carried it, so
  every GFIX log match fell back to inferring SS from `Correl_ID_S`. Now:
  - `New-LogOp` carries an `SsCode` key; `Build-GfixEvidencePlan` gains a
    `-CorrelToSs` map and a `Resolve-GfixSsForCorrel` helper that fills each log
    op's `SsCode` (empty when the correl is not in the map).
  - `ReplaceEvidence.ps1` builds `$correlToSs` from an optional `SS_CODE` mapping
    column and passes it in; `EvidenceExecutor` already forwards `op.SsCode` to
    `Find-GfixLogForCorrel`.
  - Behavior is unchanged unless the mapping actually has an `SS_CODE` column with
    a value, so this is a safe, forward-compatible addition. New unit assertions
    in `Tests/Test-EvidencePlan.ps1` cover the override and the empty-default case.

---

## 2026-06-16 - ReviewGift/Gfix/Df: open the mode sheet, review per workbook (v2.9.2)

### Fixed
- **`ReviewEvidence.ps1` now matches its documented behavior** (VerifyTool help:
  "ReviewGift/Gfix/Df open the matching sheet up front ... per-workbook prompt").
  The implementation had drifted: it always activated the *send-data* sheet
  (`送信データ`) and searched column A for each correl id, then prompted **per
  id**. For ReviewGift the ids live on `GIFT 受信結果`, not `送信データ`, so every
  id logged `[WARN] ID not found in 送信データ column A` and the operator was
  marched through the workbook one id at a time.
  - Each workbook now brings the mode's own evidence sheet to the front via the
    already-present (but previously dead) `Open-SheetForReview` + `$openSheetName`
    switch: ReviewGift -> `GIFT 受信結果`, ReviewGfix -> `GFIX 受信結果`,
    ReviewDf -> the DF compare sheet; ReviewEvidence (bit 7) leaves the default.
  - Review is now **per workbook**: one Enter marks the review bit for every id
    in the Excel_NAME group, then saves + closes (the whole mode sheet is
    reviewed in one pass, since all correls are stacked on it). `s` skips the
    whole workbook (left pending, no save); `q` quits; `-m "comment"` records a
    per-group note. Removed the per-id `Move-ToSendDataId` navigation and the
    per-id inner loop.

---

## 2026-06-16 - Replace: fix NoGfix image overlap past row 2000 (v2.9.1)

### Fixed
- **`ExcelHelpers.ps1` anchor-row math no longer caps at row 2000.**
  `Get-RowAtOrBelow` / `Get-NextAnchorRow` / `Get-PictureBottomRow` carried a
  hard-coded `maxScanRows = 2000` ceiling. On a GIFT evidence sheet the NoGfix
  block is the trailing (4th) section — excel.png, then HM/MQ per correl, then
  Jenkins per correl, then NoGfix per correl — so once a workbook held ~10+
  correl ids the running anchor crossed row 2000. Past that the
  `startScan = floor(shape.Top / 15)` approximation also exceeded the cap, so
  `Get-RowAtOrBelow`'s `while ($r -le $maxScanRows)` never ran and returned the
  2000 cap for every lookup. Result: all further NoGfix pictures stacked on the
  same rows (overlapping images) and every correl-id `text` op overwrote the
  same capped cell (ids "not entered" — only the last one survived).
  - New `Get-MaxSheetRow` returns the worksheet's real row limit (1,048,576 on
    .xlsx/.xlsm; 65,536 on legacy .xls). The three helpers now default
    `maxScanRows = 0`, meaning "use that limit". The `shape.Top / 15` start-row
    approximation keeps the scan only a handful of rows long, so dropping the
    fixed cap costs nothing in speed.
  - Added a `startScan > ceiling` guard so an over-approximated start row
    returns the ceiling instead of silently skipping the scan.

---

## 2026-06-12 - SnapVerify M1: pure detection library (v2.9.0)

### Added
- **`SnapVerify.ps1`** — pure dot-source library (no COM, no SendKeys) providing
  all shared base logic for snap-phase instant NG detection (M1 per plan):
  - `ConvertFrom-HmPageText` / `Test-HmAbend` — HM page parse + abend verdict
    (window-based, newest-wins retry logic, ok/ng/warn/ask, per spec F1 / 2.3)
  - `ConvertFrom-MqPageText` / `Test-MqRecord` — MQ parse (absorbs Parse-GiftMq.ps1
    regex) + 3-condition verdict: no-row / time window / non-zero Rtncd (spec F2 / 2.4)
  - `ConvertFrom-JenkinsListText` / `Test-JenkinsFile` — Jenkins file-list parse
    (absorbs Parse-JenkinsList.ps1) + file exists/absent verdict + NoGfix mode (F3/F4)
  - `Get-SnapPageKind` — page-type sentinel: HmResult / MqResult / MqNoData /
    JenkinsResult / OuterFrame / Empty / Unknown (spec A3 / 3.6)
  - `Resolve-SnapRunTime` — pure logic for batch time inquiry (spec 2.2):
    '' = Now, n = no-time, explicit datetime, tolerance override
- **`Tests/Test-SnapVerify.ps1`** — unit tests covering all functions above using
  real-world fixtures from Appendix A (HM) and Appendix B (MQ / outer-frame).
- **`SnapVerify`** config section in `VerifyConfig.psd1` (Enabled, ToleranceMinutes,
  SaveText, PollTimeoutSec, PollIntervalMs, NoGfixNoteColumn) + `SnapVerify` entry
  in Scripts table.
- `SnapVerify.ps1` added to dot-source whitelist in CLAUDE.md.

### Not yet wired
M2 (MqSnap) / M3 (JenkinsSnap) / M4 (HmSnap) integration pending.
`SnapVerify.Enabled=$false` to revert to pure screenshot mode.

## 2026-06-11 - SendVsGift OCR pipeline field fixes (v2.8.1)

First end-to-end field run of Stage-2 OCR surfaced a chain of bugs; all
fixed (PRs #42-#45 + follow-up branch `claude/affectionate-wozniak-vq4j71`):

### Fixed
- **PS 5.1 array-nesting bug** across the OCR stack: library functions
  returned `,@($arr)` while callers wrapped the call in `@(...)`, which
  NESTS. Symptoms fixed: empty-string `LiteralPath` binding error, section
  filter dying on `Object[]`->`Double`, every image counting as ONE OCR
  "line" (`images=8 lines=8`), multi-image OcrTool joining all paths into
  one. Convention now: these lib functions return plain arrays; callers
  keep `@(...)`.
- **Section export with Ctrl+G groups**: section membership now uses the
  top-level shape's Top (children can report group-relative Top);
  GroupItems failures warn instead of dropping a strip.
- **Export resolution**: temp chart is created at `Scale` x shape size
  (default 3.0) and the pasted picture stretched (Chart.Shapes with
  Chart.Pictures fallback) so Excel re-renders from the original media;
  GDI+ bicubic min-width upscale as belt and braces; `[DIAG]` prints the
  first export's pixel size.
- **Legacy parse errors**: raw-Japanese comments (UTF-8 no BOM -> CP932
  mojibake) replaced with ASCII in `Read-ClipboardJson.ps1`,
  `Resolve-ExpectedTime.ps1`, `Probe-Shapes.ps1`.
- **Mirror CI**: feature-branch pushes no longer ask GitLab to delete all
  other branches (push `refs/remotes/origin/*` instead of local heads).

### Added
- `OcrTool.ps1 -Diag`: per-image sweep over every installed recognizer
  language (+ user-profile engine) with line/word counts, sample line,
  pixel size and `MaxImageDimension` check.
- `Probe-Shapes.ps1` recurses into groups (indented children).
- SendVsGift warns (with the `-Diag` command) when OCR reads nothing;
  OCR failures print the script stack trace.

### Open TODO
- **OCR still returns zero lines on the upscaled evidence PNGs** even
  though export/resolution are confirmed good; see
  `docs/SendVsGift.md` -> "Troubleshooting: OCR reads nothing" for the
  next-session checklist (engine sanity test, `-Diag` sweep, dark-
  background preprocessing, en-US engine).

## 2026-06-11 - SendVsGift OCR auto-compare + standalone OcrTool (v2.8)

### Added
- **`OcrTool.ps1`** (new, standalone): command-line OCR tool over the existing
  dot-source libs so OCR is reusable outside SendVsGift. Accepts image files /
  folders / wildcards, or `-Workbook <xlsx>` (+ optional `-Sheet`) to export
  the embedded pictures first; `-Json` / `-OutFile` / `-OutDir` outputs;
  `-ListLanguages` probes the engine. Has `param()` -> call via `&`, never
  dot-source.
- **SendVsGift review-flow rework** (operator items 2-4):
  - pending rows are grouped per evidence workbook; each workbook opens ONCE
    and the cursor just moves between correls (no close/reopen inside a group);
  - per correl the cursor jumps to the `Correl_ID_S` label cell in column A of
    the send-data sheet (`Application.Goto`, scrolled into view);
  - after every console answer Excel is brought back to the foreground
    (SetForegroundWindow on the Excel hwnd; SendKeys Alt+Tab fallback);
  - new `n` answer marks `SendVsGift=2` (NG) from the manual prompt.
- **OCR verdict rules** (pure, unit-tested in `SendMetadata.ps1`):
  - 0-byte gift file: `used CYLINDERS : 0` dataset-info screen, or
    begin-of-data + end-of-data markers on the SAME image with no `000001`
    line (custom `ZeroBytePattern` still overrides);
  - data gift file: the zero-padded max row number (000003 / 004644 style)
    must appear, and first/last records (after their row labels) must match
    by exact first token or >= 80% prefix similarity (Levenshtein, first 20
    chars) to absorb OCR noise;
  - `Compare-SendGiftEvidence` -> verdict ok / ng / unknown.
- **OCR auto-mark** (`SendVsGift.AutoMark`, default on): ok -> `SendVsGift=1`,
  ng -> `SendVsGift=2` + red end-of-run NG summary, unknown -> manual prompt.
  `2` stays pending. `-NoAutoMark` keeps the verdict advisory.
- **Per-correl picture sections**: only the pictures between a correl's
  column-A label and the next label are exported/OCR'd, so multi-correl send
  sheets no longer cross-contaminate.
- **Ctrl+G group support** in `EvidenceImageExport.ps1`: grouped pictures were
  previously skipped entirely (type filter); groups are now flattened and each
  child picture is exported on its own (also dodges the OCR engine max image
  dimension on wide grouped strips).

### Fixed
- `VerifyTool.ps1` SendVsGift phase now has an OCR option: `o`/`ocr` toggle at
  the option prompt (the `-ocr` "unknown option" dead end), plus a `-Ocr`
  CLI switch; config `SendVsGift.Ocr` is the persistent default.
- `SendVsGift.ps1` standalone launch no longer dies on `mapping_.csv`:
  WorkDir/Owner fall back to `verify_session.json`, then to the single
  `mapping_*.csv` in the work folder, then to a prompt.

### Notes
- Authored in a Linux cloud env: parse/unit tests are cloud-runnable, but the
  Excel COM section export, Find/Goto cursor jumps, foreground switching and
  the WinRT OCR call need a Windows + Excel 2019 smoke test.
- Copying a child picture out of a Ctrl+G group uses `Shape.Copy` on the
  group item; if a given Excel build refuses that, ungroup once or report it
  (a whole-group export fallback can be added).

## 2026-06-11 - Full-width filename fallback

### Added
- **Reusable full-width filename resolver** in `WorkbookResolver.ps1`:
  `FullWidthFilenameResolver` normalizes full-width ASCII to half-width,
  detects candidate filenames, and enumerates matching files by filter.
- **Generic fallback helpers** (`Resolve-FullWidthFileName`,
  `Find-FullWidthFilenameCandidates`) so any not-found filename path can warn
  and optionally prompt before using a full-width candidate.
- **Workbook integration**: `Find-WorkbookByExcelName` still prefers exact and
  wildcard evidence/J4 matches, then uses the full-width fallback with
  `Prompt` / `Accept` / `Reject` policy.
- **Unit coverage** in `Tests/Test-WorkbookResolverFullWidth.ps1` for workbook
  and generic filename fallback behavior.

### Notes
- Interactive phases should keep `Prompt` so the operator approves the
  full-width filename. Non-interactive tests/batch flows should pass `Accept` or
  `Reject` explicitly.


## 2026-06-11 - SendVsGift Stage 2 OCR skeleton

### Added
- **Windows built-in OCR wrapper** `OcrWindows.ps1` (dot-source lib): calls the
  `Windows.Media.Ocr` WinRT API from PowerShell 5.1 -- the same engine family
  as the Snipping Tool text extraction, zero installs. Returns plain
  line/word objects with pixel bounding boxes; never throws at dot-source
  time on non-Windows hosts.
- **Pure send-metadata lib** `SendMetadata.ps1` (unit-tested,
  `Tests\Test-SendMetadata.ps1`): rebuilds the spacing the Japanese
  recognizer drops from word boxes, detects the 0-byte pattern, guesses row
  counts, builds `send_metadata.csv` records parallel to `gift_metadata.csv`,
  and compares the two sides (match / mismatch / unknown per field; absence
  of OCR evidence is never a mismatch).
- **Evidence picture export** `EvidenceImageExport.ps1` (dot-source lib):
  exports embedded send-sheet screenshots to PNG via a temp ChartObject,
  top-to-bottom, skipping `verifyMark_*` shapes.
- **SendVsGift `-Ocr` flow**: with `-Ocr` (CLI) or `SendVsGift.Ocr = $true`
  (config / work-folder overlay), each pending workbook gets its send-sheet
  pictures exported to `data\send_images\<Correl_ID_S>\`, OCR'd, recorded in
  `data\send_metadata.csv`, and a per-field send-vs-gift verdict is printed
  before the unchanged manual Enter-to-mark prompt. OCR failure or absence
  falls back to the manual flow; mapping semantics untouched.
- New `SendVsGift` config block (`Ocr`, `OcrLanguage`, `SendSheetName`,
  `ZeroBytePattern`) in `VerifyConfig.psd1`, reachable from the JSON overlay.

### Notes
- Authored in a Linux cloud env: parse/unit tests are cloud-runnable, but the
  Excel COM export and the WinRT OCR call need a Windows + Excel 2019 smoke
  test (`Tests\Run-Tests.ps1`, then `SendVsGift.ps1 -Ocr` on a copy).
- 0-byte pattern and row-number parsing are first-guess heuristics; tune with
  representative SEND screenshots (see docs/SendVsGift.md TODOs).


## 2026-06-09 - Work-folder config precedence + workbook prefix

### Changed
- **Project-level workbook prefix.** `Workbook.ExcelPrefix` now configures the
  fixed prefix before `_<Excel_NAME>` for J4 evidence files. New mappings no
  longer generate an `Excel_Prefix` column; existing mapping rows that already
  carry `Excel_Prefix` still override the project prefix for compatibility or
  rare per-workbook exceptions.
- **Cleaner `InitConfig`.** `InitConfig` now writes a separate
  `verify_config.README.txt` field guide next to `verify_config.json`, keeping
  the JSON valid and compact while still explaining what each field affects.
  Runtime overlay loading strips metadata keys such as `_README` / `_comment`.
- **Config precedence tightened.** `Window.NoResize`, `Align.J4BaseDir`,
  `CheckSheet.Path`, and `Workbook.ExcelPrefix` now follow the intended order:
  CLI args > work-folder JSON > defaults/session fallback.

## 2026-06-09 - Per-work-folder JSON config overlay

### Added
- **Per-work-folder config overlay (`verify_config.json`).** Each work folder
  may now carry a JSON file that is deep-merged over `VerifyConfig.psd1` at
  startup, so every case can fully customize its settings without touching the
  shared `.psd1`. Precedence is **CLI args > work-folder JSON > .psd1 defaults**.
  Because almost every phase already gets its values from the merged `$Config`
  (via `Invoke-ToolPhase`), the overlay reaches them all: owner, window size +
  crop, mark boxes (`Mark.Boxes`), mail subject/body (`Mail`), reviewer, check
  sheet, Df region, GFIX-log highlight, Replace labels, Align J4 dir,
  DeliverFiles targets, timing, and more.
  - File name is configurable via `Paths.OverlayName` (default
    `verify_config.json`); it lives in the WorkDir.
  - Japanese in the overlay (mail templates) is plain UTF-8 - no BOM and no
    `[char]` gymnastics, cleaner than the BOM'd `.psd1`.
  - The startup banner shows `Config overlay : ...` when one is loaded; a
    bad/unparseable overlay only warns, it never blocks startup.
- **`InitConfig` phase** (aliases `Config` / `MakeConfig` / `EditConfig`).
  Writes a starter `verify_config.json` into the WorkDir, pre-filled from the
  current effective config so the operator edits real values, not a blank
  template. `-Force` regenerates and keeps a `.bak`. Structural keys
  (`Scripts` / `PhaseOrder` / `Aliases`) are left out, but any `.psd1` key can
  be added by hand.
- **`ConfigOverlay.ps1`** - new pure, dot-sourced lib (no `param()`, ASCII, no
  BOM) with the deep-merge + JSON<->hashtable + generator helpers. Unit-tested
  by `Tests\Test-ConfigOverlay.ps1` (covers the PS 5.1 empty-array-as-"" and
  `\uXXXX`-escape quirks so Mark boxes survive a write/read round-trip).
- **Centralized `ExpectedTime` settings** in `VerifyConfig.psd1` (`TimeColumn` /
  `IdColumn` / `LookbackHours` / `TimeFormat`). `Resolve-ExpectedTime.ps1` now
  takes a `-TimeFormat` param instead of hard-coding the format. The time
  VALUES stay per-row in the mapping CSV (per-correl, not global) - that is the
  one deliberately non-JSON 'expected time' piece.
- **`Clone.SourceDir`** key so the Clone source folder can be set per work
  folder too (CLI > overlay > session).

### Notes
- Run `Tests\Run-Tests.ps1` on Windows to parse-check every `.ps1` and run
  the new overlay unit tests before trusting this on a live case (the cloud
  build env has no PowerShell).

## 2026-06-02 - Incremental mapping add + Excel_Prefix auto-capture

### Added
- **Incremental mapping add (`-Add`).** `Generate-HostOpenMapping.ps1 -Add`
  merges freshly-selected rows INTO an existing `mapping_<Owner>.csv` instead
  of overwriting it. Existing rows — and ALL their progress (snaps, isReplaced,
  isMarked, isReviewed, Excel_Prefix, delivery flags, comments) — are kept
  verbatim; only `Correl_ID_M` values not already present are appended. Use it
  to grow the map day by day:
  `-Add -JobNames CJODJDEU,CJODJDB5` / `-Add -CorrelIdsM JIDSC09M` /
  `-Add -ExcelNames LJRVWD64` / `-Add -WbsStartRow 2300 -WbsEndRow 2400`.
  Reachable from the menu (`add` toggle) and the `Mapping` phase wiring.
- **`-ExcelNames` selector** for Generate/VerifyTool — accepts Excel_NAME(s)
  and reverse-maps each to its JOB_NAME (index-5 `W`→`J`) to look the row up
  in GFIX. Menu key `ex`.
- **`Get-PrefixFromFilename`** in `WorkbookResolver.ps1` — inverse of
  `Get-ExcelFullStem`; recovers the J4 prefix that precedes `_<Excel_NAME>`.

### Fixed
- **JenkinsSnap keeps window ops on Edge (folded in from codex's fix).** After
  `Read-Host`, some terminals keep the console as the foreground window, so the
  old code resized/screenshotted the **CLI** instead of Edge. JenkinsSnap now
  resolves the real `msedge` MainWindowHandle (`Activate-JenkinsEdgeWindow`),
  drops the `Switch-ToEdge` Alt+Tab dance, and screenshots via the activated
  `$edgeHwnd` rather than `GetForegroundWindow()`. Returns a zero handle (and
  fails the row) instead of operating on the wrong window. New `.ps1` files
  normalized to no-BOM per encoding policy.

### Changed
- **Clone now captures `Excel_Prefix`.** When a workbook is cloned from
  `-SourceDir` (real filename `<prefix>_<Excel_NAME>.xlsx`), Clone extracts the
  prefix and writes it back into the mapping (per `Excel_NAME`, atomic write).
  This is the missing "input point" for `Excel_Prefix` — downstream phases
  (CheckSheet / DeliverMail / DeliverFiles) now get the exact J4 filename
  without any manual entry. Existing rows/progress are untouched.

## 2026-06-01 - Delivery: review check sheet fill + review-request mail

The final hand-off step. Two new phases close the workflow after review.

### Added
- **DeliverMail phase.** One Outlook **draft** per evidence Excel (grouped by
  `Excel_NAME`), built via Outlook COM `CreateItem` + `Display` — never
  auto-sent. The operator eyeballs each draft, clicks Send by hand, then presses Enter
  in the shell to set the new `isDelivered` mapping flag (`1` = sent). `s` skips,
  `q` quits, and `-m "comment"` records a note in the new `DeliverComment` column
  (per `Excel_NAME`, like `ReviewComment`). Subject =
  `【GIFT廃止対応】<Phase>レビュー依頼(<Excel_NAME>)`; body + reviewer + UNC paths
  are all config-driven (`Mail` / `Reviewer` in VerifyConfig.psd1). Outlook is
  released but never Quit (it may be the operator's live session).
- **CheckSheet phase.** Appends one row per Excel to the shared review check
  sheet (sheet `Check Sheet_J4`): A No. (continued, only if blank), B 記入日
  (today, format copied from the row above), C `JAVA`, E `J4内部ﾚﾋﾞｭｰ`,
  F full evidence filename, G owner, H reviewer. Because it is a public
  document the edit is **double-checked**: a TEMP copy is filled and opened for
  visual review; on Enter the original is re-stat'd and, only if it is unchanged
  since the preview began, the identical edits are committed — otherwise the
  write is **held** so the operator can re-check. Already-listed Excels are
  skipped unless `-Force`. Path comes from `CheckSheet.Path`; if missing the
  phase prompts (and remembers it in `verify_session.json`).

### Changed
- **MappingStore** now defaults two new columns: `isDelivered` (`0`) and
  `DeliverComment` (`''`). `isDelivered` is a `PhaseOrder` field, so it is
  auto-added to existing mappings on startup and shown in Status.

## 2026-06-01 - Review workflow: header-fill fix, MarkGfix merge, review comments, mode sheet

Host->Open focus. Five related improvements:

### Fixed
- **Replace no longer drops the N1:U1 header fill.** `Reset-SheetBelowRow`
  (ExcelHelpers.ps1) used `Range.Clear()` on `A<start>:T<last>`, which wipes the
  *entire* merged cell when a colored header banner merged above the start row
  dips into the clear range. It now snapshots the fill+value of any such
  above-anchored merge on the boundary row, clears, then restores (re-merging
  defensively). Shared by all three Replace modes, so the fix covers
  ReplaceGift/Gfix/Df uniformly.

### Changed
- **GFIX-log highlight folded into MarkGfix.** The yellow "Command:" highlight is
  now applied in the same pass that draws GFIX red rectangles (one workbook open,
  tracked by `isMarked` bit 2). The standalone `isGfixLogMarked` column and the
  `MarkGfixLog` phase entry are removed. Core logic moved to the shared
  `Invoke-GfixLogHighlight` in ExcelHelpers.ps1; `MarkGfixLog.ps1` stays as a
  by-name re-highlight utility (no mapping column).
- **Review opens the matching sheet.** ReviewGift/Gfix/Df now bring
  GIFT/GFIX jushin kekka (or the DF compare sheet) to the front on open.

### Added
- **Review comments.** At the per-workbook Review prompt, append `-m "comment"`
  (works with Enter / `s` / `q`) to record a note in the new `ReviewComment`
  column (per Excel_NAME group). Prior notes are shown when the workbook opens.
  New read-only **Comments** phase lists every recorded note.

## 2026-05-31 - Cleanup pass: paste residue, SnapConfig BOM, encoding policy, JenkinsSnap DryRun

Follow-up to the review of the shared-lib refactor. The big blockers
(GfixLogDownload mapping-truncation + failure recovery, DfSnap region capture,
plan-driven Replace) were already fixed in the prior PR; this pass closes the
gaps that survived.

### Fixed
- **Pack-LlmContext paste residue** removed from 6 files whose tails carried a
  truncated `` ` `` / ``` ``n ``` / `--- File: X ---` separator that breaks the
  PowerShell parser: `SnapConfig.psd1`, `Run-Snap.ps1`, `MqSnap.ps1`,
  `HmSnap.ps1`, `Crop-Snap.ps1`, `ExcelSnap.ps1`.
- **`SnapConfig.psd1`** holds raw Japanese labels but had NO BOM, so
  `Import-PowerShellDataFile` would mojibake them -> added the UTF-8 BOM
  (per the .psd1 rule in CLAUDE.md).
- **`Fix-Encoding.ps1`** rewritten: it used to add a BOM to *every* file,
  which violated the ".ps1 = no BOM" policy. It now enforces the documented
  policy (.ps1 -> strip BOM, warn on non-ASCII; .psd1 -> BOM only when it holds
  raw text; .json/.jsonl -> no BOM) and uses robust extension filtering instead
  of the fragile `-Include` (no `-Recurse`).
- **`JenkinsSnap.ps1 -DryRun`** no longer opens Edge or prompts for navigation;
  it now reports the per-row capture plan and skips all UI.

### Added
- **`Check-Encoding.ps1`** now also fails on Pack-LlmContext paste residue
  (`--- File:` separators / stray ``` ``n ``` markers) in any business script,
  so a truncated paste is caught in CI instead of at runtime.

## 2026-05-29 - Shared MappingStore + plan-driven Replace + recovery/monitoring

Large refactor toward stability, recoverability, encoding safety, and a
review-standard evidence layout. New pure (COM-free) libraries are unit-tested;
the rest validated by static analysis (no PowerShell/Excel in the build env).

### Added - shared libraries (dot-source, ASCII source, no BOM)
- `MappingStore.ps1` - single source of truth for mapping_<Owner>.csv:
  `Import-Mapping`, `Export-MappingAtomic` (temp file + Move with retry +
  backoff; never silent on a locked CSV), `Ensure-MappingColumns`,
  `ConvertTo-TargetIdList` (array / comma / trim), `Test-TargetRow`,
  `Get-PendingRows` (snap + bitmask), `Set-MappingBit`.
- `GfixLog.ps1` - pure GFIX receive-log matcher. `SS_CODE = Substring(4,1)`;
  fragment `/appl/<TO>/<TO>Ver1/gfix/recv/<id> <SS>`; 0 match = fail; many =
  newest by timestamp + warning; returns the whole file's lines.
- `EvidencePlan.ps1` - pure correl-major plan builders (`Build-Gift/Gfix/Df
  EvidencePlan`) encoding the review order; Jenkins / NoGfix are trailing
  sections; NoGfix pictures optional; `Select-ValidCorrelIds` drops #VALUE!/blank.
- `EvidenceExecutor.ps1` - walks a plan and performs the Excel inserts;
  reports MissingRequired / MissingOptional / Warnings.
- `ProgressLog.ps1` - append-only `status\progress.jsonl` (UTF-8 no BOM).
- `ProjectLabels.ps1` - Japanese sheet/label names built from `[char]` so
  consumers stay ASCII / codepage-agnostic.
- `ScreenRegion.ps1` - pure screen-region clamp math.
- `AlignCompare.ps1` - pure sheet compare + migration-type classification.
- `Tests\` - `Run-Tests.ps1` (parse-checks every .ps1 + runs units) and units
  for MappingStore / GfixLog / EvidencePlan / ScreenRegion / AlignCompare.
- `Check-Encoding.ps1` - enforces the no-BOM (.ps1) / BOM-required-Japanese
  (.psd1) policy and prints constructed Japanese labels.
- `Watch-MappingProgress.ps1` - read-only monitor; copies the mapping and tails
  progress.jsonl, so it never locks mapping_<Owner>.csv.
- `Align.ps1` - Align/Precheck phase: compare work evidence vs the J4 baseline
  workbook by Excel_NAME; DryRun report by default, `-Apply` syncs values
  (formats TODO). Migration-type branching (Host->Open = 3 receive sheets;
  Open->Open / Open->Host add the GIFT/GFIX send-result sheets).

### Changed
- `Generate-HostOpenMapping.ps1` - removed pasted garbage at EOF (would break
  the file); owner-match arrows built via `[char]` (raw arrows silently broke
  matching under no-BOM PS 5.1); added `isMarked` / `isGfixLogMarked` columns;
  on regenerate (`-Force`) old completion status is merged by Correl_ID_M so
  finished work is not wiped; added owner / GFIX-row match counts.
- `JenkinsSnap.ps1` - routes through MappingStore (atomic per-row write instead
  of full-CSV re-read/re-write); emits progress events. Stable Edge/URL/
  screenshot logic and operator prompts kept.
- `DfSnap.ps1` - `region` capture default (df.exe window handle is flaky); waits
  for MainWindowHandle; window mode validates the rect and falls back to region
  (not fullscreen); per-direction crop for the asymmetric shadow.
- `ReplaceEvidence.ps1` - rewritten plan-driven (correl-major), replacing the
  folder-major loop. Gift/Gfix/Df builders + executor; GFIX log via the matcher;
  `isReplaced |= bit` only when all REQUIRED pieces inserted; NoGfix optional
  (`-AllowMissingOptionalNoGfix`); per-correl misses -> progress.jsonl.
- `GfixLogDownload.ps1` - per-correl transaction; on a missing log it does NOT
  blind-navigate: interactive pauses for manual download/retry/skip,
  `-NonInteractive` stops the run. Log naming
  `<id>_<timestamp>_<orig>.log`; GFIX_log set only when a log is really found.
- `VerifyConfig.psd1` / `VerifyTool.ps1` - Df region/crop config; new `Align`
  and `WatchProgress` phases, scripts, aliases; Df region args passed through.

### Notes / open items
- Snap output paths are phase-only (`snap\<type>\<Correl_ID_S>.png`); no
  biz_code/to_code subfolders (TO_code still groups Jenkins navigation only).
- Align full branching needs two domain inputs: the FROM_sys/TO_sys literals
  meaning "Host" (`-HostSystemTypes`) and confirmation of per-type sheet sets.
  Until set, type is Unknown and Align uses the Host->Open scope with a warning.
- Replace partial state is tracked by row bits + progress.jsonl (no new columns).

## 2026-05-19 - ReviewEvidence live test + Apply-LlmPatch v3

### Added / Changed
- Added and wired `ReviewEvidence.ps1` manual review flow.
  - New review phases: `ReviewGift`, `ReviewGfix`, `ReviewDf`, and `ReviewEvidence`.
  - `isReviewed` now follows bitmask semantics: `1=GIFT`, `2=GFIX`, `4=DF`, `7=all`.
  - `ReviewEvidence.ps1` opens evidence workbooks by `Excel_NAME`, waits for manual review, then saves/closes and updates mapping.
- Updated `VerifyConfig.psd1` with review phase entries, aliases, `SaveWaitMs`, and review bit values.
- Updated `VerifyTool.ps1` so review phases route to `ReviewEvidence.ps1` with `ReviewBit`.
- Fixed `ReviewEvidence.ps1` `$pid` bug by avoiding PowerShell's built-in read-only `$PID` variable.
- Explicitly opened workbooks as read-write via `Workbooks.Open($file, 0, $false)` and added read-only detection.
- Moved cursor relocation to after manual review instead of immediately after opening.
  - After manual review, the script selects `A3` by default, or `A1` for sheets whose name contains the arrow marker.
  - It also scrolls the active window to the top-left area after selecting the target cell.
- Clarified review prompt wording: `update mapping` is used instead of `mark`, to avoid confusion with the red-rectangle Mark phase.

### Tooling
- Replaced `Apply-LlmPatch.ps1` with v3 behavior.
  - Resolves relative patch paths against the script/repo directory.
  - Trims one boundary newline inside `<search>` / `<replace>`.
  - Validates multiple patches against the in-memory updated file state.
  - Writes each changed file once, preventing later patches from overwriting earlier patches on the same file.
  - Preserves BOM and original line endings.
- Confirmed that patch text copied with Windows PowerShell must use `Get-Content -Raw -Encoding UTF8 | Set-Clipboard` to avoid Japanese mojibake.

### Live test notes
- `ReviewGift -TargetIds QJRVWD50` successfully opened the workbook and updated `isReviewed` bit `1`.
- `ReviewGift -TargetIds KJRVWD64` successfully opened the workbook, waited for manual review, saved/closed, and updated mapping.
- Current remaining issue: after `Ctrl+S` + `Esc`, Excel may still report `$wb.Saved = false`, causing an extra confirmation prompt. Next change should remove the second confirmation and use a longer wait before close.

### Pending / Next
- Change save routine to:
  1. send `Ctrl+S`;
  2. wait about 300 ms;
  3. send `Esc` to dismiss the GenBa macro prompt;
  4. wait about 5 seconds for network-share save to settle;
  5. if `$wb.Saved` is still false, log a warning only;
  6. close without second save to avoid a macro loop;
  7. update mapping and continue to the next workbook.
- Keep `Close($false)` for now. Do not switch to `Close($true)` unless the macro behavior is retested, because close-time save may trigger the same macro/prompt again.
- After the save-close flow stabilizes, run a multi-workbook review pass without `-TargetIds`.

## 2026-05-18 (Session 3) — AI Co-workflow Tooling & Encoding Fixes

### Added / Changed
- **AI Sync Infrastructure**: Added `Pack-LlmContext.ps1`, `Apply-LlmPatch.ps1`, and `Export-DailyPatch.ps1` to facilitate secure, clipboard-driven LLM interactions bypassing Genba network restrictions.
- **Encoding Fixes**: Enforced `-Encoding UTF8` across all `Get-Content` and `Set-Content` calls in the LLM toolchain to prevent Shift-JIS text corruption (文字化け).
- **Docs**: Planned deprecation of `HANDOFF.md`. Updated `CLAUDE.md` to track local build tools.

## 2026-05-18 (later2) — v2.2: Mark phase, auto column repair, shape probe

### Added

- **`Mark.ps1`** — Phase `MarkGift` / `MarkGfix` / `MarkDf`. Draws red rectangles around designated areas of each picture in the evidence workbook. Idempotent: shapes whose name starts with `verifyMark_` are cleared before drawing, so re-runs are safe.
  - Reads each picture's `AlternativeText` metadata (stamped by `ReplaceEvidence.ps1`) to know what the picture represents (`v1|<folder>|<correl>`).
  - For each picture, looks up `VerifyConfig.psd1 -> Mark.Boxes -> <folder>` to find the list of `(OffsetX, OffsetY, Width, Height)` rectangles to draw, in points, relative to the picture's top-left.
  - `isMarked` bitmask field (1=Gift, 2=Gfix, 4=Df) tracks completion. Bit only set when at least one mark was drawn and nothing failed.
  - All marks ZOrder = `msoBringToFront` so they sit on top of the pictures.
  - Line color = red BGR `0x0000FF`, line weight default `1.5` (override per-box or via `Mark.LineWeight`).

- **`Probe-Shapes.ps1`** — Calibration tool. Lists every Shape in a workbook with its name, type, `(Left, Top, Width, Height)`, and `AltText`. Color-codes output:
  - **green** = Picture with `v1|...` metadata (Mark anchors on these).
  - **yellow** = manual AutoShape (your reference rectangle — measure these).
  - Prints the calibration recipe at the end:
    ```
    OffsetX = AutoShape.Left - Picture.Left
    OffsetY = AutoShape.Top  - Picture.Top
    Width   = AutoShape.Width
    Height  = AutoShape.Height
    ```
  Usage: `.\VerifyTool.ps1 -Phase ProbeShapes -ProbeFile evidence\KJRVWD64.xlsx`

- **`ExcelHelpers.ps1`** additions:
  - `Set-ShapeMetadata` / `Get-ShapeMetadata` — store and read `"v1|<key>|<value>"` payload on a Shape's `AlternativeText`. Used to identify pictures across runs.
  - `Add-RedRectangle` — hollow red border rectangle at absolute `(Left, Top)`, configurable `Width`/`Height`/`LineWeight`/`Name`. Sets `msoBringToFront`.
  - `Remove-MarkShapes` — deletes every Shape on a sheet whose Name starts with a given prefix. Used by Mark for idempotent re-runs.

- **`ReplaceEvidence.ps1`** changed:
  - Every inserted Picture is now stamped with `Set-ShapeMetadata` immediately after insertion. ExcelSnap gets key `'excel'` and the `JOB_NAME` as value. Per-correl snaps get the source folder name and `Correl_ID_S` as value. NoGfix tail snaps stamped too. This is what makes Mark work.

- **`Ensure-PhaseColumns`** function in `VerifyTool.ps1`. Reads `PhaseOrder`, finds every non-empty `Field`, ensures the mapping CSV has that column (defaults missing cells to `'0'`). **Auto-called on every startup** so adding a new phase/field never causes `(value missing)` issues. Existing data is never modified.
  - Manual entry point: `.\VerifyTool.ps1 -Phase RepairMapping`.
  - Solves: "I shouldn't lose progress just because I changed the mapping module" — old mapping CSVs upgrade in place.

- **`VerifyConfig.psd1`** additions:
  - New `Mark` section with `NamePrefix`, `LineWeight`, and `Boxes` table per source folder. Default offsets are placeholders (all `OffsetX=0, OffsetY=0`); calibrate via Probe.
  - New `Scripts` entries: `Mark`, `Probe`.
  - New `PhaseOrder` entries: `MarkGift` (bit=1), `MarkGfix` (bit=2), `MarkDf` (bit=4), `RepairMapping`, `ProbeShapes`.
  - New aliases: `Mark` / `Mgift` / `Mgfix` / `Mdf` / `RepairMapping` / `Repair` / `EnsureCols` / `Probe`.

### How to calibrate Mark (first-time setup)

1. Run `ReplaceGift` once on a single Excel: `.\VerifyTool.ps1 -Phase ReplaceGift -TargetIds KJRVWD64`.
2. Open `evidence\KJRVWD64.xlsx` manually.
3. Switch to the `GIFT受信結果` sheet.
4. Draw a red rectangle by hand around the area you want Mark to circle on one of the pictures (the area depends on which picture's source folder — e.g. for `GIFT_HM` picture, circle the `処理状態` cell).
5. Save the workbook (leave it closed when you run probe).
6. `.\VerifyTool.ps1 -Phase ProbeShapes -ProbeFile <full path to that xlsx>` — note the Picture position and the AutoShape position.
7. Compute `OffsetX = AutoShape.Left - Picture.Left`, `OffsetY = AutoShape.Top - Picture.Top`, `Width = AutoShape.Width`, `Height = AutoShape.Height`.
8. Edit `VerifyConfig.psd1 -> Mark -> Boxes -> 'GIFT_HM'` (etc.) with the measured values.
9. Repeat per folder you care about (`GIFT_MQ`, `GIFT_Jenkins`, etc.).
10. Run `.\VerifyTool.ps1 -Phase MarkGift` — it will draw the same rectangle at the right place on every picture across every evidence workbook.

For folders you don't want marked (like `GIFT_noGfixfile` and `excel`), leave the list empty `@()`.

### Known limitations

- All pictures from the same source folder get the same rectangle. If the row position varies inside the picture (e.g. HM occasionally has the target row on row 2 instead of row 1), the rectangle will be off. OCR support is the planned escape hatch but not yet built — see `OcrLocate.ps1` in pending.
- `Mark` requires `ReplaceEvidence` to have run first (otherwise pictures have no metadata). If you re-run `ReplaceGift` after `MarkGift`, all marks are wiped by `Reset-SheetBelowRow`; just re-run `MarkGift` after.
- Multiple AutoShape candidates in the same Probe output (e.g. you drew 2 reference rectangles) are all listed — pick whichever is yours.

---

## 2026-05-18 (later) — v2.1: Validate phase + Claude Code context

### Added

- **`Validate.ps1`** — Phase `Validate` (aliases: `Check`, `Diagnose`). Read-only diagnostic. Scans `WorkDir` without touching Excel and reports:
  - Mapping file presence + required column check + `isReplaced` bitmask distribution (e.g. `none=4, Gift=2, all=1`).
  - Directory structure (`evidence`, all `snap/*` subfolders, `DATA`, `log`) with PNG/XLSX counts.
  - Template files (`template.xlsx`, `template_<bizcode>.xlsx`) present in WorkDir.
  - **Per Excel_NAME readiness matrix** — for each Excel_NAME, a Y/N per phase (Clone / Gift / Gfix / Df) plus a compact `missing` column showing what's blocking it (`evidence`, `excel`, `gHM`, `gMQ`, `gJk`, `GHM`, `GJk`, `DF`).
  - Aggregate readiness totals per phase (e.g. `ReplaceGift : 5/8 ready`).
  - Next suggested action with actual CLI commands to run.
  - `-Compact` suppresses the per-Excel matrix.
  - `-TargetIds` narrows the scan.

- **`CLAUDE.md`** — context file for Claude Code at home. Documents project purpose, file map, key conventions (dot-source rules, switch-flag pattern, encoding, bitmask), Excel COM rules, current state, planned phases, Misaki's preferences, and the cross-environment workflow pattern.

- **Clone next-step hints** — `Clone.ps1` now prints clear follow-up commands at the end. If anything was missed, it tells you which fallback locations were tried. If anything was cloned, it points to `Validate` and `ReplaceGift` as the next moves.

- **`VerifyConfig.psd1`** — `Scripts.Validate = 'Validate.ps1'`, `PhaseOrder` entry, aliases (`Validate`, `Check`, `Diagnose`).

### Why Validate matters

When entering a new work session, Misaki was previously guessing which phase to run next. Status only shows mapping CSV column states — it doesn't know whether the actual snap PNGs exist on disk. Validate fills the gap by cross-checking mapping rows against filesystem state in one pass, so the answer to "what should I run next?" is one command away.

---

## 2026-05-18 — v2: Clone + Replace + bitmask

### Added

- **`ExcelHelpers.ps1`** — dot-source library for Excel COM operations.
  - `New-ExcelApp` / `Close-ExcelApp` — lifecycle, ScreenUpdating off, alerts off.
  - `Open-Workbook` / `Close-Workbook` / `Get-SheetByName` / `Unhide-AllSheets`.
  - `Reset-SheetBelowRow` — deletes shapes whose top ≥ row's top, then clears values + formats + highlights in the range `A<startRow>:T<lastUsed>`. Uses `End(xlUp)` from the bottom of column A to find last used row (fixes the `xlDown` bug in the old script).
  - `Get-RowAtOrBelow` / `Get-NextAnchorRow` — anchor-row math using `shape.Top + shape.Height` against `Cells(r,1).Top`. Approximates start scan row from `shape.Top / 15` for speed.
  - `Insert-PictureSendToBack` — inserts picture at native size with z-order = `msoSendToBack` (1), so later Mark rectangles will stay on top.
  - `Write-PlainText` — plain label without bold / color / highlight.
  - `Write-LogLines` — paste multi-line log content into stacked cells.
  - `Get-BitValue` / `Set-BitValue` — read / `-bor` write on integer columns.
  - `Ensure-Column` — adds a missing column to every row before `Export-Csv` so it persists.

- **`Clone.ps1`** — Phase `Clone` (aliases: `MkExcel`, `RenameExcel`).
  - Groups mapping by `Excel_NAME`.
  - Per group, tries source paths in priority order:
    1. `<SourceDir>\<bizcode>\<Excel_NAME>.xlsx` for each bizcode candidate.
    2. `<WorkDir>\template_<bizcode>.xlsx`.
    3. `<WorkDir>\template.xlsx` (universal fallback).
  - Bizcode candidates come from `-BizCodes` if given, otherwise from row's `TO_code` then `FROM_code` (deduped).
  - Copies to `<WorkDir>\evidence\<Excel_NAME>.xlsx`.
  - Skip if dest exists unless `-Force`.
  - Prompts for `SourceDir` interactively if neither CLI flag nor session has it.

- **`ReplaceEvidence.ps1`** — Phases `ReplaceGift`, `ReplaceGfix`, `ReplaceDf`.
  - Single script, `-Mode Gift|Gfix|Df` switches behavior.
  - Per `Excel_NAME` group: open workbook → find sheet by mode → `Reset-SheetBelowRow 3` → insert ExcelSnap at B3 (Gift/Gfix only) → stacked per-correl snaps (`GIFT_HM`/`GIFT_MQ`/`GIFT_Jenkins` or `GFIX_HM`/`GFIX_Jenkins` or `DF`) → tail per mode → save.
  - Tail:
    - **Gift** — label `GFIX Jenkins フォルダ受信ファイルなし` once, then `GIFT_noGfixfile\<correl>.png` snaps stacked. NoGfix snaps may legitimately be absent; absence is logged as info, not a failure.
    - **Gfix** — per-correl label `GFIX受信log` + placeholder text. Stub function `Get-GfixLogLines` returns `<<TODO: GFIX 受信 log for <correl>>>` for now; replace with real grep against `work\log\` after `GfixLodDownload` lands (cf. old `Replace-GFIX.ps1` — `Find-RecvLogFile` / `Find-SendLogFile` / `Extract-SystemOutBlock`).
    - **Df** — nothing extra.
  - **All-OK contract** — `isReplaced |= bit` only if every required image existed and was inserted without error. Missing required images mark the run as failed for that group and the bit stays clear.

- **`isReplaced` bitmask field** in `mapping_<owner>.csv`.
  - `bit 0 (1)` = Gift done, `bit 1 (2)` = Gfix done, `bit 2 (4)` = Df done. Total 7 = all done.
  - One column instead of three flag columns, avoids schema bloat.
  - `Ensure-Column` adds the column to every row on first run so `Export-Csv` writes it.

- **`Get-FieldStats` bitmask support** in `VerifyTool.ps1`. New optional `BitValue` parameter. When `> 0`, "done" means `(value -band BitValue) -eq BitValue` instead of `eq '1'`.

- **`PhaseOrder` `BitValue` entries** in `VerifyConfig.psd1`. `Show-Status` and `Get-RecommendPhase` both read it. Status display shows `bit=N` next to the field name.

- **`CloneSourceDir` session persistence**. `VerifyTool.ps1` now has `-CloneSourceDir <path>` and `-BizCodes A,B` parameters; `CloneSourceDir` is saved into `verify_session.json` so subsequent runs reuse it without re-typing.

- **`Ask-RunOptions`** in `VerifyTool.ps1` gained two more interactive options: `d` to set `CloneSourceDir`, `b` to set `BizCodes`.

- **Aliases** in `VerifyConfig.psd1`:
  - `Clone` / `MkExcel` / `RenameExcel` → `Clone`
  - `Replace` / `ReplaceEvidence` → `ReplaceGift` (default; explicit per-mode names below)
  - `ReplaceGift` / `Rgift` → `ReplaceGift`
  - `ReplaceGfix` / `Rgfix` → `ReplaceGfix`
  - `ReplaceDf` / `Rdf` → `ReplaceDf`

- **`Replace` section** in `VerifyConfig.psd1`:
  - `StartRow = 3`, `ColAnchor = 2`, `BlankRowsBetween = 1`, `ClearEndColumn = 20`.
  - `GiftNoGfixLabel`, `GfixLogLabel`, `GfixLogTodoText`. Each can be edited without touching code. Empty values fall back to the in-code `[char]0xNNNN` defaults.

### Changed

- `Get-FieldStats` signature: added `[int]$BitValue = 0`. Default behavior (no `BitValue`) is unchanged — still `eq '1'`.
- `Show-Status` reads `BitValue` from each phase entry and prints `bit=N` next to the field name when present.
- `Get-RecommendPhase` likewise honors `BitValue` when picking the next pending phase.
- Old planned-only handler for `ReplaceEvidence` removed; `ReplaceEvidence` alias now points to `ReplaceGift` so old muscle memory still works.
- `Show-VerifyHelp` gained Clone / Replace examples and the bitmask legend.

### Conventions kept

- Switch flags resolved to `[bool]$X.IsPresent` before any dot-source, to avoid the `$Force` overwrite bug pattern.
- Tool scripts with `param()` are never dot-sourced; they are called via `& $p @args`. `ExcelHelpers.ps1` has no `param()` so dot-sourcing it from `ReplaceEvidence.ps1` is safe.
- Japanese text in `.ps1` files uses `[char]0xNNNN` for sheet names. Labels live in `VerifyConfig.psd1` as literal Japanese strings (UTF-8 BOM PSD1).
- English-only console output to avoid encoding issues in Windows shells.

### Pending / next iteration

- `Mark.ps1` — red rectangles, highlights overlaid on inserted pictures. Will dot-source `ExcelHelpers.ps1`.
- `TimeRange.ps1`.
- `OcrLocate.ps1`.
- `GfixLodDownload` real implementation (drives the `Get-GfixLogLines` stub in `ReplaceEvidence.ps1`).
- `DfSnap` real implementation.
- After `GfixLodDownload` lands: replace `Get-GfixLogLines` stub with grep against `work\log\` for recv/send log lines per `Correl_ID_S`. Reference old `Replace-GFIX.ps1` helpers (`Find-RecvLogFile`, `Find-SendLogFile`, `Extract-SystemOutBlock`).

## How to use this changelog

When iterating between browser (work) and IDE (home):

1. After making changes on either side, add an entry to this file under a new date heading describing what changed and why.
2. Sync the whole `VerifyTool` directory across.
3. The other side reads this file first to know what's new.

Suggested entry skeleton:

```markdown
## YYYY-MM-DD — short title

### Added / Changed / Fixed / Pending
- ...
```
