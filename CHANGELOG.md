## 2026-07-20 - ProcessTime: parse the OCR the recognizers ACTUALLY produce + content-validated picture candidates (v2.12.2)

Driven by the first full office-PC run (43 rows, ALL `not detected [none]`,
32 export folders): the real `.ocr.txt` dumps showed the ja recognizer reads
the HM rows COMPLETELY but injects spaces inside every time token
(`10 :58 :20`, `00 :00 : 0 1`) and garbles the status literal
(sei-TEI-shuuryo for normal-end), while the en-US recognizer reads the date
columns but drops the time-of-day entirely -- so the strict
`\d{2}:\d{2}:\d{2}` datetime anchor matched zero lines on every one of the
43 rows. Separately, 11 of the 43 rows (the hand-made `*JDLW*` workbooks;
9 unique correls -- `JIDSCS4S` / `JIDSQS4S` are each mapped twice) hit
`[MISS] no exportable HM picture`: their HM picture sits ABOVE the found
correl label (picture Tops 29-688 vs label Tops 886-1469) where neither
search tier ever looked, which is also why only 32 (not 43) per-correl
export folders were created.

### Fixed
- `ConvertFrom-ProcessTimeOcrLines` (`ProcessTimeParse.ps1`) now repairs
  OCR-injected whitespace inside time tokens before anchoring: new pure
  `ConvertTo-ProcessTimeNormalizedLine` rewrites every loose H:M:S cluster
  (spaces around the colons and even between the two digits of one field)
  to canonical zero-padded `HH:mm:ss`, leaving 14-digit datetime columns
  and `1 ,036` record counts untouched. The date-to-time gap in the
  datetime anchor is capped at 5 spaces so a date whose own time was
  dropped by OCR can never pair with a time from a later column (the
  en-US date-only rows now correctly yield NO row instead of a potential
  invented-midnight row).
- Status literal matching is fuzzy: after the exact normal-end/abend
  literals miss, a `...shuuryo` match classifies by the preceding
  characters (sei -> normal-end, i -> abend) and reports the CANONICAL
  literal, so the observed sei-TEI-shuuryo garble still counts.

### Added
- Cross-check mechanism: every parsed row now also carries
  `PageDuration` (the page's own proc-time column, first standalone
  `HH:mm:ss` after the datetimes) and `CorrelSeen` (whether the target
  correl id appears verbatim in the OCR line). `Resolve-ProcessTimeSide`
  compares the derived end-start duration against `PageDuration` and
  prints a `note:` on mismatch (derived wins); a row picked from a
  relaxed picture candidate without `CorrelSeen` is flagged.
- Partial extraction instead of blanket misses: a line with ONE readable
  datetime plus a status-ish literal becomes a `Partial` row (start
  kept, end `$null`); the phase reports `start <t>; end NOT read`, fills
  the duration from `PageDuration` when available (real on-page
  evidence, tagged in the note), writes the partial row into the
  ProcessTime workbook with source `ocr-partial[:<tier>]`, and keeps the
  per-side mapping flag at `2` (only a full start+end read sets `1`).
  New pure `Select-ProcessTimeRow` / `Get-ProcessTimeRowRank` implement
  the preference order (full beats partial, correl-seen beats unseen,
  newest StartTime among equals -- the established newest-wins rule).
- Detailed miss reporting: new pure `Get-ProcessTimeOcrMissNote`
  summarizes WHY a pooled OCR read yielded nothing ("N with a date but
  no readable time-of-day", "no date/time tokens recognized"), printed
  per candidate and folded into the final `not detected` note; a missing
  recv sheet or unfound correl label now also says so explicitly
  (previously silent `[none]`).
- Non-standard workbook layouts: `Resolve-ProcessTimeSide` now tries
  CONTENT-VALIDATED picture candidates in confidence order -- (1) the
  correl's section picture (position trusted, as before); (2) up to 2
  pictures below the label (the first is position-trusted only when the
  section itself had no picture, i.e. the old v2.12.1 retry target; the
  second must show the correl id in its OCR text); (3) up to 3 pictures
  ABOVE the label, nearest first, accepted ONLY when the correl id is
  seen (position proves nothing up there). This reaches the `*JDLW*`
  hand-made workbooks whose HM picture sits above the label / whose
  label column collapses the section to one row.
  `Export-SheetPicturesToPng` (`EvidenceImageExport.ps1`) gained an
  optional `FromBottom` flag (existing callers unchanged) for the
  nearest-above-first order. Candidate PNGs/dumps are name-tagged
  (`<side>_<correl>_belowlabel_01.png` etc.); the trusted section tier
  keeps the v2.12.1 names.
- `Tests\Test-ProcessTimeParse.ps1`: the REAL office-PC OCR lines (ja
  spaced-time + garbled-literal rows, en-US date-only row) are now
  fixtures; 46 assertions cover normalization, fuzzy status, partial
  rows, PageDuration capture, rank/selection, and the miss notes.
- `-Force` reruns no longer stack duplicate rows in the output workbook:
  `Write-ProcessTimeWorkbook` deletes any existing row whose
  (Excel_NAME, Correl_ID_S) pair is being rewritten this run before
  appending, so the workbook keeps ONE row per correl (the first
  office-PC run already wrote 43 all-blank rows; the redo replaces them
  instead of appending 43 more).
- Duplicate mapping rows for one correl (observed: `JIDSCS4S` /
  `JIDSQS4S` each mapped twice) are extracted ONCE per run; later
  duplicates mirror the first occurrence's per-side flags with a
  `[DIAG]` pointing at the mapping instead of redoing export+OCR and
  writing a second identical workbook row.

### Notes
- The 43-rows-vs-32-folders question from the run log: 32 rows exported at
  least one PNG (32 unique correls -> 32 folders under `snap\ProcessTime`);
  the 11 `[MISS]` rows never created a folder. Two of those correls are
  duplicated mapping rows (`JIDSCS4S`, `JIDSQS4S`) -- worth checking whether
  the duplication is intended.
- Pure logic green under portable pwsh 7 in this dev env
  (`Tests\Test-ProcessTimeParse.ps1` 46/46; full `Run-Tests.ps1` has only
  the 2 pre-existing Linux path-separator failures in Test-EvidencePlan,
  which pass on Windows). The COM/OCR candidate loop is static-checked
  only -- confirm on the office PC, especially: a standard workbook still
  resolves from the section picture in one OCR pass; a `*JDLW*` workbook
  now finds its above-label picture; and the GIFT side of the `*JDLW*`
  workbooks now explains itself (missing sheet vs missing label).

## 2026-07-16 - ProcessTime OCR robustness + workbook loose-match (v2.12.1)

Synthesis of the open ProcessTime follow-up PRs (#111, #112, #113); PR #112
was closed with no committed diff, so its intent is folded in here.

### Fixed
- `ProcessTime.ps1` OCR tier ran too rarely and left nothing to debug when it
  did. `Export-CorrelPicture` now prints `[DIAG]`/`[MISS]` instead of silently
  returning `$null`, and when the strict correl-section bounds find no picture
  it retries once from the correl label down to the sheet end (still capped at
  the first picture). `Resolve-ProcessTimeSide` logs each `[OCR] ... lang=...`
  invocation and writes the pooled OCR lines to a `<base>_NN.ocr.txt` sidecar
  next to the exported PNG so a zero-match run is diagnosable from the actual
  recognized text.
- `ProcessTime.ps1` `-OcrLanguage` now defaults to `''` (en-US only); set it
  to e.g. `ja` to also pool the Japanese recognizer. The startup banner and
  per-image dumps make the effective language explicit.
- The two sides of one correl no longer collide on the same
  `<correl>_NN.png` / `.ocr.txt` in the shared per-correl export dir:
  `Export-CorrelPicture` takes a `BaseName` and `Resolve-ProcessTimeSide`
  threads `GIFT_<correl>` / `GFIX_<correl>`. Stale `<base>_*.png` / `<base>_*.txt`
  from a previous run are cleared before export so a fresh MISS can't be
  masked by last run's leftovers.
- `Tests\Test-ProcessTimeParse.ps1` compares parsed datetimes via
  `.ToString('yyyy/MM/dd HH:mm:ss')` instead of the culture-dependent
  `[string]$dt` cast, so the assertions pass on a JP-locale host (PR #111).

### Added
- `Export-SheetPicturesToPng` (`EvidenceImageExport.ps1`) gained an optional
  `MaxPictures` cap (0 = unlimited, existing callers unchanged) so
  ProcessTime's "label to sheet end" retry exports only the one HM screenshot
  it needs instead of chart-exporting every picture on a busy sheet
  (PR #112 intent).
- `WorkbookResolver.ps1` gained `Convert-WorkbookNameForLooseMatch` /
  `Find-LooseWorkbookCandidates`, used as a last-resort fallback in
  `Find-WorkbookByExcelName` (after exact / wildcard / full-width all miss) to
  tolerate transposed J/W prefixes and full-width/half-width variants, emitting
  a `[WARN]` when the loose match is used (PR #113).

### Notes
- No Windows/Excel/OCR in this dev environment: `ProcessTimeParse.ps1` pure
  logic is unit-tested; the Excel COM / picture-export / real OCR paths are
  static-checked only. Confirm on an office PC that the retry export, the
  `.ocr.txt` dumps, the per-side base names, and the loose workbook match
  behave against a real evidence workbook.

## 2026-07-16 - ProcessTime phase: HM processing start/end/duration extraction (v2.12.0)

### Added
- New `ProcessTime` phase (`ProcessTime.ps1`) extracts each correl's HM
  batch processing start time / end time (and derives the duration) for
  both the GIFT and GFIX sides, then appends one summary row per correl to
  a standalone evidence workbook (`ProcessTime_<Owner>.xlsx` by default).
- Two-tier source per side, cheapest/most-accurate first: (1) the archived
  Ctrl+A page text `HmSnap.ps1` already saved at snap time
  (`snap\GIFT_HM\<correl>.txt` / `GFIX_HM`, when `SnapVerify.SaveText` was
  on), re-parsed with the existing `SnapVerify.ps1` `ConvertFrom-HmPageText`;
  (2) OCR of the HM screenshot already inserted into the evidence workbook
  by ReplaceGift/ReplaceGfix (new `Export-CorrelPicture`, which locates the
  correl's `Correl_ID_S` label in column `Replace.ColAnchor` on the
  GIFT/GFIX jushin-kekka sheet and exports the first picture in that
  section -- always the HM screenshot, even on the GFIX sheet where a log
  block follows in the same section), read via `OcrWindows.ps1`'s
  `Invoke-WinOcrFile` in both en-US and a configured secondary language.
- New pure library `ProcessTimeParse.ps1` (unit-tested in
  `Tests\Test-ProcessTimeParse.ps1`): `Get-ProcessDurationText`,
  `ConvertFrom-ProcessTimeOcrLines` (anchors on two datetime tokens + a
  status literal per OCR'd row instead of trusting column position), and
  `Get-NewestProcessTimeRow` (newest-by-StartTime).
- Three new mapping columns (`MappingStore.ps1`): `GIFT_ProcessTime` /
  `GFIX_ProcessTime` (informational per-side result) and
  `ProcessTime_Inserted` (this phase's plain 0/1 `Get-PendingRows` field).
- New `ProcessTime` config block in `VerifyConfig.psd1`, `Scripts.ProcessTime`,
  a `PhaseOrder` entry (after `ReplaceDf`, before `MarkGift`), and `Aliases`
  (`ProcessTime`/`Pt`/`ProcTime`). Wired into `VerifyTool.ps1`
  (`-Phase ProcessTime`, respects `-Force`/`-DryRun`). `ProcessTime` added to
  `ConfigOverlay.ps1`'s `excel` editor group per this repo's schema-drift
  convention. New `ProjectLabels.ps1` `SheetProcessTime` label.

### Notes
- No Windows/Excel/OCR in this dev environment: `ProcessTimeParse.ps1`'s
  pure logic is unit-tested, but the Excel COM (label-cell find,
  section-bounds, picture export, workbook create/append) and the real HM
  screenshot OCR quality are static-checked only -- confirm on an office PC
  against a real evidence workbook, including a correl whose archived
  `snap\GIFT_HM\<correl>.txt` is missing (forces the OCR tier), before
  trusting this in production.

## 2026-07-14 - Four-direction snap crop, per-side + per-snap-folder (v2.11.0)

### Added
- Snap screenshot cropping (HM/MQ/Jenkins) is now directionally controllable
  instead of a single uniform `Window.CropPx` trimmed off all four sides.
  New `Window.CropLeft`/`CropTop`/`CropRight`/`CropBottom` (default `-1` =
  inherit `CropPx` for that side) let an operator crop one edge by a
  different amount than the others. New `Window.CropByFolder` (empty `@{}`
  by default) narrows this further to one snap folder (`GIFT_HM`, `GFIX_HM`,
  `GIFT_MQ`, `GIFT_Jenkins`, `GFIX_Jenkins`, `GIFT_noGfixfile`): each entry
  may set any of `Left`/`Top`/`Right`/`Bottom` (px), falling back to the
  resolved global `Crop<Side>`/`CropPx` value for any side left out --
  mirrors the existing `Df.CropLeft/Top/Right/Bottom` pattern, applied
  per-folder.
- New pure `Resolve-DirectionalCrop` (`ScreenRegion.ps1`, unit-tested)
  resolves `CropPx` + global per-side overrides + a per-folder override into
  four concrete non-negative ints. `VerifyTool.ps1`'s new `Resolve-FolderCrop`
  wraps it per phase dispatch and threads the result as new
  `-CropLeft`/`-CropTop`/`-CropRight`/`-CropBottom` params to
  `HmSnap.ps1`/`MqSnap.ps1`/`JenkinsSnap.ps1`/`Crop-Snap.ps1` (each script's
  own `Invoke-CropPng`/`Invoke-CropDir` extended in place with the same
  `-1`-sentinel-inherits-`cropPx` fallback, so a bare `-CropPx`-only
  invocation is unchanged).
- The interactive menu's `c` option now accepts a plain number (uniform,
  unchanged) or `L,T,R,B` (e.g. `6,8,6,10`) to set the four sides
  individually; the status display shows `CropPx : 6` when uniform or
  `CropPx : 6 (L6/T8/R6/B10)` once any side diverges. New CLI params
  `-CropLeft`/`-CropTop`/`-CropRight`/`-CropBottom` (default `-1`) mirror
  `-CropPx`'s CLI > config precedence and are remembered in
  `verify_session.json`.

### Notes
- Every new field defaults to inherit-CropPx / empty-CropByFolder, so no
  existing work folder's crop behavior changes until an operator opts in.
  Documented in `ConfigOverlay.ps1`'s InitConfig readme text,
  `verify_config.example.json`, and `README.md`. `Window` was already a
  named `snap` group in `Get-ConfigOverlayGroups`, so `CropByFolder`'s
  nested per-folder hashtables walk/edit generically in the
  `-Phase InitConfig -Interactive` field walker with no further code change
  (confirmed by reading `Expand-ConfigWalkPath`'s generic hashtable
  recursion). No PowerShell/Excel in this dev environment -- static-checked
  only (parse review + hand-traced logic); confirm the resolved crop math
  against a real HM/MQ/Jenkins screenshot and the `-Phase InitConfig` walker
  on an office PC.

## 2026-07-10 - Mark image-match: PadWidth/PadHeight box-size margin (v2.10.10)

### Added
- `Find-MarkBoxByImage` (`Mark.ps1`) now accepts optional `PadWidth`/
  `PadHeight` keys on a `Mark.Boxes` entry, alongside the existing `PadX`/
  `PadY`: a constant amount added to the drawn box's `Width`/`Height` (fixed-
  size branch) or to the matched crop's own pixel size before scaling
  (crop-sized branch, same as `PadX`/`PadY`'s existing `2 * pad` treatment).
  Motivating case: `GFIX_Jenkins` marks a Jenkins file-list entry whose on-
  page position differs from `GIFT_Jenkins` (confirmed via `Probe-Shapes.ps1`
  -- their `Template` anchor position genuinely differs, `PadX`/`PadY` fixes
  that), and the entry's filename length varies per correl, so a single
  fixed `Width` sometimes falls short.

### Notes
- `PadWidth`/`PadHeight` are a FIXED number applied to every correl in the
  folder -- they widen/heighten the box by a constant margin, they do NOT
  make it track each correl's actual on-page content length. This is step
  one of an iterative fix (operator explicitly asked for the constant-pad
  knob first); a true per-correl auto-size still needs a measured source
  (e.g. SnapVerify's M5 `loc.json` pixel-localisation rect, or an OCR-based
  text-width measurement like `GfixLog.AutoHighlightWidth` uses for the GFIX
  log highlight) and is not wired into plain `Template` boxes yet -- tracked
  as a follow-up. Documented in `mark_templates/README.txt`'s per-box
  overrides list and `verify_config.example.json`'s `_TemplateExample`.
  Static-checked only (no Windows/Excel in this dev environment) -- confirm
  on an office PC.

## 2026-07-09 - MarkDf: Template image-match now works on cell-range (CellCols) boxes (v2.10.9)

### Fixed
- **A `Mark.Boxes` entry with BOTH `CellCols` and `Template` silently ignored
  the `Template`**: the box-placement loop in `Mark.ps1` branched on
  `CellCols` FIRST and only ran the image-recognition path
  (`Find-MarkBoxByImage`, v2.9.23) in the non-cell branch, so a cell-range
  box could never opt into template matching -- the operator's DF config
  (`CellCols='AW:BC'; RowsFromBottom=2; Template='DfSame.png'`) kept
  producing plain `[MARK]` lines at the fixed cell position. This matters
  for DF specifically: since v2.9.31 `Df.CaptureMode` defaults to `window`,
  so the df.exe same-content button's on-image position moves with however
  the operator sized the window -- a cell-anchored rectangle cannot track
  it (confirmed on a real workbook: marks landed off the button). The
  template match is now tried FIRST for both box kinds; a cell-range box
  falls back to its legacy cell placement (and an offset box to its fixed
  offset) when there is no `Template`, the file is missing, or no match --
  behavior for boxes without `Template` is unchanged. `[MARK-IMG] ...
  (live)` in the console confirms the match fired; `Find-MarkBoxByImage`
  itself is untouched (DF snap PNGs already live at `snap\DF\<correl>.png`,
  so the source-image lookup just works).
- `mark_templates/README.txt` documents the cell-range case + the DF
  example.

### Notes
- Sizing on a match keeps the v2.9.31 rules: this DF box has no
  `Width`/`Height`, so the drawn rectangle takes the template crop's own
  size (scaled to sheet points; add `PadX`/`PadY` for margin). Add
  `Width`/`Height` to the box to force a fixed size with the match as
  anchor. Static-checked only -- confirm `[MARK-IMG]` placement on the DF
  sheet on an office PC.

## 2026-07-09 - CheckSheet date root cause fixed (PS COM binder cast) + DeliverMail filename prefix fallback (v2.10.8)

### Fixed
- **FillCheckSheet column-B date -- actual root cause identified from the
  office-PC log**: the write threw a managed InvalidCastException ("Unable
  to cast object of type 'System.Double' to type 'System.String'", Japanese
  .NET message in the log) out of `$cell.Value2 = <OADate double>` while the
  five STRING columns in the same rows wrote fine. That cast never happens
  inside Excel -- it is PowerShell 5.1's COM binder: the dynamic call site
  for the `Value2` setter caches a conversion rule from a previous (string)
  binding and force-casts the next value through it. Fixes, layered:
  1. New `Set-RangeValue2` helper routes every cell write: normal assignment
     first, and on any exception one retry via IDispatch `InvokeMember`
     (`SetProperty`), which bypasses the PS binder and its cached rule
     entirely. A genuinely bad cell (protected sheet etc.) fails the retry
     too and the ORIGINAL error message propagates.
  2. Last-resort tier for the date only: write it as TEXT (`yyyy/MM/dd`) --
     the cell already carries the date NumberFormat (v2.10.7 sets it before
     the value), so Excel parses the text into a real date; the verify
     accepts the parsed serial via the new `AcceptSerial` argument. Tier-1's
     warning is only surfaced when this tier also fails, and a recovered
     date prints an `[INFO] ... recovered by writing the date as text` line.
  3. Verify warnings now also report which write path ran (`via assign` /
     `via invokemember`).
- **DeliverMail body listed a bare workbook filename** (e.g. `KJODWWB5.xlsx`
  instead of `J4検証資料(...)_KJODWWB5.xlsx`) whenever neither the mapping
  row's legacy `Excel_Prefix` nor `Workbook.ExcelPrefix` was set -- the
  on-disk prefix fallback FillCheckSheet gained in v2.9.29 was never applied
  to the mail body. The fallback is now a shared
  `Resolve-ExcelPrefixWithDisk` (WorkbookResolver.ps1, unit-tested,
  non-interactive `FullWidthFallback Reject`): legacy row column ->
  `Workbook.ExcelPrefix` -> the prefix the real evidence file already
  carries on disk. FillCheckSheet's inline copy was replaced by the shared
  helper; DeliverMail now uses it for the `{3}` body filename, with the same
  `[INFO] ... using prefix found on disk` console note.

### Notes
- The proper per-project fix remains setting `Workbook.ExcelPrefix` in the
  work folder's `verify_config.json` -- the on-disk fallback is a safety
  net, not the primary source. The legacy mapping `Excel_Prefix` column is
  no longer generated anywhere and is safe to delete from existing mapping
  CSVs; it is still honored as a per-row override when present.
- Pure logic (prefix recovery) is unit-tested; the COM write paths are
  static-checked only. Confirm on an office PC: a real date in column B
  (watch for `via invokemember` / the text-fallback INFO in the console),
  and the DeliverMail body now carrying the full prefixed filename.

## 2026-07-09 - CheckSheet date-write hardening + config layering consolidation (v2.10.7)

### Fixed
- **FillCheckSheet column-B (date) write**: two robustness fixes for the
  operator-reported "date write failed" WARN (office-PC log still pending;
  these address the two failure modes reproducible from the code alone):
  1. The date NumberFormat is now applied BEFORE the value is written.
     Writing the OADate serial into a cell still formatted `@` (text)
     stores the digits as text, and re-formatting afterwards does not
     convert it back -- the cell keeps showing `46212`-style text.
     Format-first makes the serial land as a real date in one pass. A
     failed format-set is now itself a visible `[WARN]` instead of a bare
     `catch {}`.
  2. The date format mirrored from the row above is no longer trusted
     blindly: `@` (text) and `General` are rejected (either one renders
     the serial as a bare number / text -- the exact blank/garbled column-B
     symptom) and the configured `CheckSheet.DateFormat` fallback is used
     instead.
- **FillCheckSheet write-verify diagnostics**: `Set-CellChecked`'s WARN
  line now reports the written value, the raw readback (value + .NET type),
  and the cell context -- address, NumberFormat, merged-cell flag, sheet
  `ProtectContents` -- so the next failure log pinpoints WHICH of the known
  causes (protected sheet / merged cell / text format / silent no-op) fired,
  instead of only "did not verify". The numeric verify no longer does a
  blind `[double]` cast on the readback (a text readback used to throw
  inside the check and surface as a misleading "write failed" exception);
  it parses and compares, so number-stored-as-text is reported as a
  mismatch with the actual readback shown.

### Changed
- **CheckSheet path now persists to the WORK FOLDER, not the global session
  file**: the first-run prompt's answer is written into this work folder's
  `verify_config.json` as `CheckSheet.Path` via the new
  `Save-ConfigOverlayValue` (ConfigOverlay.ps1) -- the check sheet is
  project-scoped, and remembering it in `verify_session.json` (old
  behavior) leaked one project's path into every other work folder that had
  no explicit config. The session file remains a legacy read fallback and
  the fallback store when the JSON cannot be written (locked/open file).
  `verify_session.json` is now reserved for machine/operator state
  (WorkDir pointer, Owner, window size, DfExePath -- df.exe location is a
  property of the PC, so it deliberately stays session-first).

### Added
- `Save-ConfigOverlayValue` (ConfigOverlay.ps1): persist ONE dotted-path
  value into a sparse `verify_config.json`, creating the file when absent,
  preserving every operator value and the `_README`/`_SCHEMA` metadata
  keys, and refusing to touch an unparseable file or overwrite through a
  non-object ancestor. Unit-tested in `Tests\Test-ConfigOverlay.ps1`.
- `docs/Configuration.md`: the config layering reference -- the three
  files (`VerifyConfig.psd1` = shipped defaults only / `verify_config.json`
  in WorkDir = single source of truth for everything project-scoped /
  `verify_session.json` = machine+operator ephemera), runtime precedence,
  the travels-with-the-work-folder rule of thumb, a field-by-field
  inventory of paths/prefixes, and the known sharp edges (array-wholesale
  overlay merge, global-session leaks, editor-group drift guard).

### Notes
- Static-checked only (no Windows/Excel in this dev environment). Confirm
  on an office PC: the CheckSheet first-run prompt writing
  `CheckSheet.Path` into `verify_config.json`, and column B landing as a
  real date on the shared check sheet. If the date WARN fires again, the
  new diagnostic suffix (`addr=... fmt=... merged/sheetProtected`) in the
  console log identifies the cause -- send that log.

## 2026-07-08 - Repo hygiene: untrack IDE workspace + session state; generalization roadmap + sanitization audit (v2.10.6)

### Changed
- Stopped tracking `.metadata/` (a whole committed Eclipse/RAD workspace: IDE
  logs carrying the corporate proxy hostname and `C:\Users\<employee-id>`
  paths, server configs, caches), `.project`, and `verify_session.json` (live
  per-operator session state with real work paths / owner name / internal UNC
  share). Files stay on disk; a new `.gitignore` keeps them and other runtime
  or evidence artifacts (`snap/`, `status/`, `bk/`, `mapping_*.csv`,
  workbooks, screenshots, `mark_templates/*.png`) from ever being committed
  again. First step (M0) of the main-branch sanitization plan.
- **Pull caveat for existing clones**: pulling this change tries to DELETE the
  now-untracked files where they are unmodified and raises modify/delete
  conflicts where they are modified. `verify_session.json` is safe to lose
  (regenerated; re-asks WorkDir/Owner once). Close Eclipse/RAD first and copy
  `.metadata` aside if its workspace state matters; being gitignored, a
  restored copy is not re-tracked.

### Added
- `docs/Generalization-Roadmap.md`: staged plan for evolving the tool into a
  generic, profile-driven "evidence workbench" -- branch strategy
  (`spec/gift-gfix` snapshot + progressively sanitized `main`, fresh-history
  repo for any public release since git history is not sanitizable in place),
  tip-sanitization checklist S1-S10 from the 2026-07 repo audit, target
  core/engine/adapters/profiles layering with a file-by-file map, AI
  onboarding (`Describe` phase) design, and milestones M0-M6.
- `docs/Sanitization-Audit.md`: the full audit report (masked literals):
  repo-visibility headline + flip-private recommendation, tip findings
  mapped to S-items, why git history can never be published, module
  coupling classification for all files, split-now vs defer-to-M6
  strategy analysis.

## 2026-07-07 - MarkGfix log highlight: fix auto-width measurement (DPI + GDI/GDI+ mismatch) + per-row diagnostics (v2.10.5)

### Fixed
- The MarkGfix / MarkGfixLog auto-width yellow highlight
  (`GfixLog.AutoHighlightWidth`) kept computing a wrong width on the office
  PC even after the v2.9.24 GenericTypographic fix. Two independent math
  defects in the measurement chain:
  1. **Hardcoded 96-DPI conversion (highlight too LONG on scaled
     displays).** `Get-AutoHighlightColEnd` converted the measured pixel
     width to points with a fixed `* 0.75`, but `Get-TextPixelWidth`
     measures on a `Graphics` from a fresh `Bitmap(1,1)`, which inherits the
     PROCESS'S SCREEN DPI -- and powershell.exe is DPI-aware, so on a
     125%/150%-scaled laptop (the common office setup) the pixels came back
     at 120/144 DPI and the fixed conversion inflated the width by
     1.25x/1.5x. `Get-TextPixelWidth` now pins its bitmap to 96 DPI
     (`SetResolution`), making the documented px->points contract (x 0.75)
     exact on every display scale.
  2. **GDI+ vs GDI renderer mismatch (highlight too SHORT).** Excel renders
     cell text with GDI, whose hinted/bitmap advance widths for classic
     Japanese fonts like MS Gothic run wider than the ideal typographic
     advance (8 px vs 7.33 px per half-width char at 11pt/96dpi -- MS Gothic
     carries embedded bitmap strikes at these sizes). Measuring with GDI+
     `GenericTypographic` (the only path until now) therefore undershot ~8%
     on a long `Command:` line, ending the highlight before the text does.
     New `Get-TextPointWidthInfo` (`ExcelHelpers.ps1`) measures with GDI
     first (`System.Windows.Forms.TextRenderer`, `NoPadding|SingleLine`,
     converted with the REAL screen DPI), falling back to the 96-DPI-pinned
     GDI+ path when WinForms is unavailable.
- Whichever tier measured, the result is now floored at the ideal
  fixed-pitch character-cell width -- new pure `Get-TextCellUnits`
  (half-width chars incl. halfwidth katakana = 0.5 em, full-width = 1.0 em)
  times the font size in points -- so the highlight can never end short of
  the text's nominal advance ("mark enough cells"); the `HighlightColEnd`
  cap still bounds it above, exactly as before. A total measurement failure
  still falls back to the fixed `HighlightColEnd` (legacy full width).
- The COM column-width read's fallback default was corrected from 59.0 to
  48.0 points (the real width of Excel's standard 8.43-char column at
  96 DPI: 64 px x 0.75); it is only used when the per-column `.Width` read
  throws.

### Added
- Per-row width diagnostics: every `Get-AutoHighlightColEnd` decision or
  fallback (no text in the cell, measurement failure, and the full success
  breakdown: char count, font/size/bold, points, measurement source
  `gdi`/`gdiplus`/`floor` (+`+floor` when the char-cell floor won), DPI,
  pixels, floor points, final column range/pad/cap, any failed column-width
  reads) is returned in a new `Diag` array from `Invoke-GfixLogHighlight`
  and printed by both callers (`[GfixLog width]` in Mark.ps1, `[width]` in
  MarkGfixLog.ps1). Previously every failure path silently returned the
  fixed `HighlightColEnd`, which is indistinguishable from AutoWidth being
  off -- same lesson as the v2.10.2 `[rowinfo]` diagnostics.
- `Mark.ps1` / `MarkGfixLog.ps1` now `Add-Type System.Windows.Forms`
  (try/catch, warn-only) for the TextRenderer tier; a failed load just
  drops that tier.
- `Tests\Test-ExcelHelpers.ps1`: unit tests for `Get-TextCellUnits`
  (empty/null, ASCII, trailing space, full-width A/katakana/kanji,
  halfwidth katakana, mixed, ASCII Command: line = len/2 units).

### Notes
- No new config fields; `GfixLog.AutoHighlightWidth` / `HighlightPadCols` /
  `Replace.GfixLogFontName/Size` drive it exactly as before.
- Pure logic verified under portable pwsh 7 in this dev env
  (`Tests\Test-ExcelHelpers.ps1` 19/19 green; full `Run-Tests.ps1` parse
  check clean -- the 2 EvidencePlan failures are pre-existing Linux
  path-separator artifacts, confirmed identical on the unmodified tree).
  The GDI/GDI+/COM paths are static-checked only -- confirm the measured
  width on an office PC and read the new `[GfixLog width]` line: `Source`
  says which renderer answered and `dpi` exposes the display scale that
  used to corrupt the math.

## 2026-07-06 - GIFT_MQ row-info: fall back to the "Number of records" header when the per-record regex misses (v2.10.4)

### Fixed
- Real-world `.ocr.txt` dumps (v2.10.3) showed the actual bug behind the
  wrong GIFT_MQ box position: OCR read `Number of records 1` cleanly (the
  `en-US` pass), yet `ConvertFrom-MqPageText`'s strict 9-field-per-line regex
  never matched the record line itself (mangled by OCR spacing/noise), so
  `ConvertTo-MarkMqRowInfo` returned `$null` for every one of 4 correls in a
  workbook -- all four were genuine 1-record correls, so the box should have
  shifted up one row (`(1-2) * 63.8 = -63.8`) but instead silently kept the
  row-2 baseline, landing on the wrong spot.
- `ConvertFrom-MqPageText` already extracts this same header into
  `$parsed.NumRecords` (`Number of records\s+(\d+)`), it just was never
  consulted as a fallback. `ConvertTo-MarkMqRowInfo` now falls back to it
  whenever the per-record row match comes up empty: individual records can't
  be told apart without a parsed row, so this assumes the target -- always
  the LAST/newest one per this project's convention -- is the final row
  (`NumRecords`). `NumRecords = 1` is the common, fully unambiguous case:
  there is only one row to point at, no assumption needed. Source is tagged
  `<tier>-header` (e.g. `ocr-header`) so it's visible in `[ROW ]`/`[rowinfo]`
  output that this came from the header count rather than a matched record.
- This applies to both the `.txt` and `ocr` tiers (they share this parsing
  tail); the `.mqrow.json` sidecar tier is unaffected (already exact, since
  it is computed live at snap time).

## 2026-07-06 - GIFT_MQ OCR tier: dump reconstructed rows for debugging (v2.10.3)

### Added
- `Get-MarkMqRowInfoFromOcr` now writes every reconstructed row from both
  `en-US` and `ja` to `<WorkDir>\snap\GIFT_MQ\<correl>.ocr.txt` (labeled per
  language, including a `FAILED` marker if a language's OCR call threw), and
  prints a `[rowinfo] ocr: dumped reconstructed rows to ...` line pointing at
  it. Same idea as SendVsGift Stage 2's per-correl `_ocr.txt` dump.

### Notes
- First real-world OCR run (v2.10.2) showed OCR genuinely reading
  substantial text (`en-US`: 17 rows/506 chars, `ja`: 16 rows/593 chars) yet
  matching 0 MQ records -- with counts alone there was no way to tell whether
  that's mostly non-table page furniture (title/buttons/column headers)
  crowding out the ~4 real record lines, a single OCR misread breaking
  `ConvertFrom-MqPageText`'s strict anchored regex (it has no fuzzy-match
  tolerance, unlike `Compare-SendRecordCheck`), or a row-reconstruction
  tolerance issue splitting one real record line into several fragments.
  This dump makes that diagnosable directly from the actual text instead of
  guessing further regex/tolerance changes blind.

## 2026-07-06 - GIFT_MQ row-info fallback chain: diagnostics on every failure path (v2.10.2)

### Fixed
- All three GIFT_MQ row-position fallback tiers (`Get-MarkMqRowInfoFromSidecar`
  / `-FromArchivedText` / `-FromOcr`, `Mark.ps1`) failed completely silently:
  a missing sidecar/`.txt`/PNG, a read/parse error, or an OCR engine failure
  each just returned `$null` with a bare `catch {}`, so an operator seeing the
  `[WARN] ... row info unavailable` line had no way to tell which of the 3
  tiers were tried or why each one came up empty (first real-world report:
  all 4 correls in a workbook hit the WARN with zero visibility into whether
  it was a missing sidecar, a missing archive, or a failed OCR call).
- Every failure path now prints a `[rowinfo]` diagnostic line before falling
  through to the next tier: sidecar/txt report the exact path checked and
  whether it was missing/unreadable/had no matching row; the OCR tier
  additionally reports the PNG path, and on a successful OCR call, the actual
  recognizer language used, its text-read strategy, and the character count
  it extracted -- so it's visible when the English pack isn't installed and
  `Get-WinOcrEngine` silently fell back to the user-profile (e.g. Japanese)
  recognizer, or when OCR ran but read too little/garbled text to match any
  MQ record row.

### Added
- The OCR tier (`Get-MarkMqRowInfoFromOcr`) now applies the two lessons
  already learned the hard way in the SendVsGift Stage 2 OCR work (see
  `docs/SendVsGift.md` "Troubleshooting: OCR reads nothing"), since MQ's raw
  page text is the same shape of problem (9 TAB-separated fields per
  record): (1) the recognizer can fragment ONE wide row into SEVERAL OCR
  "lines", so it no longer trusts the engine's own line breaks -- it reuses
  `SendMetadata.ps1`'s `ConvertTo-SendRowLines` to reconstruct true rows by
  re-clustering word boxes on vertical position; (2) the `ja` recognizer
  garbles ASCII digit runs on this font family while `en-US` reads them
  cleanly (and vice versa), so it now OCRs with BOTH `en-US` and `ja` and
  pools every reconstructed row from both languages before parsing -- a
  garbled read from one language just fails to match and is ignored, a clean
  read from the other produces the hit. `[rowinfo]` now prints one line per
  language attempted (engine, strategy, row count, char count).
- No behavior change to box placement -- purely additive/improved fallback
  logic and console output.

### Changed
- `Mark.Boxes.GIFT_MQ.RowHeight` changed from `0` (disabled) to `63.8`, so the
  v2.10.0 row-position feature is now active by default instead of needing a
  manual opt-in. Measured via `Probe-Shapes.ps1` against a real J4 evidence
  workbook: two 2-record correls (`JIDSM01S`, `JIGPM01S`) had their
  hand-verified mark box at `OffsetY` 177.0 / 176.9 (matching the existing
  `BaseRow=2` default almost exactly), while a 1-record correl (`JIGPMA1S`)
  landed at `OffsetY` 113.2 -- one row up, a delta of 63.75, essentially
  identical to the box's own `Height` (63). `BaseRow` stays `2`.
- `verify_config.example.json` updated to match.

### Notes
- **Per-work-folder `verify_config.json` overlays do NOT auto-pick this up.**
  `Mark.Boxes.<folder>` is an array; `Merge-ConfigHashtable`
  (`ConfigOverlay.ps1`) replaces arrays wholesale rather than merging their
  contents, and `-Phase InitConfig` repair treats a whole array as one atomic
  schema field -- so a `verify_config.json` that already has a `GIFT_MQ`
  entry (from any earlier `InitConfig` run) keeps overriding this default,
  `BaseRow`/`RowHeight` included, until the operator adds those two keys to
  that JSON entry by hand (or deletes the `GIFT_MQ` key from their overlay's
  `Boxes` object to fall through to this default).
- Still static-checked only in this dev environment -- confirm the box lands
  correctly on a real 1-record and 3+-record correl on an office PC, and
  re-measure via `Probe-Shapes.ps1` if it looks off (a 3+-record sample would
  further validate the linear `RowHeight` assumption beyond the 1-vs-2-record
  data point above).

## 2026-07-06 - GIFT_MQ Mark: row-position aware red-box placement (v2.10.0)

### Added
- `Mark.Boxes.GIFT_MQ` (and any other box, opt-in) can now carry `BaseRow` /
  `RowHeight` keys. Some correls show only 1 MQ record on the transfer-status
  page, most show 2, and some show 3 or more retries; the red box must always
  land on the LAST/newest record for that correl (the same row
  `Test-MqRecord`'s newest-wins verdict already picked), not always the row
  the box's fixed `OffsetY` was originally calibrated against (`BaseRow`,
  default 2 = the common 2-record case). When `RowHeight` (points, > 0) is
  set, `Mark.ps1` adds `(actualRow - BaseRow) * RowHeight` to `OffsetY`
  before drawing.
- The actual row index + total record count are resolved through a 3-tier
  fallback chain, every tier reading only files already saved under
  `WorkDir\snap\<folder>\<correl>.*` (no Excel/Edge involved, never blocks
  Mark):
  1. `<correl>.mqrow.json` -- a new sidecar `MqSnap.ps1` writes right after
     the F2 verdict, using the exact same pure `Get-MatchedRowIndex`
     (`SnapVerify.ps1`) that already backs the M5/F5 `.loc.json`
     pixel-localisation sidecar, but decoupled from that feature's pixel
     geometry calibration (`SnapVerify.Localize.MqRow1Top/RowHeight/...`) --
     this path only needs the row's parse-order position, not its pixel
     rect, so it works with zero extra config.
  2. A re-parse of the archived Ctrl+A page capture `<correl>.txt` (already
     saved whenever `SnapVerify.SaveText` is on) via the same pure
     `ConvertFrom-MqPageText` + `Get-MatchedRowIndex` (`Expected = $null` ->
     newest-overall, matching the common `SnapVerify.TimeCheck = $false`
     default), for evidence snapped before this feature existed or where the
     sidecar is otherwise missing.
  3. Last resort: English Windows OCR (`OcrWindows.ps1`) of the source PNG,
     parsed the same way -- MQ records are ASCII/numeric, a reasonable OCR
     target. Only reached when both the sidecar and the archived `.txt` are
     missing.
- New `Mark.ps1` functions: `Get-MarkMqRowInfoFromSidecar`,
  `Get-MarkMqRowInfoFromArchivedText`, `Get-MarkMqRowInfoFromOcr`, the
  dispatcher `Get-MarkMqRowInfo`, and shared parser `ConvertTo-MarkMqRowInfo`.
  `Mark.ps1` now dot-sources `SnapVerify.ps1` and `OcrWindows.ps1` (both
  already no-`param()` dot-source-safe) for this chain; either file being
  missing only disables its own fallback tier and never blocks marking.

### Notes
- Ships with `RowHeight = 0` on `GIFT_MQ` in both `VerifyConfig.psd1` and
  `verify_config.example.json`, which keeps the legacy fixed-offset behavior
  byte-for-byte unchanged until an operator sets a real value. The actual
  row-to-row point spacing (and whether `BaseRow = 2` needs adjusting) must
  be measured on an office PC -- e.g. via `Probe-Shapes.ps1` against a sample
  evidence workbook with a 1-record or 3-record correl -- since this dev
  environment has no Windows/Excel/MQ page access to derive it. Static-
  checked only: confirm the `.mqrow.json` write, the `.txt`/OCR fallback
  parsing, and the resulting box position on an office PC once `RowHeight`
  is set.

## 2026-07-06 - Mark image-match anchor-only sizing + snap-time template-hit sidecar; DfSnap window default; FillCheckSheet write verification (v2.9.31)

### Fixed
- **A `Mark.Boxes` entry with `Template` sized the drawn red box from the
  template crop's own pixel dimensions whenever a match was found, ignoring
  the box's own `Width`/`Height`.** This was fine for a box whose whole point
  IS the crop (`GIFT_noGfixfile`'s `StampImage`), but broke as soon as
  `Template` was added to an existing fixed-size box (`GIFT_Jenkins`/
  `GFIX_Jenkins`, previously `Width=288.8`/`Height=18.8` via
  `OffsetX`/`OffsetY`): the box shrank to the anchor crop's own size instead
  of staying `288.8x18.8`. `Find-MarkBoxByImage` (`Mark.ps1`) now checks for
  `Width`/`Height` on the box first -- when either is present, `Template`
  only supplies the anchor (top-left corner, still shiftable via
  `PadX`/`PadY`) and the configured `Width`/`Height` are used as-is; a box
  with neither keeps the old crop-derived-size behavior.
- **`FillCheckSheet.ps1`'s cell writes could fail silently.** Each of the 6
  written columns (date + 5 others) was a bare `try { ... } catch {}` around
  `.Value2 = ...` with no readback, so an exception or a silent no-op write
  was swallowed with zero diagnostic and the run still printed `[OK] wrote N
  row(s)` even when column B (date) came back blank. New `Set-CellChecked`
  writes then reads back every cell and compares; a mismatch/exception is now
  logged (`[WARN] <label> write failed/did not verify (row N)`) and the final
  `[OK]` summary + `status\progress.jsonl` event reflect actual outcomes.

### Added
- **Snap-time template-hit sidecar.** `JenkinsSnap.ps1` (GiftJenkins /
  GfixJenkins / GiftJenkinsNoFile) accepts `-MarkBoxes`
  (`Config.Mark.Boxes[<folder>]`, threaded from `VerifyTool.ps1`) and, right
  after each screenshot is saved, runs the same `Locate-ByImage.ps1` match
  for every box carrying a `Template` key against the page as just captured,
  caching hits to `snap\<folder>\<correl>.tplhit.json`
  (`SnapLocalize.ps1`'s new `Write-MarkTemplateHits`). `Mark.ps1` reads this
  sidecar first (`Get-MarkTemplateHitFromSidecar`) instead of re-scanning the
  archived PNG, falling back to a live match when the sidecar is missing, has
  no entry for that box, names a different Template, or was recorded against
  a different-sized PNG. Zero new config fields -- reuses the existing
  `Mark.Boxes[<folder>].Template` opt-in key. Console output now tags which
  anchor source fired: `[MARK-IMG]`/`[STAMP-IMG]` each print `(sidecar)` or
  `(live)`.
- `Df.CaptureMode` default changed from `'region'` to `'window'`: DF snap now
  auto-fits the actual df.exe window size instead of assuming a fixed
  `1250x657` rectangle. The existing fallback to `region` on an invalid
  handle/rect is unchanged and is now the safety net rather than the
  default.

### Notes
- Static-checked only (no Windows/Excel/Edge in this dev environment):
  confirm the anchor-only sizing once real `GIFT_Jenkins`/`GFIX_Jenkins`
  `Template` crops are added, confirm the `.tplhit.json` sidecar is written
  and actually consumed (not silently falling back to live every run),
  confirm `window` capture mode against a real df.exe window, and confirm
  `FillCheckSheet`'s verified writes on the real check-sheet workbook.

## 2026-07-03 - Document versioning policy and automation guidance (v2.9.30)

### Added
- Added `docs/Versioning.md` to document the repository version format, bump decision guide, changelog expectations, and a lightweight PowerShell-friendly release automation approach.
- Linked the versioning policy from `README.md` and `CLAUDE.md` so operators and future LLM sessions can find the rule quickly.

### Notes
- Documentation-only change; no runtime behavior changed.

## 2026-07-03 - FillCheckSheet: on-disk prefix fallback + CheckSheetPath remembered (v2.9.29)

### Fixed
- **Check-sheet row F (review target) could list a filename that didn't
  match the real workbook.** `FillCheckSheet.ps1` built the filename purely
  from `Resolve-ExcelPrefix` (mapping row's legacy `Excel_Prefix` column,
  else `Workbook.ExcelPrefix`) with no way to tell "deliberately no prefix"
  from "nothing configured for this row" -- a row lacking both (e.g. an
  older-vintage mapping row, or a run where `Workbook.ExcelPrefix` isn't set
  in `verify_config.json`) silently produced a bare, unprefixed name even
  when every sibling row in the same run got the full
  `J4...(...)_<Excel_NAME>.xlsx` prefix from their own legacy column value.
  `FillCheckSheet.ps1` now takes a new `-EvidenceDir` param (default
  `<WorkDir>\evidence`, same convention as `DeliverFiles`/`DeliverMail`) and,
  whenever the resolved prefix comes back blank, looks up the real evidence
  file on disk (`Find-WorkbookByExcelName`) and recovers its actual prefix
  via the existing `Get-PrefixFromFilename` (previously only used by
  `DeliverFiles`'s bare-name-fallback path, read here in the opposite
  direction) -- so the check sheet always lists the filename that is
  actually on disk. `VerifyTool.ps1`'s `CheckSheet` dispatch threads
  `State.EvidenceDir` through.
- **`CheckSheetPath` prompted on every single run even after answering it.**
  The path prompt lived inside `FillCheckSheet.ps1` itself; whatever the
  operator typed there was a local variable that vanished when the script
  returned -- `VerifyTool.ps1`'s session save (`verify_session.json`) had
  already run *before* dispatch, so the answer was never persisted and the
  next run prompted again from scratch (this was silent: config *was* being
  read correctly, `CheckSheet.Path` was just genuinely unset). The prompt
  moved up into `VerifyTool.ps1`'s `CheckSheet` dispatch itself (mirroring
  the existing `DfExePath` first-run-prompt-then-remember pattern): State
  (CLI/session/menu `k`) wins, then config `CheckSheet.Path`, and only if
  both are empty does it prompt once and immediately save the answer to
  `verify_session.json` with a `[TIP]` pointing at `-Phase InitConfig` to set
  `CheckSheet.Path` permanently instead.

### Added
- The check-sheet row preview (`Show-Plan`) now prints column B's date
  (`yyyy/MM/dd`) alongside each planned row, so the operator can see the
  date that's about to be written (today, real date value, format mirrored
  from the row above -- unchanged behavior, equivalent to pressing Ctrl+;)
  without having to inspect the workbook directly.

### Notes
- Static-checked only (no Windows/Excel in this dev environment); confirm
  the on-disk prefix fallback and the CheckSheetPath remember-once flow on
  an office PC.

## 2026-07-03 - InitConfig: fix GetNewClosure() losing ConfigOverlay.ps1 functions (v2.9.28)

### Fixed
- `-Phase InitConfig`'s repair/full-snapshot writer (`$writeOverlay`, built
  via `{...}.GetNewClosure()` so it can be invoked from more than one call
  site with its own captured `$dest`/`$dryRunFlag`/etc.) threw
  `Get-ConfigOverlayJson`/`Get-ConfigOverlayReadmeText : ... 用語 ... 認識
  されません` (`CommandNotFoundException`) on a real Windows/PowerShell run.
  `GetNewClosure()` snapshots *variables* from the defining scope into a
  detached session state, but does not carry over *functions* -- the two
  `ConfigOverlay.ps1` functions only exist because that file is dot-sourced
  at the top of `VerifyTool.ps1`, so they were unreachable from inside the
  closure.
- The same detachment silently dropped the script's top-level
  `$ErrorActionPreference = 'Stop'` override -- the closure's own scope
  falls back to PowerShell's true default `Continue` -- so the
  `CommandNotFoundException`, and a real follow-on `WriteAllText`
  `IOException` (verify_config.json open in another program), were both
  non-terminating: the run printed `[OK] wrote/updated work-folder config
  overlay` and `[OK] wrote config field guide` even though neither file was
  written correctly.
- Fix: capture `Resolve-ToolPath $Config 'ConfigOverlay'` (a plain string,
  which `GetNewClosure()` DOES preserve) and re-dot-source it at the top of
  the closure body -- this loads the functions into the closure's own
  scope regardless of where/when it is invoked -- and set
  `$ErrorActionPreference = 'Stop'` there too, so a real failure (e.g. a
  locked destination file) now stops and reports instead of limping on to
  a false `[OK]`.

### Notes
- Root-caused from an operator's actual console output on the first real
  Windows/PowerShell exercise of the InitConfig repair path -- confirms the
  repo's long-standing "static-checked only" caveat on InitConfig was
  hiding this. No PowerShell in this dev environment to re-run
  `Tests\Run-Tests.ps1`; confirm both the default (non-interactive) repair
  path and the `-Interactive` editor's save on an office PC, including the
  locked-file retry behavior in `Invoke-ConfigEditorSave`.

## 2026-07-03 - Mark.Boxes: StampImage image-recognition-only stamp (v2.9.27)

### Added
- **`StampImage`**: a new key usable alongside `Template` on any
  `Mark.Boxes` entry. When the `Template` crop matches on the source snap
  PNG (`Locate-ByImage.ps1`, the existing LockBits scan used by plain
  Template boxes), `StampImage` is inserted (native size, brought to front)
  at the matched+scaled location instead of a red rectangle.
- **Deliberately no fixed-offset fallback.** Unlike a plain Template box
  (falls back to `OffsetX/OffsetY` on a miss), a `StampImage` box only ever
  draws when the pattern is actually found -- a miss means nothing is
  inserted (`[SKIP-STAMP]` in the console), because for this use case "not
  found" is itself the correct, common outcome (no past-data file exists).
- New `ExcelHelpers.ps1` `Insert-PictureAtPointBringToFront` (raw
  point-coordinate insert, mirrors `Add-RedRectangle`'s shape but for a
  whole picture).
- Wired the default `Mark.Boxes.GIFT_noGfixfile` (previously `@()`) to
  `@{ Template = 'NoGfixHit.png'; StampImage = 'already_exists.png' }`.

### Notes
- This is a simpler, self-contained alternative to v2.9.26's
  `Mark.NoteStamps` for the same GIFT_noGfixfile past-data-hit annotation:
  it runs directly against the source snap PNG via live template matching
  and does not depend on `SnapVerify.Localize.Enabled` or a `.note.json`
  sidecar -- both are off/absent by default, which is why `Mark.NoteStamps`
  alone never actually fired in practice. Both mechanisms coexist without
  conflict (they trigger on different shape-metadata paths).
- Static-checked only (no Windows/Excel in this dev environment); needs a
  real `NoGfixHit.png` crop and `already_exists.png` added to
  `mark_templates/` before it does anything on an office PC.

## 2026-07-03 - Mark: NoteStamp images on verifyNote annotations (v2.9.26)

### Added
- **`Mark.NoteStamps`**: an opt-in way to insert a whole stamp image (e.g.
  `already_exists.png`, dropped into `mark_templates/`) next to a
  `verifyNote` annotation, alongside the existing red rectangle +
  `SnapVerify.NoGfixNoteColumn` text note. Keyed by the note's `Folder`
  value (the first `|` field of the AltText payload `EvidenceExecutor.ps1`
  stamps via `Set-ShapeMetadata 'verifyNote'`), so future note kinds can
  register a stamp later without touching `Mark.ps1`; only
  `GIFT_noGfixfile` (the F4/M6 past-data hit) is wired today. Each entry:
  `@{ Image; Column; RowOffset }` (default column `AF`, `RowOffset=0`).
- Reuses the pixel rect already carried in the `verifyNote` payload (from
  the snap-time `<correl>.loc.json` / `.note.json` sidecars, M5/M6) instead
  of re-scanning the source PNG for the orange Ctrl+F highlight band a
  second time at Mark time: the existing verifyNote block's scaled sheet-Y
  (`$top`) is fed through `Get-RowAtOrBelow` (with the same `-1`
  off-by-one correction `Get-PictureBottomRow` already uses, since
  `Get-RowAtOrBelow` returns the row *after* the target pixel) to get an
  Excel row, then the image is inserted at `(row + RowOffset, Column)`.
- New `ExcelHelpers.ps1` `Insert-PictureBringToFront` (same shape as
  `Insert-PictureSendToBack` but `ZOrder(0)`, since a stamp must sit on
  top of the base screenshot). The inserted picture is named
  `<NamePrefix>verifyNoteStamp_<correl>_0` so `Remove-MarkShapes` cleans
  it up on every idempotent Mark re-run like any other mark shape. A
  stamp failure (image not found, bad config) only warns; it never fails
  the surrounding verifyNote mark or blocks `isMarked`.
- Threaded `VerifyConfig.psd1` / `verify_config.json` -> `VerifyTool.ps1`
  (`-Mode Gift` only, mirroring `NoGfixNoteColumn`) -> `Mark.ps1
  -NoteStampConfig`. Documented in `mark_templates/README.txt` and
  `ConfigOverlay.ps1`'s InitConfig readme text.

### Notes
- Static-checked only (no Windows/Excel in this dev environment); confirm
  the row math and stamp placement on an office PC once `already_exists.png`
  is added to `mark_templates/`.

## 2026-07-02 - DeliverFiles unzip delivery, Review -J4 option, config field merge + _SCHEMA repair + walker rework (v2.9.25)

### Added
- **DeliverFiles: DATA unzip subfolders are delivered too** (operator item 1).
  DfSnap's isZip rows (v2.9.24) extract zips into `work\DATA\GIFT\unzip` /
  `work\DATA\GFIX\unzip`; DeliverFiles now copies those extracted files
  (matched per correl id, same `<correl>*` filter as the plain DATA files)
  into the matching `J4\DATA\GIFT\unzip` / `J4\DATA\GFIX\unzip` folders,
  creating the J4 `unzip` subfolder on first delivery. The unzip specs are
  optional: a work folder without them prints nothing. `-Backup` covers the
  overwritten J4 unzip files exactly like the plain DATA files.
- **Review phases: `-J4` / menu option `j4`** (operator item 2). All four
  review phases (`ReviewGift`/`ReviewGfix`/`ReviewDf`/`ReviewEvidence`) can
  now open the DELIVERED J4 workbook instead of the work `evidence` copy:
  new `VerifyTool.ps1 -J4` switch and a `j4` toggle in the interactive
  option prompt (shown as `ReviewJ4`), wired like `f=Force`. When ON, the
  Review dispatch resolves the J4 folder via the new canonical
  `J4EvidenceDir` (legacy fields honored, see below) and passes it as
  `ReviewEvidence.ps1 -EvidenceDir`; a missing J4 folder fails with the
  where-to-set-it message instead of opening nothing. Saves land on the J4
  file; the local mapping's review bits update as usual.
- **Config: duplicated fields merged into canonical top-level fields**
  (operator item 3a). `Mail.EvidenceFolder` and `DeliverFiles.J4EvidenceDir`
  were the same J4 evidence folder typed twice, `Reviewer.Address` the only
  address -- now there is ONE top-level `J4EvidenceDir` (read by
  DeliverFiles, BackupJ4, DeliverMail's body path `{2}`, and Review `-J4`)
  and ONE top-level `Address` (DeliverMail's To). New pure helpers
  `Get-ConfigJ4EvidenceDir` / `Get-ConfigReviewerAddress`
  (`ConfigOverlay.ps1`, unit-tested): a non-empty legacy field still WINS so
  existing configs behave unchanged, but InitConfig no longer generates the
  legacy duplicates, and a legacy value migrates into the canonical field
  when the snapshot is (re)generated (the live runtime config is never
  mutated by the scrub). `VerifyConfig.psd1`, `verify_config.example.json`,
  the editor groups (`path`/`mail` now list `J4EvidenceDir`/`Address`), the
  overlay README text, and every "set DeliverFiles.J4EvidenceDir or
  Mail.EvidenceFolder" error message were updated to the new field names.

### Fixed
- **InitConfig walker was broken: a group collapsed into ONE bogus
  "System.Object[]" field.** `Expand-ConfigWalkPath` / `Get-ConfigWalkLeaves`
  (v2.9.22) returned `,$array`; every caller wraps the call in `@(...)`, and
  that combination NESTS (exactly the PS 5.1 convention documented since
  v2.8.1: shared functions must return plain arrays). So `w` never actually
  walked the group's fields -- it prompted once for a stringified array and
  wrote garbage keys if answered. Both functions now return plain arrays;
  verified end-to-end with a scripted-input harness (group `wbs` walks its
  real 3+ fields, set/keep/delete all land on the right paths). This was
  found by actually RUNNING the editor under PowerShell 7 in this dev
  environment (portable pwsh), which previous static-check-only sessions
  could not do.
- **InitConfig repair dumped the whole snapshot into a sparse file**
  (operator item 3b: "the repair seems to repeat the whole file").
  v2.9.24's repair added every default field missing from the operator's
  file -- it could not tell "field the tool gained since" apart from "field
  the operator deliberately left out", so a sparse hand-written overlay
  ballooned into the full snapshot on the first repair. Snapshots now carry
  a `_SCHEMA` metadata key (the dotted field-path inventory at write time;
  stripped at runtime by the existing `Remove-ConfigMetadataKeys`, excluded
  from the group walker), and `Update-ConfigOverlayData` was reworked around
  it: repair appends ONLY default fields that are NOT in the file's stamp
  (i.e. genuinely new to the tool), never re-adds a field the operator
  deleted (stamp keeps it "known" via a union refresh), and a stamp-less
  file (older InitConfig or hand-written) is only STAMPED on its first
  repair -- values untouched, nothing added, with a console note pointing at
  `f=Force`/`i=Interactive` for the full field set. New pure
  `Get-ConfigSchemaPaths` + path helpers in `ConfigOverlay.ps1`; repair
  tests rewritten in `Tests\Test-ConfigOverlay.ps1` (stamp-less, stamped,
  deleted-field, second-run-idempotent cases).

### Changed
- **InitConfig `-Interactive` walker + save flow reworked to the operator's
  spec** (operator item 3c). `w` now runs a walk LOOP: pick a group (number
  or key), walk its fields one at a time (`Enter`=keep, value=set,
  `-d`=delete -- `-del` still accepted, delete confirm is a light `y/N` since
  nothing is on disk yet, `q`=stop), and when the group finishes the prompt
  offers the NEXT group / `s`=save / Enter=back to the menu instead of
  silently returning. Saving now happens INSIDE the editor through a writer
  callback (`Invoke-ConfigEditorSave`): a failed write (typically
  verify_config.json open/locked in another program) no longer crashes the
  phase and loses every edit -- the operator is told to close the file, then
  `r`=retry the write, Enter=back to the menu with all edits kept,
  `q`=quit discarding. The editor tracks unsaved edits and warns on `q`.
  The same writer serves the non-interactive paths (repair/Force/dry-run),
  so backup/README behavior is identical everywhere.

Verified in this dev environment under portable PowerShell 7: full
`Tests\Run-Tests.ps1` (all suites green; the 2 EvidencePlan "\\ vs /" fails
are Linux path-separator artifacts), an end-to-end InitConfig
snapshot/repair/overlay-load smoke, a scripted-input harness driving the
editor's walk loop + `-d` delete + save-retry + dirty-quit paths, and a real
`DeliverFiles -SkipExcel` run against a fake work folder (unzip files landed
in `J4\DATA\...\unzip`, sources untouched). Excel COM paths (the `-J4`
review open) are static-checked only -- confirm on an office PC, and re-run
the suite once under Windows PS 5.1.

## 2026-07-02 - GFIX log font size + highlight measurement fixes, InitConfig repair mode, DfSnap isZip unzip-compare (v2.9.24)

### Added
- **ReplaceGfix: forced font SIZE on the pasted GFIX log** (operator bug 1:
  "all text should be MS Gothic and size 11"). v2.9.23 forced only the font
  NAME; the size still came from the workbook default. `Write-LogLines`
  (`ExcelHelpers.ps1`) gained a `FontSize` parameter, threaded through
  `EvidenceExecutor.ps1`'s `Invoke-EvidencePlan -GfixLogFontSize` and a new
  `ReplaceEvidence.ps1 -GfixLogFontSize` parameter (default `11`). New
  `Replace.GfixLogFontSize` config field (default `11`; `0` leaves the
  workbook's default size untouched), threaded by `VerifyTool.ps1` and
  documented in `verify_config.example.json` + the ConfigOverlay README text.
- **InitConfig repair/update mode** (operator bug 3: "it always generates a
  new one; sometimes I just want to update it"). When `verify_config.json`
  already exists, a plain `-Phase InitConfig` run now REPAIRS it instead of
  rewriting a full snapshot: the operator's file is kept exactly as-is
  (values untouched, a sparse hand-written file stays sparse) and only
  config fields the tool gained since the file was written are added -- each
  added dotted path is listed on the console, and a run with nothing to add
  changes nothing. `f=Force` performs the old full-snapshot regenerate
  (loaded values still survive via the merge; a `.bak` is kept either way).
  A file that fails to parse is NOT touched (explicit error + hint instead).
  New pure `Update-ConfigOverlayData` in `ConfigOverlay.ps1` (recursive
  add-missing-keys; existing scalars/arrays kept wholesale, matching the
  runtime merge semantics), unit-tested in `Tests\Test-ConfigOverlay.ps1`.
- **DfSnap: isZip rows compare the UNZIPPED data** (operator bug 4). Rows
  whose mapping `isZip`/`isZIP` flag is `1` hold ZIP archives on both sides;
  df.exe used to be launched on the two zip binaries, which compares
  nothing meaningful. DfSnap now extracts each side's zip into
  `DATA\GIFT\unzip` / `DATA\GFIX\unzip` (named after the correl id, same
  convention as SendVsGift's `data\unzip`) and launches df.exe on the two
  extracted files. Zip discovery + entry selection mirror SendVsGift's
  proven logic (exact `<correl>.zip` first, then prefixed zips, then any
  readable prefixed file; entry matched by name/basename, else the single
  entry). A flagged row with no readable zip on one side falls back to the
  plain data file with a warning; an extraction failure (e.g. multi-entry
  zip with no matching entry) fails the row (`progress.jsonl` `unzip/fail`)
  instead of silently comparing zips. Zip helpers were exercised end-to-end
  with real archives (PS 7 on the build box); the df.exe/capture flow is
  unchanged.

### Fixed
- **GFIX-log highlight width detection** (operator bug 2: "highlight longer
  or shorter"). Three fixes in `ExcelHelpers.ps1`:
  (1) `Get-TextPixelWidth` now measures with `StringFormat.GenericTypographic`
  (+ `MeasureTrailingSpaces`) -- plain `MeasureString` pads the result with
  layout margins (roughly an em of slack), which pushed the auto-sized
  highlight several grid columns LONGER than the text on narrow-column
  evidence sheets.
  (2) `Get-AutoHighlightColEnd` / `Invoke-GfixLogHighlight` gained optional
  `FontName`/`FontSize` overrides: `VerifyTool.ps1` passes
  `Replace.GfixLogFontName/GfixLogFontSize` into MarkGfix (`Mark.ps1
  -GfixLogFontName/-GfixLogFontSize`) and `MarkGfixLog.ps1`
  (`-FontName/-FontSize`), so the width is measured with the font the log
  was actually PASTED in -- previously a log pasted before the font forcing
  existed (or a failed cell-font read) was measured with the wrong font,
  making the highlight longer OR shorter than the real text.
  (3) The per-column width read uses `$ws.Columns.Item($c)` (canonical COM
  form used everywhere else in this codebase) instead of `$ws.Columns($c)`.
  Also fixed a latent case-collision in the new override code: PowerShell
  variables are case-insensitive, so the measurement locals are named
  `$measureFont`/`$measureSize` (NOT `$fontName`/`$fontSize`, which would
  silently overwrite the `$FontName`/`$FontSize` parameters).
- GDI+/COM paths (font-size assignment, typographic measurement on the
  JP-locale host, the column-width read) are static-checked only -- no
  Windows/Excel in this dev environment; parse checks + the pure unit tests
  (`Tests\Run-Tests.ps1`) run clean under PowerShell 7 on the build box
  (2 pre-existing `Test-EvidencePlan` failures are Linux path-separator
  artifacts only). Confirm the size-11 paste, the tightened highlight, the
  InitConfig repair flow, and the DfSnap unzip-compare on an office PC.

## 2026-07-01 - Mark image-recognition placement + GFIX highlight auto-width + forced log font (v2.9.23)

### Added
- `Mark.Boxes` entries can now add a `Template` key (PNG filename, resolved
  against `Mark.TemplateDir` then `mark_templates/`). When set, `Mark.ps1`
  tries `Locate-ByImage.ps1` (existing LockBits template matcher) against the
  original snap PNG for that folder/correl before drawing the red rectangle,
  scales the pixel hit to the inserted picture's on-sheet point size, and
  falls back to the existing fixed `OffsetX/OffsetY/Width/Height` box when
  there is no Template, the file is missing, or nothing matches -- this never
  blocks Mark. Per-box `Tolerance`/`PadX`/`PadY` overrides; console lines are
  tagged `[MARK-IMG]` vs `[MARK]` so a run shows which path was used. New
  `Mark.TemplateDir` / `Mark.ImageMatch.Tolerance` config; new (empty)
  `mark_templates/` folder with a `README.txt` explaining how to capture and
  wire up a reference crop per mark target.
- The GFIX-log yellow highlight (`-Mode Gfix` in `Mark.ps1`, and the
  standalone `MarkGfixLog.ps1`) now sizes itself to the Command: row's
  ACTUAL pasted text width by default, instead of always filling the fixed
  `HighlightColStart..HighlightColEnd` range -- a long Command: path no
  longer risks running past a short fixed range, and a short one no longer
  leaves an oversized highlight band. New `ExcelHelpers.ps1` helpers:
  `Get-TextPixelWidth` (GDI+ `MeasureString`), `Get-AutoHighlightColEnd`
  (walks the sheet's real column widths to find where the measured width
  lands, capped at the existing `HighlightColEnd` so this only ever tightens
  the old behavior, then pads by `GfixLog.HighlightPadCols`), and a pure
  `Get-ColumnsForWidth` (the column-accumulation math, unit-tested in the new
  `Tests\Test-ExcelHelpers.ps1`). New `GfixLog.AutoHighlightWidth` config
  (default `$true`; `$false` restores the old fixed-width behavior).
- `ReplaceGfix` now forces a fixed font on every pasted GFIX receive-log
  line. `Write-LogLines` (`ExcelHelpers.ps1`) gained a `FontName` parameter,
  threaded through `EvidenceExecutor.ps1`'s `Invoke-EvidencePlan` and a new
  `ReplaceEvidence.ps1 -GfixLogFontName` parameter defaulting to `'MS
  Gothic'` (the ASCII-typeable name Windows/Excel resolve to the Japanese
  fixed-width font "MS ゴシック" -- kept ASCII per this repo's source-encoding
  rule). New `Replace.GfixLogFontName` config field (blank leaves the
  workbook's default font untouched).
- `VerifyTool.ps1` threads all three new settings from
  `VerifyConfig.psd1`/`verify_config.json` into the `Mark`/`MarkGfixLog`/
  `Replace` dispatch blocks. `ConfigOverlay.ps1`'s `excel` editor group now
  includes `GfixLog`; its README text and `verify_config.example.json`
  document the new fields.
- Pure logic (`Get-ColumnsForWidth`) is unit-tested. The COM/GDI+ paths
  (image-match pixel->point scaling, live column-width reads, font
  assignment, the 96-DPI pixel<->point assumption in `MeasureString`) are
  static-checked only (no Windows/Excel in this dev environment) -- confirm
  image-match placement (needs real template PNGs captured first per
  `mark_templates/README.txt`), the auto-width highlight, and the forced log
  font on an office PC.

## 2026-07-01 - InitConfig editor: grouped field walker (v2.9.22)

### Added
- `-Phase InitConfig -Interactive` gained a `w` editor command: pick a group
  (same number/key lookup as the other commands) and it walks every editable
  field in that group one at a time -- current value shown, Enter=keep,
  a value=set it, `-del`=delete it, `q`=stop walking -- so the operator never
  has to type a dotted JSON path (e.g. `Mark.Boxes.GIFT_HM.0.OffsetX`) from
  memory. `v`/`e`/`d` are unchanged for operators who already know the exact
  path they want.
- New `VerifyTool.ps1` helpers: `Expand-ConfigWalkPath` (recurses a group's
  top-level paths to leaf fields -- hashtables always recurse per key; an
  array recurses by index only when every element is a hashtable, e.g.
  `Mark.Boxes` entries / `PhaseOrder` rows, otherwise the whole array is one
  atomic edit leaf), `Get-ConfigWalkLeaves` (expands a group's `Paths`,
  including `all`'s `*`), and `Invoke-ConfigFieldWalk` (the per-field prompt
  loop; edits apply in memory and still require `s` + typed `YES` to save).
  Walking a group with more than 30 fields (e.g. `phase`) asks for
  confirmation first.
- Updated the editor's intro text, `InitConfig` phase notes,
  `ConfigOverlay.ps1`'s `_README` snapshot text + `Get-ConfigOverlayReadmeText`,
  `README.md`, and `verify_config.example.json` to document `w` alongside
  `v`/`e`/`d`.
- Static-checked only (no PowerShell in this dev environment) -- confirm the
  walk prompts and the >30-field confirmation on an office PC.

## 2026-07-01 - DeliverFiles rework + config error messages + GfixLogDownload naming (v2.9.20)

### Changed
- `DeliverFiles.ps1` no longer has a "move" mode: it only ever copies, and
  source DATA\GIFT/GFIX files are never deleted. `-MoveData` (and the
  `DeliverFiles.MoveData` config key) is removed, replaced by `-SkipExcel` /
  `-SkipData` (default: copy both the evidence Excel and the DATA files) and
  a new `-Backup` switch that copies any J4 file about to be
  overwritten/removed into `J4EvidenceDir\_bak\<name>.<timestamp>.bak` first.

### Fixed
- DeliverFiles now falls back to the bare `Excel_NAME` when the local
  evidence workbook predates the configured `Workbook.ExcelPrefix`, and
  always names the J4 copy with the resolved prefix regardless of the
  source's on-disk name.
- DeliverFiles detects a same-stem full-width-ASCII duplicate already sitting
  in J4EvidenceDir (e.g. a workbook name typed with `０` instead of `0`) and,
  after asking the operator to confirm, removes it so only the half-width
  (work-folder) copy remains -- `-Backup` saves the removed file first.
- `VerifyTool.ps1`'s DeliverFiles dispatch now fails with an explicit
  "set DeliverFiles.J4EvidenceDir (or Mail.EvidenceFolder) in
  verify_config.json -- run -Phase InitConfig" message instead of only the
  child script's bare `-J4EvidenceDir is required` error; DeliverMail's
  missing-reviewer error and phase notes got the same "where to configure
  this" treatment, including a full breakdown of every `{n}` placeholder in
  `Mail.SubjectTemplate` / `Mail.BodyLines`.
- `GfixLogDownload.ps1` downloaded logs as
  `<JobNo>_<timestamp>_<originalName>.log`, but GoAnywhere itself names the
  file after the job number, so the first and last filename fields were
  always identical. The job number is now replaced with the mapping
  `JOB_NAME`(s) that needed it (joined with `+` for the duplicate-IF_NO
  case), falling back to the job number only if no JOB_NAME is known.
- Static-checked only (no Windows/Excel/Edge in this dev environment) --
  confirm the full-width J4 prompt, the prefix fallback, and the new
  GfixLogDownload log filenames on an office PC.

## 2026-06-30 - Align same-name workbook open fix + sheet-order preserve (v2.9.17)

### Fixed
- **Align reported every sheet "missing in J4" with an empty "J4 sheets present:"
  list** (the diagnostic added in v2.9.16 made the cause visible). Real root
  cause: Excel cannot have two workbooks with the **same leaf filename** open in
  one instance, and the J4 baseline shares the *identical* filename with the work
  evidence by design (`J4...(_<Excel_NAME>).xlsx` in both the evidence dir and the
  J4 baseline dir). `Open-Workbook` opens the work file first; the second
  `Workbooks.Open` for the same-named J4 file, with `DisplayAlerts = $false`,
  has Excel **suppress** the "can't open two workbooks with the same name" error
  and return `$null` (no throw). `$j4Wb` was therefore `$null`; Align has no
  `StrictMode`, so every `Get-AlignSheetMatch $j4Wb ...` iterated `$null.Worksheets`
  (zero iterations, no error) -> all sheets "missing in J4", empty present-list,
  no `[FAIL]`. The one workbook that DID sync (`JJODWDB2`) only worked because its
  J4 file carried a full-width `W` (`..._JJODＷDB2.xlsx`), so its leaf name
  *differed* from the work file -- the only pair Excel could open at once. This
  affected every same-named pair, i.e. the normal case. `Align.ps1` now opens the
  J4 baseline through a new `Open-J4Safely`: when the J4 leaf name collides with
  the work leaf name (or it is literally the same path), it copies the J4 file to
  a uniquely-named temp file (`%TEMP%\verify_j4_<guid>.xlsx`) and opens that copy,
  so both workbooks can be open simultaneously; the temp file is removed in the
  `finally`. A `$null`-workbook guard now `throw`s (-> `[FAIL]`) instead of
  silently misreporting.
- **Synced sheets landed in the wrong position** ("the order is chaos"):
  `Sync-Sheet` deleted the work sheet first, then re-inserted the J4 copy by
  index -- but the delete shifted every later sheet's index, so the copy landed
  in the wrong slot. It now copies the J4 sheet to immediately **before** the work
  sheet and then deletes the work sheet, keeping the synced sheet at its original
  position.
- COM/Excel path -- static-checked only (no PowerShell/Excel in the cloud build
  env); confirm on an office PC + Excel 2019.

## 2026-06-30 - Align sheet-name miss diagnostics + tolerant match (v2.9.16)

### Fixed
- **Align still reported sheets "missing in J4" for some workbooks** even after
  v2.9.15. `[WARN] sheet missing in J4` fires only when the J4 workbook has no
  sheet whose name is byte-for-byte identical to the canonical label
  (`Get-SheetByName` uses an exact `-eq`). Two distinct causes were
  indistinguishable because the code never showed what the J4 file actually
  contained: (a) the opened J4 file is a blank/stale template (the exact-named
  file on disk was picked over a prepared but differently-named one -- note the
  one workbook that DID sync, `JJODWDB2`, was resolved through the full-width
  filename fallback), or (b) the sheet name differs only by stray whitespace or
  full-width vs half-width ASCII (e.g. full-width `GIFT`/`GFIX`).
  `Align.ps1` now matches sheet names with a new `Get-AlignSheetMatch`: exact
  match first, then a normalized compare (trim + `Convert-FullWidthAsciiToHalfWidth`
  from `WorkbookResolver.ps1`), so width/whitespace mismatches resolve. When a
  sheet is still missing it now prints the J4 file leaf and the full list of
  worksheet names present in that J4 file, so the operator can see immediately
  whether it is the wrong file (cause a) or a naming variant (cause b).
  COM/Excel path -- static-checked only; confirm on an office PC + Excel.

## 2026-06-30 - Align Host->Open default + J4 no-content guard + picture-aware diff (v2.9.15)

### Fixed
- **Align always failed (every sheet "missing in J4")**. Root cause: with
  `Align.HostSystemTypes` unset, `Get-MigrationType` returns `Unknown`, and the
  `Unknown` scope was the three *receive* sheets -- but J4 baselines never carry
  recv sheets (operator evidence), so every sheet was reported
  `[WARN] sheet missing in J4` and nothing synced. Align now defaults an
  unclassifiable migration to `Align.DefaultMigrationType` (default `HostToOpen`):
  it deletes the work `Soushin data` / `GIFT send result` / `GFIX send result`
  sheets and copies J4's, in order. Set `HostSystemTypes` for true per-row
  classification; the warning now prints the actual FROM_sys/TO_sys literals.
- **Picture sheets compared as "same" and were skipped.** `Compare-SheetGrid`
  only diffs cell values, so the image-based send-data sheet (few/no cell values)
  read identical even when the work copy had no screenshot. New picture-aware
  `Compare-AlignSheet` (in `AlignCompare.ps1`) also compares pasted-picture count
  (msoPicture / msoLinkedPicture / msoGroup) so the send-data sheet syncs, and an
  already-aligned sheet (equal grid + equal picture count) is correctly `[same]`
  and skipped.

### Added
- **J4 "no contents" guard.** Before replacing, Align checks the J4 sheet: the
  send-data sheet must hold >= 1 picture; the GIFT/GFIX send-result sheets must
  hold more than `Align.MinSendResultRows` (default 3) rows of text. A J4 sheet
  that is still a blank template is reported `[NO CONTENTS] ... replace skipped`
  and never overwrites the work evidence. New pure helpers `Get-AlignSheetKind`
  and `Test-J4SheetPrepared` (unit-tested in `Tests\Test-AlignCompare.ps1`);
  `Align.ps1` reads `PictureCount` / `TextRowCount` via the new COM
  `Get-SheetMetrics`. New config keys `Align.DefaultMigrationType` /
  `Align.MinSendResultRows`, threaded from `VerifyTool.ps1`. Pure logic +
  tests run via `Tests\Run-Tests.ps1`; the COM paths are static-checked only and
  need an office-PC + Excel run to confirm.

---

## 2026-06-30 - DfSnap df.exe path: configurable default + first-run prompt (v2.9.14)

### Added
- **`Df.DefaultExePath`** (`VerifyConfig.psd1`, default `C:\tools\DF\DF.exe`):
  the df.exe path the first-run prompt pre-fills (press Enter to accept). Also
  added to `verify_config.example.json`.
- **DfExePath is now remembered between runs.** VerifyTool persists the resolved
  df.exe path in `verify_session.json` (`DfExePath`) and loads it on startup, so
  the prompt fires only on the first DfSnap run and is silent afterward.

### Changed
- **DfSnap df.exe path resolution** is now CLI `-DfExePath` > `verify_session.json`
  > `Df.ExePath` > prompt(`Df.DefaultExePath`). `Df.ExePath` stays empty by
  default (= ask on first run); set it to a real path to lock it and never be
  prompted. The first-run prompt + persistence live in VerifyTool's DfSnap
  dispatch; `DfSnap.ps1` gained a `-DefaultExePath` param and pre-fills its own
  standalone prompt with it. (`VerifyTool.ps1` DfSnap dispatch + session load/save,
  `DfSnap.ps1` prompt, `VerifyConfig.psd1`/`verify_config.example.json` Df block.)
- COM/Excel parts are static-checked only (no PowerShell/Excel in the cloud env);
  confirm the prompt + session persistence on an office PC.

---

## 2026-06-29 - Snap TimeCheck menu toggle + -Add owner filter (v2.9.13)

### Added
- **`tc` menu toggle for the run-time window check** on the HM / MQ / Jenkins
  snap phases. Previously `SnapVerify.TimeCheck` could only be set in config; now
  the interactive options menu shows `TimeCheck : ON/off` and accepts `tc` to
  flip it per run. It seeds from `SnapVerify.TimeCheck` (still usually off) and
  threads `$State.TimeCheck` into the three snap dispatch blocks. (`VerifyTool.ps1`:
  `Get-PhaseOptionKeys`, `Show-PhaseNotes`, `Ask-RunOptions`, the option loop, the
  `$State` seed, and the GiftHmSnap/GfixHmSnap/GiftMqSnap/Jenkins dispatch.)
- **`Generate-HostOpenMapping -Add` now composes with the owner filter.** Explicit
  `JOB_NAME` / `Correl_ID_M` / `Excel_NAME` selectors are looked up in the WBS
  (col A) and dropped when their owner cell (col P) belongs to another operator.
  A JOB_NAME absent from the WBS is kept (a temp / not-yet-listed job the WBS
  can't judge) and reported in the warnings. The WBS-range `-Add` path already
  owner-filtered (Step C) and is unchanged.
- **New pure lib `OwnerFilter.ps1`** (`Test-OwnerMatch` + `Select-JobsByOwner`),
  unit-tested in `Tests\Test-OwnerFilter.ps1`. `Test-OwnerMatch` moved out of
  `Generate-HostOpenMapping.ps1` so both the WBS scan and the new `-Add` path
  share one implementation. The WBS scan glue (`Build-WbsJobOwnerMap`) is COM and
  static-checked only.

---

## 2026-06-25 - SnapVerify field fixes (time window, focus, NoGfix poll)

### Changed
- **Time-window check is now OFF by default** across HM / MQ / Jenkins snap phases.
  New `SnapVerify.TimeCheck` (default `$false`): detection still flags missing
  data / abends / missing files, but no run-time prompt or +-tolerance compare
  unless `TimeCheck = $true`. The window was mostly nice-to-have and the prompt
  was slowing every run.
- **NoGfix poll no longer waits out the full timeout.** `Wait-JenkinsPageReady`
  takes `-RequireTerm`; NoGfix passes `$false` so a loaded file-list page is
  "ready" as soon as it classifies as a Jenkins result (the correl is expected
  to be absent, so waiting for it always burned the whole `PollTimeoutSec` and
  re-read the clipboard ~20x per row).

### Fixed
- **Run-time input parsing.** `Resolve-SnapRunTime` now accepts time-only input
  `HH:mm:ss` / `HH:mm` (1- or 2-digit hour, anchored to today) in addition to the
  full `yyyy/MM/dd HH:mm[:ss]` forms, validates the input, and no longer lets a
  blank/garbage tolerance zero the default. New unit tests in `Test-SnapVerify.ps1`.
- **Edge focus after an operator prompt.** After the HM `o/n/s` ask (and the
  HM/MQ page-kind sentinel retry), the shell is foreground; the phase now
  `Switch-ToEdge` before continuing so subsequent keystrokes hit the Edge page
  instead of the CLI.
- **Stale NoGfix note sidecar.** A NoGfix row that now reads OK deletes any
  leftover `<correl>.note.json`, so a past-data annotation is not re-stamped once
  the file is gone.
- **MarkGift NoGfix note** uses the `過去分データー` label from `ProjectLabels.ps1`
  (was inlined `[char]`), and the `verifyNote` branch is guarded to `-Mode Gift`.

### Notes
- COM/SendKeys/Excel portions remain office-PC-validated only (no PowerShell/Excel
  in this container). Pure logic is covered by `Tests\Run-Tests.ps1`.

## 2026-06-22 - SnapVerify M6: NoGfix past-data annotation

### Added
- **NoGfix detection is wired into JenkinsSnap.** `GiftJenkinsNoFile` now uses the Jenkins page text with `Test-JenkinsFile -ExpectExists:$false`; unexpected files set `GIFT_noGfixfile_snap = 2` and can emit `<correl>.note.json` when localisation is enabled.
- **NoGfix note sidecars flow through Replace/Mark.** `ReplaceEvidence` stamps NoGfix pictures with `verifyNote` AltText from the note sidecar, and `MarkGift` converts screenshot pixels to Excel points, draws the red rectangle, and writes the past-data note to the configured column.

### Notes
- The COM/SendKeys/Excel portions still require office-PC validation; this container does not provide PowerShell/Excel.

# Changelog

Tracks iterations across Misaki's browser (work) ↔ IDE (home) workflow.
Bump the date heading whenever a new bundle is delivered.


## 2026-06-22 - JenkinsSnap: stop page-body click navigating into a queued job (v2.9.10)

### Fixed
- **Consecutive Jenkins screenshots opened a job page
  (`.../job/sc_str1_50_21_stop_appserver/`) and left the second file in a
  `TO_code` group unscreenshotted.** The per-row flow called `Click-PageBody` (a
  fixed `Left+150, Top+150` left-click) to focus the page before `Ctrl+F` and
  before the page-text read. On a Jenkins folder page `(150,150)` lands in the
  LEFT sidebar; when a build is queued the "Build Queue" (`実行予定のビルド`) widget
  shows a job hyperlink right there, so the click navigated Edge into that job.
  Every later capture in the group then shot the wrong page and the correl `Ctrl+F`
  found nothing. The queue widget only renders while something is queued, which is
  why this surfaced intermittently ("it didn't used to happen"). The `(150,150)`
  `Click-PageBody` was actually doing double duty -- focusing the page AND
  collapsing the previous row's `Ctrl+A` select-all so it was not captured in the
  next screenshot -- so it is replaced, not removed, with a new
  `Click-JenkinsPageCenter` that clicks the window centre. On these Jenkins pages
  the centre carries no hyperlink (confirmed by the operator), so the click is
  safe; it still clears the selection and focuses the page. (`Esc` does **not**
  clear an Edge text selection -- only a click does -- so an Esc-only fix would
  leave the select-all highlight in the following capture.) Used at both sites:
  before `Ctrl+F` in the per-row loop (clears the prior selection before the
  screenshot) and in `Get-JenkinsPageTextOnce` (document focus for the
  `Ctrl+A`/`Ctrl+C` read). Same approach as `MqSnap`'s `Click-MqPageCenter`;
  `MqSnap`/`HmSnap` keep their own clicks (MQ/HM pages are text, not hyperlink
  lists, so they have no navigation hazard).
- **Mojibake comment decorations removed repo-wide.** Box-drawing `─` (U+2500)
  and em-dash `—` (U+2014) characters used as comment rules render as garbage
  (e.g. `笏笏`) when Windows PowerShell 5.1 reads a no-BOM `.ps1` on the CP932
  JP-locale host. Replaced with ASCII across all 12 affected scripts:
  `JenkinsSnap.ps1` (645x), `MarkGfixLog.ps1` (267x), `Validate.ps1` (267x),
  `ExcelHelpers.ps1` (270x), `Mark.ps1` (182x), `ExcelSnap.ps1` (52x),
  `Generate-HostOpenMapping.ps1` (32x), `ReviewEvidence.ps1` (31x), `Common.ps1`
  (24x), `GfixLogDownload.ps1` / `ReplaceEvidence.ps1` (4x each) and
  `JenkinsDownload.ps1` (1x em-dash). Only the two decoration codepoints were
  changed (comment-only; BOM state and line endings preserved). Note: ~13 older
  scripts still carry *raw Japanese* comments (e.g. `# 正常終了`) -- a separate,
  larger `[char]` migration, not this box-rule garbage.


## 2026-06-19 - SnapVerify: ASCII-clean library + M5 pixel localisation (v2.9.9)

### Fixed
- **`SnapVerify.ps1` failed at runtime on the JP-locale host with
  `The variable '$script:SV_Abend' cannot be retrieved because it has not been
  set.`** The file still carried raw Japanese in ~18 comments (e.g. `# 正常終了`,
  `# 異常終了`, `# バッチ処理状況一覧 ...`). Windows PowerShell 5.1 reads a no-BOM
  `.ps1` with the system ANSI codepage (CP932), so the Shift-JIS misread of those
  multibyte comment bytes shifts tokenisation (the operator's stack even reported
  line 75 where the clean file has line 79 -- ~4 lines collapsed), which drops the
  top-of-file `$script:SV_Normal` / `$script:SV_Abend` constant assignments. Under
  `Set-StrictMode` (JenkinsSnap sets `Latest`; an interactive/profile strict mode
  hits the HM path too) reading the unset variable throws instead of returning
  `$null`. Replaced every non-ASCII character in `SnapVerify.ps1` (and the
  now-dot-sourced `Find-ActiveHighlightRow.ps1`) with ASCII per the project's
  ASCII-source rule -- the runtime Japanese values are still built from `[char]`
  code points, so behaviour is identical; the file now tokenises the same on every
  codepage. Same class of bug fixed for `Test-SnapVerify.ps1` in v2.9.8.

### Added
- **SnapVerify M5 -- object-data pixel localisation (F5).** New pure, unit-tested
  geometry/builder functions in `SnapVerify.ps1` produce a `<correl>.loc.json`
  sidecar (screenshot pixel rect) for the exact data row a verdict is about, so a
  later Mark pass can red-box it:
  - `Get-MatchedRowIndex` -- the 1-based **screen** row (parse order) of the row a
    verdict selects (newest-wins inside the time window, else newest overall),
    mirroring `Test-HmAbend` / `Test-MqRecord` so the box lands on the judged row.
  - `Get-RowPixelRect` -- HM/MQ fixed geometry `Row1Top + (n-1)*RowHeight` (same
    model as `Find-Abend.ps1`), in the cropped-PNG coordinate space.
  - `Get-JenkinsHighlightRect` -- turns a `Find-ActiveHighlightRow` orange-band
    `@{Top;Bottom}` into a rect.
  - `New-SnapLocRect` / `Save-SnapLocSidecar` -- build + persist the sidecar
    (`{ correl, source, rowIndex, x, y, w, h, imageWidth, imageHeight, created }`;
    `imageWidth` lets Mark scale pixel->point via `Shape.Width / imageWidth`).
  Tests added to `Tests\Test-SnapVerify.ps1` (HM/MQ row selection, geometry +
  crop offset + clamping + throws, highlight-band rect, sidecar JSON round-trip).
- **Localisation wiring (default OFF).** New non-pure glue `SnapLocalize.ps1`
  (`Write-SnapLocalize`) combines the pure geometry with System.Drawing image
  sizing + the Jenkins highlight scan; it swallows every error (returns `$null`)
  so it can never block snapping. `HmSnap.ps1` / `MqSnap.ps1` / `JenkinsSnap.ps1`
  dot-source it and write the sidecar after each verdict when enabled; VerifyTool
  threads `Config.SnapVerify.Localize` through as `-Localize`. New
  `SnapVerify.Localize` config block (`Enabled=$false` by default; HM/MQ geometry
  fields are 0 until calibrated with `Calibrate-HmGeometry.ps1`, so the leg is
  inert until an operator opts in -- Jenkins needs no geometry). M6 (NoGfix
  annotation: consume the sidecar via AltText -> Mark + AZ note) still remains.

### Notes
- Pure logic + tests run on Windows via `Tests\Run-Tests.ps1`. The COM/SendKeys/
  GDI+ wiring (`SnapLocalize.ps1` + the three snap scripts) is static-analysis
  only in the cloud build env and needs an office-PC + Excel run to confirm, and
  geometry calibration before the HM/MQ leg emits sidecars.


## 2026-06-19 - SnapVerify M4: HM abend detection + HmSnap migration (v2.9.8)

### Fixed
- **`Tests\Test-SnapVerify.ps1` failed to parse on the JP-locale host.** Three
  `Assert-Equal` messages embedded raw Japanese (`正常終了` / `異常終了`) *inside
  single-quoted strings*, and several comments carried raw CJK. Windows PowerShell
  5.1 reads a no-BOM `.ps1` with the system ANSI codepage (CP932), so a Shift-JIS
  lead byte swallowed the string's closing quote -- the string ran away and the
  parser only surfaced the failure much later as a bogus "missing terminator" at
  the `Set-EmptyRunTimeCells` block (hence the misleading line numbers). Replaced
  every non-ASCII character with ASCII per the project's ASCII-source rule. The
  assertions still compare against the `[char]`-built `$normal` / `$abend` values,
  so the test logic is unchanged.

### Added
- **SnapVerify M4 -- HM instant NG detection (F1).** `HmSnap.ps1` rewritten from
  the legacy bare `Import-Csv`/`Export-Csv` version to the modern stack:
  MappingStore (atomic writes), ProgressLog events, and the pure `SnapVerify.ps1`
  detection library -- while keeping HM's per-`TO_code` appl grouping (one HM page
  opened per appl). After the search it polls the page text (A2), classifies the
  page (`Get-SnapPageKind -Phase Hm`, sentinel A3), archives the Ctrl+A text as
  `snap\<Stage>_HM\<correl>.txt` (A1), screenshots, then `ConvertFrom-HmPageText`
  + `Test-HmAbend` decide the verdict (plan 2.3):
  - **ok** -> `<Stage>_HM_snap` = 1. The newest run inside the `Expected_Time` ±
    tolerance window ended normally; earlier in-window abends (a retried run)
    become "retried, last run ok" warnings.
  - **ng** -> = 2. The newest in-window run is an abend. `2` stays pending and is
    re-offered next run; an end-of-run NG summary lists them.
  - **ask** -> the operator decides `o`=OK(1) / `n`=NG(2) / `s`=skip(pending) /
    `q`=quit. Triggered when the correl has 0 rows, no rows inside the window, or
    an abend in no-time-check mode (plan 4.F1 steps 3-4).
  Out-of-window historic abends only warn (never auto-NG). A one-time batch
  run-time prompt (`Resolve-SnapRunTime`) fills empty `Expected_Time` cells on the
  pending rows (plan 2.2). Off-page kinds (OuterFrame/Empty/Unknown) stop and ask
  `r=retry / s=skip / q=quit`. Pending uses a local `Test-HmSnapDone` (done ==
  exactly `'1'`) so NG=`'2'` rows are not hidden. `SnapVerify.Enabled=$false`
  reverts HmSnap to pure screenshot (legacy behavior).
- **VerifyTool dispatch** for `GiftHmSnap` / `GfixHmSnap` now passes the
  `SnapVerify` (Enabled / ToleranceMinutes / SaveText / PollTimeoutSec /
  PollIntervalMs) and `ExpectedTime` (TimeColumn / TimeFormat) config, mirroring
  the `GiftMqSnap` block.

### Notes
- The v2.9.7 focus-safe pattern is preserved: per-row refocus is
  `Reset-FocusToBody` only; `Switch-ToEdge` runs once per appl (after the
  page-ready Read-Host) and inside the interactive branch (after
  `Bring-ShellToFront`) -- never blindly between rows.
- F1's pure functions (`ConvertFrom-HmPageText`, `Test-HmAbend`,
  `Get-SnapPageKind`) and their unit tests shipped in M1; this change is the
  COM/SendKeys wiring only. M5 (pixel localisation) and M6 (NoGfix annotation)
  remain. Authored in a Linux env without PowerShell/Excel -- run
  `Tests\Run-Tests.ps1` and smoke-test GiftHmSnap/GfixHmSnap on a copy before
  production use.


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
