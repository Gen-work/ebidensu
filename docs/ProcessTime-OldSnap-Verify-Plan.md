# ProcessTime old-snap 9→3 hand-verification — implementation plan

Status: PLANNED (design only). Hand-off spec for an implementing agent.
Author context: written after the mock-page benchmark confirmed the root cause
(2026-07). Read `CLAUDE.md` first for conventions; this plan assumes them.

## 1. Background — what we know for certain

- **The bug:** the ja `Windows.Media.Ocr` recognizer misreads MS Gothic digit
  `9` as `3` on the HM batch page (`IGPXA041`, バッチ処理状況一覧), across the
  時刻 / 14-digit データ作成日 / 処理件数 fields.
- **Root cause = capture resolution, NOT the font.** Proven with the `mock-page`
  module: OCR of a faithful mock rendered/captured at ~100% Edge zoom reproduced
  9→3 everywhere; the *same page* captured at ~200% zoom OCR'd **every 9
  correctly (zero 9→3)**. MS Gothic `9` is fine — it only collapses to `3` when
  the rasterised glyph is too small for OCR.
- **Post-hoc upscale can't fix old snaps.** ProcessTime already upscales 2×
  before OCR (`ConvertTo-ProcessTimeOcrImage`, v2.15.3); it does not help because
  interpolating an already-small capture adds no real detail. A native 2× render
  works; a 2× resize of a low-res snap does not.
- **The problem is BOUNDED and NON-RECURRING.** HmSnap captures exact page text
  (`Ctrl+A/Ctrl+C` → `snap\<Stage>_HM\<correl>.txt`, via `Read-PageText.ps1`) and
  ProcessTime prefers that `.txt` (`ConvertFrom-HmPageText`) over OCR. The `.txt`
  tier is **immune to 9→3** (it is copied text, not OCR). But the `.txt` feature
  shipped after the operator had already finished the GIFT phase and part of
  GFIX, so **those old snaps have no `.txt`** and fall back to OCR → 9→3. New
  snaps are fine. Affected backlog ≈ **500+ rows**.
- **Snap availability:** most old snaps exist as standalone
  `snap\<Stage>_HM\<correl>.png`; **some exist only as pictures embedded in the
  evidence workbook** (export via `EvidenceImageExport.ps1`).

## 2. Goal

For the finite old-snap backlog, get correct 開始日時 / 終了日時 / 処理時間 into
the `処理時間(*).xlsx` outputs with **minimal, confident human effort**. Do NOT
try to "cure" OCR — triage: auto-confirm the confident majority, route the rest
to a fast human check. Only rows whose values came from **OCR** are in scope;
rows sourced from `.txt` are already trustworthy (tag them and skip triage).

Two deliverables, in priority order:

- **D1 — correl-ID → snap hyperlink (always on, robust floor).** Every output
  row's 相関ID cell becomes a clickable hyperlink to that correl's snap image, so
  any row is one click from human confirmation. Ships regardless of D2.
- **D2 — auto pixel-diff triage (primary).** Per OCR-derived row, ask the snap
  image whether the extracted digits are right, and write a verdict column so the
  operator only hand-checks the flagged rows. **Conservative bias: never
  auto-confirm unless highly confident — when in doubt, flag** (a false flag
  costs one extra human glance; a false auto-confirm ships a wrong value).

## 3. D1 — correl-ID hyperlink (implement first; low risk)

In the ProcessTime output-write path (`Write-ProcessTimeWorkbook` in
`ProcessTime.ps1`):

1. Add a pure resolver `Resolve-OldSnapImagePath (correl, side)` →
   `snap\<Stage>_HM\<correl>.png` where `<Stage>` is `GIFT`/`GFIX` from the row's
   side. Return `$null` if absent (then D2 marks NoSnap; hyperlink omitted).
   Put the path-building (no I/O) in a pure helper; the existence check stays in
   the COM caller.
2. After writing each row, if a snap path resolved, set a hyperlink on the 相関ID
   cell: `$ws.Hyperlinks.Add($cell, $absolutePath)`. Use an absolute path (the
   output workbook and the snap dir are in known locations under WorkDir); keep
   the cell's displayed text = the correl id.
3. For embedded-only snaps (no standalone PNG), export once via
   `EvidenceImageExport.ps1` to a stable `snap\_verify\<correl>.png` and link
   that. (Guard: exported picture may be as low-res as the original — see §6.)

Config gate: `ProcessTime.OldSnapVerify.EmitHyperlink` (default `$true`).

## 4. D2 — auto pixel-diff triage

### 4.1 Phase 0 — de-risk separability in CI (node/`mock-page`, DO THIS FIRST)

Before any office-PC work, prove the core assumption in the cloud with the
existing `mock-page` renderer + `node --test`:

> Can a check reliably tell "the extracted digit matches the image" from "it does
> not (9 read as 3)", with a threshold that separates right-vs-right (accept)
> from right-vs-wrong (reject) by a safe margin — tolerant to anti-aliasing and
> small offsets?

- Build a node prototype (`mock-page/pixeldiff.mjs` + tests) that renders digit
  crops at matched size and scores similarity (grayscale → binarize → trim →
  normalize height → normalized cross-correlation or column ink-profile
  distance). Feed it right/right and right/wrong(3↔9) pairs; sweep the threshold;
  report the separation margin.
- **GO/NO-GO gate:** if separation is not robust, ship **D1 + the deterministic
  checks (§4.3) only** and drop the image comparison. Record the decision here.

### 4.2 Phase 1 — office-PC implementation (COM/GDI; static-checked in cloud)

**Recommended comparison method — per-digit 3/9 disambiguation (sharper &
more robust than whole-cell render-diff):** MS Gothic on the HM page is
effectively monospaced, so digit x-positions within a cell are computable. For
each digit the OCR read as `3` **or** `9` in 開始/終了/処理時間 (and, if in
scope, データ作成日/処理件数):
1. Crop that digit's pixel box from the snap.
2. Template-match it against a `3` template and a `9` template (MS Gothic,
   rendered at the crop's pixel size) using `Locate-ByImage.ps1` (LockBits).
3. If the image matches the *other* glyph better than the OCR'd one → the digit
   is wrong → flag the row (and optionally record the corrected digit).

This targets exactly the 3↔9 decision, needs only per-digit localization (not
whole-cell alignment), and reuses existing template-match infra. The whole-cell
"render the value, diff against the snap crop" approach (the original idea) is a
fallback if per-digit localization proves unreliable.

Per OCR-derived row:
1. **Resolve snap image** (§3 resolver; export embedded if needed).
2. **Locate the correl's row + the time cells** in the snap. Reuse
   `SnapVerify.ps1` `Get-MatchedRowIndex` / `Get-RowPixelRect` + `ScreenRegion.ps1`
   crop math + the real HM column widths
   (開始150 / 終了150 / 処理時間80 / …; from the calibrated mock template).
   Calibrate offsets in config like `Mark.Boxes`.
3. **Render 3/9 templates** (or the whole-value reference) in MS Gothic at the
   crop's pixel size via `System.Drawing` `Graphics.DrawString` with
   `Font("MS Gothic", size)`. NOTE: GDI+ rasterises differently from the Edge
   snap (DirectWrite) — the §4.1 metric MUST be AA-tolerant, or render the
   reference via a headless Edge/WebBrowser instead.
4. **Compare** (`Locate-ByImage.ps1`) → similarity/decision.
5. **Verdict** via a pure decision function (`Get-OldSnapVerifyVerdict`,
   unit-tested): `OK` (auto-confirmed) / `NG` (flag) / `NoSnap`. Threshold from
   Phase 0. Conservative default = `NG` on any uncertainty.

### 4.3 Deterministic pre-checks (robust floor — ship regardless of Phase 0)

Flag likely-wrong rows with no images, cheaply:
- **Arithmetic:** 処理時間 ≠ 終了 − 開始 → flag. (Already the output J column —
  reuse `ProcessTimeCheck.ps1`.)
- **Datestamp cross-check (pure, unit-testable):** when データ作成日 is present
  and its hh:mm differs from 開始 hh:mm by **only** a 3↔9 swap → flag (optionally
  auto-correct start via `Repair-ProcessTimeStartFromStamp`; adopt the datestamp
  hh:mm, seconds unchanged). Sparse (blank on 0-count rows) but free and certain
  when it fires.
- **Optional cross-engine:** re-read the cell with en-US; where ja/en disagree on
  a digit AND en actually read it, flag. (en drops many fields — weak but free.)

A row is **auto-confirmed only if** it passes every deterministic check AND (when
D2 enabled) the image check says OK; otherwise it is flagged and the operator
uses the D1 hyperlink.

## 5. Output workbook changes

Append after the existing check columns (`ProcessTimeCheck.ps1`
`Get-ProcessTimeCheckColumnSpec` — keep the data-driven spec pattern; add a
column, do not inline):
- **`検証` (Verify):** `OCR-OK` / `要確認`(needs check) / `txt`(from `.txt`, trusted) /
  `画像なし`(NoSnap). Japanese via `ProjectLabels.ps1` `[char]`.
- 相関ID cell carries the D1 hyperlink.
- End-of-run summary line: count of `要確認` rows (extend
  `Get-ProcessTimeCheckSummaryLine`).

## 6. Config (`VerifyConfig.psd1` + overlay)

New section `ProcessTime.OldSnapVerify`:
`Enabled` (default `$true`), `EmitHyperlink` (`$true`), `PixelDiff.Enabled`
(default `$false` until Phase 0 passes), `PixelDiff.Threshold`,
`RenderFont` (`"MS Gothic"`), `SnapDirPattern`, crop-geometry offsets
(per snap-window size, like `Mark.Boxes`), `CrossEngine.Enabled` (`$false`).

**Per `CLAUDE.md`:** also add `OldSnapVerify` to the most relevant group in
`ConfigOverlay.ps1` `Get-ConfigOverlayGroups` and mention it in
`Get-ConfigOverlayReadmeText` in the SAME change, then run `Tests\Run-Tests.ps1`
(the schema-drift guard fails otherwise).

## 7. Conventions to honor (from CLAUDE.md)

- ASCII-only `.ps1`; Japanese via `ProjectLabels.ps1` `[char]`.
- Pure (COM-free) helpers get unit tests in `Tests\` and run in CI; COM/GDI/Excel
  code is static-checked only (no PowerShell/Excel/Windows in the cloud env) —
  confirm on an office PC. Mark such code clearly.
- Capture `[bool]$Force.IsPresent` before any dot-source; dot-source only
  `no-param()` files.
- New pure helpers: `Resolve-OldSnapImagePath` (path build),
  `Get-OldSnapVerifyVerdict` (decision), `Repair-ProcessTimeStartFromStamp`
  (3↔9 datestamp correction) — all unit-tested.
- This writes to the OUTPUT workbook, not the mapping CSV — no new mapping
  columns/bitmasks expected.

## 8. Testing

- **CI (node):** Phase-0 separability prototype + `node --test`
  (`mock-page/pixeldiff.mjs`).
- **CI (pwsh):** `Resolve-OldSnapImagePath`, `Get-OldSnapVerifyVerdict`,
  `Repair-ProcessTimeStartFromStamp`, and any geometry math — unit-tested via
  `Tests\Run-Tests.ps1`; `Check-Encoding.ps1` clean.
- **Office PC:** run on a real `処理時間` folder against known old snaps; confirm
  (a) hyperlinks open the correct snap, (b) a known 9→3 row is flagged `要確認`
  (NOT auto-confirmed), (c) a known-good row auto-confirms, (d) calibrate crop
  geometry + Phase-0 threshold on real captures.

## 9. Open questions / calibration

- Exact crop geometry per snap-window size (real old snaps vary; needs office-PC
  samples — reuse the `Mark.Boxes`/`Calibrate-HmGeometry.ps1` calibration idea).
- GDI+ vs Edge rasteriser AA gap — Phase 0 must confirm the metric tolerates it;
  else render the reference via headless Edge, or drop to deterministic-only.
- Embedded-only snaps: if the exported picture is as low-res as the original,
  cropping/template-match won't help — those rows are **hyperlink-only** (mark
  `要確認`, never auto-confirm). Decide this per exported-picture resolution.
- Sequencing suggestion: ship **D1 + §4.3 deterministic checks** first (immediate
  value, low risk), then gate D2's image comparison on the Phase-0 result.
