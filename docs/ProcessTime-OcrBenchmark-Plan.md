# ProcessTime OCR Benchmark & 3/9 Fix — Plan (Phases 2 & 3)

Status: PLANNED (design only). Phase 1 (check-formula module extraction /
unified generation, v2.16.0) is tracked separately and is a pure refactor.

## Problem

The ja Windows OCR recognizer misreads digit `9` as `3` on the HM batch page
font (time-of-day, 14-digit datestamp, and record-count fields). The J-column
T/F check (`=IF(ROUND(F*86400,0)=ROUND(I*86400,0),"T","F")`, v2.15.2) and the
operator's added record-count check reliably FLAG these rows, but do not fix
them. v2.15.3 added image preprocessing (`ConvertTo-ProcessTimeOcrImage`:
upscale + grayscale + contrast) as a root-cause mitigation, but there is no
repeatable way to MEASURE its effect on 3/9 accuracy or to tune it.

A post-hoc regex check cannot catch this class of misread: a `9`->`3` swap
yields a FORMAT-VALID timestamp (`11:13:16`). Prevention-before-recognition
(clean pixels) plus deterministic cross-checks from data already on the page
are the only clean fixes.

## Hard constraints

- `Windows.Media.Ocr` is Windows-only -> the benchmark runner, page rendering,
  and any second-pass OCR run only on an office PC; they are static-checked in
  the Linux/pwsh CI, never executed there.
- A Chromium-rendered synthetic page cannot be 100% pixel-identical to the real
  Edge-captured Honda page (font rendering / DPI differ). Real captured snaps
  are therefore the PRIMARY ground truth; synthetic pages only SUPPLEMENT
  coverage of 3/9 digit combinations absent from real data.
- Target runtime is Windows PowerShell 5.1; pwsh 7 in CI proves parse + pure
  logic only. Collection expansion must go through
  `ConvertTo-ProcessTimeBucketArray` (never `@($var[index])`).

## Phase 2 (v2.17.0) — benchmark set + deterministic 3/9 fix (C1)

### 2a. Real-snap ground truth (primary)
`Export-OcrBenchmarkTruth.ps1` (new, read-only, office PC): reverse-export a
`Tests/ocr-benchmark/manifest.json` from a J-column-all-`T` confirmed
`処理時間(*).xlsx` — one entry per side `{ Png; CorrelId; Side; Start; End;
Count; Status }`, Png pointing at the real `snap\<Stage>_HM\<correl>.png`.
(Whether the PNGs are committed depends on size; otherwise the manifest holds
local paths.)

### 2b. Synthetic HTML ground truth (supplement) — see procedure below

### 2c. Benchmark runner + pure comparison
`Test-OcrAccuracy.ps1` (new, `param()`, Windows-only): for each manifest PNG,
run the SAME pipeline as ProcessTime (`Read-ProcessTimeOcrLines` preprocess +
en-US/ja pool + `ConvertFrom-ProcessTimeOcrLines` + `Select-ProcessTimeRow`),
compare parsed vs truth, emit a 3<->9 confusion matrix, per-field accuracy,
pass/fail, and `-Json`. `-Sweep` is reserved for Phase 3.
Pure helpers into `ProcessTimeParse.ps1` (CI-unit-tested): `Compare-OcrDigits`
(same-position compare -> 3/9 confusion counts), `Get-OcrBenchmarkScore`.

### 2d. C1 — datestamp hh:mm cross-correction (deterministic; NOT result-merging)
`Get-ProcessTimeDateHints` already parses the clean en-US 14-digit
`データ作成日` into a full start datetime; the existing correction
(`ProcessTimeParse.ps1`, date-hint block) only fires when the time-of-day
matches EXACTLY, so a 9->3 read (hh:mm differ) never triggers. New pure
`Repair-ProcessTimeStartFromStamp`: when a hint's hour:minute differs from the
ja start hh:mm ONLY by a 3<->9 swap, adopt the datestamp's hour/minute (seconds
left unchanged — `開始日時` seconds and `作成日` seconds legitimately differ),
tag `TimeCorrected`. Covers the START time only (no second on-page source for
END). Unit test: `11:13:16` + hint `11:19:06` -> `11:19:16`; a non-3/9 mismatch
is left untouched.

### 2b procedure — synthetic HTML screenshot benchmark (needs operator CSS + testing)
Roles: operator supplies CSS + does the office-PC visual/OCR calibration; the
scripts do skeleton / inject / crop / compare.
1. Operator supplies: the real HM full-screen screenshot (have it), plus the
   page font family/size, teal header RGB, alt-row color, column widths (as
   available; otherwise drafted from the screenshot).
2. `Tests/ocr-benchmark/hm-template.html`: faithful 10-column
   「バッチ処理状況一覧」layout with `{{placeholder}}` values and a first-draft CSS.
3. Operator refines CSS in Edge until the rendered `3`/`9` are visually
   indistinguishable from the real page (not pixel-identical — that is not
   achievable cross-browser).
4. `Build-OcrBenchmarkImages.ps1` (office PC/Edge): inject a truth CSV (with
   deliberately 3/9-dense samples) -> Edge screenshot -> crop reusing
   `ScreenRegion.ps1`'s `Resolve-DirectionalCrop` + `Window.CropPx` (same params
   as production) -> PNG + manifest.
5. Operator runs `Build-...` then `Test-OcrAccuracy.ps1` on the synthetic set;
   remaining 3/9 errors -> back to step 3 or on to Phase 3.

## Phase 3 (v2.17.x) — preprocessing tuning + optional Plan A (C3)

### C2 — tune preprocessing from the benchmark
`Test-OcrAccuracy.ps1 -Sweep` over `OcrPreprocessScale {1,2,3}` ×
`OcrPreprocessContrast {1.0,1.3,1.6}` × optional binarization threshold on the
real + synthetic sets; set the winning combo as `VerifyConfig.psd1` defaults.
If binarization helps, wire it into `ConvertTo-ProcessTimeOcrImage` via the
`ProcessTime.OcrPreprocessBinarize` / `OcrPreprocessThreshold` config (reserved
in Phase 1d).

### C3 — bounding-box second-pass (Plan A; ONLY if C1+C2 fall short of 100%)
Trigger already exists: the J-column computes `F` (or datestamp vs start hh:mm
disagree). For a triggered row, crop the numeric cell using the OCR
`Words[].BoundingRect` (`OcrWindows.ps1`) and re-OCR that fragment with en-US
only, overwriting the ja read. `-NonInteractive` falls back to the C1+C2
result. Not implemented until the benchmark proves it necessary.

## Verification
- CI (pwsh): `Compare-OcrDigits` / `Get-OcrBenchmarkScore` /
  `Repair-ProcessTimeStartFromStamp` unit-tested; `Check-Encoding.ps1` clean.
- Office PC: `Export-OcrBenchmarkTruth.ps1` -> `Test-OcrAccuracy.ps1` baseline
  confusion matrix; `-Sweep` picks the preprocessing default; confirm 3/9 hits
  100% (else enable C3); re-run a real JOD `-Stage Both -Force` and confirm the
  J column no longer shows `F` caused by 3/9.
