# mock-page — HM/MQ/Jenkins page stand-ins for in-cloud testing

A **stand-alone** module that renders faithful HTML stand-ins for the Honda
snap pages and screenshots them to PNG with Chromium (Playwright). It exists so
the ProcessTime OCR and snap-parse logic can be exercised and benchmarked in a
**Linux / CI environment without an office PC**.

The first page implemented is the HM batch-processing status list
(`IGPXA041.do`, バッチ処理状況一覧) — the page behind the ProcessTime
`9`→`3` OCR problem (see `../docs/ProcessTime-OcrBenchmark-Plan.md`). Its
layout, column widths, classes, colours and font are **calibrated from the real
saved page** (`ids.css` + the list markup).

Real column order (fixed `titlerow` header + scrolling body, both `width=970`):

| 開始日時 | 終了日時 | 処理時間 | バッチID | ＳＳ | 処理状態 | データ作成日 | 処理件数 | 処理結果 | 相関ID |
|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|
| 150 | 150 | 80 | 80 | 40 | 80 | 130 | 70 | 70 | 100 |

`処理件数` right-aligned; blank `データ作成日` renders as a full-width space;
`処理結果` is a `◆`; `相関ID` is a link.

### Font fidelity (important)

The real page font is **`ＭＳ ゴシック` (MS Gothic) 10pt** — the exact font whose
digit glyphs cause the ja OCR `9`↔`3` misread. MS Gothic is proprietary and
Windows-only, so:

- An **office-PC Edge render** of this template is production-identical (MS Gothic present) → use it for the real OCR benchmark.
- A **Linux/Chromium render** falls back to another CJK font → faithful **layout**, not faithful **digit glyphs**. Use it for authoring, layout checks and manifests — not for claiming a 3/9 result.

## What runs where

| Step | Runs in Linux/CI | Notes |
|------|:---:|------|
| Render mock page → PNG (this module) | ✅ | node + Playwright + the pre-installed Chromium |
| Ground-truth manifest (PNG ↔ correct values) | ✅ | we generate the page, so we already know the answer |
| Feed known / deliberately-corrupted text to the pure parse+repair logic | ✅¹ | pure PowerShell — no OCR needed to test the 3↔9 repair |
| Run the **real** `Windows.Media.Ocr` engine over the PNGs | ❌ | Windows-only; the actual 9→3 bug lives in that engine — measure it on an office PC |

¹ Executing the PowerShell unit tests needs `pwsh`, which this container's egress
policy blocks from installing (GitHub release CDN / Microsoft apt repo return
403). The tests are authored here and run in the project's existing CI /
office PC, same as every other pure lib in this repo.

**Key point:** a Linux stand-in OCR (e.g. Tesseract) is a *different engine with
a different error profile* — it will NOT faithfully reproduce the ja
`Windows.Media.Ocr` 9↔3 confusion, so it must not be used to claim a 3/9 result.

## Usage

```bash
npm install            # one-time; playwright-core (browsers already on the image)
node gen.mjs           # renders samples/sample-truth.json -> out/sample-truth.png (+ .manifest.json)

node gen.mjs --truth samples/sample-truth.json --name run1 --scale 2
node gen.mjs --snap    # real window-clipped body (height 500, scroll) instead of all rows
node gen.mjs --full    # whole-page screenshot instead of a tight crop
```

Flags: `--truth <file>` `--out <dir>` `--name <stem>` `--scale <dpi factor>` `--snap` `--full`.
Default capture shows **every** ground-truth row in one image; `--snap` reproduces
the real scrolled-window clip.

### Scoring an OCR read (`compare.mjs`)

The manifest is the correct answer. Score an OCR read against it:

```bash
node compare.mjs                       # scores samples/sample-ocr-read.json vs out/sample-truth.manifest.json
node compare.mjs --manifest <m> --read <r> [--json]
```

It matches rows by `correlId` and reports per-field digit accuracy and a
**3↔9 confusion matrix** — the ruler for the OCR problem. Example output:

```
field      exactRows  digitAcc  9->3  3->9  other
start      9/10       99.3%     1     0     0
...
OVERALL digit accuracy: 484/489 (99.0%)
3<->9 confusion: 9->3 = 4, 3->9 = 1, other digit swaps = 0
```

The `--read` file lists `{ correlId, start, end, duration, datestamp, count }`
per row. Here it's a hand-made fixture with injected errors; **on an office PC it
is the real `Windows.Media.Ocr` output** dumped to JSON. The scorer is pure
(string compare + counting), unit-tested with `node --test` (`tests/`), so it
runs in CI with no OCR/Windows/Excel.

## Ground-truth JSON schema (`samples/*.json`)

```jsonc
{
  "title": "バッチ処理状況一覧",
  "meta":  "IGPXA041  test",
  "rows": [
    {
      "start":     "2026/07/23 03:59:07",  // 開始日時
      "end":       "2026/07/23 03:59:52",  // 終了日時
      "duration":  "00:00:45",             // 処理時間 (HH:mm:ss)
      "batchId":   "IGPLB073",             // バッチID
      "ss":        "S",                     // ＳＳ (single char)
      "status":    "normal",                // normal -> 正常終了 | abend -> 異常終了 | processing -> 処理中
      "datestamp": "20260723035907",        // データ作成日 (14-digit; "" -> full-width space)
      "count":     "59,476",                // 処理件数 (comma-grouped, right-aligned)
      "correlId":  "JIGPB1S"                // 相関ID (rendered as a link)
    }
  ]
}
```

Every value is **ground truth** — what the OCR must read back. Pack the
time / datestamp / count fields with `9`s and `3`s to probe the confusion.

## Calibration status

Structure, column widths, classes, colours and font-family are already
calibrated from the real saved page (`ids.css` + the list markup) — the sample
rows are transcribed from a real `IGPXA041` screenshot. What remains:

- **Font glyphs on Linux** — MS Gothic is Windows-only, so the exact `3`/`9`
  pixels only appear in an office-PC Edge render (see "Font fidelity" above).
  This is a platform limit, not a missing input; the real OCR benchmark runs on
  Windows regardless.
- **Optional visual polish** — the HONDA top bar / search form are omitted (not
  OCR-relevant). The title strip colour is approximated. Add them only if a
  future check needs the full window chrome to match the crop geometry.

## Relationship to the OCR benchmark plan

This module is the Linux-runnable implementation of step **2b** in
`../docs/ProcessTime-OcrBenchmark-Plan.md` (the synthetic HTML supplement). It
deliberately renders with Chromium here instead of office-PC Edge so page
authoring and the render pipeline iterate in-cloud; the real-OCR measurement
(`Test-OcrAccuracy.ps1`, step 2c) still runs on Windows against these PNGs plus
the primary real-snap ground truth.

## Next steps

- **Office PC:** render this template in Edge, run the real `Windows.Media.Ocr`
  over the PNGs, dump the read to a `--read` JSON, and run `compare.mjs` to get
  true 3/9 numbers against production font rendering.
- **PS port for the production pipeline:** mirror `compareDigits` /
  `scoreBenchmark` into `../ProcessTimeParse.ps1` as `Compare-OcrDigits` /
  `Get-OcrBenchmarkScore` (per the benchmark plan) so the ProcessTime OCR path
  can self-score; validate on the office PC (no pwsh in this container).
- **Per-row integrity check (replaces the dropped datestamp repair):**
  `処理時間 == 終了 − 開始` holds on *every* row (unlike データ作成日, which is
  often blank), so it can gate/flag a misread row without a second OCR pass.
- Add MQ and Jenkins page templates once HM is proven.
