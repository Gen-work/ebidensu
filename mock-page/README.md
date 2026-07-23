# mock-page — HM/MQ/Jenkins page stand-ins for in-cloud testing

A **stand-alone** module that renders faithful HTML stand-ins for the Honda
snap pages and screenshots them to PNG with Chromium (Playwright). It exists so
the ProcessTime OCR and snap-parse logic can be exercised and benchmarked in a
**Linux / CI environment without an office PC**.

The first page implemented is the HM batch-processing status list
(`I???XA041.jsp`, バッチ処理状況一覧) — the page behind the ProcessTime
`9`→`3` OCR problem (see `../docs/ProcessTime-OcrBenchmark-Plan.md`).

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
node gen.mjs --full    # whole-page screenshot instead of a tight table crop
```

Flags: `--truth <file>` `--out <dir>` `--name <stem>` `--scale <dpi factor>` `--full`.

## Ground-truth JSON schema (`samples/*.json`)

```jsonc
{
  "title": "バッチ処理状況一覧",
  "meta":  "I???XA041  (mock)  ページ 1 / 1",
  "rows": [
    {
      "correlId": "JIGPC06S",
      "system":   "JIG",
      "jobName":  "JIGPC06SXA",
      "start":    "2026/07/01 11:19:16",   // 開始日時
      "end":      "2026/07/01 11:19:28",   // 終了日時
      "status":   "normal",                 // normal -> 正常終了 | abend -> 異常終了
      "datestamp":"20260701111906",         // データ作成日 (14-digit yyyyMMddHHmmss)
      "count":    "29,264",                 // 処理件数 (comma-grouped)
      "rc":       "0000"
    }
  ]
}
```

Every value is **ground truth** — what the OCR must read back. Pack the
time / datestamp / count fields with `9`s and `3`s to probe the confusion.

## Calibration — make the digits faithful (needs operator input)

The mock is a stand-in for a specific font's pixels; until it's calibrated the
rendered `3`/`9` are only a guess. Everything marked `/* CALIBRATE */` in
`templates/hm-batch-status.html` is a placeholder. To calibrate, supply (in
priority order):

1. **A real full-screen screenshot of `I???XA041.jsp`** — the visual ground
   truth to match. *(Most valuable; enables everything below.)*
2. **The page CSS / `.jsp` `<style>`** — real font-family, font-size, teal
   header RGB, alt-row colour, column widths. These decide how `9` vs `3`
   render, which is the entire point.
3. **The `.jsp` source** — exact `<table>` DOM: real column order, all headers,
   any colspans. The current 10 columns are drawn from what the parser reads,
   not from the page markup.

Refine the CSS in a browser until the rendered `3`/`9` are visually
indistinguishable from the real page (pixel-identical is not achievable
cross-browser — real captured snaps stay the primary ground truth per the
benchmark plan).

## Relationship to the OCR benchmark plan

This module is the Linux-runnable implementation of step **2b** in
`../docs/ProcessTime-OcrBenchmark-Plan.md` (the synthetic HTML supplement). It
deliberately renders with Chromium here instead of office-PC Edge so page
authoring and the render pipeline iterate in-cloud; the real-OCR measurement
(`Test-OcrAccuracy.ps1`, step 2c) still runs on Windows against these PNGs plus
the primary real-snap ground truth.

## Next steps

- Calibrate the HM template against a real screenshot (operator inputs above).
- Consume the emitted `*.manifest.json` from the (Windows) OCR accuracy runner
  so synthetic pages contribute to the 3↔9 confusion matrix.
- Add MQ and Jenkins page templates once HM is proven.
