# Office-PC test guide (no Node.js required)

The office PC has PowerShell 5.1 + Excel + Edge but **no Node.js**, so use the
pre-built HTML in this folder — it has the sample data baked in and needs no
tooling. `*.html` opens directly in Edge; `*.manifest.json` is the matching
ground truth (the correct values OCR must read back).

Files:
- `hm-sample-allrows.html` — every sample row in one page (easiest to eyeball).
- `hm-sample-snap.html` — the real scrolled-window clip (height ~500).
- `hm-sample-*.manifest.json` — ground truth for each.

## Step 1 — fidelity check (do this first)

1. Copy `hm-sample-allrows.html` to the office PC and **double-click → opens in Edge**.
2. Put it side by side with the real `IGPXA041` page.
3. Confirm: column order/widths, slate header `#778899`, light-green alt rows,
   raised (outset) cell borders, right-aligned 処理件数, blank データ作成日, and
   — most important — the **font**. Edge uses real MS Gothic, so the digit
   shapes (esp. `9` vs `3`) should match the real page.

If anything's off, tell me the diff and I'll adjust the template.

## Step 2 — real OCR (the actual 9→3 measurement)

1. In Edge, screenshot the page with your normal snap tool (the MS-Gothic pixels
   are what matter — do **not** use a Linux/Chromium screenshot for this).
2. Run your existing Windows OCR over that PNG — the same engine ProcessTime
   uses (`OcrTool.ps1` / `OcrWindows.ps1`, `Windows.Media.Ocr`).
3. For each row read back: `start, end, duration, datestamp, count` (+ `correlId`
   from 相関ID). Save them as a `--read` JSON, same shape as
   `../samples/sample-ocr-read.json`.

## Step 3 — score it

Bring the `--read` JSON and the matching `*.manifest.json` to any machine with
Node (this dev session works) and run:

```
node ../compare.mjs --manifest hm-sample-allrows.manifest.json --read <your-read>.json
```

You get per-field digit accuracy and the **3↔9 confusion matrix** — the real
number for how often MS Gothic `9` is misread as `3`, and vice versa.

## Regenerating these files (on a Node machine)

```
node ../gen.mjs --html-only --out . --name hm-sample-allrows
node ../gen.mjs --html-only --snap --out . --name hm-sample-snap
```

Edit `../samples/sample-truth.json` (or point `--truth` at your own) to change
the rows.
