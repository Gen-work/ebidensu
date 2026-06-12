# SendVsGift phase handoff plan

`SendVsGift` is a new manual bridge phase for comparing SEND data evidence against downloaded GIFT data.

## Stage 1 (implemented MVP)

Command:

```powershell
.\VerifyTool.ps1 -Phase SendVsGift
```

The phase does the following:

1. Scans every file under `<WorkDir>\DATA\GIFT` (or `<WorkDir>\data\GIFT` when that lowercase folder exists).
2. Writes exact file metadata to `<WorkDir>\data\gift_metadata.csv`. When a mapping row has `isZIP`/`isZip` set to `1`, the matching GIFT zip archive is extracted to `<WorkDir>\data\unzip\<Correl_ID_S>` and that extracted same-name file is also written to metadata for comparison.
3. Ensures the mapping CSV has a `SendVsGift` column.
4. For each pending mapping row (`SendVsGift` not `1` -- pending covers empty, `0` and the NG value `2`), prints the matching GIFT file metadata in the console.
5. Pending rows are **grouped per evidence workbook**: the workbook is opened once, and for each of its correl rows the cursor jumps to the `Correl_ID_S` label cell in **column A of the send-data sheet** (the evidence pictures sit right below the label). After every console answer Excel is brought back to the foreground (direct `SetForegroundWindow` on the Excel hwnd; `Alt+Tab` SendKeys fallback) so the operator lands on the next correl without switching windows by hand.
6. Operator checks the workbook and console data. `Enter` sets `SendVsGift=1`; `n` sets `SendVsGift=2` (NG -- stays pending and is listed at the end of the run); `s` skips; `q` quits. The workbook is saved (cursors reset to `-CursorCell`) and closed only after its **last** correl row is answered -- no close/reopen between correls of the same workbook.

When run standalone (`.\SendVsGift.ps1`), `-WorkDir`/`-Owner` fall back to `verify_session.json`, then to the single `mapping_*.csv` in the work folder, then to a prompt -- the old `mapping_.csv` failure is gone.

Metadata columns written to `gift_metadata.csv`:

- `FileName`, `FullName`, `SourceZip`
- `SizeBytes`, `SizeDisplay`
- `MaxRowNumber`
- `MinRecordLength`, `MaxRecordLength`
- `FirstRecordLength`, `LastRecordLength`
- `FirstRecordToken`, `LastRecordToken`
- `FirstRecord`, `LastRecord`
- `LastWriteTime`, `MetadataVersion`

### Record-length rule

Record length is calculated as the PowerShell string length of each line after removing only the line ending. This means CR/LF bytes are not counted. `MaxRecordLength` is the longest line in the file; `MinRecordLength` is included to quickly show mixed record types such as a long header/details line and shorter continuation line.

### TL;DR record display decision

The phase stores both forms:

- Full `FirstRecord` and `LastRecord` are kept for exact comparison and audit.
- `FirstRecordToken` and `LastRecordToken` are printed as short identifiers for fast visual scanning.

The first/last non-space token alone is not enough for final comparison because it can miss changes in fixed-position fields. It is useful as a console summary, but the full first/last records remain the source of truth.

## Stage 2 (OCR auto-compare, implemented)

Enabled with `-Ocr` on `SendVsGift.ps1` / `VerifyTool.ps1`, with the `o` option at
the VerifyTool phase prompt, or persistently via `SendVsGift.Ocr = $true` in
`VerifyConfig.psd1` / the `verify_config.json` work-folder overlay.

### Engine choice

The OCR engine is the Windows built-in `Windows.Media.Ocr` WinRT API -- the same
engine family behind the Snipping Tool text extraction / PowerToys Text Extractor.
It is called directly from PowerShell 5.1 (`OcrWindows.ps1`), so there is nothing
to install and no GUI tool to automate. Japanese works when the `ja` language pack
is present (default on JP-locale hosts). Note: the API exposes no per-word
confidence score, so confidence is heuristic (see below).

### Pipeline

1. The correl's **section** on the send-data sheet is located from its column-A
   label cell: the section spans from the label down to the next non-empty cell
   in column A (the next correl's label), or to the end of the sheet.
2. `EvidenceImageExport.ps1` exports only the pictures inside that section to
   `<WorkDir>\data\send_images\<Correl_ID_S>\<Correl_ID_S>_NN.png`, top-to-bottom
   then left-to-right. **Ctrl+G groups are flattened**: each child picture is
   exported on its own (a grouped strip exported as one composite PNG could
   exceed the Windows OCR engine's max image dimension; per-child export also
   keeps the left-to-right capture order). (Export goes through a temp
   ChartObject and clobbers the clipboard.)
3. `OcrWindows.ps1` OCRs each PNG and returns plain line/word objects with
   bounding boxes. The Japanese recognizer drops spaces between tokens;
   `SendMetadata.ps1` rebuilds them from the word X/Width boxes so first/last
   tokens and approximate fixed positions survive.
4. `Compare-SendGiftEvidence` (pure, unit-tested by `Tests\Test-SendMetadata.ps1`)
   applies the operator's review rules:
   - **gift file is 0 bytes** -> some image must show 0-byte evidence:
     the dataset-info screen with `used CYLINDERS : 0`, **or** the
     begin-of-data and end-of-data markers on the *same* image with no
     `000001` record line (a non-empty file never shows both markers on one
     screenshot). A `000001` record anywhere is a positive `ng`.
   - **gift file has data** -> the zero-padded max row number from
     `gift_metadata.csv` (e.g. `000003`, `004644`) must appear in the OCR
     text, and the first record (after `000001`) and last record (after the
     max label) must match the gift records: exact first space-free token
     wins; otherwise a >= 80% Levenshtein similarity of the first 20 chars
     passes as `fuzzy` (OCR noise tolerance). A present-but-failing record
     is `mismatch`.
   - Verdict: any mismatch -> `ng`; max row found + both records pass ->
     `ok`; anything thinner -> `unknown`. Absence of OCR evidence is never
     a mismatch.
5. The parsed lines are also stored as a `send_metadata.csv` record per correl
   (same columns as before) for audit.

### Auto-marking

With the default `SendVsGift.AutoMark = $true`:

- verdict `ok` -> `SendVsGift=1` automatically, no prompt;
- verdict `ng` -> `SendVsGift=2` automatically, the row is reported in red and
  listed in the end-of-run NG summary; `2` is *not* `1`, so the row stays
  pending and is re-checked on the next run;
- verdict `unknown` -> the normal manual prompt decides.

Set `AutoMark = $false` (or `-NoAutoMark`) to keep the verdict advisory-only.
OCR failures are caught and fall back to the manual flow.

### Standalone OCR tool

`OcrTool.ps1` exposes the same OCR stack as a reusable command line tool
(images, folders, wildcards, or `-Workbook <xlsx>` to export embedded pictures
first; `-Json` for structured output; `-ListLanguages` to probe the engine).
Future features should dot-source `OcrWindows.ps1` + `SendMetadata.ps1`
directly, or shell out to `OcrTool.ps1`.

### Troubleshooting: OCR reads nothing (open TODO, 2026-06-11)

Field status from the first JIDSC49S end-to-end run:

- Picture export now works end to end: per-correl section located by the
  top-level (Ctrl+G group) shape Top, children flattened, PNGs land in
  `<WorkDir>\data\send_images\<Correl_ID_S>\` and the 3x upscale is
  confirmed effective (`[DIAG] first export ...` prints the pixel size).
- **But the Windows OCR engine still returns ZERO lines** on the upscaled
  evidence PNGs (`[OCR] images=8 lines=0`), so every verdict is `unknown`
  and the flow falls back to the manual prompt. The engine itself
  initializes fine (en-US + ja listed).

Diagnosis aids now built in:

- `.\OcrTool.ps1 -Diag -Path <png|folder>` sweeps every installed
  recognizer language (plus the user-profile engine) per image and prints
  line/word counts, a sample line, the image pixel size and the engine's
  `MaxImageDimension` (flags oversized images).
- SendVsGift prints the exact `-Diag` command in its
  `no text recognized` warning.

**2026-06-12 update** -- `-Diag` run on JIDSC49S_01.png (1687x1276, max
10000) showed the engine DOES see the text: en-US lines=93 words=117,
ja lines=92 words=489 -- **but every `.Text` property read back empty**
(sample blank, pipeline `lines=0` after blank-dropping). So the engine
and image are fine; the failure is reading WinRT string properties from
PS 5.1: enumerating `OcrResult.Lines` / `OcrLine.Words` works while
`OcrLine.Text` / `OcrWord.Text` silently return null on this host.

Fallbacks now in place (untested on the office host yet):

- `Invoke-WinOcrFile` reads the aggregate `OcrResult.Text`; when all
  line/word texts come back empty but the aggregate works, plain lines
  are rebuilt from it (no word boxes -> spacing rebuild is skipped).
- `ConvertTo-SendTextLines` falls back to the line `.Text` when the
  words-rebuilt text is empty.
- `-Diag` now prints `chars=` / `rawChars=` per language plus the WinRT
  line type name, and flags the "enumerates but .Text empty" case
  explicitly.

**2026-06-12 update 2** -- `rawChars=0` as well: even the aggregate
`OcrResult.Text` reads back empty through the PS adapter, while the line
type projects correctly (`Windows.Media.Ocr.OcrLine`). So ALL HSTRING
property reads fail silently on this host. Countermeasure now in place:
`Read-WinRtText` (OcrWindows.ps1) tries, in order: adapter -> `.psbase`
-> .NET reflection (`GetProperty('Text').GetValue(...)`) -> a lazily
compiled inline C# reader (`Add-Type` referencing the in-box
`C:\Windows\System32\WinMetadata\Windows.Media.winmd` -- no SDK needed)
that casts to `OcrLine`/`OcrWord`/`OcrResult` and reads `.Text` from
compiled code, bypassing the PS adapter. `-Diag` reports which strategy
produced text (`text strategy:`) and probes the first word's bounding
box (nonzero X/W proves struct marshaling works even if strings fail).

**2026-06-12 update 3** -- strings flow again (strategy `adapter`,
ja chars=1458 per image; the layered reader stays as insurance). The
next blocker was visible in the sample text: the ja recognizer returns
one word per CHARACTER and the word-box spacing rebuild over-inserts
(`002640` -> `0 0 2 6 4 0`), so the row-label / record / CYLINDERS
matchers all missed. Every matcher now also runs against a COMPACT
(space-stripped) form of each line: row labels match at compact line
start, records are extracted from the compact line (and compared with
compact prefix similarity), and the 0-byte rules scan compact lines
too. Unit-tested in `Tests\Test-SendMetadata.ps1`.

**2026-06-12 update 4** -- compact matching alone did not flip the
verdicts. The line counts told the real story: ~187 OCR "lines" per
~40-row screen, i.e. **the engine fragments one terminal row into
several OCR lines**, so a row label and its record live in different
fragments and no per-line matcher can see them together. New
`ConvertTo-SendRowLines` (SendMetadata.ps1, unit-tested) re-clusters
ALL word boxes of an image by vertical center (tolerance 0.6 x median
word height) and rebuilds true terminal rows left-to-right; SendVsGift
and OcrTool now match on these reconstructed rows. SendVsGift also
dumps what the matcher saw to
`data\send_images\<Correl_ID_S>\<Correl_ID_S>_ocr.txt` per run.

**2026-06-12 update 5** -- row reconstruction works (one terminal row
per line in the `_ocr.txt` dump) and the ground truth exposed the last
two gaps:

- the **ja recognizer garbles digit runs** on the host terminal font:
  `002640` -> a kanji-box + `2640`, `7` -> `?`. SendVsGift therefore now
  OCRs every image with **both** the configured language and `en-US`
  (when installed) and merges the row lines; en-US reads the same digit
  runs cleanly.
- the **`使用 CYLINDERS : 0` value digit is often missed entirely** by
  the engine. New optional pixel fallback: set `SendVsGift.ZeroTemplate`
  (config/overlay) or `-ZeroTemplate <png>` to an operator-cropped
  template image of the `0` evidence; when a 0-byte verdict stays
  `unknown`, every exported section PNG is scanned with
  `Locate-ByImage.ps1` and a hit upgrades the verdict to ok.
  **Crop the template from one of the exported PNGs under
  `data\send_images\...`** (same pipeline = same pixels; do NOT crop
  from a raw screenshot -- the 3x re-render would not match).
- Note: the Snipping Tool's text extraction uses the newer OneOcr
  engine (`oneocr.dll`, Store-app private, no public API) -- noticeably
  better on this screen, but calling it would mean P/Invoking an
  undocumented DLL inside the Snipping Tool package; parked as a
  future option.
- Caveat seen in the C49S dump: the VIEW screenshots can show columns
  215-286 (`欄 215 286`), so first/last record PREFIXES are only
  matchable on the images that show column 1 -- the record checks rely
  on those; the max-row check works on any column view once the digits
  read correctly.

TODO (next office session):

- [ ] `Tests\Run-Tests.ps1` (the 2 longer-digit-run asserts pass again:
      compact fallback now requires >= 6 record chars after the label).
- [ ] `SendVsGift -Ocr -TargetIds JIDSC49S -Force`: with the en-US merge
      the max-row label should be found; check FirstRecord/LastRecord
      against the column-1 images in the `_ocr.txt` dump.
- [ ] 0-byte case: crop a `0`-evidence template from
      `data\send_images\JIDSC05S\JIDSC05S_01.png`, save e.g. as
      `<WorkDir>\zero_tpl.png`, set `SendVsGift.ZeroTemplate` in
      `verify_config.json` (or run with `-ZeroTemplate zero_tpl.png`),
      re-run and expect `ZeroByteTpl ... -> match`.

### Remaining TODOs (need representative screenshots)

- The CYLINDERS / begin-end marker fragments are first-guess OCR forms taken
  from the reference screenshots; if the recognizer mangles them on the real
  host, tune `SendVsGift.ZeroBytePattern` (it overrides both built-in rules).
- **Record length**: OCR is not trustworthy for character-exact record lengths
  (digit drops, O/0 confusion); lengths are intentionally not compared yet.
