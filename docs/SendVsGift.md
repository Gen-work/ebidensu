# SendVsGift phase handoff plan

`SendVsGift` is a new manual bridge phase for comparing SEND data evidence against downloaded GIFT data.

## Stage 1 (implemented MVP)

Command:

```powershell
.\VerifyTool.ps1 -Phase SendVsGift
```

The phase does the following:

1. Scans every file under `<WorkDir>\DATA\GIFT` (or `<WorkDir>\data\GIFT` when that lowercase folder exists).
2. Writes exact file metadata to `<WorkDir>\data\gift_metadata.csv`.
3. Ensures the mapping CSV has a `SendVsGift` column.
4. For each pending mapping row (`SendVsGift` empty or `0`), prints the matching GIFT file metadata in the console and opens the evidence workbook.
5. Operator checks the workbook and console data. Pressing Enter sets `SendVsGift=1`, sets workbook cursors to `A3` (or the selected `-CursorCell`), saves, and closes.

Metadata columns written to `gift_metadata.csv`:

- `FileName`, `FullName`
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

## Stage 2 (OCR skeleton, implemented)

Enabled with `-Ocr` on `SendVsGift.ps1`, or persistently via `SendVsGift.Ocr = $true`
in `VerifyConfig.psd1` / the `verify_config.json` work-folder overlay.

### Engine choice

The OCR engine is the Windows built-in `Windows.Media.Ocr` WinRT API -- the same
engine family behind the Snipping Tool text extraction / PowerToys Text Extractor.
It is called directly from PowerShell 5.1 (`OcrWindows.ps1`), so there is nothing
to install and no GUI tool to automate. Japanese works when the `ja` language pack
is present (default on JP-locale hosts). Note: the API exposes no per-word
confidence score, so confidence is heuristic (see below).

### Pipeline

1. `EvidenceImageExport.ps1` exports every picture on the send-data sheet
   (the ProjectLabels send-data sheet; override with `-SendSheetName`) to
   `<WorkDir>\data\send_images\<Correl_ID_S>\<Correl_ID_S>_NN.png`, top-to-bottom.
   Multi-picture groups are concatenated in that order, covering stacked
   screenshots of one SEND file. (Export goes through a temp ChartObject and
   clobbers the clipboard; the workbook is left unsaved unless Enter is pressed.)
2. `OcrWindows.ps1` OCRs each PNG and returns plain line/word objects with
   bounding boxes. The Japanese recognizer drops spaces between tokens;
   `SendMetadata.ps1` rebuilds them from the word X/Width boxes so first/last
   tokens and approximate fixed positions survive.
3. `SendMetadata.ps1` (pure, unit-tested by `Tests\Test-SendMetadata.ps1`) parses
   the lines into a record parallel to `gift_metadata.csv`, written to
   `<WorkDir>\data\send_metadata.csv`:
   `CorrelIdS, ExcelName, ImageCount, OcrLineCount, ZeroByte, RowNumberGuess,
   FirstRecord, LastRecord, FirstRecordToken, LastRecordToken, Confidence,
   MetadataVersion`.
4. `Compare-SendGiftMetadata` prints a per-field verdict (`ZeroByte`, `RowNumber`,
   `FirstRecordToken`, `LastRecordToken`): each check is match / mismatch /
   unknown -- absence of OCR evidence is always `unknown`, never `mismatch`.
   Verdict: any mismatch -> `mismatch`; zero-byte agreement or >= 2 field matches
   -> `match`; otherwise `unknown`.

### Confidence and manual fallback

Confidence is heuristic (0 / 0.4 / 0.7 / 1.0 by parsed-field coverage) because the
Windows OCR API has no native score. The verdict is advisory only: the manual
Enter-to-mark flow is unchanged and remains the source of truth, OCR failures are
caught and reported as warnings, and the `SendVsGift` mapping column semantics are
untouched.

### Remaining TODOs (need representative screenshots)

- **0-byte pattern**: the default detection regex (`0 byte/bytes`) is a guess;
  tune `SendVsGift.ZeroBytePattern` once a real 0-byte SEND screenshot exists.
- **Row-number / record parsing**: `RowNumberGuess` assumes a leading row number
  on host list lines; verify against real screen layouts and add fixed-position
  field extraction from the word bounding boxes where needed.
- **Record length**: OCR is not trustworthy for character-exact record lengths
  (digit drops, O/0 confusion); lengths are intentionally not compared yet.
