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

## Stage 2 (planned OCR/image recognition)

Future work should atomize metadata from the SEND-side evidence images when direct SEND files are not available.

Planned cases:

- **0-byte pattern**: identify the standard 0-byte image/screenshot pattern in Excel and mark it as a known comparison case.
- **Non-0-byte pattern**: detect grouped long screenshots, OCR the first 10+ lines and last 10+ lines, and parse row number, record length, and fixed-position content from stable screen positions.
- **Multi-picture groups**: support evidence where one SEND file spans several stacked screenshots.
- **Confidence and manual fallback**: write OCR confidence and parsed fields to a future metadata file, then keep the current manual Enter-to-mark flow when confidence is low.

Suggested future extension points:

- Add OCR parsing functions to `SendVsGift.ps1` without changing the `SendVsGift` mapping column semantics.
- Add a second output file such as `<WorkDir>\data\send_metadata.csv` with fields parallel to `gift_metadata.csv`.
- Compare `send_metadata.csv` and `gift_metadata.csv` in a later automated stage, but keep manual review as the safe fallback.

TODO: implement OCR/image recognition once representative SEND-side screenshots are available.
