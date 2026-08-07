# VerifyTool

[日本語](#日本語) | [English](#english)

---

## 日本語

ファイル転送基盤（GIFT→GFIX）移行案件で、証跡（エビデンス）取得作業を自動化するCLIツールです。

### 背景

移行案件では、HM/MQ/Jenkinsの画面をキャプチャしてExcelに貼り付け、該当セルに赤枠を引き、進捗をCSVで管理する証跡作成作業を手作業で行っていました。1件あたり約0.3人日かかっていました。

### 解決方法

PowerShellで実装しました。処理の流れは次のとおりです。

- 対象画面（HM/MQ/Jenkins）をキャプチャ
- エビデンス用Excelに画像を挿入
- 該当セルに赤枠を描画
- 完了状態をCSVマッピングファイルに記録

### 結果

1件あたりの作業時間が0.3人日から0.1人日未満へ短縮しました（同僚に確認した見積りによる）。

### 設計上の判断

「動くツール」と「現場に根づくツール」は別だと考え、以下を意識しました。

- **段階的な確認モード**: 最初から全自動にはせず、判定結果を都度確認できるモードを用意しました。誤検知を早期に発見できないと、利用者が結果を信用できなくなるためです。
- **カスタマイズの余地**: 枠の位置や文言を設定ファイルで調整できるようにし、案件ごとの違いや仕様変更に対応できる構成にしました。
- **利用者に使ってもらうまで**: 導入前に既存の手作業と結果を突き合わせ、差分がないことを確認したうえで展開しました。

### 使い方

```powershell
.\VerifyTool.ps1
.\VerifyTool.ps1 -Phase Status
```

---

## English

> **从一个人的证据工具，到可复用的工作流构建器。**
> *From one operator's evidence tool to a reusable workflow builder.*

A migration project produces evidence: screenshots of the batch-status page, of
the transfer queue, of the CI job that moved the file — each pasted into the
right sheet of the right workbook, annotated, cross-checked, logged, reviewed
and delivered. For a few hundred data-transfer jobs, that is weeks of careful
clicking, and a single mis-pasted screenshot is a defect that ships.

**VerifyTool turns that into one command.** It drives the browser and Excel on
the operator's own machine, captures each page, files it, marks it, verifies it
against what the page actually said, and tracks every job's state so the work
can stop and resume at any point.

It runs on a locked-down Windows remote desktop with **PowerShell 5.1 and Excel
2019 and nothing else installed** — no runtime, no package manager, no
third-party library.

### What it looks like

One entry point. It reads the current state, tells you what to do next, and
runs it:

```text
PS> .\VerifyTool.ps1

===== Current mapping status =====
  Rows: 257
  Mapping                : 257/257 done, 0 pending  [isMapped]
  GiftHmSnap             : 257/257 done, 0 pending  [GIFT_HM_snap]
  GiftMqSnap             : 251/257 done, 6 pending  [GIFT_MQ_snap]
  ReplaceGift            : 240/257 done, 17 pending  [isReplaced bit=1]
  MarkGift               : 0/257 done, 257 pending  [isMarked bit=1]
  ...

Recommended next: GiftMqSnap

Choose phase:
   1  InitConfig           work-folder config JSON (verify_config.json)
   2  Mapping              mapping 生成 / 更新
   5  GiftMqSnap           GIFT MQ 証跡
  16  ReplaceGift          GIFT 証跡置換 bit=1
  19  ProcessTime          処理時間 抽出 (証跡 Excel 生成) bit=3
  ...
   s  Status only
   h  Help
   q  Quit

phase [GiftMqSnap]:
```

*(The operator-facing labels are Japanese — this is a tool for a Japanese
site. The code itself is ASCII-only; see [Working on it](#working-on-it).)*

Every phase is reachable from this menu, remembers its own settings between
runs, and can be limited to specific jobs. Flags exist for scripting
(`-Phase`, `-TargetIds`, `-DryRun`, …) — see [`docs/Operations.md`](docs/Operations.md) —
but day to day you just run `.\VerifyTool.ps1` and answer prompts.

### The pipeline

State lives in one CSV (`mapping_<Owner>.csv`), one row per data-transfer job.
Each phase reads the rows still pending for it, does its work, and marks them.
Nothing is order-dependent beyond the arrows; interrupt anywhere and re-run.

```text
  Mapping ── generate the job list from the project's WBS workbook
     │
  Capture ── GIFT/GFIX × HM · MQ · Jenkins · DF diff  ─────► snap\*.png + page text
     │        (browser driven, each capture verified against the page's own text)
     │
  Clone ──── create one evidence workbook per deliverable
     │
  Align ──── compare against the J4 baseline before touching anything
     │
  Replace ── paste the captures into the right sheet, in review order
     │
  Mark ───── draw the red boxes on the cells a reviewer must look at
     │
  ProcessTime ─ read each batch's start/end time back off the evidence,
     │           cross-check it, and build the processing-time workbook
     │
  Review ─── walk a human through every job, one keystroke each
     │
  Deliver ── check sheet · J4 file transfer · Outlook review-request drafts
```

### Why it is built this way

The constraint that shaped everything: **the target machine cannot be changed.**
No admin rights, no installer, no internet, an old remote desktop session that
someone else owns. So:

| Need | What most tools reach for | What VerifyTool uses |
|---|---|---|
| Drive the browser | Selenium / Playwright | `SendKeys` + window activation, page text via the clipboard |
| Read text from a screenshot | Tesseract, a cloud OCR API | `Windows.Media.Ocr` — the engine already behind Snipping Tool |
| Write Excel | EPPlus, ClosedXML, openpyxl | Excel COM (the operator's own Excel) |
| Image work | ImageMagick, OpenCV | `System.Drawing` |
| Unzip | 7-Zip | `System.IO.Compression` |
| Email | an SMTP library | Outlook COM — as a **draft**, never auto-sent |

Total third-party dependencies: **zero**. Everything above ships with Windows.
Copy the folder onto the machine and it runs.

*(The one exception is `mock-page/`, a node-only test harness for developing
OCR logic off-site. It never runs on the operator's PC, and it is currently
[parked](docs/Parked-Ideas.md).)*

#### Testable where it matters

COM, `SendKeys` and screen capture cannot be exercised in CI. So every piece of
real logic is pulled out into a **pure module** — no COM, no I/O, no screen —
and unit-tested:

```powershell
.\Tests\Run-Tests.ps1     # parse-checks every .ps1, then runs the unit suites
```

Parsing an OCR'd table, deciding which log belongs to which job, resolving a
file name, planning what goes where in a workbook, judging whether a capture is
NG — all of it is pure and covered. The COM shells around them stay thin enough
to review by eye.

#### It says "I don't know" out loud

The tool automates evidence for an audit, so a confidently wrong answer is worse
than no answer. The rules it holds itself to:

- **Only act when the answer is forced.** The Japanese OCR engine confuses
  MS Gothic `9` and `3`. A digit is corrected only when exactly one substitution
  is arithmetically possible — a minute of `93` can only ever have been `33`; a
  start/end pair that disagrees with the page's own printed duration is fixed
  only if exactly one swap reconciles all three readings. Anything ambiguous is
  **left exactly as read and marked red** for a human. It never "improves" a
  plausible value.
- **Preview before touching shared documents.** The team's review check sheet is
  filled in a temp copy first, shown to you, and committed only if the original
  has not changed meanwhile.
- **Never send.** Review-request mail opens as a draft; a human clicks Send.
- **Atomic state.** The mapping CSV is written atomically, and every phase
  appends to an append-only `status\progress.jsonl` you can tail live from
  another shell without locking anything.
- **Verify the capture, not just take it.** Each screenshot is checked against
  the page's own text — right page, right job, right time window, no abend —
  and a failed check keeps the row pending instead of banking a bad screenshot.

### Status

This is a **working tool in daily production use**, and deliberately a
**readable sample project**: one real, messy, end-to-end workflow, automated
under real-world constraints, with the reasoning left in the code.

It is currently specific to one migration project. A properly decoupled
version — where the phases, the page definitions and the workbook layout are
described in configuration rather than written into the scripts — is in active
development; see [`docs/Generalization-Roadmap.md`](docs/Generalization-Roadmap.md)
for where that is going.

Already generalized: everything project-specific that could be moved out lives
in `VerifyConfig.psd1`, and any work folder can override all of it with a
`verify_config.json` overlay — window sizes, red-box coordinates, sheet names,
mail templates, expected-time rules, output routing.

### Documentation

| | |
|---|---|
| [`docs/Operations.md`](docs/Operations.md) | day-to-day reference: every phase, its options, the folder layout, the mapping columns |
| [`docs/Configuration.md`](docs/Configuration.md) | the three config layers and which one a setting belongs in |
| [`docs/Generalization-Roadmap.md`](docs/Generalization-Roadmap.md) | where the decoupled version is heading |
| [`docs/SnapVerify-Plan.md`](docs/SnapVerify-Plan.md) | how a capture is judged OK or NG |
| [`docs/SendVsGift.md`](docs/SendVsGift.md) | send-vs-GIFT metadata comparison |
| [`docs/Versioning.md`](docs/Versioning.md) | version bump rules |
| [`docs/Parked-Ideas.md`](docs/Parked-Ideas.md) | designed, then deliberately shelved — and what it would take to resume |
| [`CLAUDE.md`](CLAUDE.md) | full architecture map; read this first when opening the repo in an IDE or an LLM session |
| [`CHANGELOG.md`](CHANGELOG.md) | per-version history |

### Working on it

```powershell
.\Tests\Run-Tests.ps1        # parse check + unit tests
.\Check-Encoding.ps1         # enforce the source encoding policy
```

Two rules that are easy to trip over:

- **`.ps1` source stays ASCII.** Japanese used at runtime is built from code
  points in `ProjectLabels.ps1`. Raw Japanese in a BOM-less script mojibakes on
  a Japanese-locale host, which once silently broke owner matching.
- **Only files without a `param()` block are ever dot-sourced.** Dot-sourcing a
  script that has one overwrites the caller's switch parameters with `$false`.

Both are checked; `CLAUDE.md` has the full conventions.

#### Cross-environment workflow

Development happens away from the office PC, which has no git access:

```text
office PC  →  Pack-LlmContext.ps1    → clipboard → paste into an LLM session
LLM        →  XML patch or git diff  → clipboard → Apply-LlmPatch.ps1 → local files
end of day →  Export-DailyPatch.ps1  → clipboard → commit from home
```
