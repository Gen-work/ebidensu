# kernel

Shared ebi-dance execution infrastructure belongs here: context evaluation,
step registration, workflow execution, tracing, gates, and generated docs.

Business- or system-specific automation does not belong here. Put reusable
capabilities in `modules/` and declarative orchestration in `workflows/`.

## What is here today

- `Json.ps1` -- P1-35, the one place JSON is read or written. `Read-EbiJson`
  / `Write-EbiJson` (atomic: temp file + replace), `ConvertFrom-EbiJson` /
  `ConvertTo-EbiJson`, `Add-EbiJsonLine` / `Read-EbiJsonLines` (JSONL: trace,
  ledger), `ConvertTo-EbiHashtable`, `Test-EbiJsonSerializable`. Values come
  back as hashtables and `object[]`, never PSCustomObjects; depth is fixed at
  20 and a deeper value is refused loudly instead of truncated to a string;
  Japanese is written as characters; UTF-8 without BOM. Failures are records
  (`@{ ok; value; message }`); the depth guard in `ConvertTo-EbiJson` is the
  one exception and throws. Iron rule R8: nothing else under `modules/` or
  `kernel/` calls `ConvertFrom-Json` / `ConvertTo-Json` / `Get-Content` on a
  `.json` file (`Tests/StepContract.ps1` rule `direct_json`).
- `Trace.ps1` -- append-only `run/<runId>/trace.jsonl` (P0-03), over
  `Json.ps1`. `Read-TraceEvents` returns hashtables; a malformed line in the
  middle of a trace is reported once, a half-written final line is not.
- `Registry.ps1` -- P1-02, the step registry. Discovery
  (`Find-EbiStepFiles`: every `modules/<group>/<group>.<verb>.ps1`;
  `Get-EbiStepCatalog`: manifest-only scan in a throwaway scope, the list
  `ebi lint` / `ebi help` / `Docs.ps1` start from), loading for running
  (`Import-EbiStep`: dot-source, capture `Invoke-Step` at once, remove the
  bare name -- **must itself be dot-sourced**, `. Import-EbiStep -Registry
  $r -Use 'x.y'`, so the step's helper functions land in the caller's
  scope; any other call is refused), the `with` -> `$In` pipeline
  (`Resolve-EbiStepInputs`: `as` lifted out, `Test-EbiStepInputs` checks
  required / type / enum / default / unknown parameters and names each
  offending parameter, then session names become instances) and the
  section 3.1 return check (`Test-EbiStepReturn`). A bad call is
  `input_invalid`; a manifest the runner cannot check against (an input of
  unknown type) is `contract_violation`. `path` inputs are type-checked
  only; a relative path resolves under the work dir through
  `Native.ps1`'s `Resolve-EbiWorkPath`, which every step that touches a
  file calls (settled by P1-18/P1-20).
- `Worklist.ps1` -- P1-03, the in-memory worklist (`@{ path; columns;
  rows }`, a Session resource of kind `worklist`) and the ONE row filter
  `Select-EbiWorklistRows` that the runner's `source.select` and the
  `table.select` step share: the five `pendingWhen` forms, `verdict.values`
  translation both ways, bitmask by name or number, `Sort-EbiWorklistRows`
  (stable, groupBy then orderBy). Pure.
- `Runner.ps1` -- P1-03, the runner's main body (P0-07 spike grown up).
  `Invoke-EbiWorkflow` validates the workflow's shape (`workflow_invalid`;
  `schema: 1` is required), runs `setup`, then `each` once per selected
  worklist row, then `teardown` in a `finally`; expands `{{...}}` in `with`
  per call (`Context.ps1`), evaluates `when` (a skipped step has every
  output null plus `skipped=true`), replays `once: "group"` outputs to the
  later items of the group, ends only the failing item on a step failure
  and turns `operator_quit` into the reserved `cancelled`. Loads steps
  through `Registry.ps1`, owns `$Ctx` (including `$Ctx.Session`) and the
  resource channel of `docs/ebi-dance/spec/STEP-CONTRACT.md` section 3.4
  point 7. P1-04 adds `Invoke-EbiStepWithPolicy` around every call:
  ledger replay on `-Resume` (never for a `provides` / `releases` step,
  which always runs again), the confirm gate before a `destructive` step
  (unless `"confirm": false`), then attempts under the `onError` policy
  (`retry` only for a transient failure, doubling backoff, exhausted ->
  `ask`; `ask` r/s/q; `skip`; `fail`; `byFailure` and a per-call `onError`
  override). Questions go through one `-AskHandler` scriptblock whose two
  question shapes are documented on `Invoke-EbiDefaultAsk`; the default
  is a console prompt that answers itself under DryRun or when stdin is
  not a console. `once: "groupEnd"` runs after a group's last item. Every
  completed or when-skipped `each` step is appended to the ledger with its
  outputs; `run.json` holds `run.*` and the arguments. `-Profile` is passed
  in as a hashtable (loading `profiles/<name>/` is P2-01).
- `Gate.ps1` -- P1-05, the one ASCII gate panel: `Format-EbiGatePanel`
  (pure; WHAT HAPPENED / NEXT / EVIDENCE / ACTIONS in an 80-column box),
  `Read-EbiGateAnswer` (pure; r / s / q / `m <note>` / Enter = default /
  numbered choices), `Show-EbiGate` (render + read until valid; reader
  injectable; takes the `-Auto` action when nobody can answer) and
  `Invoke-EbiGateAsk`, the runner's default `-AskHandler`.
- `Docs.ps1` -- P1-06, manifests -> `docs/ebi-dance/CATALOG.md` (people)
  and `catalog.json` (Agents). Both are committed; `Tests/Test-Catalog.ps1`
  regenerates them and fails on drift. Regenerate after any manifest
  change: `. .\kernel\Docs.ps1; Write-EbiCatalog`.
- `Help.ps1` -- P1-07, `ebi help`: `Format-EbiHelpList` (every step by
  group, one line each) and `Format-EbiHelpStep` (one manifest in full),
  every line at most 80 characters.
- `Lint.ps1` -- P1-08, `ebi lint`: `Invoke-EbiLint` runs the static checks
  of `WORKFLOW-SCHEMA.md` section 9 over a workflow, the catalog and the
  profile without running anything (the runner's shape checks first, then
  use / with / templates / session resources / idempotency / byFailure /
  mustRelease; warnings for fallback tier, confirm:false, needs, missing
  profile). `Get-EbiMustReleaseKinds` parses the kind table out of the spec.
- `Explain.ps1` -- P1-09, `ebi explain`: `Format-EbiExplain` renders the
  execution plan in ASCII (Plan.md section 9's shape): page as
  role(label), source, one line per step with its effects tag and what it
  touches, a footer with onError / gates / destructive / fallback counts.
- `Profile.ps1` -- `Read-EbiProfile`: `profiles/<name>/*.json`, one file
  per top-level key, with `<WorkDir>/ebi.local.json` deep-merged on top
  (`Merge-EbiHashtable`); `Resolve-EbiProfileDir` turns a name or a path
  into the directory. The profile CONTENT is P2-01.
- `Ledger.ps1` -- P1-04, `run/<runId>/ledger.jsonl` (keys `item:<key>|<step>`
  and `group:<group>|<step>`, last record per key wins, appended through
  `Json.ps1`) and `run/<runId>/run.json` (the `run.*` scope, workflow
  id/version, arguments, `finished`, the result), plus
  `Find-EbiUnfinishedRuns` for `ebi run --resume` without a run id. `ebi.ps1 run|dryrun <workflow> -WorkDir <dir>` wraps it
  (the P1-10 seed); it can also be driven by hand:

  ```powershell
  . .\kernel\Runner.ps1
  Invoke-EbiWorkflow -Path .\workflows\x.json -WorkDir C:\work [-DryRun]
  ```

  The result is a hashtable (`ok`, `runId`, `steps`, `session`); the run's
  trace is at `<WorkDir>\run\<runId>\trace.jsonl`.


- `Context.ps1` -- P1-01, pure `{{...}}` template evaluation per
  `docs/ebi-dance/spec/WORKFLOW-SCHEMA.md` section 4: `New-EbiTemplateScope`,
  `Resolve-EbiPath`, `Expand-EbiTemplate`, plus `Test-EbiTemplateString` /
  `Get-EbiTemplateReferences` for `ebi lint`. Failures are records naming the
  unresolved segment, never exceptions. Also the four `when` forms of
  section 5 (`ConvertFrom-EbiWhen` / `Test-EbiWhen`, P1-03). Wired into
  `Runner.ps1` since P1-03.
- `Native.ps1` -- P1-11, the one Win32 / SendKeys / clipboard binding.
  `Get-EbiNative` compiles `EbiNative` lazily (nothing is touched on a dry
  run); `Set-EbiForeground` restores + `SetForegroundWindow` + **verifies
  with `GetForegroundWindow`** (one retry) and returns `@{ ok; message }`,
  which every key-sending step turns into `foreground_lost`;
  `Get-EbiWindowRect`, `Send-EbiKeys`, `Read-EbiPageText` (Ctrl+A, Ctrl+C,
  Esc, clipboard), `Invoke-EbiClick`, `Set-/Get-EbiClipboardText`,
  `Get-EbiVirtualScreen`; pure `Resolve-EbiWorkPath` (relative -> under the
  work dir, separators normalized) and `Write-EbiTextFile` (UTF-8, no BOM).
  Steps dot-source it with `. (Join-Path $PSScriptRoot
  '..\..\kernel\Native.ps1')` -- a step may not call another step, but
  every step may share a kernel library.
- `Image.ps1` -- P1-18/P1-20, the one GDI+ binding: `Save-EbiScreenRegionPng`,
  `Invoke-EbiCropPng` (per-side crop, atomic write; the legacy `-CropPx`
  + `-1`-inherits convention kept via `Resolve-EbiCropSides`, so HmSnap /
  MqSnap / JenkinsSnap / Crop-Snap call it in place of the four
  `Invoke-CropPng` copies they carried), `Get-EbiPngSize`; pure
  `Get-EbiCropGeometry`, `Resolve-EbiScreenRegion` (clamp + which edges
  moved). **Every impure entry point is a pure check plus a `*Core`
  function that alone names `System.Drawing` types**: on Linux pwsh the
  first call of any function whose body mentions `System.Drawing` throws
  `PlatformNotSupported` before a statement runs, so the dry-run and
  file-not-found branches must never share a function with GDI+ code.
- `Key.ps1` -- P1-27, the ONE place keys are normalized and compared:
  full-width folding, the `" / "` display form, the `_`-joined file-safe
  form (`Context.ps1`), and the matching half -- `Get-EbiKeyRules` (the
  profile's `confirmedRules`, or the default stamp-suffix / fullwidth /
  case-insensitive three), `ConvertTo-EbiKeyForm` per tier (exact >
  stripped > fullwidth > case; an undeclared rule's tier never yields a
  new hit), `Get-EbiKeyMatchTier`, `Get-EbiKeyPartsTier` (composite keys
  column by column), `Find-EbiKeyMatches` (the best tier that hit, every
  record in it), `Find-EbiKeySafeCollisions` (table.load's refusal) and
  `New-EbiCandidateList` (the P0-R4 shape every ambiguous step returns).
  `file.find`, `table.key`, `table.set`, `flow.checkpoint` and
  `verify.match_record` all call it; `Tests/Test-Steps.ps1` fails on a
  step that compares a key with its own `-eq`.
- `Table.ps1` -- P1-24, the worklist CSV: `Read-EbiCsv` (hashtable rows,
  header order kept), `Write-EbiCsvAtomic` (UTF-8 WITH BOM for Excel, CRLF,
  every field quoted, temp file + Move with retries -- the encoding is
  fixed here because `Export-Csv -Encoding UTF8` means BOM on PS 5.1 and no
  BOM on pwsh 7), `New-EbiWorklist`, `Save-EbiWorklist` (the flush every
  writing table step does before returning).
- `Parse.ps1` -- P1-30/P1-31, page text -> records: the four grammars of
  PROFILE-SCHEMA section 4 (`ConvertFrom-EbiGrammar`: delimited / labeled /
  columns / regex), every one returning the lines it did NOT recognise so
  `verify.parse_text` can report them, plus `ConvertTo-EbiDateTime` with
  the single-digit-hour formats (`H:mm:ss`) that the old parsers lacked.
