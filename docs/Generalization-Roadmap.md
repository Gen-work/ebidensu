# Generalization Roadmap -- from VerifyTool to a generic evidence workbench

This document turns the vision in `README.md` ("未来展望: 从个人证据工具到可复用工作流构建器")
into a concrete, staged plan: move the work-specialized version to a branch, keep
`main` sanitized and progressively generic, and grow toward an AI-assisted,
profile-driven evidence automation workbench.

Status of each milestone is tracked here. Update this file whenever a milestone
lands.

## 1. Guiding principles

1. **Pure-function core, fixture-first testing.** Every rule that can be tested
   without COM/browser/network lives in a `param()`-less dot-source library with
   unit tests (the repo already does this well -- keep it the law).
2. **Minimal dependencies.** Windows PowerShell 5.1 + built-in .NET only. No
   modules to install, nothing that needs admin rights on an office PC.
3. **Profile over code.** Project-specific knowledge (sheet names, labels, page
   sentinels, NG rules, box offsets, phase pipeline) becomes *data* in a project
   profile, not `if`-branches in shared code. The existing `verify_config.json`
   overlay + `VerifyConfig.psd1` section split is the seed of this.
4. **Operator safety by default.** Approval gates, backups, append-only progress
   logs, never auto-send, never silently overwrite -- as spelled out in the
   README vision section.
5. **Sanitized main.** No person, client-company, hostname, or internal share
   path identifiers on the `main` tip. Client-flavored *defaults* live in
   per-work-folder overlays or the spec branch, not in committed defaults.

## 2. Branch & repository strategy

### Audit facts (2026-07 repo audit)

Full findings, masked excerpts, and the module-coupling classification live in
`docs/Sanitization-Audit.md`. Headline addition after the first pass: the
GitHub repo itself was **public** at audit time while the GitLab mirror is
private -- see that report's section 0 for the flip-private recommendation and
why it does not affect the mirror or the office-PC sync.

- **Git history is not sanitizable in place.** Most early commits are authored
  with a real name + employee-ID corporate e-mail; blobs before commit
  `9aa5886` ("Remove committed personal defaults") contain a colleague's real
  name and internal address, internal file-server UNC paths, and internal mail
  defaults; one commit message names a colleague. Rewriting would mean
  `git filter-repo` across ~50 remote branches plus the GitLab mirror --
  disproportionate and error-prone.
- **The GitHub Action `mirror-to-gitlab.yml` force-mirrors every branch** to a
  personal GitLab account; whatever is pushed here replicates there.
- The tip carried a committed Eclipse/RAD workspace (`.metadata/`, with an
  internal proxy hostname and `C:\Users\<employee-id>` paths) and the live
  `verify_session.json` (real work paths, owner name, internal UNC share).
  Both were untracked in the M0 hygiene commit.

### Strategy

| Line | Role | Visibility |
|------|------|------------|
| `spec/gift-gfix` | Frozen snapshot of the last fully-specialized state (v2.10.4) + hotfix line for daily production work during the transition. No feature work here. | private |
| `main` | Single line of development. Tip-sanitized (M1), then progressively restructured into core/adapter/profile layers (M2+). The GIFT/GFIX logic does not leave `main` -- it becomes the first *profile*. | private |
| future public repo | If/when the generic workbench is worth publishing: seed a **new repository from a sanitized snapshot** (fresh history, curated file set, neutral author identity). **Never publish this repository's history.** | public |

Rationale: a solo operator cannot afford two diverging feature lines. `spec/*`
is insurance + hotfix space only; everything else happens on `main`, where
"de-specialization" and "modularization" are the same refactor. History
sanitization is deferred to the fresh-repo export (M6) instead of an in-place
rewrite.

### Effect of the M0 hygiene commit on existing clones

Pulling a commit that stops tracking `.metadata/` and `verify_session.json`
will try to **delete those files** in any clone where they are unmodified, and
will raise modify/delete conflicts where they are modified (likely, since
Eclipse and VerifyTool both touch them):

- `verify_session.json`: safe to lose -- VerifyTool regenerates it and re-asks
  WorkDir/Owner once.
- `.metadata/`: close Eclipse/RAD first; if the workspace state matters, copy
  `.metadata` aside before pulling and restore it after (it is gitignored, so
  restoring it will not re-track it).

## 3. Sanitization checklist (main tip)

Severity from the audit: **high** = identifies a person/company/host; **medium**
= internal jargon or document names that hint at the client.

| # | Item | Where | Severity | Status |
|---|------|-------|----------|--------|
| S1 | Committed Eclipse workspace (`.metadata/`: proxy host, user paths, RSE prefs) | `.metadata/` | high | **done (M0)** -- untracked + gitignored |
| S2 | Live session state (user paths, owner kanji, internal UNC share) | `verify_session.json` | high | **done (M0)** -- untracked + gitignored |
| S3 | Client company named in project docs | `CLAUDE.md` line 7 | high | M1 -- replace with neutral wording |
| S4 | Internal project name in mail defaults (`【GIFT廃止対応】...`), route codes (`JRV→IDS,IGP_J4`), check-sheet filename default | `VerifyConfig.psd1` Mail/CheckSheet sections | medium | M1 -- blank the committed defaults; ship them in `verify_config.example.json` as commented examples instead |
| S5 | Real HM screen Ctrl+A dump (screen ID, correl IDs) in design doc | `docs/SnapVerify-Plan.md` | medium | M1 -- anonymize the sample (fake screen ID / correl IDs, keep the shape) |
| S6 | Internal workbook/sheet names (`wipGFIX一覧.xlsx`, `GFIX送受信一覧`) in defaults and docs | `VerifyConfig.psd1`, `CLAUDE.md`, `README.md`, scripts | medium | M3 -- moves into the gift-gfix profile naturally |
| S7 | Client-order terms across CHANGELOG/CLAUDE prose (J4, GIFT廃止対応 mentions) | `CHANGELOG.md`, `CLAUDE.md` | medium | accept while private; the M6 export ships fresh docs, not this changelog |
| S8 | Git history (author identities, pre-9aa5886 blobs, colleague names) | all history + GitLab mirror | high | unfixable in place -- covered by the M6 fresh-repo rule; repo stays private until then |
| S9 | Regression guard: nothing re-introduces person/host/company identifiers | CI / Tests | -- | M1 -- `Check-Sensitive.ps1` scanner + test hook |
| S10 | Colleague given name hardcoded in comments / LLM header / CLI example | `FillCheckSheet.ps1`, `DeliverMail.ps1`, `ReviewEvidence.ps1`, `Pack-LlmContext.ps1`, `README.md` | high | M1 -- replace with neutral wording |

`Check-Sensitive.ps1` (M1) scans tracked files for: employee-id patterns
(`JP\d{6}`), corporate mail domains, UNC server shares (`\\Fs-*` and general
`\\host\share`), `C:\Users\<id>` paths, proxy/internal hostnames, and a small
denylist of person names -- wired into `Tests\Run-Tests.ps1` so a regression
fails the build. (Same spirit as the existing `Check-Encoding.ps1`.)

## 4. Target architecture

### Layers

```
core/       zero domain knowledge, param()-less dot-source libs
engine/     phase runner, plan walker, mapping store -- generic mechanisms
adapters/   one folder per evidence source: how to drive/parse a system
profiles/   per-project DATA: labels, sheet sets, boxes, NG rules, pipeline
tools/      dev/meta utilities (LLM bridge, encoding checks, calibration)
Tests/      unchanged location; fixtures stay client-clean
```

### Current file -> layer map

| Layer | Files (today) | Notes |
|-------|---------------|-------|
| core | `Common.ps1` (WinAPI/screenshot/SendKeys), `ExcelHelpers.ps1`, `OcrWindows.ps1`, `EvidenceImageExport.ps1`, `ScreenRegion.ps1`, `Locate-ByImage.ps1`, `ConfigOverlay.ps1`, `ProgressLog.ps1`, `WorkbookResolver.ps1`, `Read-PageText.ps1`, `Find-ActiveHighlightRow.ps1` | already domain-free or nearly so |
| engine | `VerifyTool.ps1` (menu/router/status), `MappingStore.ps1`, `EvidenceExecutor.ps1`, `Watch-MappingProgress.ps1`, `Validate.ps1` | mapping columns are profile-declared data; the bitmask mechanics are generic |
| adapters | `HmSnap.ps1`, `MqSnap.ps1`, `JenkinsSnap.ps1`, `JenkinsDownload.ps1`, `GfixLogDownload.ps1`, `DfSnap.ps1`, `ExcelSnap.ps1`, `Read-ClipboardJson.ps1`, page parsers inside `SnapVerify.ps1` | capture/drive mechanics are generic per system *type* (host terminal page, MQ status page, Jenkins list, GoAnywhere list, file-diff tool); their sentinels/fields come from the profile |
| profile: gift-gfix | `ProjectLabels.ps1`, `GfixLog.ps1`, `GfixJobList.ps1`, `EvidencePlan.ps1` (review order), `AlignCompare.ps1` (sheet sets), `SendMetadata.ps1` (verdict rules), `OwnerFilter.ps1` (WBS conventions), `Generate-HostOpenMapping.ps1`, `Mark.Boxes`/`PhaseOrder`/`Aliases` config, mapping column schema | the actual client-specific knowledge; today interleaved with the generic mechanisms above |
| mixed (to split in M3) | `Mark.ps1`, `ReplaceEvidence.ps1`, `ReviewEvidence.ps1`, `SendVsGift.ps1`, `Align.ps1`, `Clone.ps1`, `FillCheckSheet.ps1`, `DeliverMail.ps1`, `DeliverFiles.ps1`, `BackupJ4.ps1`, `SnapVerify.ps1`, `SnapLocalize.ps1`, `Run-Snap.ps1` | generic mechanism (draw box, paste picture, walk plan, draft mail, sync sheets) entangled with gift-gfix rules (which sheets, which columns, which labels) |
| tools | `Pack-LlmContext.ps1`, `Apply-LlmPatch.ps1`, `Export-DailyPatch.ps1`, `Check-Encoding.ps1`, `Fix-Encoding.ps1`, `Probe-Shapes.ps1`, `Calibrate-HmGeometry.ps1`, `Sample-HighlightColor.ps1`, `Find-Abend.ps1`, `Crop-Snap.ps1`, `Parse-GiftMq.ps1`, `Parse-JenkinsList.ps1`, `Resolve-ExpectedTime.ps1`, `OcrTool.ps1`, `MarkGfixLog.ps1` | calibration & dev aids; some retire into adapters |

### The project profile

A profile is a folder, not a branch: `profiles/<name>/` holding

- `profile.json` -- formalized successor of today's `verify_config.json`
  overlay: phase pipeline (`PhaseOrder`), mapping column schema, page
  sentinels/expected fields/NG rules per adapter, sheet sets, labels
  (Japanese text lives in JSON, which is encoding-safe -- this also retires
  much of the `[char]` gymnastics), box definitions, delivery rules.
- optional `rules.ps1` -- pure functions for verdicts too complex for data
  (e.g. today's `Test-HmAbend`, 0-byte SendVsGift rules), unit-tested.
- `fixtures/` -- anonymized page-text samples backing the rule tests.

`VerifyConfig.psd1` keeps only tool-level defaults (timing, window, paths
schema); everything client-flavored moves into the profile. Precedence stays
CLI > work-folder overlay > profile > tool defaults.

## 5. AI entry point (Describe phase)

Goal: a new project becomes usable without reading the codebase.

1. **Interview + samples in, profile out.** A `Describe` phase collects: what
   pages exist (operator captures each once -- Ctrl+A text via the existing
   `Read-PageText.ps1` and a screenshot), what "OK" looks like, key identifiers,
   time-window rules, target Excel layout. The AI (via the existing clipboard
   LLM bridge `Pack-LlmContext.ps1`/`Apply-LlmPatch.ps1` for offline offices, or
   an API/CLI where allowed) drafts `profile.json` + parser fixtures from those
   samples.
2. **Plan preview before automation.** The tool renders the proposed pipeline
   (phases, adapters used, mapping columns, approval gates) as text; the
   operator approves before anything runs -- mirroring the README vision's
   "reviewable execution plan".
3. **Fixture-first validation.** Generated fixtures immediately become unit
   tests, so a profile can be largely verified at home without the office
   environment -- the same trick the repo already uses for SnapVerify/GfixLog.

## 6. Milestones

Each milestone is one PR: tests green (`Tests\Run-Tests.ps1`), CHANGELOG entry
per `docs/Versioning.md`, this file's status column updated.

| # | Scope | Size | Status |
|---|-------|------|--------|
| M0 | Hygiene: untrack `.metadata`/`.project`/`verify_session.json`, add `.gitignore`; this roadmap; create `spec/gift-gfix` snapshot branch | small | **this PR** |
| M1 | Tip sanitization: S3/S4/S5 edits; `Check-Sensitive.ps1` + test hook (S9) | small | open |
| M2 | Folder restructure into `core/ engine/ adapters/ profiles/ tools/` -- file moves + `Scripts` path table + dot-source paths only, zero behavior change | medium | open |
| M3 | Profile extraction: move gift-gfix literals (labels, sheet sets, SS_CODE rules, Mark.Boxes, mapping schema, job-name regexes) into `profiles/gift-gfix/`; engine/adapters read the profile | large, several PRs | open |
| M4 | Proof of decoupling: a `profiles/demo/` built only from anonymized fixtures + engine contract docs (`docs/ProfileSchema.md`) | medium | open |
| M5 | AI onboarding: `Describe` phase (interview -> draft profile -> fixtures -> plan preview) | large | open |
| M6 | Public export decision: curated sanitized snapshot into a fresh-history repo; this repo stays the private spec home | small mechanically, needs S1-S9 all green | open |

### Working agreement during the transition

- Production hotfixes: land on `main` as usual **and** cherry-pick to
  `spec/gift-gfix` only if `main` has already diverged in that area; once M2
  restructures folders, hotfix on `spec/gift-gfix` first, then port.
- No new client-flavored defaults in committed files (S9 guard enforces).
- Every M2/M3 PR must keep `Tests\Run-Tests.ps1` green and the office smoke
  flow (Status -> one snap phase -> one mark phase) intact.

## 7. Risks

- **Two-line drift** (spec vs main): mitigated by hotfix-only policy on spec.
- **M2 restructure breaking office runs**: it is a mechanical move, but the
  office PC runs from a synced folder -- schedule the restructure right after a
  delivery lull; keep `spec/gift-gfix` as the rollback.
- **Encoding regressions when moving Japanese into JSON**: JSON is UTF-8
  no-BOM by policy and `ConvertFrom-Json` handles it on any codepage -- safer
  than `.ps1`, but keep `Check-Encoding.ps1` in the loop.
- **The GitLab mirror**: everything pushed replicates there; the fresh-repo
  export (M6) must be a *new* remote, not a branch of the mirrored repo.
