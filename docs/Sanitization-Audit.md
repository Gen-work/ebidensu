# Sanitization Audit Report (2026-07)

Full-repo audit backing `docs/Generalization-Roadmap.md`'s S1-S9 checklist and
the branch/visibility strategy. Run 2026-07-08 with seven parallel audit legs:
docs/meta files, committed config + core libs, domain libs + snap phases, main
phase scripts + tests, git history, a whole-tree mechanical regex sweep, and a
completeness critic.

**This report deliberately masks the sensitive literals it describes** (IDs as
`JP2*****`, hosts as `proxy01.***`, share roots as `\\Fs-*`, people as "operator
surname" / "colleague A"). Locations are given as `file:line` so each item is
findable without repeating the value. Do not paste the raw values back into any
tracked file.

## 0. Headline: repository visibility (checked 2026-07-08)

| Remote | Visibility | Notes |
|--------|-----------|-------|
| GitHub `Gen-work/ebidensu` | **PUBLIC** at audit time | 0 forks / 0 stars / 0 watchers; created 2026-05-21. Everything in sections 1-2 below -- including git history -- was publicly reachable for ~7 weeks. |
| GitLab mirror (`Tokumei_M/...`) | private | Anonymous access redirects to sign-in; API returns 404. Office-PC sync pulls from here with credentials. |

**Immediate recommendation: flip the GitHub repo to private** until the M6
fresh-history export exists:

```
gh repo edit Gen-work/ebidensu --visibility private --accept-visibility-change-consequences
```

Verified consequences: none for daily work. The GitLab mirror Action runs inside
the repo and keeps working regardless of visibility; the office PC syncs from the
*private* GitLab mirror with credentials, so it is unaffected; home/worktree
pushes are authenticated. With 0 forks/stars there is no social cost. Reversible
with the same command (`--visibility public`).

## 1. Findings -- current tip

Severity: **high** = identifies a person/company/host; **medium** = internal
jargon, real production identifiers, or document names that hint at the client;
**low** = cosmetic / already-scrubbed leftovers.

### 1.1 High (all fixed in M0 except S3)

| Where | What | Status |
|-------|------|--------|
| `verify_session.json` (was tracked) | `C:\Users\JP2*****` work paths; `Owner` = operator surname kanji; `J4BaseDir`/`CloneSourceDir` = full internal UNC share (`\\Fs-*` + departmental hierarchy + internal request-ID/project folder name) | untracked in M0 (S2) |
| `.metadata/` (was tracked, 86 files) | Eclipse/RAD workspace: `.log` with corporate proxy FQDN + `C:\Users\JP2*****`; server configs; JDK/IDE install paths; workspace name ending in the client abbreviation | untracked in M0 (S1) |
| `.project` | Eclipse project name = 2-letter client abbreviation | untracked in M0 (S1) |
| `CLAUDE.md:7` | Client company named in full | **open -- M1 (S3)**: replace with neutral wording |
| `FillCheckSheet.ps1:10`, `DeliverMail.ps1:6`, `ReviewEvidence.ps1:4,228`, `Pack-LlmContext.ps1:8`, `README.md:327` | A colleague's given name hardcoded in code comments describing the manual-review flow, in the LLM context header string, and as a CLI example owner | **open -- M1 (S10)**: replace with "the operator" / neutral placeholder |

### 1.2 Medium (M1 edits vs M3 profile extraction)

| Where | What | Fix |
|-------|------|-----|
| `VerifyConfig.psd1:266-274` | Mail defaults: internal project name in `SubjectTemplate` (`【…廃止対応】`), route codes in `Mail.Phase`, real check-sheet filename in `CheckSheetFile` | M1: blank the committed defaults; move real values to `verify_config.example.json` comments / per-work-folder overlay |
| `README.md:484`, `CHANGELOG.md` (one entry) | The same mail subject template documented verbatim | M1: document with `【<ProjectName>】` placeholder |
| `README.md:309` | Example `-J4BaseDir` path reproducing the real internal share structure + route code | M1: fully generic example path |
| `README.md:327` | `-Owner <colleague-given-name>` used as a CLI example | M1: neutral placeholder (`-Owner op1`) |
| `docs/SnapVerify-Plan.md:405-440` | Verbatim Ctrl+A dumps of real HM/MQ production pages (screen IDs, program IDs, job names, timestamps) | M1: replace with synthetic dumps of the same shape (they are also test-fixture documentation, keep structure identical) |
| `docs/SendVsGift.md:122,142,226,251` | Real correl IDs from live debugging runs | M1: synthetic IDs |
| `CLAUDE.md` / `CHANGELOG.md` throughout | Real production job/correl IDs, IF numbers, one real GoAnywhere job number, real workbook names (`wipGFIX一覧.xlsx` etc.) | accept while private (S7); the M6 export ships fresh docs + changelog, not these |
| `VerifyConfig.psd1:11,40` + scripts | Internal tracking-workbook / sheet names as functional defaults | M3: becomes profile data (S6) |
| Script comments (`GfixLog.ps1:13`, `GfixLogDownload.ps1:14,144,177`, `Parse-GiftMq.ps1:15`, `ExcelSnap.ps1:78`, others) | Real job IDs / BIZ codes / internal shell paths in explanatory comments and examples | M3: sweep comments to synthetic IDs when each file moves layers |
| Usage examples (`VerifyTool.ps1:119-140`, `Generate-HostOpenMapping.ps1:20-100`, `Clone.ps1:22`, `Align.ps1:32`) | Real production job names / BIZ codes / share-structure fragments in `-Add`/clone/align examples | M1 (cheap) or M3: synthetic IDs |
| `Tests/` fixtures (`Test-GfixLog.ps1:26`, `Test-SnapVerify.ps1:28,53-57`, `Test-SendMetadata.ps1:152-166`, `Test-GfixJobList.ps1:19`, `Test-EvidencePlan.ps1:22`) | Fixture strings copied from real runs: internal shell path + real receive path, real program/node names, real dataset names (`L***.C.VER.*`), a real GoAnywhere list row | M3: re-synthesize fixtures when files move layers (values are load-bearing for tests -- change them together with expected assertions) |

### 1.3 Low (no action or covered elsewhere)

- `HmSnap.ps1:445` / `MqSnap.ps1:426`: application URLs already use
  `<hm-host>`/`<mq-host>` placeholders on the tip -- good; path fragments remain
  but identify nothing.
- Reviewer/mail identity fields in `VerifyConfig.psd1` and
  `verify_config.example.json` are all empty strings (the v2.9.20 scrub held; no
  e-mail addresses match anywhere in the tracked tree).
- `Sample-HighlightColor.ps1:5`: hardcoded `C:\workspace\<client-abbrev>` dev
  path -- parameterize whenever touched (M1 nice-to-have).
- `.github/workflows/mirror-to-gitlab.yml`: pseudonymous personal GitLab
  namespace hardcoded; token correctly injected via secrets. Keep, but drop the
  workflow from any M6 public export.
- WBS jargon, `J4` stage code, `GIFT`/`GFIX` terms: judged client-proprietary
  *terminology* but not re-identifying on their own; they exit `main` naturally
  in M3 (profile extraction) rather than by string-hunting now.

## 2. Findings -- git history (why it can never be published)

The history audit (two independent passes) confirmed, beyond the tip findings:

1. **Author identities**: most early commits are authored with the operator's
   real name + employee-ID corporate e-mail (immutable commit metadata; fixable
   only by rewriting every commit). Later commits use pseudonymous identities.
2. **Colleague PII**: parents of `9aa5886` ("Remove committed personal
   defaults", ~24 files) still expose a colleague's real name and internal
   address, the operator's internal address, and full internal file-server UNC
   paths in `VerifyConfig.psd1`/`SnapConfig.psd1`/`README.md` blobs.
3. **A real internal application FQDN** under the client's domain in historical
   `HmSnap.ps1` revisions (the tip has a placeholder).
4. **One commit message** names a colleague; one early commit message names the
   client project outright.
5. `verify_session.json` and `.metadata/` exist in many historical trees.

Rewriting requires `git filter-repo` over ~50 remote branches plus the GitLab
mirror, and cannot fix mirrored/indexed copies already fetched while the repo
was public. **Conclusion (roadmap S8/M6): the public artifact must be a
fresh-history repository seeded from a sanitized snapshot. This repository
stays private permanently.**

## 3. Module coupling classification

Agent-verified classification (feeds the roadmap section-4 layer map; the
roadmap table stays the canonical plan -- differences noted there):

| Classification | Files |
|----------------|-------|
| generic-core (reusable as-is) | `Common.ps1`, `ScreenRegion.ps1`, `OcrWindows.ps1`, `Locate-ByImage.ps1`, `Read-PageText.ps1`, `Find-ActiveHighlightRow.ps1`, `Crop-Snap.ps1` |
| generic-core (trivial cleanups) | `ProgressLog.ps1` (event field names mirror the mapping schema) |
| mixed -- generic mechanism worth extracting | `MappingStore.ps1` (hardcoded status-column list + key columns), `ConfigOverlay.ps1` (group definitions + readme text name client sections), `ExcelHelpers.ps1` (GFIX-log highlight rules + `Command:` pattern), `WorkbookResolver.ps1` (Excel_NAME/prefix conventions), `EvidenceImageExport.ps1` (`verifyMark_*` name filter), `EvidencePlan.ps1`, `EvidenceExecutor.ps1`, `AlignCompare.ps1`, `SnapVerify.ps1`, `SnapLocalize.ps1`, `SendMetadata.ps1`, `JenkinsDownload.ps1`, `DfSnap.ps1`, `Find-Abend.ps1` |
| domain-adapter (client/system rules) | `VerifyConfig.psd1`, `SnapConfig.psd1`, `ProjectLabels.ps1`, `OwnerFilter.ps1`, `GfixLog.ps1`, `GfixJobList.ps1`, `HmSnap.ps1`, `MqSnap.ps1`, `JenkinsSnap.ps1`, `GfixLogDownload.ps1`, `ExcelSnap.ps1`, `Run-Snap.ps1`, `Parse-GiftMq.ps1`, `Parse-JenkinsList.ps1` |
| tooling | `OcrTool.ps1`, `Calibrate-HmGeometry.ps1`, `Sample-HighlightColor.ps1` |

| Classification | Files (phase-script leg) |
|----------------|--------------------------|
| generic-core | `Probe-Shapes.ps1`, `Resolve-ExpectedTime.ps1`, `Read-ClipboardJson.ps1` |
| mixed | `VerifyTool.ps1` (menu/router generic; phase table + dispatch args are profile data), `Mark.ps1`, `Clone.ps1`, `ReviewEvidence.ps1`, `DeliverMail.ps1`, `BackupJ4.ps1`, `Watch-MappingProgress.ps1` |
| domain-adapter | `Generate-HostOpenMapping.ps1`, `Align.ps1`, `ReplaceEvidence.ps1`, `MarkGfixLog.ps1`, `SendVsGift.ps1`, `FillCheckSheet.ps1`, `DeliverFiles.ps1`, `Validate.ps1` |
| tooling | `Pack-LlmContext.ps1`, `Apply-LlmPatch.ps1`, `Export-DailyPatch.ps1`, `Fix-Encoding.ps1`, `Check-Encoding.ps1` |
| test | all 16 `Tests/` files (fixtures carry real-run strings -- see 1.2) |

Notable takeaway for M2/M3: the *snap* and *phase* scripts classify as
domain-adapters mostly because their sentinels, column layouts, sheet sets, and
verdict rules are inline -- exactly what the profile schema externalizes. Their
window/capture/keystroke/Excel mechanics are generic.

## 4. Strategy analysis: when to split public vs private

Two paths were considered for "make the sanitized result the real public
ebidensu, keep the personal one private":

| | Path A: split now | Path B: defer to M6 (recommended) |
|---|---|---|
| Prerequisite work | S3-S5 edits + fresh-history export + doc curation, immediately | flip GitHub private today (one command); everything else on the existing M1-M5 cadence |
| GitLab mirror | untouched either way (mirror Action is repo-internal; target private) | untouched |
| Office-PC sync | untouched (pulls the private GitLab mirror) | untouched |
| Risk | repo-name reuse: renaming this repo and creating a new public repo under the old name **breaks GitHub's rename redirect** -- any stale hardcoded remote (home clones, scripts) silently retargets to the new public repo | none now; same rename hazard applies at M6 and is documented there |
| Benefit | public identity exists early | split happens once, against a polished, fully-sanitized generic core; no pressure while the personal tool is still iterating |

Path B matches the operator's stated preference (keep iterating privately,
"拆" the project gradually once the personal tool is polished). The only
non-deferrable action is the visibility flip in section 0.

## 5. Coverage and residual items

- Completed legs (7 of 7): docs/meta, config + core libs, domain libs + snap
  phases, main phase scripts + tests, git history (run twice), whole-tree regex
  sweep (employee-ID patterns, e-mails, hosts/URLs, UNC, IPs, user-profile
  paths, reviewer/mail values). All deltas are folded into sections 1-3.
- Completeness-critic checks (run 2026-07-08 after the M0 untracking):
  - **no binary files remain tracked** (png/xlsx/zip/jar/db sweep of
    `git ls-files` is empty -- all previous hits were under `.metadata/`);
  - **`ProjectLabels.ps1` [char] labels decode clean**: all eleven labels are
    domain sheet/marker names (GIFT/GFIX receive-result sheets, send arrows,
    past-data note) -- no person/company/host content hidden in codepoints;
  - **window-activation titles are generic** (`Microsoft Edge`, `Excel`,
    process id) -- no internal application titles anywhere in `AppActivate`
    calls;
  - spot-verified that S3 (client name, `CLAUDE.md:7`) and S10 (colleague
    given name, 9 occurrences) are still present at tip -- both remain M1
    scope, nothing regressed;
  - `.github/workflows/` contains only the mirror workflow (section 1.3).
- Manual items no audit can close:
  - rotate nothing: no credentials/tokens were found in tree or history (the
    GitLab token lives in Actions secrets);
  - assume the 7-week public window means the history *may* have been crawled;
    that strengthens, not weakens, the fresh-history rule for M6;
  - `mark_templates/` must stay screenshot-free (now gitignored for `*.png`) --
    real template crops are per-office-PC artifacts.
