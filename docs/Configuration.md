# Configuration layering — where each setting lives, and why

VerifyTool reads configuration from three files. Each one has a distinct
role; mixing them up is how "some settings are in the work folder's JSON,
some in the repo's .psd1, some remembered invisibly in the session file"
happens. This document is the single reference for which file owns what.

## The three layers

| File | Location | Scope | Role |
|------|----------|-------|------|
| `VerifyConfig.psd1` | repo (code folder) | shipped with the tool | **Defaults only.** Every field the tool understands, with a safe/neutral default. NO personal or site-specific value is ever committed here (paths, mail addresses, prefixes are blank). |
| `verify_config.json` | **WorkDir** (work folder) | per project / work folder | **The single source of truth for everything project-scoped.** Sparse JSON, deep-merged over the .psd1 at startup (JSON wins). Generate/repair/edit with `-Phase InitConfig`. |
| `verify_session.json` | repo (code folder), gitignored | per machine + operator | **Ephemeral convenience state only**: the pointer to the last WorkDir, last Owner, window size, cursor cell, and machine-local tool paths (`DfExePath`). Losing this file must never lose project configuration — it only re-asks a question once. |

Precedence at runtime: **CLI argument > `verify_config.json` (WorkDir) >
`VerifyConfig.psd1` > `verify_session.json` fallback.**

(One deliberate exception: `DfExePath` reads session before config —
df.exe's install path is a property of the PC, not of the project, and the
session file is the per-machine store.)

## Best practice (the rule of thumb)

Ask: *if the operator zipped up the work folder and moved to another PC —
or opened a second work folder on the same PC — which values must travel
with the folder, and which must not?*

- **Travels with the work folder → `verify_config.json`.**
  Check-sheet path, J4 evidence folder, workbook prefix, reviewer/mail
  fields, Clone source, Align J4 baseline, Mark box offsets tuned for this
  project's evidence, SnapVerify options.
- **Property of the PC → `verify_session.json`** (or a real machine config
  if one ever exists). df.exe path, window size for this monitor.
- **Property of the tool → `VerifyConfig.psd1`.** Fallback defaults,
  phase order, aliases, script names. Never a real path/address.

First-run prompts should persist their answer into the layer the value
belongs to. As of v2.10.7 the CheckSheet path prompt writes
`CheckSheet.Path` into the work folder's `verify_config.json`
(via `Save-ConfigOverlayValue`, which preserves everything else in the
file including `_README`/`_SCHEMA`); it falls back to the session file
only when the JSON cannot be written. The DfSnap prompt stays on the
session file on purpose (machine-scoped).

## Field inventory — paths, prefixes, and who owns them

| Setting | Canonical home | Session fallback? | Notes |
|---------|----------------|-------------------|-------|
| `WorkDir` | session (it IS the pointer to the work folder) | — | prompted when absent |
| `Owner` | session | — | mapping file suffix |
| `Workbook.ExcelPrefix` | work JSON | no | CLI `-ExcelPrefix` wins; legacy per-row mapping `Excel_Prefix` still overrides per row |
| `CheckSheet.Path` | work JSON | legacy read fallback | first-run prompt now writes it here (v2.10.7) |
| `J4EvidenceDir` (top-level) | work JSON | no | read via `Get-ConfigJ4EvidenceDir`; legacy `DeliverFiles.J4EvidenceDir` / `Mail.EvidenceFolder` still win when non-empty |
| `Address` (top-level, reviewer To) | work JSON | no | legacy `Reviewer.Address` still wins when non-empty |
| `Clone.SourceDir` | work JSON | yes (legacy) | CLI `-CloneSourceDir` wins |
| `Align.J4BaseDir` | work JSON | via CloneSourceDir fallback | CLI `-J4BaseDir` wins |
| `Df.ExePath` / `DfExePath` | **session** (machine-scoped) | primary store | `Df.ExePath` in config locks it; `Df.DefaultExePath` seeds the prompt |
| `EvidenceDir` | derived: `<WorkDir>\evidence` | display only | relative override via `Review.EvidenceDir` in config |
| Window size / `CursorCell` | session | primary store | per-monitor / per-operator |
| Mail templates, `Mail.CheckSheetFolder/File` | work JSON | no | committed .psd1 keeps only generic templates |
| `Mark.Boxes`, `Mark.TemplateDir` | work JSON (per-project tuning) | no | .psd1 carries the calibrated GIFT/GFIX defaults |
| `SnapVerify.*`, `ExpectedTime.*`, `GfixLog.*` | work JSON overrides; .psd1 defaults | no | |

## Known sharp edges

- **`Mark.Boxes.<folder>` entries are arrays** — the JSON overlay replaces an
  array wholesale, and InitConfig repair treats a whole array as one atomic
  field. New keys added to a box in the .psd1 (e.g. `BaseRow`/`RowHeight`)
  do NOT backfill into an overlay that already carries that box; add them by
  hand or delete the box from the JSON to fall through to the default.
- **The session file is global.** A per-project value that only lives there
  (from an old version, or a menu override) silently leaks into the next
  work folder that has no explicit config for it. If a phase picks up a path
  you don't recognize, check `verify_session.json` first, then set the real
  value in that work folder's `verify_config.json` to shadow it.
- **New top-level .psd1 sections must be added to `Get-ConfigOverlayGroups`**
  (and mentioned in `Get-ConfigOverlayReadmeText`) in the same change, or
  they are invisible in the InitConfig grouped editor —
  `Tests\Test-ConfigOverlay.ps1` has a drift guard.
