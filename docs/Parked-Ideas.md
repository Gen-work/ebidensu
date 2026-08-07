# Parked ideas

Work that was designed (sometimes partly built) and then deliberately set
aside. Nothing here is on the active TODO list in `CLAUDE.md` — that list is
for what the project is working on **now**. Items move back out of this file
only when someone decides to pick them up again.

Each entry says what it was for, how far it got, and what it would take to
resume, so parking costs nothing but the reading.

---

## D2 image check for the ProcessTime 3↔9 verification

- **Design doc:** `docs/ProcessTime-OldSnap-MockMatch-Plan.md` (status: PARKED)
- **Code that exists:** `PixelDigitMatch.ps1` (pure scorer, unit-tested),
  `OldSnapPixelVerify.ps1` (GDI+ glue, static-checked only, OFF by default via
  `ProcessTime.OldSnapVerify.PixelDiff.Enabled`), `mock-page/` (node-only
  Phase-0 harness: `pixeldiff.mjs`, `compare.mjs`, HTML templates)
- **Parked:** 2026-08-04

### What it was

A pixel-level check that would confirm whether an OCR'd `3` or `9` really is
the digit shown in the snap image, by rendering a reference row and comparing
it against the captured screenshot. Two generations were designed:

1. **v2.17.0 (implemented, off by default):** GDI+ renders MS Gothic `3`/`9`
   per-digit templates; each digit box is cropped from the snap and scored by
   normalized cross-correlation.
2. **MockMatch plan (designed, never built):** render the reference row with
   **Edge** from `mock-page/templates/hm-batch-status.html` instead — same
   rasterizer, font and CSS as the real page — and whole-field template-match
   it via `Locate-ByImage.ps1`, removing both the engine mismatch and the
   per-digit coordinate calibration.

### Why it is parked

It never reached a usable state, and the cheaper checks landed first and turned
out to be enough for day-to-day work:

- Both generations need an **office PC** to calibrate (crop geometry per snap
  window size, or the Edge render + match threshold). That calibration session
  never happened, so `PixelDiff` has always shipped disabled.
- The deterministic checks now carry the load: impossible-position digit repair
  and the start/end/page-duration arithmetic disambiguation
  (`TimeDigitVerify.ps1`), plus the red conditional formatting that puts every
  ambiguous `3`/`9` in front of the operator. Those need no calibration and no
  image at all.

### To resume

1. Pick generation 2 (the Edge/mock-page one) — generation 1's engine mismatch
   is a known dead end, and `OldSnapPixelVerify.ps1` stays off meanwhile so two
   competing D2 paths never run at once.
2. Read `docs/ProcessTime-OldSnap-MockMatch-Plan.md` §10 for the 2026-07-27
   review and the intended landing order.
3. Book an office-PC session: render the reference, capture a matching snap,
   and calibrate the match threshold against known-good and known-bad rows.
4. `Get-OldSnapVerifyVerdict` already takes `-PixelResult` / `-PixelEnabled`,
   so wiring a working checker back in needs no change to the triage logic.

---

## `Repair-ProcessTimeStartFromStamp` (datestamp-based 3↔9 correction)

- **Where:** `OldSnapVerify.ps1`, unit-tested in `Tests\Test-OldSnapVerify.ps1`
- **Status:** implemented, **never called by any phase**
- **Parked:** 2026-08-04

Adopts the clean 14-digit datestamp's `HH:mm` for the start time when the OCR'd
start differs from it by nothing but a 3↔9 swap. It was written for the
benchmark plan and wired nowhere, so the "datestamp cross-check" the release
notes described was never actually running.

It is kept because the idea is sound and the function is tested, but it is
**not** the project's answer to 3↔9 any more: the arithmetic disambiguator in
`TimeDigitVerify.ps1` (`Resolve-ProcessTimeDurationConflict`) uses three
independent readings instead of two and refuses to act when the fix is not
unique, which is strictly safer. If anyone revives this, it should feed
`Get-OldSnapVerifyVerdict -DatestampSwap` as a **flag**, not rewrite a value on
its own.
