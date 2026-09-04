# legacy

Retired implementations retained temporarily for migration or historical-data
cleanup belong here. Legacy code is excluded from the ebi-dance step catalog
and must not be referenced by new workflows.

Each migrated component must document its retirement condition. New reusable
capabilities belong in `modules/`, never here.

## Why these four are here

They all exist to serve one finite job: clearing the backlog of OLD HM
snapshots that were captured as a low-resolution PNG only, with no immune
Ctrl+A page text next to them. Those rows have to be read by OCR, and the
Japanese recognizer misreads MS Gothic `9` as `3`. Everything below is
scaffolding around that single defect.

**Once that backlog is cleared, delete them.** They are not a capability the
tool keeps -- they are a debt it is paying off. Per `docs/ebi-dance/Plan.md`
section 6.3, the one idea worth keeping is generalized instead, as the
`verify.crosscheck` step: read a fact from every source that carries it, and
when the sources disagree, stop and ask a human rather than picking one.

| File | State today | Retirement condition |
|------|-------------|----------------------|
| `TimeDigitVerify.ps1` | Live. Dot-sourced by `ProcessTime.ps1`; ships the v2.21.0 deterministic 3/9 rules and the red conditional format. | No `ProcessTime` run still falls back to OCR, i.e. every remaining correl has a `<correl>.txt`. Its cross-reading rule is superseded by `verify.crosscheck`. |
| `OldSnapVerify.ps1` | Live. Dot-sourced by `ProcessTime.ps1`; supplies the D1 hyperlink and the verification column verdict. | Same as above -- the verification column exists only to triage OCR'd rows. |
| `PixelDigitMatch.ps1` | Parked, off by default (`OldSnapVerify.PixelDiff.Enabled`). See `docs/Parked-Ideas.md`. | Delete outright unless the D2 image check is ever revived; it never had its office-PC calibration session. |
| `OldSnapPixelVerify.ps1` | Parked, off by default. GDI+ glue for the above. | Same as `PixelDigitMatch.ps1`. Do not run it alongside a competing D2 path. |

## Rules while they are still here

- Not in the step catalog. No new workflow may declare a dependency on them.
- Do not extend them. A bug that only affects the old-snap backlog gets the
  smallest possible fix; anything larger is a signal the backlog should be
  cleared instead.
- Their unit tests stay in `Tests/` (`Tests/Run-Tests.ps1` keeps running
  them) so a move never silently drops coverage. Delete the test with the
  file.
