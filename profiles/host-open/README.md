# profiles/host-open

The Host -> Open migration evidence work, as ebi-dance profile data
(PROFILE-SCHEMA.md). P2-01 skeleton; grammars and rules for the two pages
whose text the repo already had fixtures for (P2-04), the other three pages
are shape only until real page text is tuned with `ebi grammar tune` (P2-03).

Column names in worklist.json are the LEGACY mapping CSV names
(GIFT_MQ_snap, GIFT_HM_snap, ...) with verdict.values ok=1 / ng=2 /
pending=0 / unknown='' -- the mixed-run rule (P0-R11): the new engine writes
the same cells the old phase scripts read.

Known holes, to fill on the office PC (never guess them here):
- pages.*.url is empty (the person opens the page; human.prepare shows openHint)
- fingerprint.expired / fingerprint.loading are empty for every page -- the
  session-expired page text is not in the repo
- fileList / jobList / reportPreview have no fixture yet
- layout.json pictures / boxes need `ebi probe` on the office PC (P4-20)

Fixtures are the synthetic page texts the old SnapVerify tests already used
(Tests/Test-SnapVerify.ps1); expected.json entries with a `#suffix` reuse a
file under another key / time window.
