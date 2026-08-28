# progress

Reusable run-reporting steps belong here: end-of-run status summaries, pending
counts, and anything that tells the operator where a run stands. `progress.status`
is the canonical example (`docs/ebi-dance/spec/WORKFLOW-SCHEMA.md` section 8 calls
it from `teardown`).

Trace writing itself belongs to `kernel/`, not here, and the meaning of any
particular column or verdict belongs in a profile.
