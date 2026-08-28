# flow

Steps that drive the run itself rather than the outside world belong here:
`flow.checkpoint` (write the worklist field plus the ledger entry) and
`flow.call` (inline a sub-workflow). Both are invoked through `use` like any
other step, so they need manifests and live here.

Not everything named `flow.*` is a step. `flow.foreach` and `flow.group_by`
(`docs/ebi-dance/spec/WORKFLOW-SCHEMA.md` sections 7.1 and 7.2) are runner
constructs -- `each` iterating `source`, and `source.groupBy` plus `once` --
named for symmetry but never written as a `use`. Do not create files for them.

Ordering policy and which field a checkpoint writes belong in the workflow
JSON, not in these implementations.
