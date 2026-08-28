# kernel

Shared ebi-dance execution infrastructure belongs here: context evaluation,
step registration, workflow execution, tracing, gates, and generated docs.

Business- or system-specific automation does not belong here. Put reusable
capabilities in `modules/` and declarative orchestration in `workflows/`.
