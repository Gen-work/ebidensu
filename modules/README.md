# modules

Reusable, capability-oriented ebi-dance steps belong in the group directories
below. Every step must follow `docs/ebi-dance/spec/STEP-CONTRACT.md`.

Workflow ordering, project-specific values, and legacy implementations do not
belong here; use `workflows/`, `profiles/`, and `legacy/` respectively.

## Files that are not steps yet

The refactor parks pre-conversion libraries here before they are rewritten to
`spec/STEP-CONTRACT.md`. `Tests/Test-StepContract.ps1` lists them each run and
exempts them from the contract check, on two conditions: the file is not named
`<group>.<verb>.ps1`, and it defines neither `$Manifest` nor `Invoke-Step`.
Defining either one opts the file into the check whatever it is called, so a
real step cannot hide from the contract by being misnamed.

Being on that exemption list is a TODO, not a resting place.
