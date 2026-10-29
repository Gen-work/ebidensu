# Versioning policy

This repository uses practical semantic versioning for operator-visible releases.
Each `CHANGELOG.md` release heading carries its version, for example `v2.9.29`.
A branch never writes that number itself: it writes `vNEXT` and main assigns
the number on merge (see "Who assigns the number" below).

## Version format

Use `MAJOR.MINOR.PATCH`:

- `MAJOR`: incompatible workflow changes, large rewrites, or brand-new feature
  families that significantly change how operators use the tool.
- `MINOR`: backward-compatible feature additions, new phases, new automation
  flows, or sizeable improvements to existing phases.
- `PATCH`: backward-compatible bug fixes, small feature improvements, docs,
  config/readme clarifications, and low-risk internal refactors.

In other words, prefer:

- small fixes / small improvements: `x.x.(x+1)`
- meaningful new compatible capability: `x.(x+1).0`
- large refactor, incompatible behavior, or a new product-level capability:
  `(x+1).0.0`

Do not use `x.x+1.x` as a generic "big phase" rule. In semantic versioning,
`x.(x+1).0` means a MINOR bump and should reset PATCH to `0`. Likewise, a MAJOR
bump should reset both MINOR and PATCH to `0`.

## Decision guide

Choose the smallest version bump that honestly describes the change:

| Change type | Bump | Example |
| --- | --- | --- |
| typo/docs only, tests only, safe bug fix | PATCH | `2.9.29 -> 2.9.30` |
| small backward-compatible enhancement to an existing phase | PATCH | `2.9.29 -> 2.9.30` |
| new option or config key that keeps defaults compatible | MINOR or PATCH, depending on user impact | `2.9.29 -> 2.10.0` for notable capability; `2.9.29 -> 2.9.30` for tiny helper |
| new phase or notable workflow automation | MINOR | `2.9.29 -> 2.10.0` |
| breaking config/CLI behavior, renamed/removing phases, migration required | MAJOR | `2.9.29 -> 3.0.0` |
| large rewrite with same interface and same operator behavior | MINOR if risky/notable, otherwise PATCH | `2.9.29 -> 2.10.0` |
| experimental/internal-only work not released to operators | no version bump until release | keep current version |

When in doubt, optimize for operator expectations rather than code size: if the
operator must learn a new workflow, use MINOR; if old commands/configs stop
working, use MAJOR.

## Changelog rules

Every released change should add a new top entry to `CHANGELOG.md`, with the
bump KIND instead of a number:

```markdown
## YYYY-MM-DD - Short release title (vNEXT patch)

### Added
- ...

### Fixed
- ...

### Notes
- ...
```

`(vNEXT patch)`, `(vNEXT minor)` or `(vNEXT major)` -- the decision guide above
picks the kind; a bare `(vNEXT)` means patch. Keep sections that apply and omit
empty sections. Put the newest release at the top. If multiple commits are part
of one unreleased release, keep updating the same top entry instead of creating
many tiny version headings.

Commit messages and PR titles carry no version number: the number does not
exist until the merge.

## Who assigns the number

Branches used to pick their own number. Two branches cut from the same main
then picked the same one, and both merged: `v2.41.1` and `v2.41.2` each head
two entries in `CHANGELOG.md`. Those two duplicates stay as they are (an old
delivery is not redone under a new rule); the stamper knows them and reports
any other duplicate.

Now `.github/workflows/stamp-version.yml` runs on every push to `main`:

1. `tools/version/stamp_changelog.py` replaces each `(vNEXT <kind>)` heading
   with the next number after the HIGHEST real one in the file (not the top
   one); several pending headings are numbered bottom-up, oldest first.
   Only `## ` headings outside code fences, and only a tag that ends the
   heading, are touched.
2. The workflow commits that to `main` as `CHANGELOG: stamp vX.Y.Z` and tags
   `vX.Y.Z`, pushed atomically.
3. Runs share one `concurrency` group, so they queue; each run checks out the
   latest `main`, never the pushed SHA. If `main` moved during a run, the
   push is rejected and the next queued run stamps the same headings.
4. `mirror-to-gitlab.yml` mirrors again when the stamp run completes (a push
   made with `GITHUB_TOKEN` starts no workflow, so without this the office PC
   would not see the number until the next push).

Until the stamp run lands (seconds), `main` -- and a GitLab pull taken in that
window -- shows `(vNEXT ...)`; the next mirror fixes it.

The stamper's unit tests: `python3 -m unittest discover -s tools/version -p
"test_*.py"` (the workflow runs them before stamping).

Version numbers still mark operator-facing releases, not every commit: only a
heading someone wrote gets a number, never a commit without one.
