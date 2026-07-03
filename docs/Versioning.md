# Versioning policy

This repository uses practical semantic versioning for operator-visible releases.
The current project version appears in each `CHANGELOG.md` release heading, for
example `v2.9.29`.

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

Every released change should add a new top entry to `CHANGELOG.md`:

```markdown
## YYYY-MM-DD - Short release title (vMAJOR.MINOR.PATCH)

### Added
- ...

### Fixed
- ...

### Notes
- ...
```

Keep sections that apply and omit empty sections. Put the newest release at the
top. If multiple commits are part of one unreleased release, keep updating the
same top entry instead of creating many tiny version headings.

## Automation recommendation

For this PowerShell-only repository, keep automation simple:

1. Store the current version in one plain text file, `VERSION`.
2. Keep `CHANGELOG.md` as the human release record.
3. Use a release helper script that:
   - reads `VERSION`;
   - bumps `major`, `minor`, or `patch`;
   - writes the new `VERSION`;
   - verifies the top `CHANGELOG.md` heading contains the same `vX.Y.Z`;
   - creates a git tag such as `v2.10.0` after tests pass.
4. In CI, fail if `VERSION` and the latest changelog heading disagree.

Recommended command shape:

```powershell
.\Release-Version.ps1 -Bump patch -Title "FillCheckSheet prefix fallback"
.\Tests\Run-Tests.ps1
# review, commit, then tag after merge/release
# git tag v2.9.30
```

Avoid fully automatic version bumps on every commit. Version numbers should mark
operator-facing releases, not every internal checkpoint.
