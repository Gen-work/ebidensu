#!/usr/bin/env python3
"""Stamp real version numbers into CHANGELOG.md release headings.

Branches never pick a version number: two branches cut from the same main
picked the same one (v2.41.1 and v2.41.2 each appear twice). A branch writes
its heading as

    ## 2026-10-29 - Short release title (vNEXT patch)

with `patch`, `minor` or `major` (no word = patch). When the branch reaches
main, .github/workflows/stamp-version.yml runs this script: every vNEXT
heading gets the next number after the highest real one in the file, the
lowest heading in the file first (the oldest entry), and the run is queued
by `concurrency`, so two merges can never get the same number.

Pure text in, text out; no git. The workflow commits, tags and pushes.
See docs/Versioning.md.
"""

import argparse
import re
import sys

HEADING = re.compile(r'^## ')
# The tag ends the heading; the lookahead leaves a CRLF file's \r in place.
REAL = re.compile(r'\(v(\d+)\.(\d+)\.(\d+)\)[ \t]*(?=\r?$)')
NEXT = re.compile(r'\(vNEXT(?:[ \t]+(patch|minor|major))?\)[ \t]*(?=\r?$)', re.IGNORECASE)

# Headings that were already duplicated before stamping existed. History is
# not rewritten (an old delivery is not redone under a new rule), so these
# are not reported again; any OTHER duplicate is.
KNOWN_DUPLICATES = {(2, 41, 1), (2, 41, 2)}


def bump(version, kind):
    major, minor, patch = version
    if kind == 'major':
        return (major + 1, 0, 0)
    if kind == 'minor':
        return (major, minor + 1, 0)
    return (major, minor, patch + 1)


def fmt(version):
    return 'v%d.%d.%d' % version


def release_headings(lines):
    """Yield (index, line) for each '## ' heading outside ``` fences."""
    fenced = False
    for i, line in enumerate(lines):
        if line.lstrip().startswith('```'):
            fenced = not fenced
            continue
        if not fenced and HEADING.match(line):
            yield i, line


def stamp(text):
    """Return (new_text, assigned, duplicates).

    assigned   -- version strings given to vNEXT headings, in the order given
    duplicates -- real versions written on more than one heading (after
                  stamping), minus KNOWN_DUPLICATES
    """
    lines = text.split('\n')
    real = []
    pending = []
    for i, line in release_headings(lines):
        m = REAL.search(line)
        if m:
            real.append(tuple(int(g) for g in m.groups()))
            continue
        n = NEXT.search(line)
        if n:
            pending.append((i, (n.group(1) or 'patch').lower(), n))

    current = max(real) if real else (0, 0, 0)
    assigned = []
    for i, kind, m in reversed(pending):
        current = bump(current, kind)
        line = lines[i]
        lines[i] = line[:m.start()] + '(' + fmt(current) + ')' + line[m.end():]
        assigned.append(fmt(current))
        real.append(current)

    seen = set()
    duplicates = []
    for v in real:
        if v in seen and v not in KNOWN_DUPLICATES and fmt(v) not in duplicates:
            duplicates.append(fmt(v))
        seen.add(v)
    return '\n'.join(lines), assigned, duplicates


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    p.add_argument('changelog', help='path to CHANGELOG.md (rewritten in place)')
    p.add_argument('--github-output', help='append versions=... to this file')
    args = p.parse_args(argv)

    # newline='' keeps the file's own line endings untouched.
    with open(args.changelog, encoding='utf-8', newline='') as f:
        text = f.read()
    new_text, assigned, duplicates = stamp(text)

    for d in duplicates:
        print('::warning file=%s::%s is on more than one heading' % (args.changelog, d))
    if assigned:
        with open(args.changelog, 'w', encoding='utf-8', newline='') as f:
            f.write(new_text)
        print('stamped: ' + ' '.join(assigned))
    else:
        print('no vNEXT heading; nothing to stamp')
    if args.github_output:
        with open(args.github_output, 'a', encoding='utf-8') as f:
            f.write('versions=%s\n' % ' '.join(assigned))
    return 0


if __name__ == '__main__':
    sys.exit(main())
