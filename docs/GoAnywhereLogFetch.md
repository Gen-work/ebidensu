# GoAnywhereLogFetch (one-off)

Fetch the GoAnywhere receive-job log of every correl in a hand-written list
and copy each one into its receive system's LOG folder. Standalone script,
not a phase; no mapping CSV involved.

```
powershell -File GoAnywhereLogFetch.ps1 -ConfigPath C:\work\logfetch\logfetch.json
```

## Config (local only, never committed)

The config holds the share path, so it lives next to your work files, not
in the repo. UTF-8 JSON; Japanese in paths is fine.

```json
{
  "destRoot": "\\\\server\\share\\...\\LOG",
  "nameFormat": "{file}",
  "systems": {
    "AAA":  { "goFolder": "/aaausr/",  "dest": "AAA"  },
    "BBB2": { "goFolder": "/bbb2usr/", "dest": "BBB2" }
  },
  "jobs": [
    { "correl": "JAAAL25M", "if": "0001", "recv": "AAA" },
    { "correl": "JBBBK16M", "if": "0006", "recv": "BBB2" },
    { "correl": "JBBBC04M", "if": "0001,0002", "recv": "BBB2" }
  ]
}
```

- `if`: the IF number the GoAnywhere project name carries (`..._IF0001_...`).
  VER1 / DEV1 and the sender prefix are ignored. Several, comma separated,
  when a job was received under one and resent under another: extra
  candidates are safe, the log content still decides the owner.
- `recv` -> `systems`: `goFolder` is the GoAnywhere Folder prefix of the
  receive job (`/aaausr/Receive`), `dest` the subfolder under `destRoot`.
- `nameFormat`: default `{file}` = the name as downloaded. Other tokens:
  `{orig}` (that name without extension), `{correl}`, `{jobNo}`, `{start}`
  (yyyyMMdd_HHmmss).

## Run

1. Edge: GoAnywhere completed jobs, date range covering the test days,
   **100 rows per page**, first filter (e.g. project name `LPS`). Enter.
2. The script captures the page, shows each correl's candidates and
   downloads the ones on THIS page (Enter; `s` skips). Ctrl+F can only find
   a job number on the visible page, so download happens per filter.
3. Set the next filter in Edge (`ALC`, `CRY`, ...), then `c`. Repeat.
4. Enter finishes: each correl's log is picked by content
   (`/gfix/recv/<correl>`, last character free: `...M` and `...S` are the
   same correl), the newest rerun wins, and it is copied to
   `destRoot\<dest>\`. When no log names the correl, its candidates are
   listed: a number copies that one (`MANUAL`), Enter leaves it `NG`.
5. `goanywhere_logs\report.csv` (next to the config) lists every correl:
   OK / MANUAL / EXISTS / NG, job number, start time, destination, note.

Downloaded logs stay in `goanywhere_logs\<jobNo>\` under their own name; a
rerun does not download a staged job again and never overwrites a
destination file without `-Force`.

Options: `-DryRun` (plan only), `-ListTextFile a.txt,b.txt` (saved list-page
text instead of a live capture; only staged logs are matched), `-Force`.
