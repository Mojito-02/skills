---
name: rtk-token-saver
description: Compress noisy shell output with RTK (rtk.exe, rtk-ai.app) before it enters the context window — test runs, build logs, git diffs, directory trees, greps, big dependency files. Use when a command is likely to produce long or repetitive output and only the meaningful lines matter.
---

# RTK — Token-Optimized CLI

`rtk` is installed at `F:\rtk\rtk.exe` and is on the user PATH. It wraps common
commands and strips boilerplate before the output reaches you: RTK reports 56%
average output reduction, up to 90% on the noisiest commands. Test failures,
errors, diffs and stack traces are preserved.

DSH cannot auto-rewrite shell commands (its hooks can only allow or deny), so
call `rtk` explicitly when the output is likely to be large.

## Usage

Prefix the command with `rtk`:

| Instead of | Run |
|---|---|
| `pytest -q`, `npm test`, `cargo test` | `rtk pytest -q` / `rtk test <cmd>` |
| `git diff`, `git status`, `git log` | `rtk git diff` / `rtk git status` / `rtk git log` |
| `ls -la` | `rtk ls -la` |
| `find . -name '*.py'` | `rtk find . -name "*.py"` |
| `grep -rn foo .`, `rg foo` | `rtk grep -rn foo .` / `rtk rg foo` |
| `tsc --noEmit`, `eslint`, `prettier --check` | `rtk tsc`, `rtk lint`, `rtk prettier` |
| `docker ps`, `kubectl get pods` | `rtk docker ps` / `rtk kubectl get pods` |
| `cat huge.json` | `rtk json huge.json` (compact) or `--keys-only` |
| reading a long file | `rtk read <file>` |
| any noisy command | `rtk err <cmd>` (errors only) / `rtk summary <cmd>` |

Do not use it when the exact raw output matters (byte-for-byte fixtures,
`--json` output you must parse programmatically, checksums).

## Checking savings

```bash
rtk gain              # tokens saved so far, by command
rtk hook check --agent claude "git status"   # what a rewrite would look like
rtk config            # configuration at %APPDATA%\rtk\config.toml
```

## Notes

- Other agents on this machine (Claude Code, Codex, Copilot CLI) already rewrite
  automatically through magpie's `library rtk` switch; this skill covers DSH.
- Telemetry is off by default.
