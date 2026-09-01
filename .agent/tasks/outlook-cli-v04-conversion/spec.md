# Task: outlook-cli-v04-conversion

## Task Statement

Ship outlook-cli v0.4.0 as a niche utility conversion pass for Windows users and agent reviewers.

## Acceptance Criteria

**AC1:** README has a first-five-minutes setup matrix and exact Windows commands near the top.
- Verify: inspect README for setup matrix, `Unblock-File`, `--version`, `doctor`, and `folders` commands.

**AC2:** A standalone setup proof doc states supported environments, unsupported environments, non-Windows verification limits, and Windows-only COM requirements.
- Verify: inspect `docs/windows-setup-proof.md`.

**AC3:** Example walkthrough exists for the first Windows commands.
- Verify: inspect `examples/first-five-minutes.md` and examples index.

**AC4:** Version surfaces identify v0.4.0.
- Verify: inspect `olctl.ps1`, `tests/cli-args.ps1`, Git tag/release after release.

**AC5:** Non-Windows validation is honest and passes available static checks.
- Verify: run grep/static checks; if PowerShell is unavailable, record that COM and argument-layer tests require Windows/PowerShell.

## Constraints

- Do not add Graph/Azure auth or cross-platform claims.
- Do not claim live Outlook COM testing unless run on Windows with Outlook Classic.
- Keep no-send/no-hard-delete safety positioning intact.

## Non-Goals

- No new Outlook features.
- No package manager work.
- No Linux/WSL support.

## Verification Approach

Inspect docs, run static checks, confirm version strings, release through PR/tag/GitHub, and explicitly record that full runtime proof requires Windows + Outlook Classic.
