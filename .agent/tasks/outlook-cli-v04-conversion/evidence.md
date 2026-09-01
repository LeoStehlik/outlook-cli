# Evidence - outlook-cli-v04-conversion

## Build Summary

Outlook CLI v0.4.0 is a conversion pass for the Windows-only utility. The README now has a first-five-minutes setup matrix, docs/windows-setup-proof.md documents supported environments and honest non-Windows verification limits, examples/first-five-minutes.md gives the first commands, and version strings moved to 0.4.0.

## Checks Run

```bash
git diff --check
grep -R "0.4.0" -n olctl.ps1 tests/cli-args.ps1 README.md docs examples
grep -nE "[^A-Za-z]\.Send\s*\(" olctl.ps1
```

Observed:

```text
git diff --check: clean
version strings found in olctl.ps1, tests/cli-args.ps1, and docs/windows-setup-proof.md
NO_SEND_CALL
PWSH_UNAVAILABLE on the Linux dev host
```

## AC1 - PASS

README has a first-five-minutes matrix and copy-paste Windows commands for Unblock-File, --version, doctor, and folders.

## AC2 - PASS

docs/windows-setup-proof.md states supported/unsupported environments, non-Windows review limits, Windows + Outlook Classic COM requirements, and honest limitations.

## AC3 - PASS

examples/first-five-minutes.md exists and examples/README.md links it.

## AC4 - PENDING RELEASE

olctl.ps1 and tests/cli-args.ps1 identify 0.4.0; GitHub tag/release verification is pending.

## AC5 - PASS WITH LIMITATION

Static checks pass. PowerShell is unavailable on the Linux dev host, so argument-layer and COM validation require Windows/PowerShell and a configured Outlook Classic profile.
