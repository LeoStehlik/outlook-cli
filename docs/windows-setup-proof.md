# olctl Windows Setup Proof

`olctl` converts best when the visitor can decide in under a minute whether it fits their machine.

## Fit Matrix

| Environment | Supported | Why |
|---|---:|---|
| Windows 10/11 + Outlook Classic configured | yes | PowerShell can talk to Outlook COM/MAPI as the logged-in user. |
| Windows 10/11 + New Outlook only | no | New Outlook has no COM interface. Install/open Classic once. |
| Windows + Outlook Classic but Group Policy blocks scripts | blocked | Group Policy beats `RemoteSigned` and the `.cmd` bypass. Ask for a signed script or policy exception. |
| WSL/Linux/macOS | no | COM/MAPI must run in a Windows process attached to Outlook Classic. |
| Agent on Linux controlling a Windows host | possible | The command must execute on Windows, for example through a Windows shell, not inside WSL. |

## First Commands

```powershell
Unblock-File .\olctl.ps1
.\olctl.ps1 --version
.\olctl.ps1 doctor --pretty
.\olctl.ps1 folders --folder inbox --depth 1 --pretty
```

Expected version shape:

```json
{ "olctl_version": "0.4.0" }
```

Expected handled failure when Outlook Classic is unavailable:

```json
{ "ok": false, "code": "no_outlook", "error": "could not start or attach to Outlook Classic over COM" }
```

## What Agents Can Verify Without Outlook

On a non-Windows review host, an agent can still verify:

- `README.md` names the Windows-only requirement and setup path.
- `examples/*.json` describe command payload shapes without live mail data.
- `olctl.ps1` contains `$script:OLCTL_VERSION = '0.4.0'`.
- `tests/cli-args.ps1` expects `0.4.0` for the argument-layer version test.
- No command calls `Send()` or performs hard delete behavior.

## What Requires Windows + Outlook Classic

These require a real configured profile:

```powershell
.\olctl.ps1 doctor --pretty
.\olctl.ps1 list --folder inbox --limit 5 --pretty
.\olctl.ps1 draft-new --to you@example.com --subject "olctl test" --text "draft only"
```

The final command saves a draft. It does not send mail.

## Limits To Keep Honest

- This is not a Graph API client.
- This is not cross-platform.
- This cannot bypass enterprise Group Policy.
- Folder names are mailbox-local and may be localized.
- Message refs change after moves; use the returned `new_ref`.
- `doctor` should stay read-only and must not resolve unsafe default folder IDs.
