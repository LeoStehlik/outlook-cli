# Deploying olctl to many users

olctl installs **per user**, with **no admin rights**: files go into the user's
profile, the PATH change is the user's own, and the registry key is under
HKCU. One package contains the tool, its agent guide and the Claude Code
skills.

## Build the package

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File packaging\build.ps1
```

This runs the argument tests and the static contract tests, then produces:

* `dist\olctl-<version>\`: the folder to package;
* `dist\olctl-<version>.zip`: the same folder zipped.

## Intune (Win32 app)

1. Wrap the staged folder:
   `IntuneWinAppUtil.exe -c dist\olctl-<version> -s install.ps1 -o <output folder>`
2. App settings:

   | Setting | Value |
   |---|---|
   | Install command | `powershell.exe -NoProfile -ExecutionPolicy Bypass -File install.ps1` |
   | Uninstall command | `powershell.exe -NoProfile -ExecutionPolicy Bypass -File uninstall.ps1` |
   | Install behavior | **User** |
   | Device restart behavior | No specific action |
   | Return codes | `0` success, `1` failure |

3. Detection rule, either of:
   * Registry: `HKEY_CURRENT_USER\Software\olctl`, value `Version`, String
     comparison **Equals** `<version>`. This is the better choice, because a
     new version is detected as not installed.
   * File: `%LOCALAPPDATA%\Programs\olctl\olctl.ps1` exists.
4. Requirements: Windows 10/11 with **Outlook Classic** installed. olctl needs
   a Classic profile; it does not work with New Outlook alone.
5. Assignment: **Available** to the group of users who should have it, so it
   appears in Company Portal and users install it themselves.

**Updates:** build the new version, create a new Win32 app with the new
detection value, and set it to **supersede** the previous one (with
"Uninstall previous version" off: the installer updates in place). Roll it out
to a pilot group first.

Where Intune manages the version, stop users updating on their own, or the
detection value no longer matches: add `"update"` to `disabled_commands` in
the machine policy (below).

## ConfigMgr

Same commands in a Script Installer deployment type, installation behaviour
**Install for user**, detection method: the same registry value.

## Optional: machine policy

To restrict what olctl may do (for example read-only for most users), deploy
`%ProgramData%\olctl\policy.json` separately as a device-context file or
script, since writing to ProgramData needs admin rights:

```json
{ "read_only": true, "disabled_commands": ["rules-add", "rules-delete", "rules-run", "folder-add"] }
```

`olctl doctor` shows which policy is active. See "Machine policy" in the main
README for exactly what it blocks.

## What the installer changes

| Where | What |
|---|---|
| `%LOCALAPPDATA%\Programs\olctl\` | `olctl.ps1`, `olctl.cmd`, `olctl` (bash), `guide\`, `uninstall.ps1`, `get-olctl.ps1`, `install.log`, `.olctl-install` marker |
| User PATH | that folder, added once |
| `%USERPROFILE%\.claude\skills\olctl`, `...\olctl-mail-scan` | Claude Code skills, each with a `.olctl-managed` marker. Existing folders of the same name without the marker are left alone. |
| `HKCU\Software\olctl` | `Version`, `InstallPath`, `SkillsRoot`, `InstalledAt`, and `ExecutionPolicyBefore` if it changed the policy |
| CurrentUser execution policy | set to `RemoteSigned`, only if it would otherwise block olctl and Group Policy does not set it |

It refuses to install into a non-empty folder it did not create. Uninstall
removes exactly the items above, and keeps `%USERPROFILE%\.olctl` (the user's
config, audit log and scan marks).

## Execution policy

The package is not signed. The install and uninstall commands above pass
`-ExecutionPolicy Bypass`, so they run whatever the policy is.

For day-to-day use, the installer makes sure the user's own shells can run
olctl. If the policy that applies to the user is `Restricted` (the Windows
client default) or `AllSigned`, it sets the **CurrentUser** scope to
`RemoteSigned`. That needs no admin rights and affects only that user. It
records the previous value under `HKCU\Software\olctl` and `uninstall.ps1`
restores it. `install.ps1 -NoExecutionPolicy` skips this.

If Group Policy sets the execution policy (`MachinePolicy` or `UserPolicy` in
`Get-ExecutionPolicy -List`), the installer leaves it alone, since nothing
below Group Policy can override it, and logs that. olctl then still works from
Git Bash, which Claude Code uses, and from cmd, because both entry points pass
`-ExecutionPolicy Bypass`. Only running `olctl.ps1` directly in PowerShell is
blocked.

If the machines enforce `AllSigned` through Group Policy, or AppLocker/WDAC
script rules, the script would need to be signed first.
