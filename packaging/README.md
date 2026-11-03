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
| `%LOCALAPPDATA%\Programs\olctl\` | `olctl.ps1`, `olctl.cmd`, `olctl` (bash), `guide\`, `uninstall.ps1`, `install.log`, `.olctl-install` marker |
| User PATH | that folder, added once |
| `%USERPROFILE%\.claude\skills\olctl`, `...\olctl-mail-scan` | Claude Code skills, each with a `.olctl-managed` marker. Existing folders of the same name without the marker are left alone. |
| `HKCU\Software\olctl` | `Version`, `InstallPath`, `SkillsRoot`, `InstalledAt` |

It refuses to install into a non-empty folder it did not create. Uninstall
removes exactly the items above, and keeps `%USERPROFILE%\.olctl` (the user's
config, audit log and scan marks).

## Execution policy

The package is not signed. The install and uninstall commands above pass
`-ExecutionPolicy Bypass`, so they run whatever the policy is.

Day-to-day use depends on the shell:

| Shell | Runs | Under the Windows default (`Restricted`) |
|---|---|---|
| Git Bash, which Claude Code on Windows uses | `olctl` (bash), which calls `powershell -ExecutionPolicy Bypass` | works |
| cmd.exe | `olctl.cmd`, which also passes `Bypass` | works |
| PowerShell | `olctl.ps1` directly | **blocked**: "running scripts is disabled" |

So on machines still at the Windows client default, PowerShell users (and
agents using a PowerShell tool) need the policy relaxed. Either IT sets
`RemoteSigned` through Intune or GPO, which is common on managed fleets, or
each user runs `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, which
needs no admin. The installer deliberately does not change the policy itself.

If the machines enforce `AllSigned` or AppLocker/WDAC script rules, the
script would need to be signed first.
