# Installing olctl (instructions for an AI agent)

You have been asked to install olctl, a command-line tool that lets you read
the user's Outlook mail, for the user on this Windows machine. Follow these
steps exactly. Do not improvise: if a step fails, stop and show the user the
JSON or the error from that step.

## What it does and does not do

- Installs into `%LOCALAPPDATA%\Programs\olctl` for this user only. **No admin
  rights** are needed and none should be requested.
- Adds that folder to the user's PATH, installs two Claude Code skills into
  `%USERPROFILE%\.claude\skills`, and records the version under
  `HKCU\Software\olctl`.
- If the user's PowerShell execution policy would block olctl, it sets the
  user's own policy to `RemoteSigned`. That needs no admin, and
  `uninstall.ps1` puts it back.
- olctl itself can never send or delete mail.

## Requirements

- Windows with **Outlook Classic** installed and a mail profile set up (it has
  been opened at least once). New Outlook alone is not enough.
- Internet access to `github.com` and `api.github.com`.

## Steps

1. Download the installer script to the user's temp folder and run it. In
   PowerShell:

   ```powershell
   $get = Join-Path $env:TEMP 'get-olctl.ps1'
   Invoke-WebRequest -UseBasicParsing -OutFile $get -Uri 'https://raw.githubusercontent.com/LeoStehlik/outlook-cli/master/get-olctl.ps1'
   powershell -NoProfile -ExecutionPolicy Bypass -File $get
   ```

   From Git Bash (Claude Code on Windows), the same thing:

   ```bash
   curl -fsSL -o "$TEMP/get-olctl.ps1" https://raw.githubusercontent.com/LeoStehlik/outlook-cli/master/get-olctl.ps1
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(cygpath -w "$TEMP/get-olctl.ps1")"
   ```

   The last line of output is one JSON object. `"ok": true` with
   `"installed": true` means it worked; show the user `current` (the
   installed version) and `install_output`. `"ok": false` means it did not:
   stop and show the user `code` and `error`.

2. The PATH change only reaches new terminals, so in this session call olctl
   by its full path:

   ```bash
   "$(cygpath "$LOCALAPPDATA")/Programs/olctl/olctl" doctor --quick
   ```

   (PowerShell: `& "$env:LOCALAPPDATA\Programs\olctl\olctl.ps1" doctor --quick`.)

   Expect JSON with `outlook_version`, `profile` and the user's mailbox under
   `stores`. `no_outlook` means Outlook Classic is not installed or has never
   been opened: ask the user to open it once, then run this step again.

3. Read the guide and follow it from now on:

   ```bash
   "$(cygpath "$LOCALAPPDATA")/Programs/olctl/olctl" guide
   ```

4. Tell the user it is installed, which version, and that a **new** Claude Code
   session (or terminal) is needed for the bare `olctl` command and the
   `olctl` skills to be available.

## Later

- `olctl update --check` reports whether a newer release exists; `olctl update`
  installs it.
- To remove it: `powershell -NoProfile -ExecutionPolicy Bypass -File "%LOCALAPPDATA%\Programs\olctl\uninstall.ps1"`.
  The user's settings and audit log in `%USERPROFILE%\.olctl` are kept.
