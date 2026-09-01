# First Five Minutes

Run on Windows PowerShell from the repository folder.

```powershell
Unblock-File .\olctl.ps1
.\olctl.ps1 --version
.\olctl.ps1 doctor --pretty
.\olctl.ps1 folders --folder inbox --depth 1 --pretty
```

If `doctor` returns `no_outlook`, open Outlook Classic once and let the profile finish syncing. If PowerShell refuses before the script starts, check `docs/windows-setup-proof.md`.
