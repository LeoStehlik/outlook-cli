# olctl - a local Outlook CLI

`olctl` drives the **Outlook Classic** desktop client on your own machine over
COM, as you, using the mail profile Outlook already has.

No Azure app registration, no admin consent, no app password, no tenant
approval. Outlook is already authenticated; the tool just drives it.

**Nothing to install.** It is one PowerShell script, and PowerShell ships with
Windows. Python is no longer involved anywhere.

It never sends mail and never hard-deletes anything. Replies are saved as drafts
for you to review and send.

## Files

| File | Role |
|---|---|
| `olctl.ps1` | everything — COM engine plus CLI. This is the tool. |
| `olctl.cmd` | entry point for native Windows `cmd.exe` shells, and a fallback where PowerShell's execution policy is the default `Restricted` |
| `tests/cli-args.ps1` | regression harness for the argument layer |
| `agent-config/` | optional example config for driving `olctl` from an AI coding agent — see below |

This is a native Windows PowerShell tool: COM/MAPI is a Windows-only
mechanism, so it has to run as a Windows process talking directly to
`outlook.exe`. There is no WSL wrapper and no Linux/cross-platform path — if
an AI coding agent is driving it, it needs to be running on the Windows side
(or reaching it via `powershell.exe`), not inside a Linux VM or WSL distro.

## Install

Copy this folder anywhere on the Windows filesystem, then unblock the script
once — Windows marks anything copied from another machine or downloaded from
the internet, and PowerShell refuses to run a marked script until you clear it:

```powershell
Unblock-File .\olctl.ps1
.\olctl.ps1 doctor --pretty
```

To call it as a bare `olctl` from any directory, put the folder on `PATH`.

**From a PowerShell prompt, prefer `.\olctl.ps1` over `.\olctl.cmd` (or a bare
`olctl` resolved from `PATH`).** The `.cmd` adds a `cmd.exe` hop, and
PowerShell 5.1 strips quotes it thinks are unnecessary when calling a legacy
batch file — so an argument containing `&`, `<`, `>`, `|` or `^` can be
re-parsed by `cmd.exe` as redirection or command separation. A folder or
subject containing `&` is entirely plausible here. The `.cmd` is for real
`cmd.exe` shells, or for a user who would rather not deal with PowerShell's
execution policy at all (see below); from PowerShell, call the `.ps1` directly.

### If PowerShell refuses to run it at all

`Unblock-File` fixes a script's own "downloaded from another machine" mark, but
a machine's **execution policy** is a separate gate and can block even an
unblocked local script outright (`... cannot be loaded because running scripts
is disabled on this system`). Windows client machines default to `Restricted`
(no scripts at all) unless someone has changed that. Fix it for your user only,
no admin rights required:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

If *that* command is itself refused (`...is overridden by a policy defined at
a more specific scope`), the policy is enforced by Group Policy, not just left
at the default — run `Get-ExecutionPolicy -List` and look for `MachinePolicy`
or `UserPolicy` set to anything other than `Undefined`. Group Policy sits above
both `-Scope CurrentUser` and the `-ExecutionPolicy Bypass` that `olctl.cmd`
passes internally, so **neither the command above nor `olctl.cmd` can get
around a real Group-Policy-enforced lock** — that needs an exception from IT
(or a signed script). Where there is no such lock — the far more common case
of a machine simply sitting on the Windows default — both fixes above are
genuine, and `olctl.cmd` is a real convenience: it passes `-ExecutionPolicy
Bypass` for you, so someone who doesn't want to touch `Set-ExecutionPolicy` at
all can just run it directly.

## Prerequisites

1. **Outlook Classic installed with a configured profile.** Keep using New
   Outlook day to day — both are views onto the same mailbox, so anything
   `olctl` does through Classic shows up there within seconds. New Outlook has
   no COM interface at all, which is why we go through Classic. If Classic has
   never been opened on this machine, open it once and let the initial sync
   finish.
2. **Classic Outlook running.** COM will launch it otherwise, but it is more
   reliable running. It can sit minimised.

## Smoke test

```powershell
.\olctl.ps1 doctor --pretty
.\olctl.ps1 folders --folder projects --depth 3 --pretty
(.\olctl.ps1 list --folder "projects/inbox" --unread --since 3d --limit 10 |
    ConvertFrom-Json).items |
  ForEach-Object { "{0}  {1}  {2}" -f $_.received, $_.from_email, $_.subject }
```

Every command prints exactly one JSON object on the pipeline, so it can be
captured directly:

```powershell
$r = (.\olctl.ps1 list --folder "projects/inbox" --limit 1 | ConvertFrom-Json).items[0].ref
```

Exit 0 = success, 1 = a handled error carrying `code` and a human-readable
`error`. A successful payload has no `ok` key; a failure always has
`"ok": false`, so `if ($result.ok -eq $false)` is the reliable in-process test
(`$LASTEXITCODE` also works).

## Command surface

| Command | What it does |
|---|---|
| `doctor` | Outlook version, profile, stores, cached mode, protected folders |
| `folders [--folder P] [--depth N]` | folder tree with item and unread counts |
| `list --folder P [filters]` | messages, newest first, as JSON rows |
| `get REF [--body text\|html\|none] [--headers]` | one message in full |
| `save-attachments REF --out DIR [--pattern "*.pdf"]` | write attachments to disk |
| `move REF --to P` | move between folders |
| `mark REF read\|unread` | read state |
| `flag REF [--text T] [--clear]` | follow-up flag |
| `categorize REF [--add C] [--remove C] [--set A B]` | categories |
| `draft-reply REF --text ...` | reply saved to **Drafts**, not sent |
| `draft-new --to X --subject S --text ...` | new mail saved to **Drafts**, not sent |

`list` filters: `--unread`, `--read`, `--flagged`, `--since`, `--until`,
`--sender`, `--subject`, `--category`, `--has-attachments`, `--limit`,
`--oldest-first`, `--count-only`, `--no-preview`, `--any-class`, `--scan-max`.

`--since` / `--until` take `30m`, `6h`, `3d`, `2w`, `2026-08-18`, or
`2026-08-18T09:30`.

Global: `--pretty`, `--dry-run`, `--version`, `--help`. Both flag positions
work (`olctl --pretty doctor` and `olctl doctor --pretty`).

`--text` accepts literal text, a file path, or `-` for stdin.

### Folder paths

`Inbox`, `Inbox/Vendors`, `Inbox\Vendors`, or store-qualified
`projects/inbox`. A literal slash in a folder name is escaped `\/`. When a
path does not resolve, the error lists the folders that *do* exist at that
level, so an agent can self-correct in one step.

**Localised mailboxes.** Folder names differ per mailbox — `Inbox`,
`Posteingang`, `Doručená pošta`. Use the real name, or use the English tokens
(`inbox`, `drafts`, `sent items`, `deleted items`, `junk email`, `outbox`),
which resolve through that store's own `GetDefaultFolder`:

```powershell
.\olctl.ps1 list --folder "SA_AutomationAgent/inbox"  # -> Doručená pošta
.\olctl.ps1 list --folder "Regional Office/inbox"     # -> Posteingang
```

Resolution happens **inside the named store**, never against your default
mailbox. `doctor` prints each store's `default_folders` map.

**`GetDefaultFolder()` creates.** It is not a lookup: if the default folder for
a type does not exist in that store, Outlook creates it. An earlier version of
this tool called it with code 31 during `doctor` and thereby created an empty
*Quick Step Settings* folder in four mailboxes, including shared ones. Anything
added here that resolves a folder by type is a potential write, even in a
command that looks read-only — and `doctor` is not covered by the audit log.
Codes 3, 4, 5, 6, 16 and 23 are safe because those folders always exist.

`archive` is deliberately **not** a token. `OlDefaultFolders` 31 is documented
as `olFolderArchive`, but this Outlook build returns *Quick Step Settings* for
it — so resolving it by type would file mail into a settings folder. Address a
real Archive folder by its own name (`projects/Archive`). `doctor` reports
`archive` by name when the store has one.

### Refs

`ref` is an opaque `EntryID!StoreID` pair. Two things matter:

* **A ref changes when the item moves.** `move` returns `new_ref`; use that. A
  stale ref gives a clean `not_found` rather than acting on the wrong mail.
* Refs are stable across restarts as long as the item stays put.

## Recipes

Worked examples against the fictional `projects` mailbox described above. Full
sample payloads for every command, including these, live in
[`examples/`](examples/) — the snippets below are trimmed to the fields that
matter for the recipe.

**Triage unread mail from the last few days.** Bound every scan with `--since`
and `--limit`; the preview is usually enough to decide what needs a closer look:

```powershell
.\olctl.ps1 list --folder "projects/inbox" --unread --since 3d --limit 25
```

```json
{
  "returned": 3,
  "truncated": false,
  "items": [
    { "subject": "Conference room booking kiosk — pilot feedback", "from_email": "alex.rivera@contoso.com", "has_attachments": true, "flag": null },
    { "subject": "Warehouse handheld scanner rollout — timeline", "from_email": "sam.patel@contoso.com", "has_attachments": false, "flag": "flagged" },
    { "subject": "Guest Wi-Fi self-service portal — access request", "from_email": "morgan.ito@northwind-supply.example", "has_attachments": true, "flag": null }
  ]
}
```

Full response: [`examples/list.json`](examples/list.json).

**Pull attachments from an external sender to disk.** `--sender` matches
against both the display name and address, applied after Outlook's own filter:

```powershell
.\olctl.ps1 list --folder "projects/inbox" --sender northwind-supply --has-attachments --since 3d
# -> one match: the Guest Wi-Fi portal message, ref C0FFEE...
.\olctl.ps1 save-attachments "C0FFEE...!CAFEBABE..." --out C:\Users\jane.doe\attachments --pattern "*.pdf"
```

```json
{ "attachments": 1, "saved": [{ "name": "wifi-portal-spec.pdf", "path": "C:\\Users\\jane.doe\\attachments\\wifi-portal-spec.pdf", "bytes": 152004 }], "skipped": [] }
```

**Flag something for follow-up, then mark it reviewed once handled.** Flag and
category are independent — flag for "needs action", category for bookkeeping:

```powershell
.\olctl.ps1 flag "FEEDFACE...!CAFEBABE..." --text "Confirm rollout order with Sam"
# ... later, once actually reviewed:
.\olctl.ps1 categorize "FEEDFACE...!CAFEBABE..." --add "Reviewed"
.\olctl.ps1 flag "FEEDFACE...!CAFEBABE..." --clear
```

**Draft a reply without hand-typing into the shell.** Write the body to a file
and pass the path — `--text` treats an existing file path as a file, anything
else as literal text:

```powershell
"Thanks Morgan - the SSID handoff section looks right. One question: does the portal expect the guest VLAN tag before or after captive-portal redirect?" |
  Out-File -Encoding utf8 reply-draft.txt
.\olctl.ps1 draft-reply "C0FFEE...!CAFEBABE..." --text reply-draft.txt
```

```json
{ "draft_ref": "BAADF00D...!CAFEBABE...", "saved_to": "Drafts", "sent": false }
```

It lands in **Drafts**, unsent, for you to review — see [Safety
model](#safety-model). Full response: [`examples/draft-reply.json`](examples/draft-reply.json).

**Move a handled item out of the inbox — and what happens if the destination
is protected.** `move` returns `new_ref`; the old one goes stale immediately:

```powershell
.\olctl.ps1 move "FEEDFACE...!CAFEBABE..." --to "projects/Inbox/Processed"
```

```json
{ "from": "projects/Inbox", "to": "projects/Inbox/Processed", "new_ref": "FACEFEED...!CAFEBABE...", "note": "the ref changed; use new_ref for further commands" }
```

Aim it at `Deleted Items` (or any other protected folder) instead and every
store's guard fires without `--allow-protected`:

```json
{ "ok": false, "code": "protected_folder", "error": "destination folder 'Deleted Items' is protected: it is a protected default folder (projects/Deleted Items). Pass --allow-protected if that is really intended." }
```

**When a folder path doesn't resolve.** The error lists what actually exists
at that level, so a script or an agent can self-correct in one step rather
than guessing:

```powershell
.\olctl.ps1 list --folder "Proejcts/inbox"
```

Full response: [`examples/error-not-found.json`](examples/error-not-found.json)
— note `available_stores` and `well_known` alongside `children`, so the error
carries enough to recover without another round trip.

## Safety model

* No `send` and no `delete` command exists. Not gated — absent.
* Deleted Items, Junk Email, Sent Items and Outbox **of every store in the
  profile** are refused as `move` destinations without `--allow-protected`.
  They are resolved per store via `GetDefaultFolder`, so `Gelöschte Elemente`
  and `Odstraněná pošta` are covered without being listed anywhere. `doctor`
  prints `protected_folders_resolved` — a guard you cannot see is a guard you
  cannot trust.
* Every mutation is appended to `%USERPROFILE%\.olctl\audit.jsonl`.
* `--dry-run` on any mutating command.

Override in `%USERPROFILE%\.olctl\config.json`:

```json
{
  "protect_default_folders": true,
  "protected_folders": ["Deleted Items", "Junk Email", "Sent Items", "Outbox"],
  "max_items": 200,
  "default_body_chars": 20000,
  "preview_chars": 240
}
```

`OLCTL_HOME` moves the config and audit log elsewhere.

## Wiring it into an AI coding agent

`olctl` has no dependency on any particular agent or AI vendor — it is a
standalone CLI, usable from a plain shell with no AI in the loop at all.
`agent-config/` is an optional, self-contained example of wiring it into
[Claude Code](https://claude.com/claude-code) specifically, since that is what
this example was built and tested against. Nothing in it is required to use
`olctl`, and nothing here is installed automatically — copy in only what you
want:

| File in `agent-config/` | Copy to | Purpose |
|---|---|---|
| `CLAUDE.md` | your project root | house rules for the agent: read-before-write, ask before batch mutations, never invent a ref, etc. `CLAUDE.md` is Claude Code's own convention for a file it loads automatically; other harnesses look for their own equivalent (e.g. `AGENTS.md`) — the content itself is plain guidance, not Claude-specific, so adapt it freely. |
| `settings.json` | `.claude/settings.json` | a read-only-first permission allowlist: `list`/`get`/`folders`/`doctor`/`--dry-run` are allowed, every mutating command is denied so you get a prompt instead of a surprise. Move lines from `deny` to `allow` once you trust it. |
| `commands/mail-scan.md` | `.claude/commands/mail-scan.md` | an on-demand slash command that scans a mailbox and triages new requests vs. noise |
| `skills/project-intake/` | `.claude/skills/project-intake/` | the same scan, packaged as a skill for a scheduled/recurring run |

If you use a different harness, treat these as worked examples of the prompting
pattern (bound every scan with `--since`/`--limit`, read the `code`/`error`
fields on failure, never touch a store you weren't told to) rather than files
to copy verbatim.

## Known limits

* **New Outlook has no COM.** If Classic is ever removed from the machine, the
  alternatives are Graph (needs the tenant approval you do not have) or browser
  automation against OWA.
* **Startup cost.** Each command spawns `powershell.exe` and attaches to
  Outlook: expect roughly half a second to a second before any work happens.
  This is the floor for the approach, so prefer one `list` returning many rows
  over many small calls.
* **Hybrid profiles.** COM does not care which organization a mailbox lives in.
  But secondary and cross-org mailboxes are often mounted in **online mode**
  rather than Cached Exchange Mode, where every `Restrict` and `Items.Count` is
  a server round trip — far slower, and `--scan-max` starts truncating.
  `doctor` reports `cached_mode` per store. Also, `--shared`
  (`GetSharedDefaultFolder`) needs delegate rights resolvable inside one
  organization and generally fails across a hybrid boundary; if the mailbox is
  already mounted as a store, use the store-path form instead. Moving an item
  between stores is a copy-then-delete under the hood, so it is slow for large
  items and always changes the ref.
* **Programmatic access guard.** Outlook can prompt about a program accessing
  e-mail addresses if antivirus is not reporting healthy to Windows Security
  Center. On a managed box it normally is.
* **Python-side filters are gone; all filtering now runs in-process** during the
  scan, but `--sender`, `--subject`, `--category` and `--has-attachments` are
  still applied *after* Outlook's own `Restrict`, so pair them with `--since`.
* **Date literals.** Outlook's filter syntax wants US-format dates regardless of
  the Windows locale, and the tool formats them with InvariantCulture — which
  matters on your German/Czech machine, where the local culture would render
  AM/PM in a form `Restrict` rejects. Do not hand-build filters.

## Reading output on Windows

PowerShell 5.1's `>` and `Out-File` default to UTF-16, which mangles the
Czech and German folder names when the file is read back elsewhere. Use:

```powershell
.\olctl.ps1 doctor --pretty | Out-File -Encoding utf8 doctor.json
```

## Tests

```bash
pwsh -NoProfile -File tests/cli-args.ps1     # or powershell.exe on Windows
```

This covers the argument layer: every validation error, both flag positions,
Unicode folder names, and that well-formed commands reach the COM attach. It
deliberately stops there — **the COM layer cannot be tested off Windows.** It
was developed against a mock of the Outlook object model and syntax-checked, so
expect first-contact rough edges; the error envelopes are designed to say
exactly what went wrong.
