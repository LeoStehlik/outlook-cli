# olctl - a local Outlook CLI

`olctl` drives the **Outlook Classic** desktop client on your own machine over
COM, as you, using the mail profile Outlook already has.

This is for the cases where Graph/Azure is the wrong door or not available:
no app registration, no admin consent, no app password, no tenant approval.
Outlook is already authenticated; the tool just drives it.

**No admin rights, no dependencies.** It is one PowerShell script, and
PowerShell ships with Windows. The optional installer is per-user and needs no
elevation either. Python is deliberately not involved. The old Python path was
too slow for day-to-day mailbox triage.

It never sends mail and never hard-deletes anything. Replies are saved as drafts
for you to review and send. The one thing it ever moves to Deleted Items on its
own is an unsent draft you explicitly replace (`draft-reply --supersede`).

**v0.5 focus:** deployment to many users: a per-user installer with no admin
rights, a bundled agent guide (`olctl guide`), an optional machine policy, and
contract tests for the JSON every command returns.

## Files

| File | Role |
|---|---|
| `olctl.ps1` | the Outlook CLI script |
| `olctl.cmd` | wrapper for running it from `cmd.exe` |
| `olctl` | wrapper for bash: Git Bash (the shell Claude Code uses on Windows) and WSL on the same machine; calls `olctl.ps1` through `powershell.exe` |
| `guide/GUIDE.md` | the agent guide that `olctl guide` prints: house rules and command reference |
| `skills/` | Claude Code skills the installer puts in `~\.claude\skills` |
| `install.ps1`, `uninstall.ps1` | per-user install and removal, no admin rights |
| `packaging/` | builds the deployment package; notes for IT (Intune / ConfigMgr) |
| `tests/cli-args.ps1` | checks CLI argument parsing without needing Outlook |
| `tests/contract.ps1` | the JSON contract of every command, checked against `examples/` and, with `-Live`, a real mailbox |
| `agent-config/` | optional example config for agent-driven mailbox triage |

This is a native Windows PowerShell tool: COM/MAPI is a Windows-only
mechanism, so it has to run as a Windows process talking directly to
`outlook.exe`. The bash wrapper only forwards to `powershell.exe`: it works
from Git Bash and from WSL on the same Windows machine, but not from a Linux
VM, macOS or a remote box.

## First 5 Minutes

| Situation | Run | Expect |
|---|---|---|
| You have Windows + Outlook Classic configured | `./olctl.ps1 doctor --pretty` | JSON with profile, stores, default folders, and protected folders |
| PowerShell blocks the script mark | `Unblock-File .\olctl.ps1` | The script can run from this checkout |
| User execution policy blocks scripts | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` | User-scope scripts can run without admin rights |
| Group Policy blocks scripts | `Get-ExecutionPolicy -List` | `MachinePolicy` or `UserPolicy` explains why local bypass is not enough |
| Agent is reviewing from Linux/macOS | read [`docs/windows-setup-proof.md`](docs/windows-setup-proof.md) | COM cannot be proven there; only docs/examples/static checks can be verified |

Copy-paste smoke path on Windows:

```powershell
Unblock-File .\olctl.ps1
.\olctl.ps1 --version
.\olctl.ps1 doctor --pretty
.\olctl.ps1 folders --folder inbox --depth 1 --pretty
```

For the conversion proof and hard limits, read [`docs/windows-setup-proof.md`](docs/windows-setup-proof.md).

## Install

### Per-user installer

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1
```

From a checkout or an unpacked package (`packaging\build.ps1` makes one), this:

* copies olctl to `%LOCALAPPDATA%\Programs\olctl` and adds that folder to the
  **user** PATH, so `olctl` works in new PowerShell, cmd and Git Bash windows;
* installs the bundled Claude Code skills (`olctl`, `olctl-mail-scan`) into
  `%USERPROFILE%\.claude\skills`. A skill folder of the same name that olctl
  did not create is left alone and reported;
* records the version under `HKCU:\Software\olctl` for deployment tools.

No admin rights are needed at any point. Running it again updates in place.
`uninstall.ps1` (also copied into the install folder) removes all of it and
keeps the user's own `%USERPROFILE%\.olctl` (config, audit log, scan marks).
For deploying to many users, see [`packaging/README.md`](packaging/README.md).

### Manual

Copy this folder anywhere on the Windows filesystem, then unblock the script
once. Windows marks anything copied from another machine or downloaded from
the internet, and PowerShell refuses to run a marked script until you clear it:

```powershell
Unblock-File .\olctl.ps1
.\olctl.ps1 doctor --pretty
```

To call it as a bare `olctl` from any directory, put the folder on `PATH`.

**From a PowerShell prompt, prefer `.\olctl.ps1` over `.\olctl.cmd` (or a bare
`olctl` resolved from `PATH`).** The `.cmd` adds a `cmd.exe` hop, and
PowerShell 5.1 strips quotes it thinks are unnecessary when calling a legacy
batch file. An argument containing `&`, `<`, `>`, `|` or `^` can be
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
at the default. Run `Get-ExecutionPolicy -List` and look for `MachinePolicy`
or `UserPolicy` set to anything other than `Undefined`. Group Policy sits above
both `-Scope CurrentUser` and the `-ExecutionPolicy Bypass` that `olctl.cmd`
passes internally, so **neither the command above nor `olctl.cmd` can get
around a real Group-Policy-enforced lock**. That needs an exception from IT
(or a signed script). Where there is no such lock, which is the far more common
case of a machine simply sitting on the Windows default, both fixes above are
genuine, and `olctl.cmd` is a real convenience: it passes `-ExecutionPolicy
Bypass` for you, so someone who doesn't want to touch `Set-ExecutionPolicy` at
all can just run it directly.

## Prerequisites

1. **Outlook Classic installed with a configured profile.** Keep using New
   Outlook day to day. Both are views onto the same mailbox, so anything
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
| `guide` | the agent guide: house rules, commands, this machine's policy (plain text) |
| `doctor [--quick]` | Outlook version, profile, stores, cached mode, protected folders; `--quick` skips the per-store folder walk |
| `folders [--folder P] [--depth N]` | folder tree with item and unread counts |
| `list --folder P [--folder P2 ...] [filters]` | messages, newest first, as JSON rows; several folders are merged into one ordered listing |
| `get REF [--body text\|html\|none] [--strip-quoted] [--headers]` | one message in full; `--strip-quoted` cuts the quoted history off a text body |
| `thread REF [--also P ...] [--since T] [--body text\|none] [--keep-quoted]` | every message of the conversation, oldest first, quoted history stripped |
| `save-attachments REF --out DIR [--pattern "*.pdf"]` | write attachments to disk |
| `move REF --to P` | move between folders |
| `mark REF read\|unread` | read state |
| `flag REF [--text T] [--clear]` | follow-up flag |
| `categorize REF [--add C] [--remove C] [--set A B]` | categories |
| `batch [--file F \| -] [--stop-on-error]` | many commands on one Outlook attach; one JSON argv array per line |
| `draft-reply REF --text ... [--all] [--html] [--subject S] [--attach F] [--from ADDR] [--save-to P] [--supersede OLD_REF]` | reply saved to **Drafts**, not sent |
| `draft-new --to X --subject S --text ... [--cc Y] [--html] [--attach F] [--from ADDR] [--save-to P]` | new mail saved to **Drafts**, not sent |
| `draft-update DRAFT_REF [--text ...] [--whole-body] [--html] [--subject S] [--from ADDR] [--attach F]` | rewrite an existing draft in place |
| `folder-add --folder P --name A[/B]` | create folders (a chain of names) |
| `rules-list [--store S]` | inbox rules of a store with conditions, actions, `is_client_rule` |
| `rules-add --name N --move-to P [--from a@b,c@d] [--subject-contains w1,w2] [--header-contains d1,d2] [--mark-read] [--stop] [--disabled] [--replace]` | new receive rule |
| `rules-delete --name N [--store S]` | delete a rule by name |
| `rules-run [--name N\|--all] [--folder P] [--include-subfolders] [--store S]` | run rules now on a folder (the store's Inbox by default); reports unread counts before and after |

`list` filters: `--unread`, `--read`, `--flagged`, `--since`, `--until`,
`--since-mark`, `--save-mark`, `--sender`, `--subject`, `--category`,
`--has-attachments`, `--body-contains`, `--exclude-sender`, `--exclude-subject`,
`--preset`, `--include-subfolders`, `--after-cursor`, `--limit`,
`--oldest-first`, `--count-only`, `--no-preview`, `--any-class`, `--scan-max`.

`--since` / `--until` take `30m`, `6h`, `3d`, `2w`, `2026-08-18`, or
`2026-08-18T09:30`.

Global: `--pretty`, `--dry-run`, `--json-out FILE`, `--version`, `--help`. Both
flag positions work (`olctl --pretty doctor` and `olctl doctor --pretty`).

`--text` accepts literal text, a file path, or `-` for stdin.

`rules-add` notes from the field: current Outlook builds hand back
`olConditionSenderAddress` conditions as an unsettable base-class object, so
sender matching is done either by exact address (`--from`, resolved through
the GAL) or by domain via `--header-contains`, which matches the raw RFC822
headers and therefore also catches per-campaign rotating sender addresses.
The `--mark-read` flag is attempted in a second pass after the rule is saved
and reported as `mark_read_applied`; some builds refuse to enable that action
on any rule object, in which case moved mail stays unread and the checkbox in
the Rules Wizard is the workaround. `MoveToFolder` needs a late-bound
property put through `InvokeMember` - the plain PowerShell setter silently
no-ops and the rule then fails to save.

### Listing: truncation, paging, several folders

Every listing says whether it is complete. `truncated` is `true` when it
stopped with more to look at, and `truncated_by` says why:

* `limit`: `--limit` rows were returned and there are more items after them.
* `scan_max`: `--scan-max` items were examined without filling the page.
  Items were **skipped unread**, not just left for later, so do not report
  this as "all mail in the window".

Either way `next_cursor` is set. Pass it back unchanged with `--after-cursor`
(and the same filters and `--oldest-first`) to get the next page; pages join
without gaps or overlaps, even across mails received in the same second.

Repeat `--folder` to list several folders as one ordered listing (each row's
`folder` says where it lives), or add `--include-subfolders` to take every
mail folder below the given one. With `--count-only`, `per_folder` breaks the
total down.

**Server-side body search.** `--body-contains TEXT` filters on the body in
Outlook's own `Restrict` (a DASL `LIKE`), so it is a server query rather than
a scan. It switches the whole filter to `@SQL=` syntax; the `filter` field in
the response shows exactly what was sent.

**Noise.** `--exclude-sender SUB` and `--exclude-subject SUB` (both
repeatable, substring, case-insensitive) drop rows after the scan, and the
response counts them in `excluded`. Recurring sets go into `config.json` as
presets and are applied with `--preset NAME`:

```json
{ "presets": { "noise": { "exclude_sender": ["noreply@", "openproject"], "exclude_subject": ["Freigabe"] } } }
```

**Scan marks.** `--save-mark NAME` records when this listing started, and
`--since-mark NAME` starts the next one from there, so a daily scan never
needs a hand-kept "last run" time. On the first run, when the mark does not
exist yet, `--since` is used instead. A truncated listing does not move the
mark: page on with `--after-cursor` and the same `--save-mark`, and the last
page saves it, stamped with the time the first page started. Marks live in
`%USERPROFILE%\.olctl\marks.json`; nothing in the mailbox changes.

```powershell
.\olctl.ps1 list --folder "projects/inbox" --since-mark intake --since 7d --save-mark intake --limit 100 --preset noise
```

### Threads

`thread REF` returns every message of the conversation, oldest first: what
Outlook's conversation view knows for the item's store (Sent Items included),
plus a subject search over the item's folder, its store's Inbox and Sent
Items, and any `--also` folders. Use `--also` for a different store's Sent
Items, e.g. when replies go out from a personal mailbox. Each message has
`match`: `conversation` when it shares the conversation ID, `topic` when only
the subject matched. Some clients start a new conversation ID on reply, so
`topic` matches are kept but labelled; decide for yourself whether they
belong. Bodies have their quoted history stripped (`--keep-quoted` to keep it)
and are capped at 4000 characters each (`--body-chars`).

### Drafts: sender, location, revisions

By default a draft is created in the default mailbox's Drafts, from the
default account. For a reply that should go out from a shared mailbox:

* `--from ADDR` sets the sender: a profile account with that address is used
  directly, otherwise it becomes *on behalf of* (shown as From in Outlook;
  Exchange sends it as that mailbox when you have Send-As).
* `--save-to P` files the draft in another folder, e.g. `projects/drafts`. The
  draft is moved there after saving, so `draft_ref` is the moved copy.
* `"from_by_store": {"projects": "projects@contoso.com"}` in `config.json` does
  both automatically for `draft-reply` to mail in that store.

The response says what happened: `from`, `from_mode` (`account` /
`on_behalf`) and the real `saved_to` folder; `get` on the draft shows
`sent_on_behalf_of` / `send_using_account`. Drafts written without `--html` are
plain-text format: plain text in an HTML item gets its `<mailto:...>` links
duplicated on every rewrite.

To revise a draft instead of piling up new ones:

* `draft-update DRAFT_REF --text new.txt` replaces only the text olctl wrote
  and keeps the quoted original under it. This works for drafts written by
  0.4.1 or later and not edited in Outlook since; otherwise it fails with
  `cannot_locate_text`. `--whole-body` replaces the entire body instead.
* `draft-reply REF --text new.txt --supersede OLD_DRAFT_REF` builds a fresh
  reply and moves the old draft to its Deleted Items (recoverable). The old ref
  must be an unsent mail sitting in a Drafts folder, or nothing happens.

### Batch

`batch` reads one JSON array of argv strings per line from `--file` (or `-`,
stdin), checks every line first, and if all are well-formed runs them on a
single Outlook attach. Blank lines and `#` comments are skipped.

```text
["mark", "<ref>", "read"]
["categorize", "<ref>", "--add", "Reviewed"]
["move", "<ref>", "--to", "projects/Inbox/Processed"]
```

`--dry-run` on the batch applies to every line, so the same file can be
previewed and then run. The result lists each line's `ok` and `data` or
`error`; `--stop-on-error` stops at the first failure (`not_run` counts the
rest). Each mutation is audited as if run on its own.

### Output to a file

`--json-out FILE` writes the response as UTF-8 (no BOM) and prints only
`{"json_out": ..., "bytes": ...}`. Use it instead of `>` in Windows
PowerShell, which writes UTF-16. Errors are still printed in full and leave the
file alone.

### Folder paths

`Inbox`, `Inbox/Vendors`, `Inbox\Vendors`, or store-qualified
`projects/inbox`. A literal slash in a folder name is escaped `\/`. When a
path does not resolve, the error lists the folders that *do* exist at that
level, so an agent can self-correct in one step.

**Localised mailboxes.** Folder names differ per mailbox: `Inbox`,
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
command that looks read-only. `doctor` is not covered by the audit log.
Codes 3, 4, 5, 6, 16 and 23 are safe because those folders always exist.

`archive` is deliberately **not** a token. `OlDefaultFolders` 31 is documented
as `olFolderArchive`, but this Outlook build returns *Quick Step Settings* for
it. Resolving it by type would file mail into a settings folder. Address a
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
[`examples/`](examples/). The snippets below are trimmed to the fields that
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
    { "subject": "Conference room booking kiosk - pilot feedback", "from_email": "alex.rivera@contoso.com", "has_attachments": true, "flag": null },
    { "subject": "Warehouse handheld scanner rollout - timeline", "from_email": "sam.patel@contoso.com", "has_attachments": false, "flag": "flagged" },
    { "subject": "Guest Wi-Fi self-service portal - access request", "from_email": "morgan.ito@northwind-supply.example", "has_attachments": true, "flag": null }
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
category are independent: flag for "needs action", category for bookkeeping:

```powershell
.\olctl.ps1 flag "FEEDFACE...!CAFEBABE..." --text "Confirm rollout order with Sam"
# ... later, once actually reviewed:
.\olctl.ps1 categorize "FEEDFACE...!CAFEBABE..." --add "Reviewed"
.\olctl.ps1 flag "FEEDFACE...!CAFEBABE..." --clear
```

**Draft a reply without hand-typing into the shell.** Write the body to a file
and pass the path. `--text` treats an existing file path as a file, anything
else as literal text:

```powershell
"Thanks Morgan - the SSID handoff section looks right. One question: does the portal expect the guest VLAN tag before or after captive-portal redirect?" |
  Out-File -Encoding utf8 reply-draft.txt
.\olctl.ps1 draft-reply "C0FFEE...!CAFEBABE..." --text reply-draft.txt
```

```json
{ "draft_ref": "BAADF00D...!CAFEBABE...", "saved_to": "Drafts", "sent": false }
```

It lands in **Drafts**, unsent, for you to review. See [Safety
model](#safety-model). Full response: [`examples/draft-reply.json`](examples/draft-reply.json).

**Move a handled item out of the inbox, including protected destination
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

Full response: [`examples/error-not-found.json`](examples/error-not-found.json).
Note `available_stores` and `well_known` alongside `children`, so the error
carries enough to recover without another round trip.

## Safety model

* No `send` and no `delete` command exists. Not gated. Absent. The only
  path to Deleted Items is `draft-reply --supersede`, and only for an unsent
  draft in a Drafts folder.
* Deleted Items, Junk Email, Sent Items and Outbox **of every store in the
  profile** are refused as `move` destinations without `--allow-protected`.
  They are resolved per store via `GetDefaultFolder`, so `Gelöschte Elemente`
  and `Odstraněná pošta` are covered without being listed anywhere. `doctor`
  prints `protected_folders_resolved`. A guard you cannot see is a guard you
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
  "preview_chars": 240,
  "from_by_store": { "projects": "projects@contoso.com" },
  "presets": { "noise": { "exclude_sender": ["noreply@"], "exclude_subject": [] } }
}
```

`OLCTL_HOME` moves the config, audit log and scan marks elsewhere.

### Machine policy

An administrator can restrict what olctl does on a machine with
`%ProgramData%\olctl\policy.json`:

```json
{ "read_only": true, "disabled_commands": ["rules-add", "rules-delete", "rules-run", "folder-add"] }
```

* `read_only` blocks every command that changes mail, drafts, folders or
  rules (`move`, `mark`, `flag`, `categorize`, `draft-*`, `folder-add`,
  `rules-add`/`-delete`/`-run`), including inside `batch`. Their `--dry-run`
  still works, so an assistant can still show what it would do.
* `disabled_commands` blocks the named commands outright.
* Blocked commands fail with `disabled_by_policy`. `doctor` reports the active
  policy, and `olctl guide` tells the agent what it may do on this machine.
* A policy file that exists but cannot be read makes every command fail with
  `bad_policy`, rather than running unrestricted.
* `OLCTL_POLICY_PATH` names an extra policy file that can only **tighten**
  (flags are OR-ed, lists merged). It exists for testing and cannot loosen the
  machine policy.

The policy keeps an assistant inside agreed limits. It is not a security
boundary: a per-user install is writable by its user.

## Wiring it into an AI coding agent

`olctl` has no dependency on any particular agent or AI vendor. It is a
standalone CLI, usable from a plain shell with no AI in the loop at all.

**`olctl guide`** prints the guide for whoever drives it: the house rules
(read before write, ask before changing anything, drafts are never sent,
never invent a ref), the commands with examples, and what this machine's
policy allows. It ships with the tool, so it always matches the installed
version. Any agent can be told "run `olctl guide` first".

The installer also puts two Claude Code skills in `~\.claude\skills`: `olctl`
(points Claude at `olctl guide` and pre-approves the read-only commands) and
`olctl-mail-scan` (a read-only inbox triage). Nothing else in the user's
Claude Code setup is touched.

`agent-config/` is an optional, self-contained example of wiring it into
[Claude Code](https://claude.com/claude-code) specifically, since that is what
this example was built and tested against. Nothing in it is required to use
`olctl`, and nothing here is installed automatically. Copy in only what you
want:

| File in `agent-config/` | Copy to | Purpose |
|---|---|---|
| `CLAUDE.md` | your project root | house rules for the agent: read-before-write, ask before batch mutations, never invent a ref, etc. `CLAUDE.md` is Claude Code's own convention for a file it loads automatically; other harnesses look for their own equivalent (e.g. `AGENTS.md`). The content itself is plain guidance, not Claude-specific, so adapt it freely. |
| `settings.json` | `.claude/settings.json` | a read-only-first permission allowlist: `list`/`get`/`thread`/`folders`/`doctor`/`rules-list`/`--dry-run` are allowed; every command that changes mail, drafts, folders or rules (including `batch`) is denied. Move lines from `deny` to `allow` once you trust it. |
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
  a server round trip. That is far slower, and `--scan-max` starts truncating.
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
  scan, but `--sender`, `--subject`, `--category`, `--has-attachments` and the
  `--exclude-*`/`--preset` exclusions are still applied *after* Outlook's own
  `Restrict`, so pair them with `--since`. `--body-contains` is the exception:
  it runs on the server.
* **`--strip-quoted` is a heuristic.** It cuts at the first reply separator it
  recognises (English, German and Czech Outlook headers, `-----Original
  Message-----`, `On ... wrote:`, `>` blocks). A forwarded mail's content counts
  as quoted, so "FYI, see below" loses what is below; `quoted_chars_removed`
  shows how much was cut.
* **Date literals.** Outlook's filter syntax wants US-format dates regardless of
  the Windows locale, and the tool formats them with InvariantCulture. This
  matters on a German or Czech machine, where the local culture would render
  AM/PM in a form `Restrict` rejects. Do not hand-build filters.

## Reading output on Windows

PowerShell 5.1's `>` and `Out-File` default to UTF-16, which mangles the
Czech and German folder names when the file is read back elsewhere. Use
`--json-out`:

```powershell
.\olctl.ps1 doctor --pretty --json-out doctor.json
```

## Tests

```bash
pwsh -NoProfile -File tests/cli-args.ps1     # or powershell.exe on Windows
```

This checks CLI argument handling: every validation error, both flag
positions, Unicode folder names, batch validation, cursors, marks, presets,
the guide and the machine policy, and that well-formed commands get past
argument checking. It runs against a
throwaway `OLCTL_HOME`, and on a machine with Outlook the "accepted" checks
only run read-only `--count-only` / `doctor --quick` calls. **The COM layer
itself is not covered** by this file.

```powershell
powershell -File tests\contract.ps1           # examples\*.json, no Outlook needed
powershell -File tests\contract.ps1 -Live     # plus real read-only runs against your inbox
```

`tests/contract.ps1` holds the **JSON contract**: the keys and types of every
command's response. Skills, scripts and agents depend on these shapes, so the
rule is: adding a key is a minor change; removing or retyping one is a
breaking change and needs a major version. `examples/` must always pass it.
`-Live` only runs `doctor --quick`, `list`, `get`, `thread` and a `batch` of
`get` against your own inbox. Check anything that changes mail by hand,
`--dry-run` first.

`packaging\build.ps1` runs both test files before it builds a package.
