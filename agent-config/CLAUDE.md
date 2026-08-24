# Working with Outlook via `olctl`

`olctl` is a local PowerShell CLI that drives the Outlook Classic client on this
machine.
Use it for anything involving my mail. Do not try to reach Outlook any other
way — no Graph, no PowerShell COM of your own, no OWA automation.

## Ground rules

1. **Read before you write.** Establish what a message actually is with `list`
   and `get` before moving, flagging or drafting anything.
2. **The tool cannot send mail and cannot delete mail.** Replies go to Drafts.
   Never present a draft as "sent" — say it is waiting in Drafts for review.
3. **Ask before mutating more than 5 items in one go.** Show me the list first.
4. **Use `--dry-run` when you are unsure** which items a plan would touch, then
   report the plan before executing it.
5. **Refs change when an item moves.** After `move`, use the `new_ref` from the
   response. A `not_found` error usually means you are holding a stale ref —
   re-run `list` rather than guessing.
6. **Never invent a ref.** They only come from `list` or a prior command's
   output.

## Invocation

Call `.\olctl.ps1` directly from PowerShell rather than `olctl.cmd` — the
`.cmd` routes through `cmd.exe`, which re-parses `&`, `<`, `>`, `|` and `^` in
arguments and can corrupt folder names and subjects.

If a command fails with something like "running scripts is disabled on this
system", that is PowerShell's execution policy, not an `olctl` error. Tell me
rather than silently switching to `olctl.cmd` — that only routes around the
*default* Restricted policy (it runs `powershell -ExecutionPolicy Bypass`
under the hood) and does nothing if the policy is actually locked down by
Group Policy, in which case switching quietly would waste both our time.

## Output contract

Every command prints one JSON object. Exit 0 = ok. Exit 1 = handled error with
`code` and `error` fields — read them, they are actionable (`not_found` includes
the folder names that do exist; `bad_argument` explains the accepted formats).
Parse the JSON, do not scrape it.

## Scope: the `projects` mailbox

Unless I say otherwise, work in the **`projects`** store. It is mounted in
the profile, so plain paths work and `--shared` is not needed:

```powershell
.\olctl.ps1 list --folder "projects/inbox" --unread --since 3d
.\olctl.ps1 folders --folder projects --depth 2
```

Do not touch `jane.doe@contoso.com` (my personal mailbox) or the other
stores unless I ask for them by name in that message. If a task seems to need
another mailbox, say so and stop rather than reaching for it.

We are still exploring. That means: prefer reading, describe what you found
before changing anything, and treat every mutation as something I want to see
proposed first. `--dry-run` liberally.

## This profile has six mailboxes — always name the store

Folder names are in three languages here. Never assume "Inbox" means what you
want: an unqualified well-known name resolves against the **default** mailbox
(`jane.doe@contoso.com`). Prefix the store name for anything else.

| Store | What it is | Inbox is called | Where |
|---|---|---|---|
| `jane.doe@contoso.com` | Jane's own mailbox (the default) | `Inbox` | EXO, **cached** |
| `projects` | **the working mailbox** | `Inbox` | EXO, online mode |
| `SA_AutomationAgent` | service account | `Doručená pošta` | EXO, online mode |
| `Contoso Config Tool` | shared | `Posteingang` | on-prem, online mode |
| `Regional Office` | shared | `Posteingang` | on-prem (`EXCH01`), online mode |
| `Api` | shared | `Inbox` | EXO (onmicrosoft), online mode |

The English tokens (`inbox`, `drafts`, `sent items`, `deleted items`,
`junk email`, `outbox`) work directly under any store name and resolve inside
that store, so `--folder "SA_AutomationAgent/inbox"` is correct and portable.
`archive` is NOT a token — use the folder's real name. Run `olctl doctor` if you
need the exact localised names or the subfolder tree.

**Only Jane's own mailbox is cached.** Every other store, including
`projects`, is in online mode, so each filter is a server round trip.
Always bound a list with `--since` and `--limit`, expect `truncated: true` on
big folders, and prefer `--count-only` when a number is all that is needed.
Do not use `--shared`; these mailboxes are mounted as stores, so use the store
path.

## Typical commands

```powershell
.\olctl.ps1 doctor                                      # check the connection first if anything is odd
.\olctl.ps1 folders --depth 2                           # find the exact folder path
.\olctl.ps1 list --folder "projects/inbox" --unread --since 3d --limit 20
.\olctl.ps1 list --folder "projects/inbox" --sender supplier --has-attachments
.\olctl.ps1 list --folder "projects/inbox" --count-only
.\olctl.ps1 get "<ref>" --body text
.\olctl.ps1 save-attachments "<ref>" --out C:\Users\<you>\attachments --pattern "*.pdf"
.\olctl.ps1 mark "<ref>" read
.\olctl.ps1 flag "<ref>" --text "Follow up"
.\olctl.ps1 categorize "<ref>" --add "Reviewed"
.\olctl.ps1 move "<ref>" --to "projects/Inbox/Processed"
.\olctl.ps1 draft-reply "<ref>" --text drafts/reply.txt
```

## Efficiency

- Always bound a `list` with `--since` and `--limit`. Unbounded scans of a large
  Inbox are slow and get truncated anyway (`truncated: true` in the response).
- `--sender`, `--subject`, `--category` and `--has-attachments` are applied
  after Outlook's own filter, so combine them with `--since`.
- Use `--no-preview` when you only need metadata, and `--count-only` when you
  only need a number.
- Every command spawns powershell.exe and attaches to Outlook, costing roughly
  half a second to a second. Prefer one `list` that returns many rows over
  several small calls — batching matters more here than it would for a local
  in-process tool.
- Do not call `get` on every row of a listing. The listing already has subject,
  sender, date, flags, categories and a body preview — that is usually enough to
  decide. Call `get` for the ones that matter.

## When drafting replies for me

Write the reply body to a file and pass the path to `--text`, rather than
embedding long text in a shell argument. Match the language of the incoming
mail (German for German, Czech for Czech, English otherwise). Keep it short and
plain. Tell me the `draft_ref` so I can find it.
