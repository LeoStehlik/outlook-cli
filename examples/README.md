# Example payloads

Hand-written, fully fictional sample output for every `olctl` command — not a
live capture. They exist so you can see the exact JSON shape a command returns
without needing Outlook, a mail profile, or a live run to check it.

The scenario behind all of these: a fictional company `Contoso`, a shared
working mailbox called `projects`, and your own default mailbox
`jane.doe@contoso.com` — the same setup used in `README.md` and
`agent-config/CLAUDE.md`. `ref` values use obviously-fake hex (`DEADBEEF`,
`CAFEBABE`, `FEEDFACE`, ...) rather than plausible-looking real ones, and are
much shorter than a real `EntryID!StoreID` pair, which typically runs to a few
hundred hex characters on each side of the `!`.

| File | Command |
|---|---|
| `doctor.json` | `olctl doctor` |
| `folders.json` | `olctl folders --folder projects --depth 2` |
| `list.json` | `olctl list --folder "projects/inbox" --unread --since 3d` |
| `get.json` | `olctl get "<ref>" --body text` |
| `save-attachments.json` | `olctl save-attachments "<ref>" --out ...` |
| `move.json` | `olctl move "<ref>" --to "projects/Inbox/Processed"` |
| `mark.json` | `olctl mark "<ref>" read` |
| `flag.json` | `olctl flag "<ref>" --text "Follow up"` |
| `categorize.json` | `olctl categorize "<ref>" --add "Reviewed"` |
| `draft-reply.json` | `olctl draft-reply "<ref>" --text ...` |
| `draft-new.json` | `olctl draft-new --to ... --subject ... --text ...` |
| `error-not-found.json` | any command given a folder path that doesn't resolve |

Every successful command prints just the object under `data` here — `olctl`
unwraps it on success so `olctl list ... | jq .items` works directly, and only
wraps the response in `{"ok": false, ...}` on failure. `error-not-found.json`
shows that failure shape.

## Setup Walkthrough

- [`first-five-minutes.md`](first-five-minutes.md) shows the first Windows commands and expected failure path.
