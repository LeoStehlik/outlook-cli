# olctl guide

olctl reads and organises the user's mail in Outlook Classic on this computer,
as the user, with the rights the user already has. It prints one JSON object
per command. It cannot send mail and cannot delete mail: replies are saved as
drafts for the user to review and send.

Read this once per session before using olctl.

## House rules

1. Read before you write. Use `list`, `get` or `thread` to establish what a
   mail is before moving, flagging or drafting anything.
2. Ask the user before changing anything: moving, marking, flagging,
   categorising or creating a draft. Show what you will do; `--dry-run` shows
   it without doing it.
3. More than 5 changes at once: show the list first, then put them in one
   `batch` file, run it with `--dry-run`, and only then for real.
4. Never say a draft was sent. Say it is waiting in Drafts for the user to
   review, and give its `draft_ref`.
5. Never invent a ref. Refs only come from olctl output. A ref changes when a
   mail is moved; use `new_ref` from the `move` response. `not_found` usually
   means a stale ref: list again rather than guessing.
6. Bound every listing with `--since` and `--limit`.
7. If a command fails, read `code` and `error` in the JSON: they say what to do
   (`not_found` lists the folders that do exist).

## Folders

Plain names work for the user's own mailbox: `inbox`, `sent items`, `drafts`,
`deleted items`, `junk email`, and subfolders like `inbox/Projects`. Folder
names in the user's language (Posteingang, Dorucena posta) also work. There is
no `archive` token: use the folder's real name. `olctl folders --depth 2`
shows the tree.

## Reading mail

```
olctl list --folder inbox --since 3d --limit 30
olctl list --folder inbox --unread --since 1w --limit 50
olctl list --folder inbox --since 2w --sender "acme" --has-attachments
olctl list --folder inbox --since 4w --body-contains "invoice 4711"
olctl list --folder inbox --folder "sent items" --since 1w --limit 50
olctl list --folder inbox --count-only --unread
olctl get "<ref>" --body text --strip-quoted
olctl thread "<ref>"
```

- Times for `--since`/`--until`: `30m` (minutes), `6h`, `3d`, `2w`,
  `2026-08-18` or `2026-08-18T09:30`. There is no month unit: use `4w`.
- A listing row already has subject, sender, date, flags, categories and a
  preview. Decide from it; only `get` the mails that matter.
- `--strip-quoted` removes the quoted history of a reply. A forwarded mail's
  content counts as quoted too: check `quoted_chars_removed`, and read without
  `--strip-quoted` if what was forwarded is the point.
- For "what happened in this conversation", use one `thread` call instead of
  several `get`s. Messages with `match: topic` only share the subject; say so
  if you rely on one.
- `truncated: true` means there is more. Continue with
  `--after-cursor <next_cursor>` and the same filters. Do not present a
  truncated listing as complete.
- Noise: `--exclude-sender` and `--exclude-subject` drop rows and count them in
  `excluded`. Report that count instead of listing those mails.
- Recurring scans: `--since-mark NAME --since 7d --save-mark NAME` continues
  from where the last complete scan of that name started.

## Changing mail (ask first)

```
olctl mark "<ref>" read
olctl flag "<ref>" --text "Follow up"
olctl categorize "<ref>" --add "Done"
olctl move "<ref>" --to "inbox/Projects"
olctl batch --file plan.jsonl --dry-run
```

A batch file has one JSON array per line, for example:

```
["mark", "<ref>", "read"]
["move", "<ref>", "--to", "inbox/Projects"]
```

Moving to Deleted Items, Junk, Sent Items or Outbox is refused.

## Drafting replies (ask first)

Write the reply text to a file and pass the path; do not put long text on the
command line.

```
olctl draft-reply "<ref>" --text reply.txt
olctl draft-reply "<ref>" --text reply.txt --all
olctl draft-update "<draft_ref>" --text reply-v2.txt
```

- Match the language of the incoming mail. Keep it short and plain.
- When the user asks for changes, use `draft-update` on the draft you made,
  not a new `draft-reply`, so there is only one version. If it reports
  `cannot_locate_text`, ask the user before using `--whole-body`.
- Tell the user the `draft_ref` and that the draft is in Drafts, unsent.

## Attachments

```
olctl save-attachments "<ref>" --out C:\Users\<user>\Downloads\mail --pattern "*.pdf"
```

## When something is odd

`olctl doctor --quick` checks that Outlook Classic is reachable. `no_outlook`
means Outlook Classic is not installed or has no profile: the user should open
Outlook Classic once and let it finish setting up.

`olctl update --check` tells you whether a newer olctl exists. Ask the user
before running `olctl update`, which installs it.
