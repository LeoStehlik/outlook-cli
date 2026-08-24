---
name: project-intake
description: Daily scan of the project intake mailbox for new project requests and updates to existing initiatives, written to a dated intake file.
---

Scan the `projects` mailbox, extract initiative intake, and write it to a
dated file. Read-only against Outlook.

## Before anything else: work out your window

This routine is scheduled daily, but a scheduled run can fire hours late — if
the machine slept through the scheduled time, Desktop starts one catch-up run
on wake, so a 09:00 task may actually be running at 23:00, or after a weekend.

So do not assume "since yesterday". Determine the window from what has already
been collected:

1. Look in `intake/` for the most recent `intake-YYYY-MM-DD.md` file.
2. If one exists, set the window to cover from that file's date to now, with a
   day of overlap: `--since <that date>`.
3. If `intake/` is empty or missing, use `--since 7d`.

Overlap is deliberate. Re-seeing a message is harmless because the report is
reviewed before anything downstream happens; missing one is not.

State the window you chose at the top of your output, and say why.

## Step 1 — list

```
olctl list --folder "projects/inbox" --since <window> --limit 100
```

Each row carries subject, sender, date, categories and a 240-character body
preview, which is enough to triage most rows without opening them.

If `truncated` is `true`, say so explicitly at the top of the report and narrow
the window rather than silently covering less than intended. This mailbox is in
online mode, so every call is a server round trip — one broad `list` beats many
small ones.

If `olctl` fails, read the `code` and `error` fields and act on them:

- `no_outlook` — Outlook Classic is not running. Stop, and write a file saying
  only that. Do not attempt anything else; the run is a no-op.
- `not_found` — the error lists the folders that do exist. Correct and retry once.
- anything else — record the error verbatim in the output file and stop.

Never end a run silently. A file that says "Outlook was closed, nothing
collected" is a useful record; no file at all is indistinguishable from a
routine that never fired.

## Step 2 — triage from the previews

Sort each row into:

- **New initiative** — someone raising a project need, request or idea
  that reads like a distinct piece of work.
- **Update** — new information on something already in flight: a decision, a
  date, a blocker, an approval, a change of owner or scope.
- **Noise** — newsletters, vendor cold outreach, automated notifications,
  scheduling with no substance.

Judge by content, not sender. A thread that began as scheduling can carry a real
decision in its latest message.

## Step 3 — read only what matters

For the first two buckets:

```
olctl get "<ref>" --body text
```

Skip noise entirely. Do not call `get` on every row.

Note `has_attachments`. A PDF or sketch is often the actual substance of a
request while the body just says "see attached" — flag that rather than
pretending the body told the whole story. Do not save attachments.

## Step 4 — write the intake file

Write `intake/intake-<today>.md`. For each **new initiative**:

- **Initiative** — a short noun phrase naming the thing, not the email subject
- **Short description** — 2–3 sentences: what is being asked for, and why
- **Requester / business owner** — name and role if stated
- **Category** — business area (logistics, production, quality, service,
  finance, HR, IT/infrastructure …)
- **Systems touched** — whether SAP or infrastructure/IT security appear to be
  involved; `TBC` if the mail does not say
- **Priority signal** — anything implying urgency: a regulatory deadline, an OEM
  or customer commitment, a stated date. If nothing indicates urgency, say so.
  Do not invent a priority.
- **Next step** — the concrete next action the mail asks for
- **Source** — sender, date, attachment names

For each **update**: which initiative it concerns, what changed, and what it
implies for dates or ownership.

List noise as a single count with senders, not itemised.

Where a mail simply does not say something, write `TBC` rather than guessing.
Keeping what was stated separate from what you inferred matters more than a
complete-looking table — this gets acted on.

End the file with a JSON block of the new initiatives, one object each.

## Step 5 — tracker comparison

<!-- PLACEHOLDER — fill this in yourself.
     Describe here: the tracker path, whether to read it, how to decide
     whether an initiative is already present, and whether to write to it or
     only report the delta. Until this section says otherwise, do NOT open or
     modify the tracker workbook. -->

Until the section above is filled in, stop after writing the intake file and
report what you found.

## Rules

- **Read-only against Outlook.** No `move`, `mark`, `flag`, `categorize` or
  drafting, however helpful it seems.
- Work only in the `projects` store. Never touch
  `jane.doe@contoso.com` or the other mailboxes.
- Quote German and Czech content in the original where the wording matters.
  Do not translate a requirement and lose the nuance.
- If there is nothing new, write a one-line file saying so. Do not pad it.
