---
description: Scan the project intake mailbox for new project requests and updates to existing initiatives
argument-hint: "[time window, e.g. 3d, 1w, 2026-08-01 — defaults to 7d]"
allowed-tools: Bash(olctl list*), Bash(olctl get*), Bash(olctl folders*), Bash(olctl doctor*)
---

Scan the `projects` mailbox and extract initiative intake, read-only.

Window: `$1` if given, otherwise `7d`.

## Step 1 — list

```
olctl list --folder "projects/inbox" --since <window> --limit 100
```

Read the JSON. Each row already carries subject, sender, date, categories and a
240-character body preview — enough to triage most rows without opening them.

If `truncated` is `true`, say so explicitly in your report and narrow the window
rather than silently covering less than I asked for. This mailbox is in online
mode, so each call is a server round trip: one broad `list` beats many small ones.

## Step 2 — triage from the previews

Sort each row into one of:

- **New initiative** — someone raising a project need, request or idea
  that reads like a distinct piece of work.
- **Update** — new information on something already in flight: a decision, a
  date, a blocker, an approval, a change of owner or scope.
- **Noise** — newsletters, vendor cold outreach, automated notifications,
  scheduling back-and-forth with no substance.

Judge by content, not by sender. A thread that started as scheduling can carry a
real decision in its latest message.

## Step 3 — read only what matters

For rows in the first two buckets:

```
olctl get "<ref>" --body text
```

Skip this for noise. Do not call `get` on every row — the preview is usually
enough to know whether a row deserves it.

Note `has_attachments`: an attached PDF or sketch is often the actual substance
of a request, and the body just says "see attached". Flag that in your report
rather than pretending the body told the whole story. Don't save attachments
unless I ask.

## Step 4 — report

For each **new initiative**, give me:

- **Initiative** — a short noun phrase naming the thing, not the email subject
- **Short description** — 2–3 sentences: what is being asked for and why
- **Requester / business owner** — name and role if stated
- **Category** — the business area it belongs to (logistics, production,
  quality, service, finance, HR, IT/infrastructure …)
- **Systems touched** — whether SAP or infrastructure/IT security appear to be
  involved, and say `TBC` if the mail does not indicate it
- **Priority signal** — anything in the mail implying urgency: a regulatory
  deadline, an OEM or customer commitment, a stated date. If nothing indicates
  urgency, say so; do not invent a priority.
- **Next step** — the concrete next action the mail is asking for
- **Source** — sender, date, and attachment names if any

For each **update**, give me: which initiative it appears to concern, what
changed, and what it implies for dates or ownership.

List noise as a single count with senders, not itemised.

Where the mail simply does not say something, write `TBC` rather than guessing.
Distinguishing what was stated from what you inferred matters more to me than a
complete-looking table — I have to act on this.

Finish with a compact JSON block of the new initiatives, one object per
initiative, so I can pipe it somewhere without retyping.

## Rules

- **Read-only.** No `move`, `mark`, `flag`, `categorize` or drafting, even if it
  seems helpful. Ask first.
- Work only in the `projects` store. Never touch
  `jane.doe@contoso.com` or the other mailboxes unless I name one.
- Quote German and Czech content in the original where the wording matters —
  don't translate a requirement and lose the nuance.
- If `olctl` returns an error, read the `code` and `error` fields and act on
  them. `not_found` on a folder lists the folders that do exist.
