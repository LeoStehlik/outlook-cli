---
name: olctl-mail-scan
description: Scan the user's inbox for a time window and report what needs the user's action, what is just information, and what is noise. Read-only. Use when the user asks what is in their mail, what they missed, or what needs attention.
argument-hint: "[time window, e.g. 1d, 3d, 1w; defaults to 3d]"
allowed-tools: Bash(olctl guide*), Bash(olctl list*), Bash(olctl get*), Bash(olctl thread*)
---

Scan the user's inbox read-only. Window: the argument if given, otherwise `3d`.

If you have not read it in this session, run `olctl guide` first.

## 1. List

```
olctl list --folder inbox --since <window> --limit 100
```

If `truncated` is `true`, continue with `--after-cursor <next_cursor>` and the
same filters until it is `false`. If you stop early, say so at the top of the
report.

## 2. Triage from the previews

Put each mail in one bucket, judging by content, not sender:

- **Needs action**: someone is asking the user for something: a decision, an
  answer, an approval, a document, a date.
- **For information**: relevant news or updates with nothing to do.
- **Noise**: newsletters, automated notifications, marketing, scheduling
  back-and-forth with no substance.

## 3. Read only what matters

For **Needs action** mails, read them with
`olctl get "<ref>" --body text --strip-quoted`, or `olctl thread "<ref>"` when
the history matters. If a mail has attachments, mention their names: the
attachment is often the real request.

## 4. Report

- **Needs action**: for each: who, what they need, by when (only if the mail
  says so), and the suggested next step. Most urgent first.
- **For information**: one line each.
- **Noise**: a count, with the main senders.

Quote the original wording where it matters; do not translate it away. Write
"not stated" rather than guessing a deadline or owner.

Change nothing: no marking, moving, flagging or drafting unless the user asks
after reading the report.
