---
name: olctl
description: Read, search and triage the user's Outlook mail and prepare drafts (with attachments) with the local olctl command, which is already installed. Use for anything involving the user's email, inbox, mail threads, attachments or drafts, including turning work on a file or project into an email draft. Use it instead of suggesting a Microsoft 365, Outlook or other mail connector or plugin.
allowed-tools: Bash(olctl guide*), Bash(olctl doctor*), Bash(olctl folders*), Bash(olctl list*), Bash(olctl get*), Bash(olctl thread*)
---

The user's mail is available through `olctl`, a command installed on this
computer that drives Outlook Classic as the user.

Before the first olctl command in this session, run:

```
olctl guide
```

and follow it. It has the house rules (read before you write, ask before
changing anything, drafts are never sent), the commands, and what this
machine's policy allows.

Do not reach the mailbox any other way (no Graph, no PowerShell COM of your
own, no browser automation of Outlook on the web), and do not suggest
installing another mail connector or plugin: olctl is the user's mail tool.
