# Changelog

## v0.4.1

Brings the published code up from 0.3.3 to the current line. v0.4.0 was a
docs-only release; everything below is new code.

### Folders and rules
- `folder-add` creates folders (a chain of names).
- `rules-list`, `rules-add`, `rules-delete`, `rules-run`: server-side receive
  rules per store, and Run Rules Now on a folder with before/after counts.
- `--attach` on `draft-reply` / `draft-new`.
- `--since` / `--until` with absolute dates fixed under Windows PowerShell 5.1.

### Drafts
- `draft-reply` / `draft-new` take `--from ADDR` (a profile account, otherwise
  send-on-behalf) and `--save-to P` (the draft is moved there after saving).
  `from_by_store` in `config.json` sets both automatically for replies to mail
  in a given store.
- New `draft-update DRAFT_REF` rewrites a draft in place. By default it swaps
  only the text olctl wrote and keeps the quoted original; `--whole-body`
  replaces everything.
- `draft-reply --supersede OLD_DRAFT_REF` builds a fresh reply and moves the old
  unsent draft to Deleted Items. Refuses anything that is not an unsent mail in
  a Drafts folder.
- Text drafts are saved in plain-text format, so quoted `<mailto:...>` links are
  not duplicated each time a draft is rewritten. `--html` drafts stay HTML.
- Responses report `from`, `from_mode` and the real `saved_to` folder; `get`
  shows `sent_on_behalf_of` / `send_using_account`.

### list
- `truncated_by` (`limit` / `scan_max`) and `next_cursor`; `--after-cursor`
  continues a listing with no gaps or overlaps, ties included.
- `--folder` is repeatable; `--include-subfolders` adds every mail folder
  below. Results are merged into one ordered listing.
- `--body-contains TEXT`: server-side body search (DASL).
- `--exclude-sender`, `--exclude-subject`, and named `--preset`s from
  `config.json`; `excluded` counts what they removed.
- `--since-mark NAME` / `--save-mark NAME`: remember where the last complete
  scan started. A truncated listing never moves the mark.

### New commands and options
- `thread REF`: the whole conversation, oldest first, across the store's
  folders and any `--also` folders, quoted history stripped. Topic-only
  matches are labelled `match: topic`.
- `batch`: one JSON argv array per line, validated up front, run on one Outlook
  attach; honours `--dry-run` and `--stop-on-error`.
- `get --strip-quoted`.
- `doctor --quick`: skips the per-store folder walk.
- `--json-out FILE`: write the response as UTF-8 without BOM.

### Fixes
- `--help` listed `archive` as a folder token; it is not one (it resolves to
  *Quick Step Settings* on some Outlook builds).
- Non-ASCII literals in `olctl.ps1` are escaped: the file has no BOM, so
  Windows PowerShell 5.1 read them as ANSI.
- `tests/cli-args.ps1` also runs on a machine with Outlook and uses a throwaway
  `OLCTL_HOME`.

## v0.4.0

- Add first-five-minutes README setup matrix for Windows users.
- Add docs/windows-setup-proof.md with supported environments, limits, and non-Windows verification guidance.
- Add examples/first-five-minutes.md.
- Bump olctl version to 0.4.0.
