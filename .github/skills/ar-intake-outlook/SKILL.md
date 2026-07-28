---
name: ar-intake-outlook
description: >
  Extract ARs (action items) from Outlook emails, normalize fields, and prepare
  rows for tracker ingestion.
user-invocable: true
---

# AR Intake From Outlook

## What this skill does

Scans Outlook email content (inbox and sent items) for AR-style tasks and maps each AR into
tracker-ready fields:
- Task
- Owner
- Status
- ETA
- Remark
- Channel (`Outlook:Inbox`, `Outlook:Sent`, or `Outlook:Search`)
- Source (unique dedup key)

Default lookback window: last 7 days (configurable via `--lookback-days`).

## When to use

- User asks to gather ARs from mailbox threads.
- User asks to convert email follow-ups into tracker rows.
- User asks to scan sent items for follow-ups issued to others.
- User asks to combine Outlook ARs with Teams ARs.

## Two extraction modes

### Inbox mode (`--outlook-json`)
Fetches emails from the Outlook inbox. Skips:
- Emails where user is CC-only and not named in the body.
- OOO / auto-reply emails (`Automatic reply`, `Out of Office`, `Auto:`, etc.).
- Emails where the greeting addresses someone else entirely.

### Sent Items mode (`--outlook-search-json` with `sent` or `followup` in filename)
Files named with `sent` or `followup` in the stem (e.g. `outlook_sent_followups.json`) are
automatically labeled `Outlook:Sent`. The extractor finds messages where **you** asked someone
for data, assistance, or a task — and assigns **them** as the owner.

Signals for sent-mail ARs:
- Imperative requests: `please provide`, `can you share`, `by EOD`, `send me`
- Self-commitments: `I will`, `we need to`, `I need to`

## Suggested extraction pattern

Look for phrases like:
- AR, Action Item, Please close, Due by, Owner, Follow-up
- `can you share`, `please provide`, `please update`, `by EOD`, `by this week`

If explicit fields are not available:
- Status = Open
- ETA = TBD
- Remark includes email subject + sender + snippet

## Output schema

| Field | Rule |
|---|---|
| Task | Required, remove noisy prefixes |
| Owner | Parse named owner; else TBD |
| Status | Default Open |
| ETA | WW format if parsed; else TBD |
| Remark | Source summary and original context |
| Channel | `Outlook:Inbox`, `Outlook:Sent`, or `Outlook:Search` based on input file |
| Source | `Channel:12-char-sha1` — e.g. `Outlook:Inbox:d18838414214` |

The Source key is built as: `sha1("Outlook|YYYY-MM-DD|task[:80]")[:12]` prefixed with the Channel.
This allows safe dedup and re-upsert without duplicates.

## Validation checklist

- De-duplicate by Source key (Channel-prefixed hash).
- Keep unresolved owner as TBD.
- Do not infer ETA from weak signals.
- Preserve traceability to original email.
- Sent-items files: bypass self-sender filter (extractor already handles this).

## Typical next step

Pass normalized rows to `ar-excel-reminder-sync` for upsert and reminder run.
