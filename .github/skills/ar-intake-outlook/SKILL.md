---
name: ar-intake-outlook
description: >-
  Extract ARs (action items) from Outlook emails, normalize fields, and prepare
  rows for tracker ingestion.
user-invocable: true
---

# AR Intake From Outlook

## What this skill does

Scans Outlook email content for AR-style tasks and maps each AR into tracker-ready
fields:
- Task
- Owner
- Status
- ETA
- Remark
- Source

Default lookback window: last 7 days.

## When to use

- User asks to gather ARs from mailbox threads.
- User asks to convert email follow-ups into tracker rows.
- User asks to combine Outlook ARs with Teams ARs.

## Suggested extraction pattern

Look for phrases like:
- AR
- Action Item
- Please close
- Due by
- Owner
- Follow-up

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
| Source | `Outlook:<message-id>` |

## Validation checklist

- De-duplicate by message-id + task text hash.
- Keep unresolved owner as TBD.
- Do not infer ETA from weak signals.
- Preserve traceability to original email.

## Typical next step

Pass normalized rows to `ar-excel-reminder-sync` for upsert and reminder run.
