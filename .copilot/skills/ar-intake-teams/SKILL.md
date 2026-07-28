---
name: ar-intake-teams
description: >-
  Extract ARs (action items) from Microsoft Teams chats/channels, normalize fields,
  and prepare rows for the weekly update tracker Excel.
user-invocable: true
---

# AR Intake From Teams

## What this skill does

Pulls recent Teams messages, finds AR-style entries, and converts them into
structured tracker rows:
- Task
- Owner
- Status
- ETA
- Remark
- Source

Default lookback window: last 7 days.

## When to use

- User asks to collect ARs from Teams messages/channels.
- User asks to convert meeting ARs into tracker rows.
- User asks to merge Teams ARs into weekly reminders.

## Suggested extraction pattern

Look for phrases like:
- AR:
- Action item:
- Owner:
- ETA:
- Due:

If structured fields are missing:
- Set Status = Open
- Set ETA = TBD
- Put raw text in Remark

## Output schema

| Field | Rule |
|---|---|
| Task | Required, concise action statement |
| Owner | Parse from message mention/name; else TBD |
| Status | Default Open if missing |
| ETA | Convert to WWxx/WWxxpx if possible; else TBD |
| Remark | Keep source context and original sentence |
| Source | `Teams:<chat/channel>:<message-id>` |

## Validation checklist

- Do not create duplicate rows if same source id already exists.
- Normalize owner names to `[OWNERS]` mapping keys where possible.
- Keep unknown owner as TBD instead of guessing.
- Keep ETA as TBD when uncertain.

## Typical next step

Hand normalized rows to the `ar-excel-reminder-sync` skill to upsert into Excel
and run reminders.
