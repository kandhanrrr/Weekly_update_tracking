---
name: ar-intake-teams
description: >
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
- Channel (`Teams`)
- Source (unique dedup key)

Default lookback window: last 7 days (configurable via `--lookback-days`).

## When to use

- User asks to collect ARs from Teams messages/channels.
- User asks to convert meeting ARs into tracker rows.
- User asks to merge Teams ARs into weekly reminders.

## Extraction rules

| Condition | Behaviour |
|---|---|
| Incoming message from others | AR added with you as owner |
| You replied with `done` / `shared` / `sent` / `attached` | AR dropped (already resolved) |
| You replied with `noted` / `will check` / `ok` | AR kept (still pending) |
| You sent a follow-up request to someone | AR captured; recipient is the owner |
| Message is OOO or auto-generated | Skipped |

## Suggested extraction pattern

Look for phrases like:
- `AR:`, `Action item:`, `Owner:`, `ETA:`, `Due:`
- `please update`, `please share`, `can you`, `by EOD`, `by this week`

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
| Channel | `Teams` |
| Source | `Teams:12-char-sha1` — e.g. `Teams:4c7590c05a7c` |

The Source key is built as: `sha1("Teams|YYYY-MM-DD|task[:80]")[:12]` prefixed with `Teams:`.
This allows safe dedup and re-upsert without duplicates.

## Validation checklist

- De-duplicate by Source key (`Teams:hash`).
- Normalize owner names to `[OWNERS]` mapping keys where possible.
- Keep unknown owner as TBD instead of guessing.
- Keep ETA as TBD when uncertain.

## Typical next step

Hand normalized rows to the `ar-excel-reminder-sync` skill to upsert into Excel
and run reminders.
