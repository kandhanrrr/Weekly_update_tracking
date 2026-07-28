---
name: ar-excel-reminder-sync
description: >
  Upsert normalized AR rows into Excel tracker, then run reminder logic for
  upcoming and overdue items.
user-invocable: true
---

# AR Excel And Reminder Sync

## What this skill does

Takes normalized AR rows from Teams/Outlook extraction and:
1. Upserts them into the **AR_Tracking** sheet in the tracker workbook
2. Avoids duplicates using the `Source` key (`Channel:hash`)
3. Updates existing rows (Status/ETA/Remark) if the key already exists
4. Runs `weekly_update_tracker.py` dry-run or live send

## AR_Tracking sheet schema

| Column | Description |
|---|---|
| Date | Date the AR was detected |
| Channel | `Outlook:Inbox`, `Outlook:Sent`, `Outlook:Search`, or `Teams` |
| Task | Action item description |
| Owner | Who must complete it (or TBD) |
| Status | Default `Open`; update manually when resolved |
| ETA | Work-week or date if parsed; `TBD` otherwise |
| Remark | Summary + sender context |
| Source | Unique dedup key — `Channel:12-char-hash` |

## Source key format

Source keys follow the pattern `Channel:hash` where hash = first 12 chars of `sha1("source|YYYY-MM-DD|task[:80]")`:

| Channel | Example Source key |
|---|---|
| `Outlook:Inbox` | `Outlook:Inbox:d18838414214` |
| `Outlook:Sent` | `Outlook:Sent:2d0c5de8c883` |
| `Outlook:Search` | `Outlook:Search:41e496aa99a3` |
| `Teams` | `Teams:4c7590c05a7c` |

## Required inputs

- Excel file path (from config.ini `[TRACKER].excel_file`)
- JSON file with extracted AR rows (produced by `parse_pending_from_graph_exports.py` `--output-json`)

## Upsert commands

```powershell
# Upsert Outlook + Teams ARs
python weekly_update_tracker.py --upsert-ar scripts/ar_rows_latest.json

# Upsert sent-mail follow-ups
python weekly_update_tracker.py --upsert-ar scripts/ar_rows_sent_followups.json
```

## Upsert policy

- If Source already exists in sheet: update Status/ETA/Remark + backfill Channel
- Else insert a new row
- Preserve existing manual rows that have no Source key

## Reminder trigger

```powershell
# Dry-run (preview only)
python weekly_update_tracker.py

# Live send
python weekly_update_tracker.py --send
```

## Output summary table expectations

Always summarize with tables:
- Inserted rows count
- Updated rows count
- Skipped rows count
- Due soon tasks count
- Overdue tasks count
- TBD ETA count

## Safety rules

- Never send live emails unless explicitly requested.
- Keep unknown owner as TBD; do not auto-map to random owner.
- Keep unparseable ETA as TBD.
