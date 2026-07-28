---
name: ar-excel-reminder-sync
description: >-
  Upsert normalized AR rows into Excel tracker, then run reminder logic for
  upcoming and overdue items.
user-invocable: true
---

# AR Excel And Reminder Sync

## What this skill does

Takes normalized AR rows from Teams/Outlook extraction and:
1. Upserts them into the tracker workbook
2. Avoids duplicates using source key
3. Runs weekly_update_tracker dry-run or live send

## Required inputs

- Excel file path (from config.ini `[TRACKER].excel_file`)
- Sheet name (`auto` or explicit)
- Normalized rows with fields:
  Task, Owner, Status, ETA, Remark, Source

## Upsert policy

- If Source already exists in sheet: update Status/ETA/Remark
- Else insert a new row
- Preserve existing manual rows that have no Source key

## Reminder trigger

- Dry-run: `python weekly_update_tracker.py`
- Live: `python weekly_update_tracker.py --send`
- Force weekly summary: `python weekly_update_tracker.py --send --weekly-summary`

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
