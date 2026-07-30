# AR Tracking (Task 2)

Start point: [../START_HERE.md](../START_HERE.md)

This README is only for AR tracking.

## Goal

Create and maintain a dedicated AR tracker workbook from Teams plus Outlook Inbox and Outlook Sent data.

## Rule

For AR tracking, workbook creation is mandatory.
- Always create your own AR workbook from source data.
- Do not depend on user-provided workbook for this task.

## Source data window

- Use last 1WW (7 days) data.

## Step A - Extract AR rows from Teams + Outlook Inbox + Sent

```powershell
python scripts/parse_pending_from_graph_exports.py `
  --outlook-json "scripts/outlook_inbox_fresh.json" `
  --outlook-search-json "scripts/outlook_search_ar.json;scripts/outlook_sent_followups.json" `
  --teams-json-glob "scripts/teams_*.json;scripts/t_*.json" `
  --lookback-days 7 --strict `
  --output-json scripts/ar_rows_latest.json
```

## Step B - Must create AR workbook from extracted data

```powershell
python scripts/create_excel_from_ar_data.py `
  --task ar `
  --input-json "scripts/ar_rows_latest.json;scripts/ar_rows_sent_followups.json" `
  --lookback-days 7
```

Generated workbook:
- artifacts/ar/AR_Tracking_Auto.xlsx
- Sheet: AR_Tracking

Columns:
- Date, Channel, Task, Owner, Status, ETA, Remark, Source

Source format:
- Teams:<hash>
- Outlook:Inbox:<hash>
- Outlook:Sent:<hash>
- Outlook:Search:<hash>

## Step C - Optional sync to configured tracker

If you also want these AR rows inside configured workbook:

```powershell
python weekly_update_tracker.py --upsert-ar scripts/ar_rows_latest.json
python weekly_update_tracker.py --upsert-ar scripts/ar_rows_sent_followups.json
```

## Expected result

- AR workbook is always created by the pipeline.
- Data is strictly from last 1WW Teams + Outlook exports.
- Standard output workbook is `artifacts/ar/AR_Tracking_Auto.xlsx`.
- HTML tab view report is generated at `artifacts/weekly/weekly_tracker_report.html`.

## Output format (tab-style, mandatory)

Use this output structure when presenting AR run results:

- Tab: Overview
  - Lookback window
  - Source files used
  - Total pending extracted
  - Final artifact path
- Tab: Source Scan Summary
  - Teams scanned, Outlook scanned, pending extracted counts
- Tab: Pending Tasks by Owner
  - Owner-wise count table
- Tab: Upsert Summary
  - Inserted, Updated, Skipped, Total processed
- Tab: Reminder Outcome
  - Due soon, Overdue, TBD ETA counts
- Tab: Overdue Owners
  - Owner-wise overdue counts (if any)

Hover tooltip details (mandatory for tab visuals):

- On cursor hover, show detailed stats for the selected point/row.
- Minimum fields to show:
  - Category labels (Module/Subflow/Owner context)
  - Type context (for example TEST_TYPE or Status bucket, where available)
  - Max
  - Upper fence
  - Q3
  - Median
  - Mean
  - Q1
  - Lower fence
- If a tab is rendered as a plain table (non-chart), provide the same details as an expanded row/details panel.
