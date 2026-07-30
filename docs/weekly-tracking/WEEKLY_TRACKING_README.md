# Weekly Tracking (Task 1)

Start point: [../START_HERE.md](../START_HERE.md)

This file is only for weekly tracking.

## Scope

- Use a user-provided workbook.
- Do not auto-create any weekly workbook.
- Process only active statuses: WIP, Open, Not yet started.
- Treat equivalent wording as active: In Progress, Not Started, Pending.

## Required workbook fields

- Task
- Owner
- Status
- ETA
- Remark

## Configuration

Set workbook details in config.ini:

- [TRACKER] excel_file = <your workbook>
- [TRACKER] sheet_name = <your sheet>

## Commands

Dry run (no emails sent):

python weekly_update_tracker.py

Live reminder send (only when explicitly requested):

python weekly_update_tracker.py --send

Force weekly summary send:

python weekly_update_tracker.py --send --weekly-summary

## Expected output

Dry run / live run prints:

- Run mode (DRY-RUN or SEND)
- Active tasks processed
- Due-soon reminders
- Overdue follow-ups
- TBD or unparseable ETA count
- Weekly summary behavior

Weekly status summary output (module-wise):

- Include a module-level summary table in weekly reporting.
- Use status buckets such as Completed, Open, WIP, Not yet started, and Dropped.

Example format:

| Module | Completed | Open | WIP | Not yet started | Dropped | Total |
|---|---:|---:|---:|---:|---:|---:|
| PMAX | 6 | 2 | 3 | 1 | 0 | 12 |
| DTS | 4 | 1 | 2 | 2 | 1 | 10 |
| BGR | 8 | 0 | 1 | 1 | 0 | 10 |
| Unassigned | 1 | 2 | 0 | 1 | 0 | 4 |

Weekly output path:

- Weekly updates are written back to the configured tracker workbook path in `config.ini`:
	- `[TRACKER] excel_file = <your workbook path>`
- No auto-generated weekly artifact file is created under `artifacts/`.

## Reminder behavior

- Reminder starts 2 days before ETA.
- Overdue follow-up continues daily while task is still active.
- Reminder stops when status is moved to a closed/completed state.

## Boundary

- Weekly task does not create workbooks.
- AR workbook creation belongs only to Task 2 in [../ar-tracking/AR_TRACKING_README.md](../ar-tracking/AR_TRACKING_README.md).
