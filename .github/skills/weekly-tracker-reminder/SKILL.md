---
name: weekly-tracker-reminder
description: >-
  Run weekly tracker reminders and follow-up using a user-provided workbook,
  with active-status normalization and ETA-based reminder rules.
user-invocable: true
---

# Weekly Tracker Reminder

## Purpose

Run weekly reminder/follow-up logic from the user-provided Excel tracker.

## Scope

- Weekly tracking only.
- Does not auto-create workbook.

## Active status logic

Treat these as active after normalization:
- WIP
- Open
- Not yet started

Equivalent wording examples:
- In Progress, In-Progress, Active, Ongoing
- Not Started, Todo, To Do, Planned
- Pending, New, Reopened

## Commands

```powershell
# Dry run
python scripts/weekly/run_weekly_tracker.py

# Live send (explicit only)
python scripts/weekly/run_weekly_tracker.py --send

# Force weekly summary
python scripts/weekly/run_weekly_tracker.py --send --weekly-summary
```

## Safety

- Never send live emails unless explicitly requested.
- Continue follow-up until status is closed/completed equivalent.
