# Weekly Update Tracker

Flexible Excel-based task reminder automation for any team. The tracker reads a workbook,
detects a usable sheet and column headers, then sends reminder or follow-up emails based
on ETA and status.

---

## What It Does

- Sends a reminder when ETA is within 2 days.
- Sends a daily overdue follow-up when ETA has passed and the task is still open.
- Sends a weekly summary to the team on Fridays.
- Skips closed tasks automatically.

The script is now config-driven and more generic:
- Sheet name can be configured or auto-detected.
- Column order can vary.
- The workbook name can vary.
- Owner email mapping is kept in `config.ini`.

---

## Files

```
Weekly_update+tracking/
├── weekly_update_tracker.py     # Main tracker script
├── config.ini                   # Tracker, email, and owner settings
├── save_password.py             # Save email password to Windows Credential Manager
├── setup_scheduler.bat          # Register the daily Task Scheduler job
└── README.md                    # This file
```

---

## Setup

1. Open [config.ini](config.ini) and update the `[TRACKER]` and `[EMAIL]` sections.
2. Set the Excel file path in `excel_file`.
3. Set `sheet_name` to a specific tab name, or leave it as `auto` to let the script
   find a sheet with recognizable headers.
4. Set `credential_service` to a unique name for the team or machine.
5. Fill in `smtp_user`, `team_email`, and the `[OWNERS]` mappings.
6. Run:

```bash
python save_password.py
```

This stores the password in Windows Credential Manager for the current Windows account.

7. Optional: double-click `setup_scheduler.bat` to register the weekday schedule.

---

## Usage

```bash
# Dry run, no email sent
python weekly_update_tracker.py

# Send real emails
python weekly_update_tracker.py --send

# Force weekly summary on any day
python weekly_update_tracker.py --send --weekly-summary
```

---

## Excel Requirements

The workbook does not need a fixed filename or a fixed sheet name, but it must contain
headers that the tracker can recognize.

Required logical fields:
- Task
- Status
- ETA

Recommended fields:
- Owner
- Remark
- Priority

Supported header aliases include:
- Task: `Task`, `Item`, `Title`, `Description`, `Work Item`
- Status: `Status`, `State`, `Progress`
- ETA: `ETA`, `Due`, `Due Date`, `Target`, `Target Date`, `Deadline`
- Owner: `Owner`, `Assignee`, `Responsible`, `Person`, `Lead`

Recommended status values:
- `WIP`
- `Not yet started`
- `Open`
- `Done`
- `Dropped`
- `NA`
- `Review Done/Closed`

Recommended ETA format:
- `WW30`
- `WW25p5`
- `WW25p3`
- `TBD`

If a workbook uses different labels or formats, update the headers or extend the aliases
in the script.

---

## Behavior

- Active tasks are those with status `WIP`, `Not yet started`, or `Open`.
- Closed tasks are skipped silently.
- `TBD` or unparseable ETA values are logged, but no email is sent.
- Co-owned tasks send to every mapped owner.
- Friday summary goes to the configured team distribution list.

On the Intel network, the current SMTP relay uses port 25.

---

## Team Handoff

If another team copies this folder:
- Keep the script and support files.
- Update [config.ini](config.ini) for their workbook, sheet, SMTP, and owner mappings.
- Run [save_password.py](save_password.py) once in their Windows account.
- Do not reuse another person's Windows Credential Manager entry.

The folder can be copied, but each team should treat `config.ini` and the saved password
as local setup, not shared configuration.

---

## Copilot Skill

The installed skill is:

`~/.copilot/skills/weekly-update-tracker/SKILL.md`

Use it by asking naturally:
- "Check overdue tasks"
- "Run the weekly tracker"
- "Show what emails would go out"
- "Send this week's reminders"

---

## Notes

- Run a dry run first before sending live emails.
- Update the Excel data regularly so reminders stop automatically when tasks are closed.
- If you need a different workbook layout, the current loader already supports a
  different sheet name and column order as long as the headers are recognizable.
