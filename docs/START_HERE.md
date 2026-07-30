# Start Here

Use this file first. Do not choose a README at random.

## Update intent (mandatory)

- If user says `update`, treat it as `complete update` by default.
- Complete update means run the full workflow for the selected task and produce all README-defined outputs/deliverables.
- Do not run partial update unless user explicitly asks for a partial step.

## Pick one task

| If your goal is... | Follow this file only |
|---|---|
| Weekly reminders/tracking (user-provided Excel required) | [weekly-tracking/WEEKLY_TRACKING_README.md](weekly-tracking/WEEKLY_TRACKING_README.md) |
| AR tracking from Teams + Outlook Inbox + Sent (last 1WW) | [ar-tracking/AR_TRACKING_README.md](ar-tracking/AR_TRACKING_README.md) |

## Quick rule

- Weekly Tracking task:
  - If a user wants to maintain their own tracking in Excel, they can provide their workbook. The skill or agent will follow up on WIP, Open, and Not yet started items, send reminder emails starting 2 days before ETA, and continue follow-ups until the task status is marked Closed or Completed.
  - Equivalent status wording also works (for example: `In Progress`, `Not Started`, `Pending`).
  - Weekly tracking does not auto-create a workbook.
- AR Tracking task:
  - System must create AR workbook from Teams + Outlook data for last 1WW.

## Output files by task

- Weekly task output path: `[TRACKER] excel_file` configured in `config.ini` (updated in place; no auto-generated artifact file)
- Weekly HTML tab report: artifacts/weekly/weekly_tracker_report.html
- AR task auto workbook: artifacts/ar/AR_Tracking_Auto.xlsx
- AR workflow also refreshes shared HTML tab report: artifacts/weekly/weekly_tracker_report.html

## Grouped file map

Weekly tracking files:
- docs/weekly-tracking/WEEKLY_TRACKING_README.md
- .github/agents/weekly-update-tracker.agent.md
- .github/agents/weekly-update-tracker.agent.yaml
- .github/skills/weekly-tracker-reminder/SKILL.md
- scripts/weekly/run_weekly_tracker.py

AR tracking files:
- docs/ar-tracking/AR_TRACKING_README.md
- .github/agents/ar-multisource-reminder.agent.md
- .github/agents/ar-multisource-reminder.agent.yaml
- .github/skills/ar-intake-outlook/SKILL.md
- .github/skills/ar-excel-reminder-sync/SKILL.md
- scripts/ar/extract_ar_rows.py
- scripts/ar/create_ar_workbook.py
- artifacts/ar/

## If you are unsure

Start with AR tracking first using [ar-tracking/AR_TRACKING_README.md](ar-tracking/AR_TRACKING_README.md), then run weekly tracking.
