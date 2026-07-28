# AR Intake and Reminder Workflow (Teams + Outlook)

This document describes the dedicated AR pipeline that collects action items
from Teams and Outlook, syncs them into the tracker Excel, and triggers
reminders.

---

## Scope

- Source systems: Microsoft Teams and Outlook
- Storage: Excel tracker used by the weekly update flow
- Reminder engine: weekly_update_tracker.py

---

## Separate Workbook For This Task

Created workbook:
- `AR_Intake_Tracker.xlsx`

Sheet name:
- `AR_Intake`

Included columns:
- `Task`, `Owner`, `Status`, `ETA`, `Remark`, `Source`, `Last_Updated`

Included starter content:
- Sample AR rows from Teams and Outlook context
- Status dropdown values: `WIP`, `Not yet started`, `Open`, `Done`, `Dropped`, `NA`, `Review Done/Closed`
- ETA dropdown starter values including `TBD` and WW examples

To use this workbook for reminders, set in `config.ini`:

```ini
[TRACKER]
excel_file = AR_Intake_Tracker.xlsx
sheet_name = AR_Intake
```

---

## Default Lookback Window

| Source | Default lookback | What is scanned |
|---|---:|---|
| Teams | Last 7 days | Channel and chat messages containing AR-like text |
| Outlook | Last 7 days | Email threads/messages containing AR-like text |

Notes:
- The 7-day range is based on local system date/time where the job runs.
- If needed, this can be extended in future versions.

---

## AR Extraction Rules

| Item | Rule |
|---|---|
| Detection keywords | AR, Action item, Owner, ETA, Due, Follow-up |
| Missing Status | Set to Open |
| Missing ETA | Set to TBD |
| Missing Owner | Set to TBD |
| Source traceability | Save source id in Source/Remark |

Expected source tags:
- Teams source format: Teams:<chat-or-channel>:<message-id>
- Outlook source format: Outlook:<message-id>

---

## Normalized Row Schema

| Field | Required | Rule |
|---|---|---|
| Task | Yes | Clear action statement |
| Owner | Yes | Parsed owner or TBD |
| Status | Yes | Default Open |
| ETA | Yes | WWxx/WWxxpx if parseable, else TBD |
| Remark | No | Context snippet and source notes |
| Source | Yes | Unique source key for dedupe |

---

## De-duplication and Upsert

| Case | Behavior |
|---|---|
| Existing Source key found | Update existing row |
| New Source key | Insert new row |
| Same message appears again | Skip duplicate insert |
| Unclear owner/ETA | Keep TBD (do not guess) |

---

## Reminder Rules After Sync

Once AR rows are in Excel, reminder behavior is identical to the main tracker.

| Condition | Action |
|---|---|
| Status WIP / Not yet started / Open and ETA within 2 days | Reminder email |
| Status WIP / Not yet started / Open and ETA overdue | Overdue follow-up email |
| Status Done / Dropped / NA / Review Done/Closed | No reminder/follow-up |
| ETA TBD or unparseable | No reminder/follow-up; logged |

---

## Run Modes

| Mode | Command | Result |
|---|---|---|
| Dry-run | python weekly_update_tracker.py | No emails sent; actions printed |
| Live send | python weekly_update_tracker.py --send | Sends reminder/follow-up emails |
| Force summary | python weekly_update_tracker.py --send --weekly-summary | Sends weekly summary even on non-Friday |

---

## Example Outcome Format

Use this format when reporting AR intake and reminder run results.

### 1. Intake and Sync Summary

| Metric | Example Value |
|---|---:|
| Teams messages scanned (last 7 days) | 120 |
| Outlook emails scanned (last 7 days) | 85 |
| AR candidates extracted | 34 |
| Inserted rows | 12 |
| Updated rows | 18 |
| Skipped duplicates | 4 |

### 2. Reminder Outcome Summary

| Metric | Example Value |
|---|---:|
| Active AR tasks processed | 30 |
| Due soon reminders (<=2 days) | 6 |
| Overdue follow-ups | 5 |
| TBD/unparseable ETA (no email) | 9 |

### 3. Owner-wise Overdue Summary

| Owner | Overdue Count | Example Tasks |
|---|---:|---|
| Rishi | 2 | PMAX dashboard mismatch; session follow-up |
| Shivagiri | 1 | Thermal signoff closure |
| Partap | 2 | BGR config AR; review closure |

### 4. Final Run Status

| Field | Example |
|---|---|
| Run mode | Dry-run |
| Workbook used | AR_Intake_Tracker.xlsx |
| Sheet used | AR_Intake |
| Live emails sent | No |
| Next recommended action | Resolve TBD owners and ETAs, then run --send |

---

## Related Skill and Agent Files

| Type | File |
|---|---|
| Teams intake skill | .copilot/skills/ar-intake-teams/SKILL.md |
| Outlook intake skill | .copilot/skills/ar-intake-outlook/SKILL.md |
| Excel/reminder sync skill | .copilot/skills/ar-excel-reminder-sync/SKILL.md |
| AR orchestration agent | .copilot/ar-multisource-reminder.agent.yaml |

---

## Operational Recommendations

| Topic | Recommendation |
|---|---|
| Run frequency | Daily on weekdays |
| Safety | Always run dry-run before live send |
| Data hygiene | Keep status values consistent and update completed items promptly |
| Ownership gaps | Resolve TBD owners quickly to avoid reminder misses |

---

## Troubleshooting Quick Checks

| Symptom | Check |
|---|---|
| No AR rows added | Verify source lookback range and AR keyword presence |
| Unexpected duplicates | Verify Source key is being captured and persisted |
| No reminder sent | Check Status is active and ETA is parseable |
| Too many overdue alerts | Confirm closed items were moved to closed statuses |
