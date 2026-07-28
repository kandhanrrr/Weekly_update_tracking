# Weekly Update Tracker

Two-part automation for any team:

1. **Excel Tracker** — reads a workbook, detects headers, and sends ETA-based reminder/follow-up emails.
2. **AR Extractor** — scans Outlook and Microsoft Teams via Graph API to produce a pending Action Request list, with correct ownership detection and smart completion filtering.

---

## What It Does

### Excel Tracker
- Sends a reminder when ETA is within 2 days.
- Sends a daily overdue follow-up when ETA has passed and the task is still open.
- Sends a weekly summary to the team on Fridays.
- Skips closed tasks automatically.
- Sheet name can be configured or auto-detected; column order can vary.

### AR Extractor (`scripts/parse_pending_from_graph_exports.py`)
- Parses Outlook inbox + search results + Teams chats exported via Microsoft Graph API.
- Detects genuine action requests addressed to **you** — based on greeting in the email/message body (`Hi Kandhan`, `@Kandhan`, `Hi All`), not just TO-field membership.
- Skips emails where you are CC-only and not explicitly named.
- Skips Teams ARs you have already convincingly replied to (`done`, `shared`, `sent`, etc.).
- Captures self-sent follow-ups (`can you share`, `by EOD`, etc.) and assigns them to the recipient.
- Outputs two separate sections: **Outlook ARs** and **Teams ARs**, each sorted newest-first.
- Fully identity-driven — zero hardcoded names; works for any user after editing `config.ini`.

---

## Files

```
Weekly_update+tracking/
├── config.ini                            # All settings — identity, email, owners, tracker
├── fivr_bgr_tracker.py                   # Main Excel tracker script
├── save_password.py                      # Save email password to Windows Credential Manager
├── setup_scheduler.bat                   # Register the daily Task Scheduler job
├── scripts/
│   └── parse_pending_from_graph_exports.py  # AR extractor (Outlook + Teams → pending list)
└── README.md                             # This file
```

---

## Setup

### 1. Fill in Your Identity (required for AR extractor)

Open [config.ini](config.ini) and fill in the `[IDENTITY]` section:

```ini
[IDENTITY]
first_name   = Kandhan         # Your first name as it appears in greetings
last_name    = Rajakumar       # Your last name
display_name = Rajakumar, Kandhan   # As shown in Teams/Outlook sender list
```

The AR extractor builds all name-matching patterns dynamically from these three values.  
**If you give this tool to someone else, only these three lines need to change.**

### 2. Fill in Email Settings

```ini
[EMAIL]
smtp_host    = smtp.intel.com
smtp_port    = 25
smtp_user    = kandhan.rajakumar@intel.com   # Your email address
team_email   = your-team-dl@intel.com
```

### 3. (Optional) Excel Tracker Settings

```ini
[TRACKER]
excel_file   = PTH_progress_tracking.xlsx
sheet_name   = auto     # or exact tab name
```

### 4. Save Password

```bash
python save_password.py
```

Stores your email password in Windows Credential Manager.

### 5. (Optional) Schedule Daily Run

Double-click `setup_scheduler.bat` to register the weekday schedule.

---

## Usage

### Excel Tracker

```bash
# Dry run — no email sent
python fivr_bgr_tracker.py

# Send real emails
python fivr_bgr_tracker.py --send

# Force weekly summary on any day
python fivr_bgr_tracker.py --send --weekly-summary
```

### AR Extractor

#### Step 1 — Export data from Microsoft Graph (via Copilot / Graph API)

Fetch your Outlook inbox, any keyword search results, and Teams chat messages, saving each as a JSON file.

#### Step 2 — Run the parser

```powershell
python scripts/parse_pending_from_graph_exports.py `
  --outlook-json        "<path to inbox export>/content.json" `
  --outlook-search-json "<path to search1>/content.json;<path to search2>/content.json" `
  --teams-json          "scripts\t_chat1.json;scripts\t_chat2.json;scripts\t_grp.json" `
  --lookback-days 14 `
  --top 50 `
  --strict
```

| Flag | Description |
|---|---|
| `--outlook-json` | Inbox export JSON (required) |
| `--outlook-search-json` | `;`-separated list of keyword-search result JSONs; treated as pre-filtered ARs |
| `--teams-json` | `;`-separated list of Teams chat/channel export JSONs |
| `--lookback-days` | How many days back to scan (default: 7) |
| `--top` | Max rows per section in output (default: 15) |
| `--strict` | Only keep high-confidence actionable asks |
| `--exclude-owner` | Skip ARs assigned to this owner (repeatable) |

---

## AR Extractor — How Ownership Is Determined

The extractor never assigns you as owner just because your name is in the `To:` field.
Ownership is determined strictly from the **message body**, after stripping all embedded
email header lines (`To:`, `From:`, `Cc:`, `Sent:`) to avoid false matches.

| Body contains | Owner assigned |
|---|---|
| `Hi Kandhan` / `@Kandhan` / `Rajakumar, Kandhan` (as greeting) | You |
| `Hi All` / `Hello All` / `Dear All` | You (group ask) |
| `Hi PK` / `Hi Thomas` / `Dear Maji` | That person |
| No recognizable greeting | TBD |

### Teams-specific rules

| Condition | Behaviour |
|---|---|
| Incoming message from others | AR added with you as owner |
| You replied with `done` / `shared` / `sent` / `attached` | AR dropped (already resolved) |
| You replied with `noted` / `will check` / `ok` | AR kept (still pending) |
| Message is from you (`Rajakumar, Kandhan`) with follow-up words | AR captured; recipient is the owner |

### Outlook-specific rules

| Condition | Behaviour |
|---|---|
| You are only CC'd, not named in body | Email skipped entirely |
| You are in `To:` but body greets someone else | Owner = that person (not you) |
| Email retrieved via keyword search | Treated as a pre-confirmed AR; greeting still checked for ownership |

---

## Output Format

```
# Source scan summary     ← counts scanned vs extracted

# Pending tasks by owner  ← summary counter per responsible person

# Outlook Action Items (N)
| Date | From | Task / Subject | Short Summary | Responsible Owner | ETA |

# Teams Action Items (N)
| Date | From | Task / Subject | Short Summary | Responsible Owner | ETA |

# Reminder outcome summary
```

- **Short Summary**: professional 2-sentence summary; HTML stripped, greetings/sign-offs/forward headers removed, capped at 240 chars.
- **Responsible Owner**: who must act — you, a named colleague, or TBD.
- **ETA**: extracted from body (`WW30`, `Q3/Q4`, `by Jul 30`); defaults to TBD.

---

## Giving This Tool to Someone Else

Only three things need to change in `config.ini`:

```ini
[IDENTITY]
first_name   = John
last_name    = Smith
display_name = Smith, John

[EMAIL]
smtp_user    = john.smith@yourcompany.com
```

Everything else — all regexes, all ownership logic, all self-detection — is built automatically
from these values at runtime. No code changes required.

---

## Excel Requirements

The workbook does not need a fixed filename or a fixed sheet name. The tracker
auto-detects the best header row and maps columns by logical name.

### Fields

| Logical Field | Required | Notes |
|---|---|---|
| Task | ✓ | The action item description |
| Status | ✓ | See status values below |
| ETA | ✓ | Work week or date; `TBD` is valid |
| Owner | Recommended | Used to route reminder emails |
| Remark | Optional | Context or progress note |
| Priority | Optional | Not used for send logic |
| Module | Optional | Grouping label |

### Supported Header Aliases

The tracker recognises any of these column names — exact match is not required:

| Logical Field | Accepted Column Names |
|---|---|
| Task | `Task`, `Item`, `Title`, `Description`, `Work Item` |
| Status | `Status`, `State`, `Progress` |
| ETA | `ETA`, `Due`, `Due Date`, `Target`, `Target Date`, `Deadline` |
| Owner | `Owner`, `Assignee`, `Responsible`, `Person`, `Lead` |
| Remark | `Remark`, `Remarks`, `Notes`, `Note`, `Comment`, `Update` |

### Status Values and Auto-Normalization

The tracker automatically normalises common informal variants to canonical values —
no need to fix every cell manually:

| Canonical Value | Auto-accepted variants | Behaviour |
|---|---|---|
| `WIP` | `In Progress`, `In-Progress`, `Active`, `Ongoing`, `Started`, `In Review` | Active — reminder/overdue emails sent |
| `Not yet started` | `Not Started`, `Not_yet_started`, `Todo`, `To Do`, `Backlog`, `Planned` | Active — reminder/overdue emails sent |
| `Open` | `Pending`, `New`, `Reopened`, `Re-opened` | Active — reminder/overdue emails sent |
| `Done` | `Completed`, `Complete`, `Finished`, `Done - verified`, `Closed`, `Resolved` | Closed — no emails |
| `Dropped` | `Cancel`, `Cancelled`, `Canceled`, `Not required`, `Deferred` | Closed — no emails |
| `NA` | `N/A`, `Not Applicable`, `No Action`, `Not needed` | Closed — no emails |
| `Review Done/Closed` | `Review Done`, `Review Complete`, `Review Closed` | Closed — no emails |

### ETA Formats

| Format | Example | Notes |
|---|---|---|
| Work week | `WW30`, `WW30p3` | `p3` = Wednesday of that week |
| TBD | `TBD` | No reminder sent; logged |
| Date | `30/07/2026` | Parsed if standard format |

---

## Behavior

- Active tasks are those with status `WIP`, `Not yet started`, or `Open`.
- Closed tasks are skipped silently.
- `TBD` or unparseable ETA values are logged, but no email is sent.
- Co-owned tasks send to every mapped owner.
- Friday summary goes to the configured team distribution list.

### ETA and Status Decision Rules (Important)

Email action is based on **both** Status and ETA, not ETA alone.

| Status value | ETA value | Action |
|---|---|---|
| `WIP` / `Not yet started` / `Open` | Valid ETA and within 2 days | Reminder email is sent |
| `WIP` / `Not yet started` / `Open` | Valid ETA and already passed | Overdue follow-up email is sent |
| `WIP` / `Not yet started` / `Open` | `TBD` or invalid ETA | No email; logged as TBD/unparseable |
| `Done` / `Dropped` / `NA` / `Review Done/Closed` | Any ETA (even overdue) | No reminder/follow-up email |

Notes:
- If ETA is filled but Status is a closed state, reminder/follow-up emails stop.
- Status matching uses auto-normalisation — `In Progress`, `Not Started`, `Completed` etc. are all accepted (see Excel Requirements table above).

### Example

| Task | Status | ETA | Result |
|---|---|---|---|
| FIVR Critical setup | `WIP` | `WW30` | Reminder/overdue logic applies based on date |
| FIVR Critical setup | `Done` | `WW30` | No reminder/follow-up email (closed task) |
| Bring-up checklist | `Not yet started` | `TBD` | No reminder/follow-up email; logged as TBD |

### Sample Tracker Rows (Copy/Paste Starter)

Use this as a starter in Excel to avoid status/ETA ambiguity.

| Task | Owner | Status | ETA | Remark | What the tracker does |
|---|---|---|---|---|---|
| PMAX voltage check | Rishi | WIP | WW30 | Running validation | Sends reminder when ETA is within 2 days; sends overdue follow-up after ETA passes |
| DTS corner test | Shivagiri | Not yet started | WW31p3 | Waiting for setup | Sends reminder/overdue based on ETA date |
| BGR config cleanup | Partap | Open | WW29 | In progress | Treated as active; reminder/overdue applies |
| Session review closure | Anju | Done | WW28 | Completed and reviewed | No reminder/follow-up email even if ETA is old |
| Legacy item dropped | Kandhan | Dropped | WW26 | Not required anymore | No reminder/follow-up email |
| Spec clarification | Boomika | WIP | TBD | ETA not fixed yet | No reminder/follow-up email; logged as TBD/unparseable ETA |

Tip:
- If work is completed, always change Status to a closed value (`Done`, `Dropped`, `NA`, `Review Done/Closed`).
- Keeping Status as active (`WIP`, `Not yet started`, `Open`) with any valid old ETA will continue overdue follow-ups.

### Status Entry Reference

The tracker now **auto-normalises** all common variants. You no longer need exact values
in Excel — any of the accepted variants in the table above will work.

Recommendation: Use a data-validation dropdown for new entries to keep data clean.
For existing sheets with mixed casing or informal labels, the tracker handles them automatically.

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
