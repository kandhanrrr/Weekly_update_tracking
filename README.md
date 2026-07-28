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
- Parses Outlook inbox + search results + sent items + Teams chats exported via Microsoft Graph API.
- Detects genuine action requests addressed to **you** — based on greeting in the email/message body (`Hi Kandhan`, `@Kandhan`, `Hi All`), not just TO-field membership.
- Skips emails where you are CC-only and not explicitly named.
- Skips OOO / auto-reply emails automatically (`Automatic reply`, `Out of Office`, etc.).
- Skips Teams ARs you have already convincingly replied to (`done`, `shared`, `sent`, etc.).
- Captures self-sent follow-ups from your **Sent Items** (`can you share`, `by EOD`, etc.) and assigns them to the recipient.
- Outputs two separate sections: **Outlook ARs** and **Teams ARs**, each sorted newest-first.
- Save extracted rows as JSON (`--output-json`) and upsert directly into the **AR_Tracking** sheet in Excel.
- Fully identity-driven — zero hardcoded names; works for any user after editing `config.ini`.

---

## Files

```
Weekly_update+tracking/
├── config.ini                            # All settings — identity, email, owners, tracker
├── fivr_bgr_tracker.py                   # Main Excel tracker script
├── weekly_update_tracker.py              # Excel reminder + AR upsert script
├── save_password.py                      # Save email password to Windows Credential Manager
├── setup_scheduler.bat                   # Register the daily Task Scheduler job
├── scripts/
│   ├── parse_pending_from_graph_exports.py  # AR extractor (Outlook + Teams → pending list)
│   ├── build_sent_ar.py                     # Build AR rows from sent-mail analysis
│   ├── ar_rows_latest.json                  # Latest combined AR rows (Outlook inbox + Teams)
│   └── ar_rows_sent_followups.json          # Latest sent-mail follow-up AR rows
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
python weekly_update_tracker.py

# Send real emails
python weekly_update_tracker.py --send

# Upsert AR rows from JSON into AR_Tracking sheet
python weekly_update_tracker.py --upsert-ar scripts/ar_rows_latest.json
python weekly_update_tracker.py --upsert-ar scripts/ar_rows_sent_followups.json
```

### AR Extractor

#### Step 1 — Export data from Microsoft Graph (via Copilot / Graph API)

Fetch your Outlook inbox, any keyword search results, and Teams chat messages, saving each as a JSON file.

#### Step 2 — Run the parser

```powershell
# Inbox + sent + search + Teams (combined), save to JSON
python scripts/parse_pending_from_graph_exports.py `
  --outlook-json        "scripts/outlook_inbox_fresh.json" `
  --outlook-search-json "scripts/outlook_search_ar.json;scripts/outlook_sent_followups.json" `
  --teams-json          "scripts/t_chat1.json;scripts/t_chat2.json" `
  --lookback-days 14 --strict `
  --output-json scripts/ar_rows_latest.json

# Upsert extracted rows into AR_Tracking sheet in Excel
python weekly_update_tracker.py --upsert-ar scripts/ar_rows_latest.json
```

| Flag | Description |
|---|---|
| `--outlook-json` | Inbox export JSON (optional if `--outlook-search-json` provided) |
| `--outlook-search-json` | `;`-separated JSONs: keyword-search exports OR sent-items exports |
| `--teams-json` | `;`-separated Teams chat/channel export JSONs |
| `--lookback-days` | How many days back to scan (default: 7) |
| `--top` | Max rows per section in output (default: 15) |
| `--strict` | Only keep high-confidence actionable asks; filters OOO/auto-replies |
| `--exclude-owner` | Skip ARs assigned to this owner (repeatable) |
| `--output-json` | Save extracted AR rows as JSON for Excel upsert |
| `--send-reminders` | Preview AR reminder emails (requires `--send` to actually dispatch) |

### AR_Tracking Sheet

Extracted ARs are upserted into a dedicated **AR_Tracking** sheet in the configured Excel workbook.
The sheet is created automatically on first upsert.

| Column | Description |
|---|---|
| Date | Date the AR was detected |
| Channel | Source folder: `Outlook:Inbox`, `Outlook:Sent`, `Outlook:Search`, or `Teams` |
| Task | Action item description |
| Owner | Who must complete it (or TBD) |
| Status | Default `Open`; update manually when resolved |
| ETA | Work-week or date if parsed; `TBD` otherwise |
| Remark | Summary + sender context |
| Source | Unique dedup key — `Channel:hash` (e.g. `Outlook:Sent:2d0c5de8`) |

**Dedup is automatic** — re-running upsert on the same JSON updates existing rows (Status/ETA/Remark) without inserting duplicates.

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

### Sent Mail Follow-ups

The extractor also scans your **Sent Items** to capture follow-ups you have issued to others:

| Signal | Example | AR captured |
|---|---|---|
| You asked someone for data | `Hi Saravanan, please provide the data by this week` | Owner = Saravanan |
| You requested a calculation | `Hi Rohini, could you assist with FSL S2T multiplier` | Owner = Rohini |
| You committed to an action | `We have to file a JIRA ticket` | Owner = You (self-AR) |

Naming files ending in `sent` or `followup` (e.g. `outlook_sent_followups.json`) in `--outlook-search-json`
automatically tags those rows as `Channel = Outlook:Sent` in the Excel sheet.

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

## Copilot Skills

The following skills are installed under `.github/skills/` (callable via `/` in Copilot Chat):

| Skill | Slash command use |
|---|---|
| `ar-intake-outlook` | Collect ARs from Outlook inbox + sent items |
| `ar-intake-teams` | Collect ARs from Teams chats/channels |
| `ar-excel-reminder-sync` | Upsert AR rows into Excel, then run reminder logic |
| `weekly-update-tracker` | Full weekly tracker automation (overdue, reminders, summary) |

### AR pipeline agent

A dedicated Copilot agent file is installed at `.github/agents/ar-multisource-reminder.agent.md`.
Invoke via `@ar-multisource-reminder` in Copilot Chat for the full end-to-end pipeline:
1. Collect Teams ARs
2. Collect Outlook inbox ARs (and optionally sent-items)
3. Upsert combined rows into AR_Tracking
4. Run reminder logic

Use it by asking naturally:
- "Collect ARs from Teams and Outlook and update Excel"
- "Check overdue tasks"
- "Run the weekly AR summary"
- "Show what reminder emails would go out"

---

## Notes

- Run a dry run first before sending live emails.
- Update the Excel data regularly so reminders stop automatically when tasks are closed.
- If you need a different workbook layout, the current loader already supports a
  different sheet name and column order as long as the headers are recognizable.
