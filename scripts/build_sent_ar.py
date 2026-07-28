import json
import hashlib
from pathlib import Path


def key(source, date, task):
    raw = f"{source}|{date}|{task[:80]}"
    return "Outlook:Sent:" + hashlib.sha1(raw.encode()).hexdigest()[:12]


rows = [
    {
        "source_key": key("Outlook", "2026-07-28", "Provide XCCP A0 Efficiency FER and FSL data"),
        "source": "Outlook",
        "folder": "Outlook:Sent",
        "date": "2026-07-28",
        "task": "Provide GNRD XCCP A0 Efficiency FER and FSL data for review",
        "owner": "Ranganathan, Saravanann",
        "status": "Open",
        "eta": "WW30",
        "sender": "kandhan.rajakumar@intel.com",
        "summary": "Kandhan follow-up: please provide the data by this week | Thread: GNRD XCCP A0 Efficiency FER, FSL data review",
    },
    {
        "source_key": key("Outlook", "2026-07-23", "Provide updated IA FSL S2T multiplier for 10pct GB GNR-D XCC+"),
        "source": "Outlook",
        "folder": "Outlook:Sent",
        "date": "2026-07-23",
        "task": "Provide updated IA FSL S2T multiplier for 10% guard band (GNR-D XCC+ power failure improvement)",
        "owner": "Krishnan, Rohini",
        "status": "Open",
        "eta": "TBD",
        "sender": "kandhan.rajakumar@intel.com",
        "summary": "Kandhan asked Rohini: calculate IA FSL S2T multiplier when reducing GB from 20% to 10% for GNR-D XCC+",
    },
    {
        "source_key": key("Outlook", "2026-07-22", "Share noted differences GNR-D XCC+ audit vs HCC SP"),
        "source": "Outlook",
        "folder": "Outlook:Sent",
        "date": "2026-07-22",
        "task": "Share noted differences from GNR-D XCC+ Program Audit (vs GNR-D HCC and GND-SP XCC)",
        "owner": "Rivera Valverde, Pablo",
        "status": "Open",
        "eta": "TBD",
        "sender": "kandhan.rajakumar@intel.com",
        "summary": "Kandhan asked Pablo to share noted differences from GNR-D XCC+ module-level audit review",
    },
    {
        "source_key": key("Outlook", "2026-07-22", "Help with 10pct IA FSL S2T multiplier 1.1 GNRD XCC+ Saravanan"),
        "source": "Outlook",
        "folder": "Outlook:Sent",
        "date": "2026-07-22",
        "task": "Help with 10% IA FSL S2T multiplier (1.1) calculation for GNR-D XCC+ power failure improvement path",
        "owner": "Ranganathan, Saravanann",
        "status": "Open",
        "eta": "TBD",
        "sender": "kandhan.rajakumar@intel.com",
        "summary": "Kandhan asked Saravanan to help calculate 10% IA FSL S2T multiplier (value: 1.1) per Maji request",
    },
    {
        "source_key": key("Outlook", "2026-07-14", "File JIRA Prime team Bin2719 offset calculation GNR-D XCC+"),
        "source": "Outlook",
        "folder": "Outlook:Sent",
        "date": "2026-07-14",
        "task": "File JIRA ticket to Prime team for Bin2719 offset calculation error (GNR-D XCC+ CLASSHOT yield)",
        "owner": "Kandhan",
        "status": "Open",
        "eta": "TBD",
        "sender": "kandhan.rajakumar@intel.com",
        "summary": "Bin2719 failing due to Prime template offset calculation bug (not silicon). Kandhan to file JIRA with Prime team. DOEs ongoing.",
    },
]

out = Path("scripts/ar_rows_sent_followups.json")
out.write_text(json.dumps(rows, indent=2, ensure_ascii=False), encoding="utf-8")
print(f"Saved {len(rows)} rows to {out}")
for r in rows:
    print(f"  [{r['date']}] {r['owner']:<35} | {r['task'][:65]}")
