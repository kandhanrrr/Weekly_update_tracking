from __future__ import annotations

import configparser
import html
import json
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from statistics import mean

import openpyxl  # type: ignore[import-untyped]

BASE_DIR = Path(__file__).resolve().parents[2]
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from weekly_update_tracker import STATUS_ALIASES, _build_header_map, _select_sheet, ww_to_date

CONFIG_FILE = BASE_DIR / "config.ini"
OUTPUT_HTML = BASE_DIR / "artifacts" / "weekly" / "weekly_tracker_report.html"


@dataclass
class TaskRow:
    module: str
    owner: str
    status_raw: str
    status_norm: str
    eta: str
    task: str
    remark: str
    days_delta: int | None


def _load_tracker_path() -> Path:
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_FILE)
    excel_rel = cfg.get("TRACKER", "excel_file", fallback="AR_Intake_Tracker.xlsx")
    return (BASE_DIR / excel_rel).resolve()


def _cell(row: tuple, idx: int | None, default: str = "") -> str:
    if idx is None or idx >= len(row):
        return default
    value = row[idx]
    if value is None:
        return default
    return str(value).strip()


def _load_tasks(sheet) -> list[TaskRow]:
    header_map, header_row = _build_header_map(sheet)
    task_idx = header_map.get("task")
    owner_idx = header_map.get("owner")
    status_idx = header_map.get("status")
    eta_idx = header_map.get("eta")
    remark_idx = header_map.get("remark")
    module_idx = header_map.get("module")

    if task_idx is None or status_idx is None or eta_idx is None:
        return []

    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    out: list[TaskRow] = []

    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        task = _cell(row, task_idx)
        if not task:
            continue

        status_raw = _cell(row, status_idx)
        status_norm = STATUS_ALIASES.get(status_raw.lower(), status_raw.lower())
        eta = _cell(row, eta_idx, "TBD")
        eta_date = ww_to_date(eta)
        days_delta = None
        if eta_date is not None:
            eta_date = eta_date.replace(hour=0, minute=0, second=0, microsecond=0)
            days_delta = (eta_date - today).days

        out.append(
            TaskRow(
                module=_cell(row, module_idx, "Unassigned") or "Unassigned",
                owner=_cell(row, owner_idx, "TBD") or "TBD",
                status_raw=status_raw or "Open",
                status_norm=status_norm or "open",
                eta=eta,
                task=task,
                remark=_cell(row, remark_idx),
                days_delta=days_delta,
            )
        )

    return out


def _bucket(status_norm: str) -> str:
    if status_norm == "done":
        return "completed"
    if status_norm in {"open", "wip", "not yet started", "dropped"}:
        return status_norm
    return "other"


def _module_summary(tasks: list[TaskRow]) -> tuple[list[dict], dict[str, dict[str, float | int | str]]]:
    mod: dict[str, dict[str, int]] = {}
    stats_map: dict[str, dict[str, float | int | str]] = {}

    for t in tasks:
        module = t.module
        if module not in mod:
            mod[module] = {
                "completed": 0,
                "open": 0,
                "wip": 0,
                "not yet started": 0,
                "dropped": 0,
                "other": 0,
                "total": 0,
            }
        b = _bucket(t.status_norm)
        mod[module][b] += 1
        mod[module]["total"] += 1

    rows = []
    for module in sorted(mod):
        c = mod[module]
        rows.append(
            {
                "module": module,
                "completed": c["completed"],
                "open": c["open"],
                "wip": c["wip"],
                "not_started": c["not yet started"],
                "dropped": c["dropped"],
                "total": c["total"],
            }
        )

        deltas = [
            t.days_delta
            for t in tasks
            if t.module == module and t.days_delta is not None
        ]
        if deltas:
            s = sorted(deltas)
            q1 = s[len(s) // 4]
            med = s[len(s) // 2]
            q3 = s[(len(s) * 3) // 4]
            iqr = q3 - q1
            stats_map[module] = {
                "label": module,
                "type": "ETA delta (days)",
                "max": max(s),
                "upper_fence": q3 + (1.5 * iqr),
                "q3": q3,
                "median": med,
                "mean": round(mean(s), 4),
                "q1": q1,
                "lower_fence": q1 - (1.5 * iqr),
            }
        else:
            stats_map[module] = {
                "label": module,
                "type": "ETA delta (days)",
                "max": "N/A",
                "upper_fence": "N/A",
                "q3": "N/A",
                "median": "N/A",
                "mean": "N/A",
                "q1": "N/A",
                "lower_fence": "N/A",
            }

    return rows, stats_map


def _overview(tasks: list[TaskRow], tracker_file: Path, sheet_name: str) -> dict[str, str | int]:
    active = [t for t in tasks if t.status_norm in {"wip", "open", "not yet started"}]
    due_soon = [t for t in active if t.days_delta is not None and 0 <= t.days_delta <= 2]
    overdue = [t for t in active if t.days_delta is not None and t.days_delta < 0]
    tbd = [t for t in active if t.days_delta is None]
    return {
        "run_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "run_mode": "VIEW ONLY (HTML report)",
        "workbook": str(tracker_file),
        "sheet": sheet_name,
        "active_tasks": len(active),
        "due_soon": len(due_soon),
        "overdue": len(overdue),
        "tbd_eta": len(tbd),
    }


def _overdue_owners(tasks: list[TaskRow]) -> list[dict[str, str | int]]:
    owners: dict[str, dict[str, str | int]] = {}
    for t in tasks:
        if t.status_norm not in {"wip", "open", "not yet started"}:
            continue
        if t.days_delta is None or t.days_delta >= 0:
            continue
        if t.owner not in owners:
            owners[t.owner] = {"owner": t.owner, "count": 0, "task": t.task}
        owners[t.owner]["count"] = int(owners[t.owner]["count"]) + 1
    return sorted(owners.values(), key=lambda x: int(x["count"]), reverse=True)


def _dispatch_summary() -> dict[str, str]:
    is_friday = datetime.today().weekday() == 4
    return {
        "dispatch": "No email dispatch from HTML view",
        "weekly_summary": "Friday" if is_friday else "Skipped (not Friday)",
        "force_hint": "Use: python weekly_update_tracker.py --send --weekly-summary",
    }


def _render_html(overview: dict, module_rows: list[dict], stats_map: dict, overdue_rows: list[dict], dispatch: dict) -> str:
    mod_rows_html = []
    for r in module_rows:
        stats = stats_map.get(r["module"], {})
        mod_rows_html.append(
            """
            <tr class=\"hover-row\" data-hover='{hover}'>
              <td>{module}</td><td>{completed}</td><td>{open}</td><td>{wip}</td><td>{not_started}</td><td>{dropped}</td><td>{total}</td>
            </tr>
            """.format(
                hover=html.escape(json.dumps(stats)),
                module=html.escape(str(r["module"])),
                completed=r["completed"],
                open=r["open"],
                wip=r["wip"],
                not_started=r["not_started"],
                dropped=r["dropped"],
                total=r["total"],
            )
        )

    overdue_html = "".join(
        "<tr><td>{}</td><td>{}</td><td>{}</td></tr>".format(
            html.escape(str(r["owner"])),
            r["count"],
            html.escape(str(r["task"])),
        )
        for r in overdue_rows
    )
    if not overdue_html:
        overdue_html = "<tr><td colspan='3'>No overdue owners</td></tr>"

    return f"""<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>Weekly Tracker Report</title>
  <style>
    :root {{ --ink:#0f172a; --muted:#475569; --line:#dbe2ea; --brand:#0b6bcb; --bg:#f8fafc; }}
    body {{ font-family: Segoe UI, Arial, sans-serif; margin:0; color:var(--ink); background:var(--bg); }}
    .wrap {{ max-width:1200px; margin:20px auto; padding:0 16px 24px; }}
    h1 {{ margin:0 0 8px; font-size:24px; }}
    .sub {{ color:var(--muted); margin-bottom:14px; }}
    .tabs {{ display:flex; gap:8px; flex-wrap:wrap; margin-bottom:14px; }}
    .tab-btn {{ border:1px solid var(--line); background:#fff; padding:8px 12px; border-radius:8px; cursor:pointer; }}
    .tab-btn.active {{ background:var(--brand); color:#fff; border-color:var(--brand); }}
    .tab {{ display:none; background:#fff; border:1px solid var(--line); border-radius:12px; padding:14px; }}
    .tab.active {{ display:block; }}
    .cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:10px; margin-bottom:12px; }}
    .card {{ background:#fff; border:1px solid var(--line); border-radius:10px; padding:10px; }}
    .k {{ color:var(--muted); font-size:12px; text-transform:uppercase; }}
    .v {{ font-size:26px; font-weight:700; margin-top:4px; }}
    table {{ border-collapse:collapse; width:100%; }}
    th, td {{ border:1px solid var(--line); padding:8px; text-align:left; font-size:14px; }}
    th {{ background:#f1f5f9; }}
    .hint {{ font-size:12px; color:var(--muted); margin-top:8px; }}
    #hoverBox {{ margin-top:10px; border:1px dashed var(--line); border-radius:8px; padding:10px; background:#f8fafc; }}
    #hoverBox code {{ font-family: Consolas, monospace; }}
  </style>
</head>
<body>
  <div class=\"wrap\">
    <h1>Weekly Tracker Report</h1>
    <div class=\"sub\">Tabbed view with hover details</div>

    <div class=\"tabs\">
      <button class=\"tab-btn active\" data-tab=\"overview\">Overview</button>
      <button class=\"tab-btn\" data-tab=\"module\">Module-wise Status Summary</button>
      <button class=\"tab-btn\" data-tab=\"overdue\">Overdue Owners</button>
      <button class=\"tab-btn\" data-tab=\"dispatch\">Dispatch Summary</button>
    </div>

    <section id=\"overview\" class=\"tab active\">
      <div class=\"cards\">
        <div class=\"card\"><div class=\"k\">Run Time</div><div class=\"v\">{overview['run_time']}</div></div>
        <div class=\"card\"><div class=\"k\">Active Tasks</div><div class=\"v\">{overview['active_tasks']}</div></div>
        <div class=\"card\"><div class=\"k\">Due Soon</div><div class=\"v\">{overview['due_soon']}</div></div>
        <div class=\"card\"><div class=\"k\">Overdue</div><div class=\"v\">{overview['overdue']}</div></div>
        <div class=\"card\"><div class=\"k\">TBD ETA</div><div class=\"v\">{overview['tbd_eta']}</div></div>
      </div>
      <table>
        <tr><th>Metric</th><th>Value</th></tr>
        <tr><td>Run mode</td><td>{html.escape(str(overview['run_mode']))}</td></tr>
        <tr><td>Workbook</td><td>{html.escape(str(overview['workbook']))}</td></tr>
        <tr><td>Sheet</td><td>{html.escape(str(overview['sheet']))}</td></tr>
      </table>
    </section>

    <section id=\"module\" class=\"tab\">
      <table id=\"moduleTable\">
        <tr><th>Module</th><th>Completed</th><th>Open</th><th>WIP</th><th>Not yet started</th><th>Dropped</th><th>Total</th></tr>
        {''.join(mod_rows_html)}
      </table>
      <div id=\"hoverBox\"><strong>Hover details:</strong> move cursor over a module row.</div>
      <div class=\"hint\">Hover details include label, type, max, upper fence, Q3, median, mean, Q1, and lower fence.</div>
    </section>

    <section id=\"overdue\" class=\"tab\">
      <table>
        <tr><th>Owner</th><th>Overdue Count</th><th>Example Task</th></tr>
        {overdue_html}
      </table>
    </section>

    <section id=\"dispatch\" class=\"tab\">
      <table>
        <tr><th>Item</th><th>Status</th></tr>
        <tr><td>Dispatch</td><td>{html.escape(dispatch['dispatch'])}</td></tr>
        <tr><td>Weekly summary</td><td>{html.escape(dispatch['weekly_summary'])}</td></tr>
        <tr><td>Force option</td><td><code>{html.escape(dispatch['force_hint'])}</code></td></tr>
      </table>
    </section>
  </div>

  <script>
    const buttons = document.querySelectorAll('.tab-btn');
    const tabs = document.querySelectorAll('.tab');
    buttons.forEach(btn => btn.addEventListener('click', () => {{
      buttons.forEach(b => b.classList.remove('active'));
      tabs.forEach(t => t.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.add('active');
    }}));

    const hoverBox = document.getElementById('hoverBox');
    document.querySelectorAll('.hover-row').forEach(row => {{
      row.addEventListener('mouseenter', () => {{
        const d = JSON.parse(row.dataset.hover || '{{}}');
        hoverBox.innerHTML = `
          <strong>Hover details</strong><br/>
          Label: <code>${{d.label ?? 'N/A'}}</code><br/>
          Type: <code>${{d.type ?? 'N/A'}}</code><br/>
          Max: <code>${{d.max ?? 'N/A'}}</code><br/>
          Upper fence: <code>${{d.upper_fence ?? 'N/A'}}</code><br/>
          Q3: <code>${{d.q3 ?? 'N/A'}}</code><br/>
          Median: <code>${{d.median ?? 'N/A'}}</code><br/>
          Mean: <code>${{d.mean ?? 'N/A'}}</code><br/>
          Q1: <code>${{d.q1 ?? 'N/A'}}</code><br/>
          Lower fence: <code>${{d.lower_fence ?? 'N/A'}}</code>
        `;
      }});
    }});
  </script>
</body>
</html>
"""


def main() -> int:
    tracker_file = _load_tracker_path()
    if not tracker_file.exists():
        print(f"[ERROR] Workbook not found: {tracker_file}")
        return 1

    wb = openpyxl.load_workbook(tracker_file)
    ws = _select_sheet(wb)
    tasks = _load_tasks(ws)

    overview = _overview(tasks, tracker_file, ws.title)
    module_rows, stats_map = _module_summary(tasks)
    overdue_rows = _overdue_owners(tasks)
    dispatch = _dispatch_summary()

    OUTPUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_HTML.write_text(
        _render_html(overview, module_rows, stats_map, overdue_rows, dispatch),
        encoding="utf-8",
    )
    print(f"Generated weekly HTML report: {OUTPUT_HTML}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
