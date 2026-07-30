from __future__ import annotations

import argparse
import glob
import html
import json
import re
import sys
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from statistics import mean
from typing import Any

BASE_DIR = Path(__file__).resolve().parents[2]
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from weekly_update_tracker import STATUS_ALIASES, ww_to_date

OUTPUT_HTML = BASE_DIR / "artifacts" / "ar" / "ar_tracking_report.html"


@dataclass
class ArRow:
    source_key: str
    source: str
    folder: str
    date: str
    task: str
    owner: str
    status_raw: str
    status_norm: str
    eta: str
    sender: str
    summary: str
    days_delta: int | None


def _split_paths(value: str) -> list[str]:
    return [p.strip() for p in value.split(";") if p.strip()]


def _expand_globs(patterns: str) -> list[Path]:
    out: list[Path] = []
    for pattern in _split_paths(patterns):
        resolved = pattern if Path(pattern).is_absolute() else str(BASE_DIR / pattern)
        for match in glob.glob(resolved):
            p = Path(match)
            if p.is_file():
                out.append(p)
    return sorted(set(out))


def _load_json_rows(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        return []
    return [x for x in data if isinstance(x, dict)]


def _parse_iso_date(value: str) -> datetime | None:
    value = value.strip()
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        return None


def _eta_to_delta_days(eta: str, today: datetime) -> int | None:
    eta = eta.strip()
    if not eta:
        return None

    ww_dt = ww_to_date(eta)
    if ww_dt is not None:
        return (ww_dt.replace(hour=0, minute=0, second=0, microsecond=0) - today).days

    iso_dt = _parse_iso_date(eta)
    if iso_dt is not None:
        return (iso_dt.replace(hour=0, minute=0, second=0, microsecond=0) - today).days

    return None


def _to_rows(raw_rows: list[dict[str, Any]]) -> list[ArRow]:
    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    out: list[ArRow] = []

    for r in raw_rows:
        task = str(r.get("task", "")).strip()
        if not task:
            continue

        status_raw = str(r.get("status", "Open") or "Open").strip()
        status_norm = STATUS_ALIASES.get(status_raw.lower(), status_raw.lower())
        eta = str(r.get("eta", "TBD") or "TBD").strip() or "TBD"

        out.append(
            ArRow(
                source_key=str(r.get("source_key", "")).strip(),
                source=str(r.get("source", "Unknown") or "Unknown").strip() or "Unknown",
                folder=str(r.get("folder", "Unknown") or "Unknown").strip() or "Unknown",
                date=str(r.get("date", "")).strip(),
                task=task,
                owner=str(r.get("owner", "TBD") or "TBD").strip() or "TBD",
                status_raw=status_raw,
                status_norm=status_norm,
                eta=eta,
                sender=str(r.get("sender", "")).strip(),
                summary=str(r.get("summary", "")).strip(),
                days_delta=_eta_to_delta_days(eta, today),
            )
        )

    return out


def _dedupe(rows: list[ArRow]) -> tuple[list[ArRow], int]:
    seen: set[str] = set()
    unique: list[ArRow] = []
    duplicates = 0

    for r in rows:
        key = r.source_key or f"{r.source}|{r.date}|{r.task}|{r.owner}"
        if key in seen:
            duplicates += 1
            continue
        seen.add(key)
        unique.append(r)

    unique.sort(key=lambda x: (x.date, x.source, x.owner), reverse=True)
    return unique, duplicates


def _is_active(status: str) -> bool:
    return status in {"open", "wip", "not yet started"}


def _overview(rows: list[ArRow], lookback_days: int, output_html: Path, sources: list[Path]) -> dict[str, Any]:
    active = [r for r in rows if _is_active(r.status_norm)]
    due_soon = [r for r in active if r.days_delta is not None and 0 <= r.days_delta <= 2]
    overdue = [r for r in active if r.days_delta is not None and r.days_delta < 0]
    tbd_eta = [r for r in active if r.days_delta is None]

    return {
        "run_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "lookback": f"{lookback_days} days (1WW)",
        "source_files": len(sources),
        "active_rows": len(active),
        "due_soon": len(due_soon),
        "overdue": len(overdue),
        "tbd_eta": len(tbd_eta),
        "output": str(output_html),
    }


def _source_scan_summary(rows: list[ArRow], sources: list[Path]) -> list[dict[str, Any]]:
    by_channel = Counter(r.source for r in rows)
    by_folder = Counter(r.folder for r in rows)

    items: list[dict[str, Any]] = []
    for name, count in sorted(by_channel.items()):
        items.append({"kind": "Channel", "name": name, "rows": count})
    for name, count in sorted(by_folder.items()):
        items.append({"kind": "Folder", "name": name, "rows": count})

    if not items:
        items.append({"kind": "Info", "name": "No rows extracted", "rows": 0})

    items.append({"kind": "Files", "name": "JSON files scanned", "rows": len(sources)})
    return items


def _owner_stats(rows: list[ArRow]) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
    counts = Counter(r.owner for r in rows if _is_active(r.status_norm))
    owner_rows = [{"owner": o, "count": c} for o, c in sorted(counts.items(), key=lambda x: (-x[1], x[0]))]

    hover_stats: dict[str, dict[str, Any]] = {}
    for owner, count in counts.items():
        deltas = sorted([r.days_delta for r in rows if r.owner == owner and r.days_delta is not None])
        if deltas:
            q1 = deltas[len(deltas) // 4]
            med = deltas[len(deltas) // 2]
            q3 = deltas[(len(deltas) * 3) // 4]
            iqr = q3 - q1
            hover_stats[owner] = {
                "label": owner,
                "type": "Owner ETA delta (days)",
                "max": max(deltas),
                "upper_fence": q3 + (1.5 * iqr),
                "q3": q3,
                "median": med,
                "mean": round(mean(deltas), 4),
                "q1": q1,
                "lower_fence": q1 - (1.5 * iqr),
                "active_count": count,
            }
        else:
            hover_stats[owner] = {
                "label": owner,
                "type": "Owner ETA delta (days)",
                "max": "N/A",
                "upper_fence": "N/A",
                "q3": "N/A",
                "median": "N/A",
                "mean": "N/A",
                "q1": "N/A",
                "lower_fence": "N/A",
                "active_count": count,
            }

    return owner_rows, hover_stats


def _upsert_summary(total_loaded: int, unique_loaded: int, duplicate_count: int) -> dict[str, int]:
    return {
        "total_processed": total_loaded,
        "inserted": unique_loaded,
        "updated": 0,
        "skipped": duplicate_count,
    }


def _reminder_outcome(rows: list[ArRow]) -> dict[str, int]:
    active = [r for r in rows if _is_active(r.status_norm)]
    return {
        "due_soon": sum(1 for r in active if r.days_delta is not None and 0 <= r.days_delta <= 2),
        "overdue": sum(1 for r in active if r.days_delta is not None and r.days_delta < 0),
        "tbd_eta": sum(1 for r in active if r.days_delta is None),
        "active_total": len(active),
    }


def _overdue_owners(rows: list[ArRow]) -> list[dict[str, Any]]:
    by_owner: dict[str, dict[str, Any]] = {}
    for r in rows:
        if not _is_active(r.status_norm):
            continue
        if r.days_delta is None or r.days_delta >= 0:
            continue
        if r.owner not in by_owner:
            by_owner[r.owner] = {"owner": r.owner, "count": 0, "example": r.task}
        by_owner[r.owner]["count"] += 1
    return sorted(by_owner.values(), key=lambda x: (-int(x["count"]), str(x["owner"])))


def _compact_text(value: str, limit: int = 120) -> str:
    text = re.sub(r"\s+", " ", value.strip())
    if len(text) <= limit:
        return text
    return text[: limit - 3] + "..."


def _render_html(
    overview: dict[str, Any],
    source_scan: list[dict[str, Any]],
    owner_rows: list[dict[str, Any]],
    owner_hover: dict[str, dict[str, Any]],
    upsert: dict[str, int],
    reminder: dict[str, int],
    overdue_rows: list[dict[str, Any]],
    task_rows: list[ArRow],
) -> str:
    def _hover_payload(label: str, type_context: str, **extra: str | int | float) -> str:
        payload: dict[str, str | int | float] = {
            "label": label,
            "type": type_context,
            "max": "N/A",
            "upper_fence": "N/A",
            "q3": "N/A",
            "median": "N/A",
            "mean": "N/A",
            "q1": "N/A",
            "lower_fence": "N/A",
        }
        payload.update(extra)
        return html.escape(json.dumps(payload))

    overview_rows_html = "".join(
        "<tr class='hover-row' data-hover='{hover}'><td>{metric}</td><td>{value}</td></tr>".format(
            hover=_hover_payload(str(metric), "Overview metric", median=value, mean=value),
            metric=html.escape(str(metric)),
            value=html.escape(str(value)),
        )
        for metric, value in [
            ("Due soon", overview["due_soon"]),
            ("TBD ETA", overview["tbd_eta"]),
            ("Output HTML", overview["output"]),
        ]
    )

    src_rows_html = "".join(
        "<tr class='hover-row' data-hover='{hover}'><td>{kind}</td><td>{name}</td><td>{rows}</td></tr>".format(
            hover=_hover_payload(
                str(r["name"]),
                f"{r['kind']} scan summary",
                max=r["rows"],
                q3=r["rows"],
                median=r["rows"],
                mean=r["rows"],
                q1=r["rows"],
            ),
            kind=html.escape(str(r["kind"])),
            name=html.escape(str(r["name"])),
            rows=r["rows"],
        )
        for r in source_scan
    )

    owner_table_rows = []
    for r in owner_rows:
        hover_payload = owner_hover.get(str(r["owner"]), {})
        owner_table_rows.append(
            """
            <tr class=\"hover-row\" data-hover='{hover}'>
              <td>{owner}</td><td>{count}</td>
            </tr>
            """.format(
                hover=html.escape(json.dumps(hover_payload)),
                owner=html.escape(str(r["owner"])),
                count=r["count"],
            )
        )

    if not owner_table_rows:
        owner_table_rows.append("<tr><td colspan='2'>No active owner rows</td></tr>")

    upsert_rows_html = "".join(
        "<tr class='hover-row' data-hover='{hover}'><td>{k}</td><td>{v}</td></tr>".format(
            hover=_hover_payload(str(k), "Upsert metric", median=v, mean=v),
            k=html.escape(str(k)),
            v=html.escape(str(v)),
        )
        for k, v in [
            ("Total processed", upsert["total_processed"]),
            ("Inserted", upsert["inserted"]),
            ("Updated", upsert["updated"]),
            ("Skipped (duplicates)", upsert["skipped"]),
        ]
    )

    reminder_rows_html = "".join(
        "<tr class='hover-row' data-hover='{hover}'><td>{k}</td><td>{v}</td></tr>".format(
            hover=_hover_payload(str(k), "Reminder metric", median=v, mean=v),
            k=html.escape(str(k)),
            v=html.escape(str(v)),
        )
        for k, v in [
            ("Due soon", reminder["due_soon"]),
            ("Overdue", reminder["overdue"]),
            ("TBD ETA", reminder["tbd_eta"]),
            ("Active total", reminder["active_total"]),
        ]
    )

    overdue_html = "".join(
        "<tr class='hover-row' data-hover='{hover}'><td>{owner}</td><td>{count}</td><td>{task}</td></tr>".format(
            hover=_hover_payload(
                str(r["owner"]),
                "Overdue owner summary",
                max=r["count"],
                q3=r["count"],
                median=r["count"],
                mean=r["count"],
                q1=r["count"],
                task=str(r["example"]),
            ),
            owner=html.escape(str(r["owner"])),
            count=r["count"],
            task=html.escape(_compact_text(str(r["example"]), 100)),
        )
        for r in overdue_rows
    )
    if not overdue_html:
        overdue_html = "<tr><td colspan='3'>No overdue owners</td></tr>"

    feed_rows = []
    for r in task_rows[:120]:
        feed_rows.append(
            "<tr class='hover-row' data-hover='{hover}'><td>{date}</td><td>{source}</td><td>{owner}</td><td>{eta}</td><td title='{task_full}'>{task_compact}</td></tr>".format(
                hover=_hover_payload(
                    r.owner,
                    f"{r.source} task row",
                    max=r.days_delta if r.days_delta is not None else "N/A",
                    q3=r.days_delta if r.days_delta is not None else "N/A",
                    median=r.days_delta if r.days_delta is not None else "N/A",
                    mean=r.days_delta if r.days_delta is not None else "N/A",
                    q1=r.days_delta if r.days_delta is not None else "N/A",
                    eta=r.eta,
                ),
                date=html.escape(r.date or "N/A"),
                source=html.escape(r.source),
                owner=html.escape(r.owner),
                eta=html.escape(r.eta),
                task_full=html.escape(r.task),
                task_compact=html.escape(_compact_text(r.task, 90)),
            )
        )
    if not feed_rows:
        feed_rows.append("<tr><td colspan='5'>No rows available</td></tr>")

    return f"""<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>AR Tracking Report</title>
  <style>
    :root {{
      --ink:#0f172a; --muted:#334155; --line:#dbe2ea; --bg:#f4f7fb;
      --card:#ffffff; --accent:#0c7a6a; --accent-2:#d46a07; --accent-3:#0f62c6;
    }}
    body {{ margin:0; font-family:Segoe UI, Tahoma, Arial, sans-serif; color:var(--ink); background:linear-gradient(180deg,#f8fbff 0%,#f4f7fb 100%); }}
    .top {{ padding:16px 18px; background:#111827; color:#fff; }}
    .top h1 {{ margin:0; font-size:22px; }}
    .top p {{ margin:4px 0 0; color:#d1d5db; }}
    .wrap {{ max-width:1280px; margin:16px auto 24px; padding:0 14px; }}
    .tabs {{ display:flex; gap:8px; flex-wrap:wrap; margin-bottom:12px; }}
    .tab-btn {{ border:1px solid var(--line); background:#fff; padding:8px 12px; border-radius:8px; cursor:pointer; font-weight:600; }}
    .tab-btn.active {{ background:var(--accent-3); color:#fff; border-color:var(--accent-3); }}
    .tab {{ display:none; background:var(--card); border:1px solid var(--line); border-radius:12px; padding:14px; box-shadow:0 2px 8px rgba(15,23,42,0.06); }}
    .tab.active {{ display:block; }}
    .cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:10px; margin-bottom:12px; }}
    .card {{ background:#fff; border:1px solid var(--line); border-radius:10px; padding:10px; }}
    .k {{ color:var(--muted); font-size:12px; text-transform:uppercase; }}
    .v {{ font-size:24px; font-weight:700; margin-top:4px; }}
    table {{ border-collapse:collapse; width:100%; }}
    th, td {{ border:1px solid var(--line); padding:8px; text-align:left; font-size:14px; }}
    th {{ background:#eef4fb; }}
        .hover-row {{ cursor: crosshair; }}
        #hoverTooltip {{
            position: fixed;
            display: none;
            z-index: 9999;
            min-width: 280px;
            max-width: 360px;
            background: rgba(12, 122, 106, 0.96);
            color: #fff;
            border-radius: 10px;
            padding: 10px 12px;
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.24);
            font-size: 13px;
            line-height: 1.45;
            pointer-events: none;
            transform: translate(18px, 18px);
        }}
        #hoverTooltip code {{ font-family:Consolas, monospace; color:#fff; background:rgba(255,255,255,0.16); padding:0 4px; border-radius:4px; }}
        #cursorPlus {{
            position: fixed;
            display: none;
            width: 28px;
            height: 28px;
            border: 2px solid #0c7a6a;
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.82);
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.22);
            pointer-events: none;
            z-index: 10000;
            transform: translate(-50%, -50%);
        }}
        #cursorPlus::before {{
            content: '+';
            position: absolute;
            inset: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #0c7a6a;
            font-size: 20px;
            font-weight: 800;
            line-height: 1;
        }}
    .search {{ margin-bottom:10px; }}
    .search input {{ width:min(360px,100%); padding:8px 10px; border:1px solid var(--line); border-radius:8px; }}
  </style>
</head>
<body>
  <header class=\"top\">
    <h1>AR Tracking Report</h1>
    <p>Dedicated AR page with tabs, hover detail panel, owner analytics, and searchable task feed.</p>
  </header>

  <main class=\"wrap\">
    <div class=\"tabs\">
      <button class=\"tab-btn active\" data-tab=\"overview\">Overview</button>
      <button class=\"tab-btn\" data-tab=\"source\">Source Scan Summary</button>
      <button class=\"tab-btn\" data-tab=\"owner\">Pending Tasks by Owner</button>
      <button class=\"tab-btn\" data-tab=\"upsert\">Upsert Summary</button>
      <button class=\"tab-btn\" data-tab=\"reminder\">Reminder Outcome</button>
      <button class=\"tab-btn\" data-tab=\"overdue\">Overdue Owners</button>
      <button class=\"tab-btn\" data-tab=\"feed\">Task Feed</button>
    </div>

    <section id=\"overview\" class=\"tab active\">
      <div class=\"cards\">
        <div class=\"card\"><div class=\"k\">Run Time</div><div class=\"v\">{overview['run_time']}</div></div>
        <div class=\"card\"><div class=\"k\">Lookback</div><div class=\"v\">{html.escape(str(overview['lookback']))}</div></div>
        <div class=\"card\"><div class=\"k\">Source Files</div><div class=\"v\">{overview['source_files']}</div></div>
        <div class=\"card\"><div class=\"k\">Active Rows</div><div class=\"v\">{overview['active_rows']}</div></div>
        <div class=\"card\"><div class=\"k\">Overdue</div><div class=\"v\">{overview['overdue']}</div></div>
      </div>
      <table>
        <tr><th>Metric</th><th>Value</th></tr>
                {overview_rows_html}
      </table>
    </section>

    <section id=\"source\" class=\"tab\">
      <table>
        <tr><th>Type</th><th>Name</th><th>Count</th></tr>
        {src_rows_html}
      </table>
    </section>

    <section id=\"owner\" class=\"tab\">
      <table>
        <tr><th>Owner</th><th>Active Count</th></tr>
        {''.join(owner_table_rows)}
      </table>
    </section>

    <section id=\"upsert\" class=\"tab\">
      <table>
        <tr><th>Metric</th><th>Value</th></tr>
                {upsert_rows_html}
      </table>
    </section>

    <section id=\"reminder\" class=\"tab\">
      <table>
        <tr><th>Metric</th><th>Value</th></tr>
                {reminder_rows_html}
      </table>
    </section>

    <section id=\"overdue\" class=\"tab\">
      <table>
        <tr><th>Owner</th><th>Overdue Count</th><th>Example Task</th></tr>
        {overdue_html}
      </table>
    </section>

    <section id=\"feed\" class=\"tab\">
      <div class=\"search\">
        <input id=\"feedFilter\" placeholder=\"Filter by owner, source, eta, task...\" />
      </div>
      <table id=\"feedTable\">
        <tr><th>Date</th><th>Source</th><th>Owner</th><th>ETA</th><th>Task</th></tr>
        {''.join(feed_rows)}
      </table>
    </section>
  </main>
    <div id="hoverTooltip"></div>
    <div id="cursorPlus"></div>

  <script>
    const buttons = document.querySelectorAll('.tab-btn');
    const tabs = document.querySelectorAll('.tab');
    buttons.forEach(btn => btn.addEventListener('click', () => {{
      buttons.forEach(b => b.classList.remove('active'));
      tabs.forEach(t => t.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.add('active');
    }}));

        const hoverTooltip = document.getElementById('hoverTooltip');
        const cursorPlus = document.getElementById('cursorPlus');
        let activeHover = null;

        function renderHoverContent(d) {{
            return `
                <strong>Details</strong><br/>
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
        }}

        function moveFloaters(evt) {{
            if (!activeHover) return;
            const x = evt.clientX;
            const y = evt.clientY;
            cursorPlus.style.left = `${{x}}px`;
            cursorPlus.style.top = `${{y}}px`;

            const pad = 16;
            const ttRect = hoverTooltip.getBoundingClientRect();
            let left = x + 18;
            let top = y + 18;
            if (left + ttRect.width + pad > window.innerWidth) left = x - ttRect.width - 18;
            if (top + ttRect.height + pad > window.innerHeight) top = y - ttRect.height - 18;
            hoverTooltip.style.left = `${{Math.max(pad, left)}}px`;
            hoverTooltip.style.top = `${{Math.max(pad, top)}}px`;
        }}

    document.querySelectorAll('.hover-row').forEach(row => {{
            row.addEventListener('mouseenter', (evt) => {{
                activeHover = row;
        const d = JSON.parse(row.dataset.hover || '{{}}');
                hoverTooltip.innerHTML = renderHoverContent(d);
                hoverTooltip.style.display = 'block';
                cursorPlus.style.display = 'block';
                moveFloaters(evt);
      }});

            row.addEventListener('mousemove', (evt) => moveFloaters(evt));

            row.addEventListener('mouseleave', () => {{
                activeHover = null;
                hoverTooltip.style.display = 'none';
                cursorPlus.style.display = 'none';
            }});
    }});

    const filterInput = document.getElementById('feedFilter');
    if (filterInput) {{
      filterInput.addEventListener('input', () => {{
        const q = filterInput.value.toLowerCase();
        document.querySelectorAll('#feedTable tr').forEach((row, idx) => {{
          if (idx === 0) return;
          const text = row.innerText.toLowerCase();
          row.style.display = text.includes(q) ? '' : 'none';
        }});
      }});
    }}
  </script>
</body>
</html>
"""


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate AR HTML report")
    parser.add_argument(
        "--input-json",
        default="scripts/ar_rows_latest.json;scripts/ar_rows_sent_followups.json",
        help="Semicolon-separated AR JSON files",
    )
    parser.add_argument(
        "--teams-json-glob",
        default="scripts/teams_*.json;scripts/t_*.json",
        help="Semicolon-separated glob patterns for teams source scan count",
    )
    parser.add_argument(
        "--outlook-json-glob",
        default="scripts/outlook*.json",
        help="Semicolon-separated glob patterns for outlook source scan count",
    )
    parser.add_argument("--lookback-days", type=int, default=7)
    parser.add_argument(
        "--output-html",
        default=str(OUTPUT_HTML),
        help="AR HTML output path",
    )
    args = parser.parse_args()

    data_files: list[Path] = []
    for item in _split_paths(args.input_json):
        p = Path(item)
        data_files.append(p if p.is_absolute() else BASE_DIR / p)

    raw_rows: list[dict[str, Any]] = []
    for p in data_files:
        raw_rows.extend(_load_json_rows(p))

    rows = _to_rows(raw_rows)
    rows, dupes = _dedupe(rows)

    source_files = _expand_globs(args.teams_json_glob) + _expand_globs(args.outlook_json_glob)

    overview = _overview(rows, args.lookback_days, Path(args.output_html), source_files)
    source_scan = _source_scan_summary(rows, source_files)
    owner_rows, owner_hover = _owner_stats(rows)
    upsert = _upsert_summary(len(raw_rows), len(rows), dupes)
    reminder = _reminder_outcome(rows)
    overdue = _overdue_owners(rows)

    html_text = _render_html(overview, source_scan, owner_rows, owner_hover, upsert, reminder, overdue, rows)

    output = Path(args.output_html)
    if not output.is_absolute():
        output = BASE_DIR / output
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(html_text, encoding="utf-8")

    print(f"Generated AR HTML report: {output}")
    print(f"Rows loaded: {len(raw_rows)} | Unique rows: {len(rows)} | Duplicates skipped: {dupes}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
