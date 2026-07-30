from __future__ import annotations

import argparse
import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import openpyxl  # type: ignore[import-untyped]

BASE_DIR = Path(__file__).resolve().parents[1]


def _parse_json_list(path_text: str) -> list[Path]:
    return [Path(p.strip()) for p in path_text.split(";") if p.strip()]


def _load_rows(paths: list[Path]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for p in paths:
        full = p if p.is_absolute() else BASE_DIR / p
        if not full.exists():
            continue
        with full.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            for item in data:
                if isinstance(item, dict):
                    rows.append(item)
    return rows


def _parse_date(value: str) -> datetime | None:
    try:
        return datetime.strptime(value.strip(), "%Y-%m-%d")
    except Exception:
        return None


def _filter_lookback(rows: list[dict[str, Any]], lookback_days: int) -> list[dict[str, Any]]:
    if lookback_days <= 0:
        return rows
    threshold = datetime.now() - timedelta(days=lookback_days)
    keep: list[dict[str, Any]] = []
    for r in rows:
        dt = _parse_date(str(r.get("date", "")))
        if dt is None:
            continue
        if dt >= threshold:
            keep.append(r)
    return keep


def _dedupe(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: set[str] = set()
    out: list[dict[str, Any]] = []
    for r in rows:
        key = str(r.get("source_key", "")).strip()
        if not key:
            key = f"{r.get('source','')}|{r.get('date','')}|{r.get('task','')}"
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    out.sort(key=lambda x: str(x.get("date", "")), reverse=True)
    return out


def _ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def _write_ar_tracker(rows: list[dict[str, Any]], output: Path, sheet_name: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    headers = ["Date", "Channel", "Task", "Owner", "Status", "ETA", "Remark", "Source"]
    ws.append(headers)

    for r in rows:
        task = str(r.get("task", "")).strip()
        if not task:
            continue
        date_text = str(r.get("date", "")).strip()
        channel = str(r.get("folder", r.get("source", "Unknown")) or "Unknown").strip() or "Unknown"
        owner = str(r.get("owner", "TBD") or "TBD").strip() or "TBD"
        status = str(r.get("status", "Open") or "Open").strip() or "Open"
        eta = str(r.get("eta", "TBD") or "TBD").strip() or "TBD"
        summary = str(r.get("summary", "")).strip()
        sender = str(r.get("sender", "")).strip()
        remark = summary if summary else ""
        if sender:
            remark = f"{remark} | Sender: {sender}" if remark else f"Sender: {sender}"
        source_key = str(r.get("source_key", "")).strip()
        ws.append([date_text, channel, task, owner, status, eta, remark, source_key])

    ws.freeze_panes = "A2"
    widths = {
        "A": 14,
        "B": 20,
        "C": 72,
        "D": 30,
        "E": 20,
        "F": 14,
        "G": 90,
        "H": 42,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    _ensure_parent(output)
    wb.save(output)


def main() -> None:
    parser = argparse.ArgumentParser(description="Create AR Excel tracker from AR JSON rows")
    parser.add_argument(
        "--input-json",
        required=True,
        help="Semicolon-separated JSON file paths (e.g. scripts/ar_rows_latest.json;scripts/ar_rows_sent_followups.json)",
    )
    parser.add_argument(
        "--task",
        required=True,
        choices=["ar"],
        help="Target output type (AR only)",
    )
    parser.add_argument("--lookback-days", type=int, default=7, help="Include only rows from last N days")
    parser.add_argument(
        "--output",
        default="",
        help="Output workbook path. Default: artifacts/ar/AR_Tracking_Auto.xlsx",
    )
    parser.add_argument(
        "--sheet-name",
        default="",
        help="Sheet name. Default: AR_Tracking",
    )
    args = parser.parse_args()

    paths = _parse_json_list(args.input_json)
    if not paths:
        raise SystemExit("No JSON files provided")

    rows = _load_rows(paths)
    rows = _filter_lookback(rows, args.lookback_days)
    rows = _dedupe(rows)

    output = Path(args.output) if args.output else BASE_DIR / "artifacts" / "ar" / "AR_Tracking_Auto.xlsx"
    output = output if output.is_absolute() else BASE_DIR / output
    sheet_name = args.sheet_name or "AR_Tracking"
    _write_ar_tracker(rows, output, sheet_name)

    print(f"Created {args.task} tracker workbook: {output}")
    print(f"Rows written (after lookback+dedupe): {len(rows)}")


if __name__ == "__main__":
    main()
