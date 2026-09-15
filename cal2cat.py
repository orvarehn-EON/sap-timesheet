#!/usr/bin/env python3
"""Reads your macOS Calendar events for a given week and exports a
CAT2-shaped Excel file for manual review/entry into SAP timesheets.

Usage:
    python3 cal2cat.py                  # current week
    python3 cal2cat.py --week-offset -1 # previous week
    python3 cal2cat.py --week-offset 1  # next week
    python3 cal2cat.py --dry-run        # print summary, don't write a file
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from datetime import date, timedelta
from pathlib import Path

import config
import calendar_reader
import export_timesheet


def monday_of_week(week_offset: int) -> date:
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    return monday + timedelta(weeks=week_offset)


def print_summary(rows: list[dict]) -> None:
    if not rows:
        print("No events matched after filtering.")
        return
    print(f"{'Short text':42} {'Total':>6}  " + "  ".join(export_timesheet.DAY_LABELS))
    for row in rows:
        hours = [row["hours"][i] for i in range(7)]
        total = sum(hours)
        hours_str = "  ".join(
            f"{export_timesheet.fmt_hours(h):>5}" if h else "     " for h in hours
        )
        print(f"{row['short_text'][:42]:42} {export_timesheet.fmt_hours(total):>6}  {hours_str}")


def open_file(path: Path) -> None:
    if sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="cal2cat.py",
        description=(
            "Export your macOS Calendar events for a week into a "
            "CAT2-shaped Excel file (output/timesheet_<monday>.xlsx) for "
            "manual review/entry into SAP. See README.md for full details "
            "on configuration (config.py, wbs_list.txt)."
        ),
        epilog=(
            "Examples:\n"
            "  python3 cal2cat.py --dry-run          preview this week, no file written\n"
            "  python3 cal2cat.py                     export current week\n"
            "  python3 cal2cat.py --week-offset -1    export previous week\n"
            "  python3 cal2cat.py --week-offset 1     export next week\n"
            "  python3 cal2cat.py --output my.xlsx    export to a custom path\n"
            "  python3 cal2cat.py --no-open           export without auto-opening the file\n"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--week-offset",
        type=int,
        default=0,
        help="0 = current week, -1 = previous week, 1 = next week, ...",
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help="Output .xlsx path (default: output/timesheet_<monday>.xlsx)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print the summary only, don't write an output file.",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Don't auto-open the generated file (overrides config.AUTO_OPEN).",
    )
    args = parser.parse_args()

    monday = monday_of_week(args.week_offset)
    sunday = monday + timedelta(days=6)
    print(f"Reading calendar events for {monday.isoformat()} - {sunday.isoformat()}...")

    events = calendar_reader.read_events(
        monday, sunday, config.CALENDAR_NAMES, config.MY_EMAIL
    )
    print(f"Found {len(events)} raw event(s).")

    filtered = calendar_reader.filter_events(events, config)
    filtered = calendar_reader.merge_overlapping_events(filtered)
    print(f"{len(filtered)} event(s) remain after filtering/merging.")

    rows = export_timesheet.build_rows(filtered, config)
    print_summary(rows)

    if args.dry_run:
        print("\nDry run - no file written.")
        return

    if args.output:
        output_path = Path(args.output)
    else:
        output_path = Path(config.OUTPUT_DIR) / f"timesheet_{monday.isoformat()}.xlsx"

    if output_path.exists():
        output_path.unlink()
        print(f"Deleted existing {output_path}")

    export_timesheet.export_excel(rows, monday, output_path, cfg=config)
    print(f"\nWrote {output_path}")

    if config.AUTO_OPEN and not args.no_open:
        open_file(output_path)


if __name__ == "__main__":
    main()
