"""Builds a CAT2-style timesheet grid from filtered calendar events and
writes it to an .xlsx file for manual review/entry into SAP CAT2.

Column layout mirrors the CAT2 "Data Entry Area" screen exactly:
ActTyp | Rec. CCtr | Receiver WBS element | Name | Short text | Total |
MO dd.mm | TU dd.mm | WE dd.mm | TH dd.mm | FR dd.mm | SA dd.mm | SU dd.mm
"""

from __future__ import annotations

from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

DAY_LABELS = ["MO", "TU", "WE", "TH", "FR", "SA", "SU"]

HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
WARNING_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
SUCCESS_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
WEEKEND_FILL = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")

# The Mon-Sun hour columns visible in the sheet (7-13); a hidden numeric
# "mirror" column sits 9 columns to the right of each one (16-22) purely so
# Excel can SUM() real numbers for the totals - the visible cells themselves
# stay plain comma-decimal TEXT (e.g. "1,25") so copy/pasting into SAP CAT2
# always gets a literal comma, regardless of the workbook's own regional
# number settings.
HIDDEN_MIRROR_OFFSET = 9

RIGHT_ALIGN = Alignment(horizontal="right")
DASHED_BORDER = Border(
    top=Side(style="dashed", color="BFBFBF"),
    bottom=Side(style="dashed", color="BFBFBF"),
    left=Side(style="dashed", color="BFBFBF"),
    right=Side(style="dashed", color="BFBFBF"),
)


def fmt_hours(value: float | None) -> str | None:
    """Format an hour value with a comma decimal separator (e.g. "1,25")
    instead of a period, matching SAP CAT2's expected number format. Stored
    as text so the comma displays correctly regardless of the workbook's
    regional settings."""
    if value is None:
        return None
    return f"{value:.2f}".replace(".", ",")


def week_dates(monday: date) -> list[date]:
    return [monday + timedelta(days=i) for i in range(7)]


def resolve_wbs(title: str, cfg) -> tuple[str, str]:
    """Return (wbs_element, act_type) for an event title, applying the first
    matching keyword rule in cfg.WBS_MAPPING, falling back to the defaults."""
    title_lower = title.lower()
    for keywords, wbs, act_type in getattr(cfg, "WBS_MAPPING", []):
        if any(kw.lower() in title_lower for kw in keywords):
            return wbs, act_type
    return cfg.WBS_ELEMENT, cfg.ACT_TYPE


def build_rows(events, cfg) -> list[dict]:
    """Group events by (short title, wbs, act_type); sum hours per weekday
    for same-titled events, so a recurring meeting across the week becomes a
    single row with hours spread across the MO-SU columns."""
    rows: dict[tuple[str, str, str], dict] = {}
    for evt in events:
        short_text = evt.title[: cfg.SHORT_TEXT_MAX_LEN]
        wbs, act_type = resolve_wbs(evt.title, cfg)
        key = (short_text, wbs, act_type)
        row = rows.setdefault(
            key,
            {
                "act_type": act_type,
                "rec_cctr": "",
                "wbs": wbs,
                "name": "",
                "short_text": short_text,
                "notes": set(),
                "hours": {i: 0.0 for i in range(7)},
            },
        )
        weekday = evt.event_date.weekday()  # Monday=0 ... Sunday=6
        row["hours"][weekday] += evt.hours
        if getattr(evt, "location", ""):
            row["notes"].add(evt.location)

    # Sort rows by earliest day they appear on, then alphabetically.
    def sort_key(row):
        days_used = [d for d, h in row["hours"].items() if h > 0]
        return (min(days_used) if days_used else 7, row["short_text"].lower())

    result = sorted(rows.values(), key=sort_key)
    for row in result:
        row["notes"] = "; ".join(sorted(row["notes"]))
    return result


def _write_hour_row(ws, r: int, headers_len: int, hours: dict | None = None, dashed: bool = False) -> None:
    """Write the Total formula, day cells (optionally pre-filled from
    `hours`), and hidden numeric mirror formulas for one grid row. Used for
    calendar-derived rows as well as the manual-entry rows at the bottom."""
    if dashed:
        for col in range(1, headers_len + 1):
            ws.cell(row=r, column=col).border = DASHED_BORDER

    total_cell = ws.cell(row=r, column=6, value=f'=SUBSTITUTE(TEXT(SUM(P{r}:V{r}),"0.00"),".",",")')
    total_cell.alignment = RIGHT_ALIGN
    total_cell.number_format = "@"

    for i in range(7):
        col = 7 + i
        letter = get_column_letter(col)
        value = fmt_hours(round(hours[i], 2)) if hours and hours[i] else None
        cell = ws.cell(row=r, column=col, value=value)
        cell.alignment = RIGHT_ALIGN
        cell.number_format = "@"
        if i >= 5:
            cell.fill = WEEKEND_FILL
        mirror = ws.cell(
            row=r,
            column=col + HIDDEN_MIRROR_OFFSET,
            value=f'=IF({letter}{r}="",0,_xlfn.NUMBERVALUE({letter}{r},","))',
        )
        mirror.number_format = "0.00"


def export_excel(rows: list[dict], monday: date, output_path, cfg=None) -> None:
    dates = week_dates(monday)
    headers = ["ActTyp", "Rec. CCtr", "Receiver WBS element", "Name", "Short text", "Total"]
    headers += [f"{label} {d.day:02d}.{d.month:02d}" for label, d in zip(DAY_LABELS, dates)]
    headers += ["Notes"]

    wb = Workbook()
    ws = wb.active
    ws.title = "CAT2"

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        cell.fill = HEADER_FILL

    weekend_header_fill = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
    for col in (12, 13):
        ws.cell(row=1, column=col).fill = weekend_header_fill

    r = 2
    for row in rows:
        ws.cell(row=r, column=1, value=row["act_type"])
        ws.cell(row=r, column=2, value=row["rec_cctr"])
        ws.cell(row=r, column=3, value=row["wbs"])
        ws.cell(row=r, column=4, value=row["name"])
        ws.cell(row=r, column=5, value=row["short_text"])
        ws.cell(row=r, column=14, value=row.get("notes", ""))
        _write_hour_row(ws, r, len(headers), hours=row["hours"])
        r += 1

    # One "Development/Changes" row pre-filled with the default ActTyp and
    # the second WBS element from wbs_list.txt, followed by a few fully
    # blank rows so anything not picked up from the calendar can be added
    # by hand - all dashed-bordered so they read as "fill me in" rows.
    wbs_list = getattr(cfg, "WBS_LIST", []) if cfg else []
    second_wbs = wbs_list[1] if len(wbs_list) > 1 else ""
    default_act_type = getattr(cfg, "ACT_TYPE", "") if cfg else ""

    ws.cell(row=r, column=1, value=default_act_type)
    ws.cell(row=r, column=3, value=second_wbs)
    dev_changes_cell = ws.cell(row=r, column=5, value="Development/Changes")
    dev_changes_cell.comment = Comment(
        "Dashed rows are for hours booked to another WBS element/activity "
        "not picked up from your calendar.",
        "sap-timesheet-automation",
    )
    _write_hour_row(ws, r, len(headers), dashed=True)
    r += 1

    extra_blank_rows = getattr(cfg, "EXTRA_BLANK_ROWS", 3) if cfg else 3
    for _ in range(extra_blank_rows):
        _write_hour_row(ws, r, len(headers), dashed=True)
        r += 1

    last_data_row = r - 1

    # Daily totals summary row: each day's total is a live SUM() formula
    # over that column's hidden numeric mirror (so it stays accurate if you
    # edit/add hours in the blank rows by hand), rendered back into
    # comma-decimal text for consistency/SAP-paste safety.
    target_hours = getattr(cfg, "MAX_HOURS_PER_DAY", None) if cfg else 8
    summary_row = r
    label_cell = ws.cell(row=summary_row, column=5, value="Daily total")
    label_cell.font = Font(bold=True, italic=True)
    grand_total_cell = ws.cell(
        row=summary_row,
        column=6,
        value=f'=SUBSTITUTE(TEXT(SUM(P2:V{last_data_row}),"0.00"),".",",")',
    )
    grand_total_cell.font = Font(bold=True)
    grand_total_cell.alignment = RIGHT_ALIGN
    grand_total_cell.number_format = "@"
    for i in range(7):
        col = 7 + i
        mirror_letter = get_column_letter(col + HIDDEN_MIRROR_OFFSET)
        cell = ws.cell(
            row=summary_row,
            column=col,
            value=(
                f'=SUBSTITUTE(TEXT(SUM({mirror_letter}2:{mirror_letter}{last_data_row}),'
                f'"0.00"),".",",")'
            ),
        )
        cell.font = Font(bold=True)
        cell.alignment = RIGHT_ALIGN
        cell.number_format = "@"
        if i >= 5:
            cell.fill = WEEKEND_FILL

    # Conditional formatting is evaluated live by Excel against the hidden
    # numeric mirror columns (relative column reference, so it shifts day by
    # day across the range), so highlighting stays correct after manual
    # edits even though the visible cells hold text. Weekdays stay red while
    # under the daily target and turn green once it's reached.
    weekdays_range = f"G{summary_row}:K{summary_row}"
    mirror_start_row, mirror_end_row = 2, last_data_row
    if target_hours is not None:
        ws.conditional_formatting.add(
            weekdays_range,
            FormulaRule(
                formula=[f"SUM(P${mirror_start_row}:P${mirror_end_row})<>{target_hours}"],
                fill=WARNING_FILL,
            ),
        )
        ws.conditional_formatting.add(
            weekdays_range,
            FormulaRule(
                formula=[f"SUM(P${mirror_start_row}:P${mirror_end_row})={target_hours}"],
                fill=SUCCESS_FILL,
            ),
        )

    # Hide the numeric mirror columns - they exist only so Excel can SUM
    # real numbers; nothing in them is meant to be seen or pasted anywhere.
    for i in range(7):
        letter = get_column_letter(7 + i + HIDDEN_MIRROR_OFFSET)
        ws.column_dimensions[letter].hidden = True

    ws.freeze_panes = "A2"

    for col in range(1, len(headers) + 1):
        letter = get_column_letter(col)
        max_len = len(str(headers[col - 1]))
        if col != 6:  # Total column holds SUM() formulas; measure header only.
            for rr in range(2, last_data_row + 1):
                value = ws.cell(row=rr, column=col).value
                if value is not None:
                    max_len = max(max_len, len(str(value)))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 45)

    # Dropdown validation on the "Receiver WBS element" column, sourced from
    # wbs_list.txt, so manual entries (in the blank rows) can only be one of
    # your known valid codes instead of a free-typed value prone to typos.
    wbs_list = getattr(cfg, "WBS_LIST", []) if cfg else []
    if wbs_list:
        lists_ws = wb.create_sheet("Lists")
        lists_ws.sheet_state = "hidden"
        for i, code in enumerate(wbs_list, start=1):
            lists_ws.cell(row=i, column=1, value=code)
        dv = DataValidation(
            type="list",
            formula1=f"Lists!$A$1:$A${len(wbs_list)}",
            allow_blank=True,
            showDropDown=False,
        )
        ws.add_data_validation(dv)
        dv.add(f"C2:C{last_data_row}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
