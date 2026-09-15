# SAP CAT2 Timesheet Export (macOS)

Reads your macOS Calendar events for a week and generates an Excel file
laid out like the SAP CAT2 grid, ready to copy/type into SAP. It does not
write into SAP GUI directly — CAT2's grid isn't scriptable on macOS, so
this just saves you the manual work of tallying hours from your calendar.

## Requirements

- Xcode Command Line Tools (`xcode-select --install`) — for the Swift
  calendar reader.
- Your work calendar in **Calendar.app** (not "New Outlook") — add your
  Exchange account via *Calendar → Settings → Accounts* first if it isn't
  there yet.
- Python 3 + `openpyxl`: `pip3 install --break-system-packages -r requirements.txt`

## Usage

Run `python3 cal2cat.py --help` any time to see all options.

```bash
python3 cal2cat.py --dry-run          # preview, no file written
python3 cal2cat.py                    # current week -> output/timesheet_<monday>.xlsx
python3 cal2cat.py --week-offset -1   # previous week / --week-offset 1 for next
python3 cal2cat.py --output my.xlsx   # custom path
python3 cal2cat.py --no-open          # skip auto-opening the file
```

First run will prompt for Calendar access — approve it. Each run deletes
any existing file at the target path before writing a fresh one.

## What it does

- Groups events by title (and by WBS if `WBS_MAPPING` matches), sums hours
  per weekday, and fills the CAT2-style grid: ActTyp, WBS, Short text,
  Total, Mon–Sun columns, Notes.
- Merges overlapping same-titled events so hours aren't double-counted;
  warns on overlaps between different-titled events.
- Strips emojis from titles and applies `TITLE_ALIASES` to give long/odd
  titles a clean short label.
- Rounds/filters durations (`ROUND_TO_HOURS`, `MIN_EVENT_HOURS`), skips
  all-day events and events you didn't accept (configurable).
- Adds a **"Daily total"** row with live `SUM()` formulas per day (and a
  grand total), so it stays correct if you edit hours by hand afterwards.
  Weekday cells are shaded red while below `MAX_HOURS_PER_DAY` (default 8h)
  and turn green once the target is reached.
- Adds one dashed "Development/Changes" row (pre-filled with the default
  ActTyp and the second WBS from `wbs_list.txt`) plus `EXTRA_BLANK_ROWS`
  (default 3) fully blank dashed rows at the bottom, for anything not
  picked up from the calendar.
- "Receiver WBS element" is a dropdown sourced from `wbs_list.txt`.
- Weekend (Sat/Sun) columns are shaded light gray.
- Auto-opens the file after writing (`AUTO_OPEN`).

## Configuration

- **`wbs_list.txt`** — one WBS code per line (`#` = comment). First line is
  the default `WBS_ELEMENT`; second line is used in the auto-inserted
  "Development/Changes" row; the whole list feeds the Excel dropdown.
- **`config.py`** — everything else: `ACT_TYPE`, `WBS_MAPPING` (keyword →
  WBS/ActType overrides), `TITLE_ALIASES`, `CALENDAR_NAMES`,
  `EXCLUDE_KEYWORDS`, `ONLY_ACCEPTED_MEETINGS`, `SKIP_ALL_DAY_EVENTS`,
  `ROUND_TO_HOURS`, `MIN_EVENT_HOURS`, `MAX_HOURS_PER_DAY`,
  `SHORT_TEXT_MAX_LEN`, `AUTO_OPEN`, `EXTRA_BLANK_ROWS`.

"Rec. CCtr" and "Name" are left blank — fill in manually, or hardcode in
`export_timesheet.py` if always the same.

## Weekly reminder (optional)

```bash
./schedule.sh install     # runs every Friday at 16:00 (launchd)
./schedule.sh status
./schedule.sh uninstall
```

Edit `com.sap.timesheet.plist`'s `StartCalendarInterval` to change the
schedule. Scheduled runs may need a separate Calendar permission grant
(System Settings → Privacy & Security → Calendars).

## Files

- `read_calendar.swift` / `read_calendar_bin` — EventKit-based calendar
  reader (faster and more reliable than AppleScript for Exchange
  calendars).
- `calendar_reader.py` — runs the binary, filters/merges events.
- `export_timesheet.py` — builds the formatted `.xlsx`.
- `cal2cat.py` — CLI entry point.

## Troubleshooting

- **No events found**: check `CALENDAR_NAMES` matches Calendar.app exactly.
- **Calendar access denied**: System Settings → Privacy & Security →
  Calendars, grant access, re-run.
- **Slow first run**: Swift helper compiles once; delete
  `read_calendar_bin` and re-run if you edit `read_calendar.swift`.
