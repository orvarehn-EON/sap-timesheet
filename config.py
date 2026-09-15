# Configuration for the SAP CAT2 timesheet export tool.
# Edit the values below to match your own setup.

from pathlib import Path

_SCRIPT_DIR = Path(__file__).resolve().parent


def _load_wbs_list(path: Path) -> list[str]:
    """Read WBS elements from wbs_list.txt (one per line, # = comment)."""
    if not path.exists():
        return []
    codes = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            codes.append(line)
    return codes


# All valid WBS elements, read from wbs_list.txt. Edit that file to add/
# remove codes - the first one is used as the default WBS_ELEMENT below and
# as the top entry in the "Receiver WBS element" dropdown in the Excel file.
WBS_LIST = _load_wbs_list(_SCRIPT_DIR / "wbs_list.txt")

# --- SAP CAT2 fixed values -------------------------------------------------
WBS_ELEMENT = WBS_LIST[0] if WBS_LIST else "0201.P00363.100.01"  # default Receiver WBS element
ACT_TYPE = "NHTJÄ"  # Activity type code

# Keyword-based overrides: events whose title contains any of the given
# (case-insensitive) keywords are booked to a different WBS/ActType instead
# of the defaults above. Checked in order; first match wins. Each entry is
# (keywords, wbs_element, act_type).
WBS_MAPPING = [
    # (["client-x", "project-x"], "0201.P00999.100.01", "NHTJÄ"),
]

# Map meeting titles (or substrings, case-insensitive) to a clean short
# label instead of using/truncating the raw calendar title. Checked in
# order after emoji-stripping; first match wins, whole title is replaced
# with the alias. Example:
TITLE_ALIASES = {
    # "Domain Area + Sign Contract - get togeth": "Domain sync",
}

# --- Calendar source --------------------------------------------------------
# Name(s) of the macOS Calendar.app calendar(s) to read events from. Must
# match exactly what is shown in Calendar.app's sidebar. If two calendars
# share the same name (e.g. an old iCloud one and your Exchange one), both
# are read - remove the stale one from Calendar.app's sidebar if that causes
# unwanted events to show up.
CALENDAR_NAMES = ["Calendar"]

# Your own email address, used to find your RSVP status on each event.
# EventKit's "isCurrentUser" flag is unreliable for Exchange/EWS accounts,
# so we match this against each event's attendee list instead.
MY_EMAIL = "atul.sareen@eon.se"

# --- Event filtering ---------------------------------------------------------
# Events whose title contains any of these (case-insensitive) are excluded.
EXCLUDE_KEYWORDS = [
    "lunch",
    "focus time",
    "coffee-time",
]

# Only include events you've accepted (participation status == "accepted").
ONLY_ACCEPTED_MEETINGS = True

# Skip all-day events (e.g. "Out of office", holidays).
SKIP_ALL_DAY_EVENTS = True

# Round each event's duration to the nearest fraction of an hour.
# Set to None to disable rounding and use exact durations.
ROUND_TO_HOURS = 0.25

# Drop events shorter than this many hours (after rounding).
MIN_EVENT_HOURS = 0.25

# Warn (not truncate) if a single day's total booked hours exceeds this.
MAX_HOURS_PER_DAY = 8

# --- Short text -------------------------------------------------------------
# Max characters kept from the event title for the "Short text" column
# (SAP text fields are limited in length; adjust to match your CAT2 field).
SHORT_TEXT_MAX_LEN = 40

# --- Output -------------------------------------------------------------
OUTPUT_DIR = "output"

# Automatically open the generated .xlsx in the default app after writing it.
AUTO_OPEN = True

# How many extra fully-blank, dashed-bordered rows to add at the bottom
# (after the auto-filled "Development/Changes" row) for anything not
# picked up from the calendar.
EXTRA_BLANK_ROWS = 3
