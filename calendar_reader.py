"""Reads calendar events from macOS Calendar.app via a small EventKit-based
Swift helper (read_calendar.swift / read_calendar_bin).

We use EventKit directly rather than AppleScript because Calendar.app's
AppleScript "whose" filter on events is unusably slow with Exchange-backed
calendars (it must expand/scan the entire event history regardless of the
requested date range - it took several minutes for a single day in testing).
EventKit queries the same data natively in well under a second.
"""

from __future__ import annotations

import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
BINARY_PATH = SCRIPT_DIR / "read_calendar_bin"
SOURCE_PATH = SCRIPT_DIR / "read_calendar.swift"

# Covers emoji blocks (pictographs, symbols, dingbats, flags, variation
# selectors, skin-tone modifiers, ZWJ) so event titles keep only real text.
_EMOJI_RE = re.compile(
    "["
    "\U0001F300-\U0001FAFF"
    "\U00002600-\U000027BF"
    "\U0001F1E6-\U0001F1FF"
    "\U00002190-\U000021FF"
    "\U00002B00-\U00002BFF"
    "\U0000FE0F"
    "\U0000200D"
    "]+",
    flags=re.UNICODE,
)


def strip_emoji(text: str) -> str:
    """Remove emoji/pictograph characters from a string, then collapse any
    resulting double spaces and trim leading/trailing whitespace."""
    cleaned = _EMOJI_RE.sub("", text)
    return re.sub(r"\s{2,}", " ", cleaned).strip()


@dataclass
class CalendarEvent:
    start: datetime
    end: datetime
    all_day: bool
    status: str
    title: str
    location: str = ""

    @property
    def event_date(self) -> date:
        return self.start.date()

    @property
    def hours(self) -> float:
        return (self.end - self.start).total_seconds() / 3600.0


def _ensure_binary() -> Path:
    """Compile the Swift helper if the precompiled binary is missing."""
    if BINARY_PATH.exists():
        return BINARY_PATH
    print("Compiling read_calendar.swift (one-time)...", file=sys.stderr)
    subprocess.run(
        ["swiftc", "-O", str(SOURCE_PATH), "-o", str(BINARY_PATH)],
        check=True,
    )
    return BINARY_PATH


def read_events(
    start_date: date, end_date: date, calendar_names: list[str], my_email: str
) -> list[CalendarEvent]:
    """Read events between start_date and end_date (inclusive) from the given
    calendars. Requires macOS Calendar access permission for the terminal/
    process running this script (granted via a system prompt on first run).
    """
    today = date.today()
    offset_start = (start_date - today).days
    offset_end = (end_date - today).days

    binary = _ensure_binary()
    result = subprocess.run(
        [
            str(binary),
            str(offset_start),
            str(offset_end),
            ",".join(calendar_names),
            my_email,
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"read_calendar_bin failed: {result.stderr.strip()}"
        )

    events: list[CalendarEvent] = []
    for line in result.stdout.splitlines():
        if not line.strip():
            continue
        parts = line.split("|", 13)
        if len(parts) != 14:
            continue
        (
            sy, smo, sd, sh, smin,
            ey, emo, ed, eh, emin,
            all_day, status, title, location,
        ) = parts
        start = datetime(int(sy), int(smo), int(sd), int(sh), int(smin))
        end = datetime(int(ey), int(emo), int(ed), int(eh), int(emin))
        events.append(
            CalendarEvent(
                start=start,
                end=end,
                all_day=(all_day == "true"),
                status=status,
                title=title,
                location=location,
            )
        )
    return events


def merge_overlapping_events(events: list[CalendarEvent]) -> list[CalendarEvent]:
    """Merge same-titled events that overlap or are back-to-back (common with
    recurring/duplicated invites) so their hours aren't double-counted.
    Overlapping events with *different* titles are left as-is but a warning
    is printed, since we can't automatically decide which one to keep."""
    by_day: dict[date, list[CalendarEvent]] = {}
    for evt in events:
        by_day.setdefault(evt.event_date, []).append(evt)

    merged: list[CalendarEvent] = []
    for day, day_events in by_day.items():
        day_events.sort(key=lambda e: e.start)
        active: list[CalendarEvent] = []
        for evt in day_events:
            overlap = None
            for other in active:
                if evt.start < other.end and other.start < evt.end:
                    overlap = other
                    break
            if overlap is None:
                active.append(evt)
                continue
            if overlap.title.strip().lower() == evt.title.strip().lower():
                overlap.start = min(overlap.start, evt.start)
                overlap.end = max(overlap.end, evt.end)
                if evt.location and evt.location not in overlap.location:
                    overlap.location = (overlap.location + "; " + evt.location).strip("; ")
            else:
                print(
                    f"Warning: '{evt.title}' overlaps with '{overlap.title}' on "
                    f"{day.isoformat()} - hours may be double-booked, review manually.",
                    file=sys.stderr,
                )
                active.append(evt)
        merged.extend(active)

    merged.sort(key=lambda e: e.start)
    return merged


def apply_title_alias(title: str, cfg) -> str:
    """Replace the title with a clean alias if it matches a key in
    cfg.TITLE_ALIASES (case-insensitive substring match, first match wins)."""
    title_lower = title.lower()
    for key, alias in getattr(cfg, "TITLE_ALIASES", {}).items():
        if key.lower() in title_lower:
            return alias
    return title


def filter_events(events: list[CalendarEvent], cfg) -> list[CalendarEvent]:
    filtered: list[CalendarEvent] = []
    for evt in events:
        if cfg.SKIP_ALL_DAY_EVENTS and evt.all_day:
            continue
        if cfg.ONLY_ACCEPTED_MEETINGS and evt.status != "accepted":
            continue
        title_lower = evt.title.lower()
        if any(kw.lower() in title_lower for kw in cfg.EXCLUDE_KEYWORDS):
            continue

        evt.title = apply_title_alias(strip_emoji(evt.title), cfg)

        hours = evt.hours
        if cfg.ROUND_TO_HOURS:
            hours = round(hours / cfg.ROUND_TO_HOURS) * cfg.ROUND_TO_HOURS
            evt.end = evt.start + timedelta(hours=hours)

        if evt.hours < cfg.MIN_EVENT_HOURS:
            continue

        filtered.append(evt)

    # Warn (don't truncate) if a day's total exceeds the configured max.
    totals: dict[date, float] = {}
    for evt in filtered:
        totals[evt.event_date] = totals.get(evt.event_date, 0.0) + evt.hours
    for day, total in totals.items():
        if total > cfg.MAX_HOURS_PER_DAY:
            print(
                f"Warning: {day.isoformat()} has {total:.2f}h booked "
                f"(exceeds MAX_HOURS_PER_DAY={cfg.MAX_HOURS_PER_DAY}).",
                file=sys.stderr,
            )

    return filtered
