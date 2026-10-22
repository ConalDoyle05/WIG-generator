"""Step times: reading what you typed in a Time column, and writing totals.

Accepted, in a cell of a column headed "Time", "Time (s)", "Time (min)" and so on:
    45 s   45s   2 min   1.5 min   1m 30s   1 min 30 sec   1 h 5 min
    1:30 (minutes:seconds)   1:02:30 (hours:minutes:seconds), when typed as text
    45     a bare number, only when the header names the unit ("Time (s)")

A bare number under a plain "Time" header is refused rather than guessed, the
same rule the polish step applies to numbers without units.

Excel trap: type 1:30 into a cell Excel treats as a time and it stores 1 hour 30
minutes, not 1 minute 30 seconds. Times like that are read as Excel stored
them, so the generator warns about any single step over an hour.
"""

from __future__ import annotations

import datetime as _dt
import re
from typing import Optional

_UNITS = {
    "h": 3600, "hr": 3600, "hrs": 3600, "hour": 3600, "hours": 3600,
    "m": 60, "min": 60, "mins": 60, "minute": 60, "minutes": 60,
    "s": 1, "sec": 1, "secs": 1, "second": 1, "seconds": 1,
}
_PART = re.compile(r"(\d+(?:\.\d+)?)\s*(hours|hour|hrs|hr|h|minutes|minute|mins|min|m|seconds|second|secs|sec|s)(?![a-z])")
_CLOCK = re.compile(r"(?:(\d+):)?(\d+):(\d{2})")
_NUMBER = re.compile(r"\d+(?:\.\d+)?")
LONG_STEP_SECONDS = 3600


def header_unit(header: str) -> Optional[int]:
    """Seconds per unit named in a header like 'Time (s)' or 'Time (min)', or None for plain 'Time'."""
    text = (header or "").lower()
    match = re.search(r"\(([^)]*)\)", text)
    unit = (match.group(1) if match else text.replace("time", "", 1)).strip(" :-")
    return _UNITS.get(unit)


def parse_duration(value, unit_seconds: Optional[int] = None) -> Optional[float]:
    """Seconds for one Time cell, None for a blank one. Raises ValueError with a reason."""
    if value is None:
        return None
    if isinstance(value, _dt.timedelta):
        return value.total_seconds()
    if isinstance(value, _dt.time):
        return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, _dt.datetime):
        raise ValueError("it was stored as a date. Type the time as text, e.g. 45 s or 1m 30s")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if unit_seconds is None:
            raise ValueError('a bare number needs a unit: write "45 s" or "2 min", or head the column "Time (s)"')
        return float(value) * unit_seconds

    text = str(value).strip().lower()
    if not text:
        return None
    clock = _CLOCK.fullmatch(text)
    if clock:
        hours, minutes, seconds = clock.groups()
        if int(seconds) >= 60 or (hours is not None and int(minutes) >= 60):
            raise ValueError("minutes and seconds must be under 60")
        return int(hours or 0) * 3600 + int(minutes) * 60 + int(seconds)
    if _NUMBER.fullmatch(text):
        return parse_duration(float(text), unit_seconds)

    cleaned = re.sub(r"\band\b|,", " ", text)
    total, consumed = 0.0, 0
    for match in _PART.finditer(cleaned):
        total += float(match.group(1)) * _UNITS[match.group(2)]
        consumed += len(re.sub(r"\s", "", match.group(0)))
    if consumed == 0 or consumed != len(re.sub(r"\s", "", cleaned)):
        raise ValueError('write it like "45 s", "2 min", "1m 30s" or "1:30"')
    return total


def is_duration_text(text: str) -> bool:
    """True for text that is only a time with its unit ("45 s", "1m 30s", "1:30"). Bare numbers don't count."""
    text = (text or "").strip()
    if not text or _NUMBER.fullmatch(text):
        return False
    try:
        parse_duration(text)
    except ValueError:
        return False
    return True


def format_duration(seconds: float) -> str:
    """45 s, 12 min, 12 min 30 s, 1 h 05 min."""
    total = int(round(seconds))
    hours, rest = divmod(total, 3600)
    minutes, secs = divmod(rest, 60)
    if hours:
        text = f"{hours} h {minutes:02d} min"
        return f"{text} {secs} s" if secs else text
    if minutes:
        return f"{minutes} min {secs} s" if secs else f"{minutes} min"
    return f"{secs} s"
