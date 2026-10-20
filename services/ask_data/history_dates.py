"""Ask Data Phase 5: deterministic date/period-phrase resolution (EN + Hebrew).

The planner (Gemini) is ONLY ever allowed to copy a date phrase verbatim out of the user's question into a
tool argument -- it never computes, normalizes, or invents a date itself (see planner.py's history rules).
Everything here is plain, tested, deterministic Python using the REAL wall-clock `today` passed in by the
caller (never trusting any date the model might have written). An unrecognised phrase is reported as
NOT_FOUND, never guessed.

Two resolvers:
  * `resolve_single_date_phrase` -- a single point in time (for get_debt_at / compare_debt_dates).
  * `resolve_period_phrase` -- a [date_from, date_to] range (for get_debt_change / get_invoice_changes /
    get_payments_in_period / get_new_invoices_in_period).
Some phrases ("today", "yesterday") are valid for both; others are range-only ("since X", "past N days",
"this/last month") or point-only ("N days ago" as a single lookback point).
"""
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Optional

ACCEPTED = "accepted"
NOT_FOUND = "not_found"

_MONTHS_EN = {
    "january": 1, "jan": 1, "february": 2, "feb": 2, "march": 3, "mar": 3, "april": 4, "apr": 4,
    "may": 5, "june": 6, "jun": 6, "july": 7, "jul": 7, "august": 8, "aug": 8,
    "september": 9, "sept": 9, "sep": 9, "october": 10, "oct": 10, "november": 11, "nov": 11,
    "december": 12, "dec": 12,
}
_MONTHS_HE = {
    "ינואר": 1, "פברואר": 2, "מרץ": 3, "אפריל": 4, "מאי": 5, "יוני": 6, "יולי": 7, "אוגוסט": 8,
    "ספטמבר": 9, "אוקטובר": 10, "נובמבר": 11, "דצמבר": 12,
}
_DIGITS_RE = re.compile(r"\d+")


@dataclass(frozen=True)
class DateResolution:
    status: str
    value: Optional[date] = None
    label: str = ""


@dataclass(frozen=True)
class PeriodResolution:
    status: str
    date_from: Optional[date] = None
    date_to: Optional[date] = None
    label: str = ""

    def __str__(self) -> str:
        """Human-readable form used e.g. in ToolResult.scope_text()'s 'Scope: period=...' line."""
        if self.status != ACCEPTED or self.date_from is None:
            return "unresolved period"
        return f"{self.label} ({self.date_from.isoformat()} to {self.date_to.isoformat()})"


def _month_in_text(lowered: str, raw: str) -> Optional[tuple[int, str]]:
    """First recognised month name in the text (EN, case-insensitive, or Hebrew), or None. Uses word
    boundaries so a short abbreviation like 'dec'/'may'/'mar' never false-positives inside an unrelated
    word (e.g. 'decade', 'mayor', 'market')."""
    for name, number in _MONTHS_EN.items():
        if re.search(rf"\b{re.escape(name)}\b", lowered):
            return number, name
    for name, number in _MONTHS_HE.items():
        if name in raw:
            return number, name
    return None


def _month_bounds(year: int, month: int) -> tuple[date, date]:
    last_day = calendar.monthrange(year, month)[1]
    return date(year, month, 1), date(year, month, last_day)


def _week_bounds(anchor: date) -> tuple[date, date]:
    monday = anchor - timedelta(days=anchor.weekday())
    return monday, monday + timedelta(days=6)


def _past_n_days(lowered: str, raw: str) -> Optional[int]:
    """'past N days' / 'N days ago' / Hebrew 'לפני N ימים' -> N, else None."""
    match = re.search(r"(?:past|last)\s+(\d+)\s+days?", lowered) or re.search(r"(\d+)\s+days?\s+ago", lowered)
    if match:
        return int(match.group(1))
    match = re.search(r"לפני\s*(\d+)\s*ימים", raw)
    if match:
        return int(match.group(1))
    return None


def resolve_single_date_phrase(phrase: str, today: date) -> DateResolution:
    """One concrete calendar date. Range-only phrases ("since ...", "past N days", "this/last month" as a
    whole) are NOT_FOUND here -- use `resolve_period_phrase` for those instead."""
    raw = (phrase or "").strip()
    lowered = raw.lower()
    if not lowered:
        return DateResolution(NOT_FOUND)

    if lowered in ("today", "now") or raw in ("היום",):
        return DateResolution(ACCEPTED, today, "today")
    if lowered == "yesterday" or raw in ("אתמול",):
        d = today - timedelta(days=1)
        return DateResolution(ACCEPTED, d, "yesterday")

    n_days = _past_n_days(lowered, raw)
    if n_days is not None and ("ago" in lowered or "לפני" in raw):
        d = today - timedelta(days=n_days)
        return DateResolution(ACCEPTED, d, f"{n_days} day(s) ago")

    if lowered in ("a week ago", "last week", "one week ago") or raw in ("לפני שבוע", "שבוע שעבר"):
        d = today - timedelta(days=7)
        return DateResolution(ACCEPTED, d, "a week ago")
    if lowered in ("a month ago", "last month", "one month ago") or raw in ("לפני חודש", "חודש שעבר"):
        d = _subtract_month(today)
        return DateResolution(ACCEPTED, d, "a month ago")

    # "30 September" / "September 30" / "1 October" style explicit day+month (no year -> current year,
    # or the most recent PAST occurrence if that date would otherwise be in the future). Explicitly
    # excludes range-indicator words ("since", "past", "this"/"last" + week/month) so e.g. "since
    # September 1" is correctly NOT_FOUND here -- that is a period phrase, not a single point.
    _RANGE_INDICATORS = ("since", "past ", "this week", "last week", "this month", "last month")
    if any(w in lowered for w in _RANGE_INDICATORS) or "מאז" in raw or "תחילת" in raw:
        return DateResolution(NOT_FOUND)
    month_hit = _month_in_text(lowered, raw)
    if month_hit:
        month_number, _ = month_hit
        day_match = _DIGITS_RE.search(raw)
        if day_match:
            day_number = int(day_match.group())
            try:
                candidate = date(today.year, month_number, day_number)
            except ValueError:
                return DateResolution(NOT_FOUND)
            if candidate > today:
                candidate = date(today.year - 1, month_number, day_number)
            return DateResolution(ACCEPTED, candidate, candidate.isoformat())

    return DateResolution(NOT_FOUND)


def resolve_period_phrase(phrase: str, today: date) -> PeriodResolution:
    """A [date_from, date_to] range, inclusive, clipped so `date_to` never exceeds `today` (never
    extrapolate into the future)."""
    raw = (phrase or "").strip()
    lowered = raw.lower()
    if not lowered:
        return PeriodResolution(NOT_FOUND)

    if lowered in ("today", "now") or raw in ("היום",):
        return PeriodResolution(ACCEPTED, today, today, "today")
    if lowered == "yesterday" or raw in ("אתמול",):
        d = today - timedelta(days=1)
        return PeriodResolution(ACCEPTED, d, d, "yesterday")

    if lowered == "this week" or raw in ("השבוע",):
        start, _ = _week_bounds(today)
        return PeriodResolution(ACCEPTED, start, today, "this week")
    if lowered == "last week" or raw in ("שבוע שעבר", "השבוע שעבר"):
        this_start, _ = _week_bounds(today)
        last_start = this_start - timedelta(days=7)
        last_end = this_start - timedelta(days=1)
        return PeriodResolution(ACCEPTED, last_start, last_end, "last week")

    if lowered == "this month" or raw in ("החודש",) or "מתחילת החודש" in raw:
        start = date(today.year, today.month, 1)
        return PeriodResolution(ACCEPTED, start, today, "this month")
    if lowered == "last month" or raw in ("חודש שעבר",):
        prev_month_end = date(today.year, today.month, 1) - timedelta(days=1)
        start = date(prev_month_end.year, prev_month_end.month, 1)
        return PeriodResolution(ACCEPTED, start, prev_month_end, "last month")

    n_days = _past_n_days(lowered, raw)
    if n_days is not None:
        start = today - timedelta(days=n_days)
        return PeriodResolution(ACCEPTED, start, today, f"past {n_days} day(s)")

    since_match = re.search(r"since (?:the beginning of )?(.+)", lowered)
    if since_match or "מאז" in raw:
        tail = since_match.group(1).strip() if since_match else raw.split("מאז", 1)[1].strip()
        if "beginning of" in lowered or "תחילת" in raw:
            month_hit = _month_in_text(lowered, raw)
            if month_hit:
                month_number, _ = month_hit
                year = today.year
                start = date(year, month_number, 1)
                if start > today:
                    start = date(year - 1, month_number, 1)
                return PeriodResolution(ACCEPTED, start, today, start.isoformat())
            return PeriodResolution(NOT_FOUND)
        point = resolve_single_date_phrase(tail, today)
        if point.status == ACCEPTED:
            return PeriodResolution(ACCEPTED, point.value, today, point.value.isoformat())
        return PeriodResolution(NOT_FOUND)

    # A bare month name on its own ("October" / "in October" / Hebrew "באוקטובר") -> that whole month,
    # clipped to today if it is the current (or a future) month.
    month_hit = _month_in_text(lowered, raw)
    if month_hit:
        month_number, _ = month_hit
        year = today.year
        start, end = _month_bounds(year, month_number)
        if start > today:
            year -= 1
            start, end = _month_bounds(year, month_number)
        if end > today:
            end = today
        return PeriodResolution(ACCEPTED, start, end, f"{calendar.month_name[month_number]} {year}")

    point = resolve_single_date_phrase(raw, today)
    if point.status == ACCEPTED:
        return PeriodResolution(ACCEPTED, point.value, point.value, point.label)

    return PeriodResolution(NOT_FOUND)


def _subtract_month(d: date) -> date:
    month = d.month - 1 or 12
    year = d.year - 1 if d.month == 1 else d.year
    last_day = calendar.monthrange(year, month)[1]
    return date(year, month, min(d.day, last_day))
