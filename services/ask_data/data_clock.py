"""Data freshness for Ask Data: when was the CRM data behind an answer loaded?

`load_data()` in streamlit_app.py is cached (ttl 300s). Its body runs only on a cache miss, so it calls
`mark_loaded()` there; every answer then reads `loaded_at()`. This lives in a module (not in the script) because the
Streamlit script is re-executed on every interaction while imported modules persist. `load_data()`'s return tuple is
unchanged. If nothing has been recorded the timestamp is None and answers say so instead of inventing one.
"""
from __future__ import annotations

import datetime
from typing import Optional

_state: dict = {"loaded_at": None, "source": None}


def mark_loaded(source: Optional[str] = None, at: Optional[datetime.datetime] = None) -> datetime.datetime:
    stamp = at or datetime.datetime.now(datetime.timezone.utc)
    _state["loaded_at"] = stamp
    _state["source"] = source
    return stamp


def loaded_at() -> Optional[datetime.datetime]:
    return _state["loaded_at"]


def source() -> Optional[str]:
    return _state["source"]


def format_as_of(stamp: Optional[datetime.datetime]) -> str:
    if stamp is None:
        return "Data timestamp unavailable"
    if stamp.tzinfo is None:
        stamp = stamp.replace(tzinfo=datetime.timezone.utc)
    return "Data as of " + stamp.astimezone(datetime.timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
