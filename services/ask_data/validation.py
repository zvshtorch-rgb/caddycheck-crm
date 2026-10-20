"""Ask Data v2 argument resolution and validation of a planner-produced tool call.

Moved verbatim from streamlit_app.py (Ask Data Phase 1).
"""
import datetime
from typing import Optional

from config.settings import MONTH_ORDER
from services.ask_data.helpers import (
    PROJECT_STATUS_OPTIONS,
    _match_countries_from_question,
    _normalize_country,
    _safe_str,
    _suggest_best_project_match,
)
from services.ask_data.tool_specs import _ASK_DATA_TOOL_SPECS


def _resolve_project_arg(value: object, projects) -> Optional[str]:
    """Fuzzy-resolve a free-text project name to a real project name (reuses `_suggest_best_project_match`)."""
    text = _safe_str(value).strip()
    if not text:
        return None
    project_names = sorted({p.project_name for p in projects if p.project_name})
    best_name, score = _suggest_best_project_match(text, project_names)
    return best_name if best_name and score >= 0.55 else None


def _resolve_single_country(value: object, projects) -> Optional[str]:
    """Resolve a free-text country name/code (typo-tolerant) to a real project country code."""
    text = _safe_str(value).strip()
    if not text:
        return None
    known_codes = sorted({p.country for p in projects if p.country})
    text_lower = text.lower()
    for code in known_codes:
        if code.lower() == text_lower or _normalize_country(code).lower() == text_lower:
            return code
    matches = _match_countries_from_question(text, projects)
    return matches[0] if matches else None


def _resolve_ask_data_arg(kind: str, value: object, projects) -> tuple[bool, object]:
    """Validate + resolve one tool argument. Returns (ok, resolved_value)."""
    if value is None:
        return True, None
    if kind == "int":
        try:
            return True, int(value)
        except Exception:
            return False, None
    if kind == "date":
        text = _safe_str(value).strip()
        try:
            return True, datetime.date.fromisoformat(text[:10])
        except Exception:
            return False, None
    if kind == "project":
        resolved = _resolve_project_arg(value, projects)
        return (True, resolved) if resolved else (False, None)
    if kind == "country":
        resolved = _resolve_single_country(value, projects)
        return (True, resolved) if resolved else (False, None)
    if kind == "status":
        text = _safe_str(value).strip()
        if text.lower() == "all":
            return True, "All"
        for option in PROJECT_STATUS_OPTIONS:
            if option.lower() == text.lower():
                return True, option
        return False, None
    if kind == "month":
        text = _safe_str(value).strip().lower()
        for month in MONTH_ORDER:
            if month.lower() == text or month[:3].lower() == text:
                return True, month
        return False, None
    if kind.startswith("enum:"):
        allowed = kind.split(":", 1)[1].split(",")
        text = _safe_str(value).strip()
        for option in allowed:
            if option.lower() == text.lower():
                return True, option
        return False, None
    return False, None


def _validate_ask_data_tool_call(parsed: object, projects) -> tuple[Optional[str], dict, Optional[str]]:
    """
    Validate a parsed {"tool": ..., "arguments": {...}} payload against
    `_ASK_DATA_TOOL_SPECS`. Rejects unknown tools, unresolvable/mistyped
    arguments, and missing required arguments. Unknown argument keys are
    silently ignored. Returns (tool_name, cleaned_arguments, error) — tool_name
    is None on any validation failure.
    """
    if not isinstance(parsed, dict):
        return None, {}, "LLM response was not a JSON object."
    tool_name = _safe_str(parsed.get("tool")).strip()
    if tool_name not in _ASK_DATA_TOOL_SPECS:
        return None, {}, f"Unknown or missing tool name: {tool_name!r}."
    raw_args = parsed.get("arguments")
    if raw_args is None:
        raw_args = {}
    if not isinstance(raw_args, dict):
        return None, {}, "Tool 'arguments' was not a JSON object."

    spec = _ASK_DATA_TOOL_SPECS[tool_name]
    cleaned: dict = {}
    for arg_name, (kind, required) in spec.items():
        raw_value = raw_args.get(arg_name)
        ok, resolved = _resolve_ask_data_arg(kind, raw_value, projects)
        if not ok:
            return None, {}, f"Could not validate argument '{arg_name}'={raw_value!r} for tool '{tool_name}'."
        if required and resolved is None:
            return None, {}, f"Missing required argument '{arg_name}' for tool '{tool_name}'."
        cleaned[arg_name] = resolved
    # Any keys in raw_args beyond the spec are simply ignored (never executed).
    return tool_name, cleaned, None


# Public aliases for later phases.
validate_tool_call = _validate_ask_data_tool_call
