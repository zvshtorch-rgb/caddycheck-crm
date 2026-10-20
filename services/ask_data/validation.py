"""Ask Data v2 argument resolution and validation of a planner-produced tool call.

Phase 1 moved this verbatim from streamlit_app.py. Phase 2 swaps project/country/month resolution for the
canonical resolver (services/ask_data/resolver.py): accents/case/punctuation/Hebrew are normalised, scores
>= 0.90 are accepted when unambiguous, 0.60-0.90 or several close matches produce a clarification question,
and < 0.60 is "not found" (the pre-existing validation error). Everything else is unchanged.
"""
import datetime
from dataclasses import dataclass
from typing import Optional

from config.settings import MONTH_ORDER
from services.ask_data.helpers import (
    PROJECT_STATUS_OPTIONS,
    _normalize_country,
    _safe_str,
)
from services.ask_data.resolver import (
    ACCEPTED,
    AMBIGUOUS,
    resolve_country,
    resolve_month,
    resolve_project,
)
from services.ask_data.tool_specs import _ASK_DATA_TOOL_SPECS


@dataclass(frozen=True)
class ValidationResult:
    tool: Optional[str]
    arguments: dict
    error: Optional[str]
    clarification: Optional[dict] = None  # {"arg", "kind", "value", "candidates", "tool", "raw_arguments"} when ambiguous


def clarification_question(clarification: dict) -> str:
    """One short question; candidates are real project/country names, never invented."""
    kind = "country" if clarification["kind"] == "country" else "project"
    options = list(clarification["candidates"])
    if not options:
        return f"I couldn't tell which {kind} you meant by '{clarification['value']}'. Which one did you mean?"
    if len(options) == 1:
        return f"Did you mean {options[0]} (for '{clarification['value']}')?"
    return f"Which {kind} did you mean by '{clarification['value']}': " + ", ".join(options[:-1]) + f" or {options[-1]}?"


def _project_resolution(value: object, projects):
    names = {p.project_name for p in projects if p.project_name}
    return resolve_project(_safe_str(value).strip(), names)


def _country_resolution(value: object, projects):
    known_codes = {p.country for p in projects if p.country}
    return resolve_country(_safe_str(value).strip(), known_codes, _normalize_country)


def _resolve_project_arg(value: object, projects) -> Optional[str]:
    """Resolve a free-text project name to a real project name; None unless accepted unambiguously."""
    text = _safe_str(value).strip()
    if not text:
        return None
    resolution = _project_resolution(text, projects)
    return resolution.value if resolution.status == ACCEPTED else None


def _resolve_single_country(value: object, projects) -> Optional[str]:
    """Resolve a free-text country name/code (typo-, accent- and Hebrew-tolerant) to a real project country code."""
    text = _safe_str(value).strip()
    if not text:
        return None
    resolution = _country_resolution(text, projects)
    return resolution.value if resolution.status == ACCEPTED else None


def _resolve_arg_ex(kind: str, value: object, projects) -> tuple[bool, object, Optional[dict]]:
    """Validate + resolve one tool argument. Returns (ok, resolved_value, clarification)."""
    if value is None:
        return True, None, None
    if kind in ("project", "country"):
        text = _safe_str(value).strip()
        resolution = _project_resolution(text, projects) if kind == "project" else _country_resolution(text, projects)
        if text and resolution.status == ACCEPTED:
            return True, resolution.value, None
        if text and resolution.status == AMBIGUOUS:
            return False, None, {"kind": kind, "value": text, "candidates": list(resolution.candidates)}
        return False, None, None
    ok, resolved = _resolve_ask_data_arg(kind, value, projects)
    return ok, resolved, None


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
        resolved = resolve_month(_safe_str(value).strip(), MONTH_ORDER)
        return (True, resolved) if resolved else (False, None)
    if kind.startswith("enum:"):
        allowed = kind.split(":", 1)[1].split(",")
        text = _safe_str(value).strip()
        for option in allowed:
            if option.lower() == text.lower():
                return True, option
        return False, None
    return False, None


def validate_tool_call(parsed: object, projects) -> ValidationResult:
    """
    Validate a parsed {"tool": ..., "arguments": {...}} payload against
    `_ASK_DATA_TOOL_SPECS`. Rejects unknown tools, unresolvable/mistyped
    arguments, and missing required arguments. Unknown argument keys are
    silently ignored. An argument that matches several projects/countries (or only
    weakly) yields a `clarification` instead of a guess; nothing is ever executed then.
    """
    if not isinstance(parsed, dict):
        return ValidationResult(None, {}, "LLM response was not a JSON object.")
    tool_name = _safe_str(parsed.get("tool")).strip()
    if tool_name not in _ASK_DATA_TOOL_SPECS:
        return ValidationResult(None, {}, f"Unknown or missing tool name: {tool_name!r}.")
    raw_args = parsed.get("arguments")
    if raw_args is None:
        raw_args = {}
    if not isinstance(raw_args, dict):
        return ValidationResult(None, {}, "Tool 'arguments' was not a JSON object.")

    spec = _ASK_DATA_TOOL_SPECS[tool_name]
    cleaned: dict = {}
    for arg_name, (kind, required) in spec.items():
        raw_value = raw_args.get(arg_name)
        ok, resolved, clarification = _resolve_arg_ex(kind, raw_value, projects)
        if clarification is not None:
            clarification = {"arg": arg_name, "tool": tool_name, "raw_arguments": dict(raw_args), **clarification}
            return ValidationResult(None, {}, f"Ambiguous argument '{arg_name}'={raw_value!r} for tool '{tool_name}'.", clarification)
        if not ok:
            return ValidationResult(None, {}, f"Could not validate argument '{arg_name}'={raw_value!r} for tool '{tool_name}'.")
        if required and resolved is None:
            return ValidationResult(None, {}, f"Missing required argument '{arg_name}' for tool '{tool_name}'.")
        cleaned[arg_name] = resolved
    # Any keys in raw_args beyond the spec are simply ignored (never executed).
    return ValidationResult(tool_name, cleaned, None)


def _validate_ask_data_tool_call(parsed: object, projects) -> tuple[Optional[str], dict, Optional[str]]:
    """Backward-compatible tuple form: (tool_name, cleaned_arguments, error); tool_name is None on any failure
    (including a needed clarification, reported via `error`)."""
    result = validate_tool_call(parsed, projects)
    return result.tool, result.arguments, result.error
