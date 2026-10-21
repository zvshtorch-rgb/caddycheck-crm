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
from services.ask_data.history_dates import (
    ACCEPTED as DATE_PHRASE_ACCEPTED,
    resolve_period_phrase,
    resolve_single_date_phrase,
)
from services.ask_data.resolver import (
    ACCEPTED,
    AMBIGUOUS,
    resolve_country,
    resolve_month,
    resolve_network,
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
    """One short question; candidates are real project/country/network names, never invented."""
    kind = clarification["kind"] if clarification["kind"] in ("country", "network") else "project"
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


def _resolve_single_network(value: object) -> Optional[str]:
    """Resolve a free-text network/chain reference (e.g. "Edeka", "Delhaize") to a canonical network name;
    EXACT alias match only -- never fuzzy (goal 2)."""
    text = _safe_str(value).strip()
    if not text:
        return None
    resolution = resolve_network(text)
    return resolution.value if resolution.status == ACCEPTED else None


def _redirect_network_text(raw_args: dict, spec: dict) -> dict:
    """Goal 2: a bare network name ("Edeka", "AD", "Delhaize", ...) must never fuzzy-collapse onto one
    specific project, regardless of which argument key the planner used for it (it will often put a bare
    chain name into 'project', since that is the only one it knew about before this tool supported
    'network'). For any tool that declares a 'network' argument, if the text supplied for 'network' or
    'project' is an EXACT network alias, treat it as the network filter and clear 'project' so normal
    (possibly fuzzy) project resolution never runs on that text."""
    if "network" not in spec:
        return raw_args
    raw_args = dict(raw_args)
    for key in ("network", "project"):
        text = _safe_str(raw_args.get(key)).strip()
        if not text:
            continue
        resolution = resolve_network(text)
        if resolution.status == ACCEPTED:
            raw_args["network"] = resolution.value
            if key != "network":
                raw_args[key] = None
            return raw_args
    return raw_args


def _resolve_arg_ex(kind: str, value: object, projects, *, carried_over: object = None) -> tuple[bool, object, Optional[dict]]:
    """Validate + resolve one tool argument. Returns (ok, resolved_value, clarification).

    `carried_over` is this same argument's already-resolved value from the previous turn's
    ConversationContext (Phase 5 at_date/period/date_a/date_b only; see `_resolve_ask_data_arg`'s
    "echoed carry-over" handling below), or None for a standalone question / a non-historical argument.
    """
    if value is None:
        return True, None, None
    if kind in ("project", "country", "network"):
        text = _safe_str(value).strip()
        if kind == "project":
            resolution = _project_resolution(text, projects)
        elif kind == "country":
            resolution = _country_resolution(text, projects)
        else:
            resolution = resolve_network(text)
        if text and resolution.status == ACCEPTED:
            return True, resolution.value, None
        if text and resolution.status == AMBIGUOUS:
            return False, None, {"kind": kind, "value": text, "candidates": list(resolution.candidates)}
        return False, None, None
    ok, resolved = _resolve_ask_data_arg(kind, value, projects, carried_over=carried_over)
    return ok, resolved, None


def _resolve_ask_data_arg(kind: str, value: object, projects, *, carried_over: object = None) -> tuple[bool, object]:
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
    if kind == "at_date":
        # Ask Data Phase 5: `value` is the planner's VERBATIM date phrase copied from the question (e.g.
        # "today", "30 September", "a week ago") -- never a date the model computed itself. The actual
        # calendar math happens here, deterministically, anchored on the real wall-clock date.
        text = _safe_str(value).strip()
        if not text:
            return False, None
        resolution = resolve_single_date_phrase(text, datetime.date.today())
        if resolution.status == DATE_PHRASE_ACCEPTED:
            return True, resolution.value
        # Follow-up echo: ConversationContext.planner_hint() tells the model the previous turn's ALREADY-
        # RESOLVED at_date as a plain ISO string (e.g. "2026-10-20"), not a phrase the resolver above can
        # re-parse. When a follow-up ("Only Belgium?") reuses the same tool, the model commonly copies that
        # exact string straight back as its own 'at_date' argument (as instructed: "keep any filter the new
        # question does not mention"). Recognise this exact, provable echo (the resolved date's own
        # isoformat matches verbatim) and silently reuse the already-resolved date instead of failing the
        # whole tool call out to an unrelated current-state fallback -- the single most important bug this
        # guards against. A genuinely different (but unparseable) new phrase never matches here and still
        # fails validation normally.
        if carried_over is not None and hasattr(carried_over, "isoformat") and carried_over.isoformat() == text:
            return True, carried_over
        return False, None
    if kind == "period":
        # Same principle as "at_date", but resolves to a (date_from, date_to) range -- see
        # services/ask_data/history_dates.py's PeriodResolution.
        text = _safe_str(value).strip()
        if not text:
            return False, None
        resolution = resolve_period_phrase(text, datetime.date.today())
        if resolution.status == DATE_PHRASE_ACCEPTED:
            return True, resolution
        # Same follow-up echo recovery as "at_date" above, but the hint value is PeriodResolution.__str__()
        # (a human-readable label, e.g. "this month (2026-10-01 to 2026-10-21)"), not an ISO date.
        if carried_over is not None and str(carried_over) == text:
            return True, carried_over
        return False, None
    if kind == "project":
        resolved = _resolve_project_arg(value, projects)
        return (True, resolved) if resolved else (False, None)
    if kind == "country":
        resolved = _resolve_single_country(value, projects)
        return (True, resolved) if resolved else (False, None)
    if kind == "network":
        resolved = _resolve_single_network(value)
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


def validate_tool_call(parsed: object, projects, context=None) -> ValidationResult:
    """
    Validate a parsed {"tool": ..., "arguments": {...}} payload against
    `_ASK_DATA_TOOL_SPECS`. Rejects unknown tools, unresolvable/mistyped
    arguments, and missing required arguments. Unknown argument keys are
    silently ignored. An argument that matches several projects/countries (or only
    weakly) yields a `clarification` instead of a guess; nothing is ever executed then.

    `context` (optional, a ConversationContext) is used only to recognise a follow-up turn's exact echo of
    an already-resolved Phase 5 at_date/period/date_a/date_b -- see `_resolve_ask_data_arg`'s "follow-up
    echo" handling. It never widens acceptance for anything else.
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
    raw_args = _redirect_network_text(raw_args, spec)
    carried_filters = context.filters if context is not None else {}
    cleaned: dict = {}
    for arg_name, (kind, required) in spec.items():
        raw_value = raw_args.get(arg_name)
        carried_over = carried_filters.get(arg_name) if kind in ("at_date", "period") else None
        ok, resolved, clarification = _resolve_arg_ex(kind, raw_value, projects, carried_over=carried_over)
        if clarification is not None:
            clarification = {"arg": arg_name, "tool": tool_name, "raw_arguments": dict(raw_args), **clarification}
            return ValidationResult(None, {}, f"Ambiguous argument '{arg_name}'={raw_value!r} for tool '{tool_name}'.", clarification)
        if not ok:
            if kind in ("at_date", "period") and raw_value not in (None, ""):
                return ValidationResult(
                    None, {},
                    f"I couldn't determine which date(s) you mean by {raw_value!r} for '{tool_name}'. Try a "
                    "phrase like 'today', 'yesterday', 'this month', 'last month', 'since October 1', "
                    "'past 30 days', or an explicit date like '30 September'.",
                )
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
