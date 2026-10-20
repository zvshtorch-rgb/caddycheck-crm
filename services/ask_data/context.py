"""Conversation context for Ask Data v3 (chat UI, Phase 4): what carries over between turns.

A ConversationContext holds only already-RESOLVED, canonical values (a country code, a canonical project
name, a year, a category, a paid/unpaid scope, a compare_country, the previous result_id, and any pending
clarification) -- never raw text, never amounts, never ledger rows. This is exactly what is safe to show the
planner as "previous turn" context (see planner.py) and to merge into the next tool call.

Follow-up rules implemented here:
  * only filters the CURRENT tool actually accepts are inherited (`relevant_for`);
  * a value the new turn resolves explicitly always overrides the inherited one (`merge`);
  * an explicit "start over" phrase clears everything (`reset`);
  * a pending clarification can be answered by a short reply that matches exactly one of the offered
    candidates (`resolve_clarification_reply`), reusing the same resolver used for a first-time question.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

from services.ask_data.helpers import _normalize_country
from services.ask_data.resolver import ACCEPTED, resolve_country, resolve_project
from services.ask_data.tool_specs import TOOL_SPECS

# Conceptual filter slots carried across turns (goal 4): resolved project, country (+ compare_country for
# debt_by_country), year, Y1/Y2+/trial category, paid_year, paid/unpaid/cancelled scope, and (Ask Data
# Phase 5) the resolved historical period/at_date so "only Belgium?" after a historical answer keeps the
# same date range. `network` is intentionally NOT carried yet (a pre-existing, documented gap predating
# Phase 5 -- the Network Resolver phase only added network to the tools themselves, not to follow-up
# context; out of scope here).
CARRYABLE_FILTER_KEYS = ("project", "country", "compare_country", "year", "paid_year", "category", "paid",
                        "period", "at_date", "date_a", "date_b")

# Phrases that explicitly start a new question / clear all carried-over filters (goal 10: "clearing context").
_RESET_PHRASES_EN = (
    "start over", "new question", "forget that", "never mind", "nevermind", "reset", "clear context",
    "overall", "all countries", "ignore the above", "forget the above",
)
_RESET_PHRASES_HE = ("בוא נתחיל מחדש", "נתחיל מחדש", "שכח מזה", "תשכח מזה", "איפוס", "הכל", "כל המדינות")


def is_reset_phrase(question: str) -> bool:
    text = (question or "").strip().lower()
    if not text:
        return False
    return any(phrase in text for phrase in _RESET_PHRASES_EN) or any(phrase in (question or "") for phrase in _RESET_PHRASES_HE)


@dataclass
class Clarification:
    """A pending ambiguous-argument question from `validate_tool_call`, kept verbatim plus its origin tool call
    so a short follow-up reply can resubmit the SAME call with just that one argument filled in."""
    tool: str
    arg: str
    kind: str                 # "project" or "country"
    value: str
    candidates: tuple
    raw_arguments: dict

    @classmethod
    def from_validation_dict(cls, data: dict) -> "Clarification":
        return cls(
            tool=data["tool"], arg=data["arg"], kind=data["kind"], value=data["value"],
            candidates=tuple(data["candidates"]), raw_arguments=dict(data["raw_arguments"]),
        )


def resolve_clarification_reply(clarification: Clarification, reply_text: str) -> Optional[dict]:
    """If `reply_text` unambiguously picks ONE of the pending candidates, return a ready-to-validate
    {"tool", "arguments"} payload with that candidate plugged in; otherwise None (not a clarification answer --
    treat the reply as a brand new question instead)."""
    if clarification.kind == "country":
        resolution = resolve_country(reply_text, clarification.candidates, _normalize_country)
    else:
        resolution = resolve_project(reply_text, clarification.candidates)
    if resolution.status != ACCEPTED or resolution.value not in clarification.candidates:
        return None
    new_arguments = dict(clarification.raw_arguments)
    new_arguments[clarification.arg] = resolution.value
    return {"tool": clarification.tool, "arguments": new_arguments}


@dataclass
class ConversationContext:
    """Per-session Ask Data v3 state. Create one per chat session (e.g. in st.session_state); never share
    across sessions/users (see ResultStore, which has the same per-session rule)."""
    intent: Optional[str] = None                  # the last tool actually executed
    filters: dict = field(default_factory=dict)    # CARRYABLE_FILTER_KEYS subset -> resolved canonical value
    previous_result_id: Optional[str] = None
    clarification: Optional[Clarification] = None
    turn_count: int = 0

    def relevant_for(self, tool_name: str) -> dict:
        """Previously resolved filters this tool actually declares as arguments."""
        spec = TOOL_SPECS.get(tool_name, {})
        return {key: value for key, value in self.filters.items() if key in spec}

    def merge(self, tool_name: str, new_args: dict) -> dict:
        """Full argument dict for `tool_name`: inherited filters, then every non-None value from `new_args`
        overrides (an explicit answer always wins over a carried-over one -- "allow the user to override one
        filter naturally", "never carry filters that conflict with the new question")."""
        merged = self.relevant_for(tool_name)
        for key, value in new_args.items():
            if value is not None:
                merged[key] = value
        for key in TOOL_SPECS.get(tool_name, {}):
            merged.setdefault(key, None)
        return merged

    def planner_hint(self) -> Optional[dict]:
        """What is safe to tell the planner about the previous turn: tool name + already-resolved filter
        values only (country codes, canonical project names, years, categories, resolved dates/periods as
        plain ISO strings/labels) -- never amounts or rows. Internal storage (`self.filters`) keeps the
        native date/PeriodResolution objects so `merge()` can hand them straight to the executor; only this
        JSON-safe copy goes to the model."""
        if self.intent is None:
            return None
        safe_filters = {}
        for key, value in self.filters.items():
            if hasattr(value, "isoformat"):
                safe_filters[key] = value.isoformat()
            else:
                safe_filters[key] = str(value) if hasattr(value, "date_from") else value
        return {"tool": self.intent, "arguments": safe_filters}

    def record(self, tool_name: str, final_args: dict, result_id: Optional[str]) -> None:
        """Called after a tool actually executed successfully: remember the carryable filters it used."""
        self.intent = tool_name
        for key in CARRYABLE_FILTER_KEYS:
            if key in final_args and final_args[key] is not None:
                self.filters[key] = final_args[key]
        if result_id:
            self.previous_result_id = result_id
        self.clarification = None
        self.turn_count += 1

    def set_clarification(self, clarification_dict: dict) -> None:
        self.clarification = Clarification.from_validation_dict(clarification_dict)

    def reset(self) -> None:
        self.intent = None
        self.filters = {}
        self.previous_result_id = None
        self.clarification = None
        self.turn_count = 0
