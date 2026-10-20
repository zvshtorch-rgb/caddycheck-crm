"""Ask Data v3: conversational turn orchestration (chat UI, Phase 4).

answer_turn() is the v3 counterpart of router._answer_data_question_smart(): it adds a per-session
ConversationContext (follow-ups, drill-down, one short clarification instead of guessing) on top of the exact
same planner -> validator -> executor -> ToolResult pipeline used by v2. v2 itself (router.py) is completely
untouched; v3 is additive and only reachable behind the `ask_data_v3` feature flag (see streamlit_app.py).

Guarantees preserved from v2:
  * the model only ever sees the question, the fixed tool catalog, and -- for a follow-up -- the previous
    turn's tool name and already-resolved filter VALUES (country codes, canonical project names, years,
    categories): never a row, never an amount;
  * every financial answer is re-checked against its ToolResult before being shown (`UNVERIFIED_TEXT`);
  * "show invoices" / "show breakdown" never call the model at all -- they are answered purely from the
    (per-session) ResultStore, deterministically.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

import pandas as pd

from services.ask_data.context import ConversationContext, is_reset_phrase, resolve_clarification_reply
from services.ask_data.executor import execute_tool
from services.ask_data.planner import _llm_parse_data_question
from services.ask_data.tool_result import ResultStore, ToolResult, UnknownResultError
from services.ask_data.validation import clarification_question, validate_tool_call

UNVERIFIED_TEXT = (
    "I could not verify the figures in this answer against the ledger rows, so I am not showing a number. See the table below."
)
NO_PREVIOUS_RESULT_TEXT = "There's no previous answer to show rows for yet -- ask a question first."
HISTORICAL_LIMITATION_TEXT = (
    "I don't yet have a complete historical snapshot for that date, so I can't reliably calculate the change. "
    "I can instead show you the payments received since then, or the current total debt -- just ask for either "
    "of those directly."
)

# Free-text drill-down phrases (goal 6 + goal 2's "Which invoices make that up?" example). Matched as a
# substring so this NEVER calls the model: it is answered purely from the stored ToolResult.
_ROWS_PHRASES_EN = (
    "show invoices", "show the invoices", "show rows", "show the rows", "which invoices", "what invoices",
    "show me the invoices", "invoices behind", "rows behind", "see the invoices", "list the invoices",
)
_ROWS_PHRASES_HE = ("הצג חשבוניות", "אילו חשבוניות", "תראה חשבוניות", "הצג את החשבוניות")
_BREAKDOWN_PHRASES_EN = ("show breakdown", "show the breakdown", "show components")
_BREAKDOWN_PHRASES_HE = ("הצג פירוט", "פירוט", "פילוח")

# Historical/change-over-time questions (e.g. "What changed since September 1?", "How has debt changed this
# month?", "What was the debt last week?"). Phase 5 (a daily history-snapshot table) is not
# implemented yet, so there is no reliable prior-point-in-time data to answer these from -- routing them to
# an unrelated current-state tool (the production bug this guards against) would silently return a wrong,
# unrelated answer. Detected deterministically (no model call) as CHANGE-word + TIME-word, so a plain
# current-state question that merely happens to mention a month/week ("invoices from last month") does NOT
# false-positive: it needs a word that implies comparison/change too.
_HISTORICAL_CHANGE_WORDS_EN = (
    "what changed", "what's changed", "what has changed", "how has", "how much has", "what was the",
    "what were the", "has changed", "have changed", "changed since", "change since", "compared to last",
    "versus last", "vs last", "used to be",
)
_HISTORICAL_TIME_WORDS_EN = (
    "since", "last week", "last month", "yesterday", "this month", "this week", " ago", "previously", "before today",
)
_HISTORICAL_CHANGE_WORDS_HE = ("השתנה", "השתנו", "שינוי", "היה", "היו", "היתה")
_HISTORICAL_TIME_WORDS_HE = ("מאז", "לפני", "אתמול", "שבוע", "חודש")


def _matches_any(text: str, phrases_en: tuple, phrases_he: tuple) -> bool:
    lowered = (text or "").strip().lower()
    return any(p in lowered for p in phrases_en) or any(p in (text or "") for p in phrases_he)


def is_historical_question(question: str) -> bool:
    """True for a historical/change-over-time question (goal 1) -- needs both a CHANGE word and a TIME
    word, so plain current-state questions that merely mention a time period never false-positive."""
    lowered = (question or "").strip().lower()
    raw = question or ""
    has_change = any(w in lowered for w in _HISTORICAL_CHANGE_WORDS_EN) or any(w in raw for w in _HISTORICAL_CHANGE_WORDS_HE)
    has_time = any(w in lowered for w in _HISTORICAL_TIME_WORDS_EN) or any(w in raw for w in _HISTORICAL_TIME_WORDS_HE)
    return has_change and has_time


def detect_drilldown_kind(question: str) -> Optional[str]:
    """"rows", "breakdown", or None (not a drill-down request). Breakdown phrases are checked first since
    "show breakdown" would otherwise also match a looser "show ..." rows phrase."""
    if _matches_any(question, _BREAKDOWN_PHRASES_EN, _BREAKDOWN_PHRASES_HE):
        return "breakdown"
    if _matches_any(question, _ROWS_PHRASES_EN, _ROWS_PHRASES_HE):
        return "rows"
    return None


def rows_table_for_result(result: ToolResult, store: ResultStore) -> tuple[str, Optional[pd.DataFrame]]:
    """"Show invoices": the ledger rows behind `result`'s figure, via the result store (drill-down, read-only)."""
    rows = store.get_result_rows(result.result_id)
    if rows.row_count == 0:
        return "There are no rows behind that figure.", None
    return f"Here are the {rows.row_count} row(s) behind that figure, totalling €{rows.total:,.2f}.", rows.dataframe


def breakdown_table_for_result(result: ToolResult) -> tuple[str, Optional[pd.DataFrame]]:
    """"Show breakdown": the named components of `result` (e.g. countries, or charges/credits) as a table."""
    if not result.components:
        return "That answer has no further breakdown; use 'show invoices' for its rows instead.", None
    frame = pd.DataFrame([
        {"Component": c.label, "Rows": c.row_count, "Amount (€)": c.amount} for c in result.components
    ])
    return f"Here is the breakdown behind that figure ({len(result.components)} component(s)).", frame


@dataclass
class TurnOutcome:
    """What happened this turn (for the chat UI and for tests); never contains amounts beyond what's already
    in `text`/`dataframe` (both of which are the normal, grounded ToolResult display)."""
    text: str
    dataframe: Optional[pd.DataFrame]
    kind: str                       # "tool" | "drilldown" | "clarification" | "legacy" | "historical_unsupported"
    tool: Optional[str] = None
    result: Optional[ToolResult] = None


def _session(result_store: Optional[ResultStore]):
    import streamlit as st
    store = result_store if result_store is not None else st.session_state.setdefault("_ask_data_v3_result_store", ResultStore())
    return st, store


def _execute_and_ground(tool_name, final_args, projects, invoices, debt_summaries, next_invoice_number, st, store):
    st.session_state["_ask_data_v3_tool_call"] = {"tool": tool_name, "arguments": final_args}
    result = execute_tool(tool_name, final_args, projects, invoices, debt_summaries, next_invoice_number=next_invoice_number)
    store.add(result)
    st.session_state["_ask_data_v3_last_result"] = result
    text = result.display_text()
    if result.financial:
        ungrounded = result.ungrounded_amounts(text)
        if ungrounded:
            st.session_state["_ask_data_v3_grounding_error"] = f"Ungrounded amounts in answer: {ungrounded}"
            return UNVERIFIED_TEXT, result.dataframe, result
    return text, result.dataframe, result


def answer_turn(
    question: str,
    projects,
    invoices,
    debt_summaries,
    context: ConversationContext,
    *,
    legacy_normalize_question,
    legacy_answer,
    next_invoice_number=None,
    result_store: Optional[ResultStore] = None,
) -> TurnOutcome:
    """Process one chat turn against `context` (mutated in place) and return a `TurnOutcome`."""
    st, store = _session(result_store)
    st.session_state.pop("_ask_data_v3_llm_error", None)
    st.session_state.pop("_ask_data_v3_tool_call", None)
    st.session_state.pop("_ask_data_v3_grounding_error", None)

    if is_reset_phrase(question):
        context.reset()

    drilldown_kind = detect_drilldown_kind(question)
    if drilldown_kind is not None:
        if context.previous_result_id is None:
            return TurnOutcome(NO_PREVIOUS_RESULT_TEXT, None, "drilldown")
        previous = store.get(context.previous_result_id)
        if previous is None:
            return TurnOutcome(NO_PREVIOUS_RESULT_TEXT, None, "drilldown")
        if drilldown_kind == "breakdown":
            text, df = breakdown_table_for_result(previous)
        else:
            text, df = rows_table_for_result(previous, store)
        return TurnOutcome(text, df, "drilldown", tool=previous.tool, result=previous)

    if is_historical_question(question):
        # No Phase 5 history source yet -- never let the planner silently pick an unrelated current-state
        # tool for a trend/change-over-time question (the production bug this guards against). A stale
        # clarification (if any) is abandoned, since this is clearly a different, unrelated question.
        context.clarification = None
        return TurnOutcome(HISTORICAL_LIMITATION_TEXT, None, "historical_unsupported")

    if context.clarification is not None:
        pending = context.clarification
        answered = resolve_clarification_reply(pending, question)
        if answered is not None:
            validation = validate_tool_call(answered, projects)
            if validation.clarification is not None:
                context.set_clarification(validation.clarification)
                return TurnOutcome(clarification_question(validation.clarification), None, "clarification")
            if validation.tool is not None:
                final_args = context.merge(validation.tool, validation.arguments)
                try:
                    text, df, result = _execute_and_ground(
                        validation.tool, final_args, projects, invoices, debt_summaries, next_invoice_number, st, store)
                except Exception as exc:
                    st.session_state["_ask_data_v3_llm_error"] = f"Tool execution failed: {exc}"
                else:
                    context.record(validation.tool, final_args, result.result_id if result else None)
                    return TurnOutcome(text, df, "tool", tool=validation.tool, result=result)
            # the clarified call itself failed validation (rare) -- clear and fall through to a fresh question
        context.clarification = None

    parsed = _llm_parse_data_question(question, context=context.planner_hint())
    if parsed is not None:
        validation = validate_tool_call(parsed, projects)
        if validation.clarification is not None:
            context.set_clarification(validation.clarification)
            return TurnOutcome(clarification_question(validation.clarification), None, "clarification")
        if validation.tool is not None:
            final_args = context.merge(validation.tool, validation.arguments)
            try:
                text, df, result = _execute_and_ground(
                    validation.tool, final_args, projects, invoices, debt_summaries, next_invoice_number, st, store)
            except Exception as exc:
                st.session_state["_ask_data_v3_llm_error"] = f"Tool execution failed: {exc}"
            else:
                context.record(validation.tool, final_args, result.result_id if result else None)
                return TurnOutcome(text, df, "tool", tool=validation.tool, result=result)
        else:
            st.session_state["_ask_data_v3_llm_error"] = validation.error

    # Fallback: legacy keyword matcher (no conversation context -- a known Phase 4 limitation).
    normalized = legacy_normalize_question(question)
    effective_question = normalized or question
    text, df = legacy_answer(effective_question, projects, invoices, debt_summaries)
    return TurnOutcome(text, df, "legacy")
