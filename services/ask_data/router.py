"""Ask Data v2 router: planner -> validator -> executor, with an injected legacy fallback.

Moved from streamlit_app.py (Ask Data Phase 1). The original `_answer_data_question_smart` called the
legacy keyword matcher (`_llm_normalize_question` + `_answer_data_question`), which still lives in
streamlit_app.py; those two callables are now passed in explicitly so this module has no dependency on
the app. Control flow, session-state keys and error strings are unchanged.
"""
from typing import Optional

import pandas as pd

from services.ask_data.executor import execute_tool
from services.ask_data.planner import _llm_parse_data_question
from services.ask_data.tool_result import ResultStore
from services.ask_data.validation import clarification_question, validate_tool_call

UNVERIFIED_TEXT = "I could not verify the figures in this answer against the ledger rows, so I am not showing a number. See the table below."


def _answer_data_question_smart(
    question: str,
    projects,
    invoices,
    debt_summaries,
    *,
    legacy_normalize_question,
    legacy_answer,
    next_invoice_number=None,
    result_store: Optional[ResultStore] = None,
) -> tuple[str, Optional[pd.DataFrame]]:
    """
    Preferred Ask Data entry point: Gemini tool-selection first, deterministic
    Python execution second, with the original keyword-based
    `_answer_data_question` (which itself still tries an LLM template-rewrite)
    as the fallback for any failure along the way.

    Phase 2: when a project/country argument is ambiguous the router asks ONE short clarification question
    (stored in session state under `_ask_data_clarification`) instead of guessing or falling back.

    Phase 3: tools return a ToolResult. The router keeps it in the (per-session) result store so the rows behind
    a figure can be retrieved later (`ResultStore.get_result_rows`), publishes it as `_ask_data_last_result`, and
    re-checks that every euro amount in the displayed text comes from the ToolResult before showing it.
    """
    import streamlit as st

    st.session_state.pop("_ask_data_llm_error", None)
    st.session_state.pop("_ask_data_llm_tool_call", None)
    st.session_state.pop("_ask_data_clarification", None)
    st.session_state.pop("_ask_data_last_result", None)
    st.session_state.pop("_ask_data_grounding_error", None)

    parsed = _llm_parse_data_question(question)
    if parsed is not None:
        validation = validate_tool_call(parsed, projects)
        if validation.clarification is not None:
            st.session_state["_ask_data_clarification"] = validation.clarification
            return clarification_question(validation.clarification), None
        tool_name, cleaned_args, error = validation.tool, validation.arguments, validation.error
        if tool_name is not None:
            try:
                st.session_state["_ask_data_llm_tool_call"] = {"tool": tool_name, "arguments": cleaned_args}
                result = execute_tool(
                    tool_name, cleaned_args, projects, invoices, debt_summaries,
                    next_invoice_number=next_invoice_number,
                )
                store = result_store if result_store is not None else st.session_state.setdefault("_ask_data_result_store", ResultStore())
                store.add(result)
                st.session_state["_ask_data_last_result"] = result
                text = result.display_text()
                if result.financial:
                    ungrounded = result.ungrounded_amounts(text)
                    if ungrounded:
                        st.session_state["_ask_data_grounding_error"] = f"Ungrounded amounts in answer: {ungrounded}"
                        return UNVERIFIED_TEXT, result.dataframe
                return text, result.dataframe
            except Exception as exc:
                st.session_state["_ask_data_llm_error"] = f"Tool execution failed: {exc}"
        else:
            st.session_state["_ask_data_llm_error"] = error

    # Fallback: legacy keyword matcher (with its own internal LLM template-rewrite attempt).
    normalized = legacy_normalize_question(question)
    effective_question = normalized or question
    return legacy_answer(effective_question, projects, invoices, debt_summaries)


# Public alias for later phases.
answer_smart = _answer_data_question_smart
