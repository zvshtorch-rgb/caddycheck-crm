"""Ask Data v2 router: planner -> validator -> executor, with an injected legacy fallback.

Moved from streamlit_app.py (Ask Data Phase 1). The original `_answer_data_question_smart` called the
legacy keyword matcher (`_llm_normalize_question` + `_answer_data_question`), which still lives in
streamlit_app.py; those two callables are now passed in explicitly so this module has no dependency on
the app. Control flow, session-state keys and error strings are unchanged.
"""
from typing import Optional

import pandas as pd

from services.ask_data.executor import _execute_ask_data_tool
from services.ask_data.planner import _llm_parse_data_question
from services.ask_data.validation import clarification_question, validate_tool_call


def _answer_data_question_smart(
    question: str,
    projects,
    invoices,
    debt_summaries,
    *,
    legacy_normalize_question,
    legacy_answer,
    next_invoice_number=None,
) -> tuple[str, Optional[pd.DataFrame]]:
    """
    Preferred Ask Data entry point: Gemini tool-selection first, deterministic
    Python execution second, with the original keyword-based
    `_answer_data_question` (which itself still tries an LLM template-rewrite)
    as the fallback for any failure along the way.

    Phase 2: when a project/country argument is ambiguous the router asks ONE short clarification question
    (stored in session state under `_ask_data_clarification`) instead of guessing or falling back.
    """
    import streamlit as st

    st.session_state.pop("_ask_data_llm_error", None)
    st.session_state.pop("_ask_data_llm_tool_call", None)
    st.session_state.pop("_ask_data_clarification", None)

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
                return _execute_ask_data_tool(
                    tool_name, cleaned_args, projects, invoices, debt_summaries,
                    next_invoice_number=next_invoice_number,
                )
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
