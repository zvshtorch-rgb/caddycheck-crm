"""Ask Data: deterministic, read-only natural-language data assistant for the CRM.

Phase 1 is a pure extraction of the existing "Ask Data v2" tool-selection logic out of
streamlit_app.py with no behavior change. Submodules:

  helpers      shared pure helpers (moved verbatim)
  tool_specs   the read-only tool catalog (specs + descriptions + prompt catalog text)
  validation   argument resolution / validation of an LLM tool call
  executor     deterministic execution of one validated tool call
  planner      the Gemini call that turns a question into ONE tool call (question + tool catalog only)
  router       planner -> validator -> executor, with an injected legacy fallback

Nothing here writes to any data store.
"""
