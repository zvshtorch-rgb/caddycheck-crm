"""Ask Data v2 planner: turns a free-text question into ONE structured, read-only tool call via Gemini.

Moved verbatim from streamlit_app.py (Ask Data Phase 1). Gemini receives only the prompt below
(the question + the fixed tool catalog): never invoice/project/customer data.
"""
from typing import Optional

from config.settings import get_gemini_config
from services.ask_data.tool_specs import _ask_data_tool_catalog_text


def _llm_parse_data_question(question: str, context: Optional[dict] = None) -> Optional[dict]:
    """
    Ask Gemini to translate a free-text question into ONE structured,
    read-only tool call: {"tool": "<name>", "arguments": {...}}.

    Gemini only ever sees this prompt (the question + a fixed tool catalog, plus -- for a follow-up turn --
    the previous turn's tool name and already-resolved filter values from `context`) — never any invoice/
    project/customer data, never an amount, never a ledger row. It only ever picks a tool name + arguments.
    It cannot write SQL, cannot write Python, cannot select a write/update/delete action (none exist in the
    catalog), and never computes any number itself; `_execute_ask_data_tool` does that deterministically in
    Python. Returns None on any failure (missing key, network error, timeout, invalid JSON), so the caller
    falls back to the legacy `_answer_data_question`.

    `context` is `None` for a standalone question (byte-identical prompt to before Phase 4) or
    `{"tool": <previous tool name>, "arguments": {<already-resolved filter> -> <value>, ...}}` for a
    follow-up turn (see services/ask_data/context.py's `ConversationContext.planner_hint`).
    """
    gemini_cfg = get_gemini_config()
    api_key = gemini_cfg.get("api_key", "")
    if not api_key or not str(question).strip():
        return None
    try:
        import json
        import requests

        prompt = (
            "You are a strict, read-only query router for a CRM. Given the user's question, choose "
            "exactly ONE tool from the catalog below plus the arguments it needs, then reply with ONLY "
            "raw JSON (no markdown fences, no explanation) in exactly this shape:\n"
            '{"tool": "<tool_name>", "arguments": {"<arg_name>": <value>, ...}}\n\n'
            "Rules:\n"
            "- Only use tool names and argument names that appear in the catalog below. Never invent one.\n"
            "- Omit any argument you are not confident about; it will default to 'no filter'.\n"
            "- Never produce SQL, Python, or any instruction to write/update/delete/send anything — "
            "every tool in the catalog is read-only, so only ever select one of them.\n"
            "- 'year', 'invoice_number', 'min_cameras' and 'limit' must be plain integers.\n"
            "- 'year' filters an invoice's billing/maintenance year (when it was invoiced FOR). "
            "'paid_year' filters by when the invoice was ACTUALLY PAID (its payment date). For "
            "questions like 'paid in 2026', 'payments received in 2026', or 'bank transactions in "
            "2026', use 'paid_year', not 'year' — an invoice can be billed for one year but paid "
            "in a later one.\n"
            "- 'country' should be the country name/code exactly as it appears in the question.\n"
            "- 'project' should be the project name exactly as it appears in the question.\n"
            "- 'network' is for a GENERIC supermarket chain/group reference (e.g. 'Edeka', 'AD', "
            "'Delhaize', 'AD Delhaize', 'Proxy', 'Intermarché', 'Carrefour Market'/'CM', 'Albert "
            "Heijn'/'AH', 'Plus', 'Jumbo') when the question does not name one specific store/branch. "
            "Use 'network' for a bare chain name ('How much does Edeka owe?', 'AD debt', 'Edeka "
            "cameras', 'Proxy projects') and 'project' only when a specific store/branch is also named "
            "('AD Anderlecht', 'Edeka Preller-Breckerfeld'). Only set 'network' for a tool that "
            "declares it in the catalog below; for other tools, omit it rather than putting a chain "
            "name into 'project'.\n"
            "- For get_projects/count_projects/get_camera_statistics, omitting 'status' means 'active "
            "projects only' (the CRM's default reporting scope). Only set status to 'All' if the user "
            "explicitly asks to include cancelled/offline/inactive projects too.\n"
            "- Distinguish PAYMENT-EVENT/HISTORY questions from FINANCIAL-STATUS questions — this is "
            "the single most important routing decision:\n"
            "  * PAYMENT-EVENT/HISTORY: the question asks about a specific bank transfer/wire/"
            "remittance as an EVENT — WHEN it happened, or WHAT it covered/included. Trigger words: "
            "'bank payment', 'bank transfer', 'wire', 'remittance', 'last payment'/'last bank payment' "
            "used to ask for a date or contents, 'payment history'. Examples: 'when was the last bank "
            "payment for X', 'when did X last pay us', 'what was included/covered in the last payment/"
            "transfer for X', 'show the bank payment history for X', 'show the latest bank transfer "
            "allocations'. Use get_last_bank_payment for ALL of these, even if the wording also "
            "contains 'paid'/'pay'. Its underlying bank_payments/bank_payment_allocations audit log "
            "may be INCOMPLETE (not every paid invoice has a logged event) — NEVER use it to answer "
            "whether a project/invoice is paid, unpaid, fully paid, partially paid, settled, or has "
            "outstanding debt/balance.\n"
            "  * FINANCIAL-STATUS: the question asks about the current paid/unpaid state, balance, "
            "or debt, regardless of individual payment events. Trigger words: 'paid', 'unpaid', 'fully "
            "paid', 'partially paid', 'settled', 'balance', 'outstanding', 'debt', 'owe'/'owes', "
            "'payment status'. Examples: 'is X fully paid', 'has X paid us', 'does X owe us money', "
            "'what is the outstanding amount for X', 'which invoices/projects are unpaid or not fully "
            "paid', 'is invoice N paid', 'how much is still outstanding across all projects'. Use "
            "get_invoice/get_invoices/get_unpaid_invoices/get_debt/get_paid_amount/"
            "get_outstanding_amount/get_top_debt_projects/get_license_payment_status for these — the "
            "invoices table is the authoritative source of paid/unpaid/outstanding/debt status, never "
            "the bank-payment log.\n"
            "  * Tie-break rule when a question uses 'paid'/'pay' AND could go either way: if it asks "
            "WHEN something happened or WHAT a transfer contained, it's a payment-EVENT question "
            "(e.g. 'when did X last pay us' -> get_last_bank_payment, because it wants a date). If it "
            "asks whether something IS paid/settled or what balance/debt remains, it's a FINANCIAL-"
            "STATUS question (e.g. 'has X paid us' -> get_paid_amount/get_invoices, because it wants "
            "a yes/no payment-status answer, not a specific transfer's date or contents).\n"
            "- For questions about license expiry/status (e.g. 'expired licenses', 'licenses needing "
            "update', 'license for AD Anderlecht', 'licenses expiring in October'), use get_licenses. "
            "'eop_year'/'eop_month' filter by the License EOP date, not an invoice year. "
            "'license_status' is one of: Active, Expired, Missing, Update Next Month, Cancelled — it "
            "is the COMPUTED license classification (already accounts for whether the EOP date has "
            "passed). It is NOT the same field as 'project_status' (the project's own operational "
            "status: Active, Offline, Cancelled, New, Installed, etc. — same values as get_projects' "
            "'status'). Do not confuse the two: a project can be license_status='Expired' while its "
            "project_status is still 'Active' (license needs renewal but the site is still running) — "
            "that is a normal, common, and important combination, NOT a contradiction. If a question "
            "names both concepts (e.g. 'expired licenses on Active projects', 'licenses that expired "
            "but the project is still active/offline'), set license_status='Expired' AND set "
            "project_status to the mentioned operational status — never drop 'expired' just because "
            "another status word like 'Active' also appears in the question. "
            "If the user gives an exact day (e.g. 'on October 1st 2026', '2026-10-01'), use 'eop_date' "
            "(format YYYY-MM-DD) instead of eop_year/eop_month, so only that exact date matches.\n"
            "- If the question about licenses ALSO asks about invoices, payments, debt, outstanding "
            "balance, or whether projects are paid/fully paid, use get_license_payment_status instead "
            "of get_licenses — it takes the exact same arguments but also resolves, per project, the "
            "LATEST already-due annual license/maintenance invoice (based on the project's recurring "
            "Annual Payment Month, not merely its most recent or highest-numbered invoice) and whether "
            "THAT specific invoice is paid: Paid, Unpaid, Not Yet Due, or No Annual Invoice Found. "
            "All-time historical totals are included for reference only.\n"
            "- License EOP reconciliation: if the question asks whether a paid year (Year 1 / Year 2 / "
            "Y1 / Y2 ...) was reflected in the License EOP — e.g. 'projects with Year 1 paid but License "
            "EOP not updated', 'Year 2 paid but EOP not updated', 'paid licenses where the current "
            "License EOP is earlier than the EOP justified by payments', 'expected license EOP', 'which "
            "licenses need their EOP updated' — use get_license_eop_reconciliation (NOT "
            "get_license_payment_status, which only sees the latest annual invoice). Year 'N' paid but "
            "EOP not updated -> paid_year=N. 'EOP earlier than justified by payments' / 'needs EOP "
            "update' with no specific year -> needs_update='yes'. Add project/country/project_status "
            "only if the question names them. This tool is read-only.\n\n"
            f"Tool catalog:\n{_ask_data_tool_catalog_text()}\n\n"
            + ("" if context is None else (
                "This is a FOLLOW-UP to a previous turn. The previous turn resolved to this tool call "
                f"(tool and already-resolved filter values only -- no amounts, no rows):\n{json.dumps(context)}\n\n"
                "If the new question refines, narrows, or changes just one or two of those filters (e.g. "
                "'only Belgium?', 'and Germany?', 'what about Y2+?'), reuse the previous tool and keep any "
                "filter the new question does not mention or contradict; only change what the new question "
                "actually asks to change. If the new question is unrelated, ignore the previous turn "
                "entirely and answer it fresh.\n\n"
            ))
            + f"Question: {question}"
        )
        model = gemini_cfg.get("model", "gemini-flash-latest")
        request_body = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": 0,
                "maxOutputTokens": 2048,
                "thinkingConfig": {"thinkingBudget": 0},
                "responseMimeType": "application/json",
            },
        }
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
        headers = {"x-goog-api-key": api_key}
        resp = None
        last_exc: Optional[Exception] = None
        for attempt in range(2):  # transient 5xx from Gemini is common; retry once
            try:
                resp = requests.post(url, headers=headers, json=request_body, timeout=20)
                if resp.status_code >= 500:
                    last_exc = ValueError(f"{resp.status_code} Server Error from Gemini")
                    continue
                break
            except requests.exceptions.RequestException as exc:
                last_exc = exc
        if resp is None or resp.status_code >= 500:
            raise last_exc or ValueError("Gemini request failed with no response")
        resp.raise_for_status()
        payload = resp.json()
        text = " ".join(
            part.get("text", "")
            for part in payload["candidates"][0]["content"]["parts"]
        ).strip()
        if not text:
            return None
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError as exc:
            finish_reason = payload["candidates"][0].get("finishReason", "?")
            usage = payload.get("usageMetadata", {})
            raise ValueError(
                f"Gemini returned invalid/truncated JSON ({len(text)} chars, "
                f"finishReason={finish_reason}, usage={usage}): {text[:200]!r}"
            ) from exc
        return parsed if isinstance(parsed, dict) else None
    except Exception as exc:
        try:
            import re as _re
            import streamlit as st
            safe_msg = _re.sub(r"key=[^&\s]+", "key=***", str(exc))
            safe_msg = safe_msg.replace(api_key, "***")
            st.session_state["_ask_data_llm_error"] = safe_msg
        except Exception:
            pass
        return None


# Public alias for later phases.
parse_question = _llm_parse_data_question
