"""Admin-only, explicitly loaded, session-only recurring candidate preview."""
import streamlit as st

from services.monthly_candidate_preview import load_candidate


def render_monthly_candidate_preview(role):
    if role != "admin":
        return
    st.subheader("Monthly candidate preview - AD Aartselar / July 2027")
    st.info("Read-only annual-maintenance preview for project 93 only. No candidate row, "
            "invoice number, ledger invoice, PDF, or email is created. "
            "Legacy monthly invoice actions below are separate and are not enabled by this preview.")
    if st.button("Load / refresh July 2027 candidate preview", key="load_monthly_candidate_preview"):
        st.session_state.pop("_monthly_candidate_preview", None)
        try:
            st.session_state["_monthly_candidate_preview"] = load_candidate()
        except Exception as exc:
            st.error(f"Candidate preview failed; eligibility was not established: {exc}")
    candidate = st.session_state.get("_monthly_candidate_preview")
    if candidate is None:
        return
    validation = candidate["validation"]
    if validation["state"] == "blocked":
        for error in validation["errors"]:
            st.error(error)
    else:
        st.info("Eligible for preview only - not approval or permission to send.")
    st.json(candidate)
    st.caption("EUR decimal strings are the canonical unit price and total. Future PDF, ledger, "
               "and email integration must consume this payload, not recalculate its pricing. "
               "Eligibility is a point-in-time read and must be revalidated before any future write.")
