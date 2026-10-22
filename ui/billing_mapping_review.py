"""Admin-only session drafts. Deliberately contains no database write/activation path."""
import datetime as dt

import streamlit as st

from config.settings import get_email_config
from services import billing_mapping_review as review


def render_billing_mapping_review(role, actor):
    if role != "admin":
        return
    st.subheader("Bill-to Mapping Review")
    st.info("Review and validation only. Decisions are session drafts, not saved mappings. "
            "Account verification does not activate the account in this phase. "
            "No candidates, invoices, reservations, or emails are created.")
    if st.button("Load / refresh bill-to review", key="load_billto_review"):
        st.session_state.pop("_billto_review", None)
        st.session_state.pop("_billto_mapping_draft", None)
        st.session_state.pop("_billto_account_draft", None)
        try:
            st.session_state["_billto_review"] = review.load_review()
        except Exception as exc:
            st.error(f"Could not load bill-to review: {exc}")
    data = st.session_state.get("_billto_review")
    if not data:
        return
    all_rows = []
    for group, rows in data["groups"].items():
        st.markdown(f"### {group} ({len(rows)})")
        st.dataframe([{k: v for k, v in row.items() if k != "approval_blocked"} for row in rows],
                     hide_index=True, width="stretch")
        all_rows.extend(rows)
    st.caption("Latest ledger ordering uses year then invoice number, not a proven issue timestamp. "
               "Operational status/license expiry do not establish future payer authorization.")
    selected = st.selectbox("Project to review", all_rows,
                            format_func=lambda r: f"{r['project_id']} - {r['project_name']}",
                            key="billto_project")
    if selected["approval_blocked"]:
        st.warning("AD Aalter: maintenance-year evidence is inconsistent (Y3 recurring invoice 8681 "
                   "versus Y2 adjustment 8687). Future approval is blocked until reconciled.")
    accounts = data["accounts"]
    if not accounts:
        st.error("No billing accounts exist. Mapping review cannot be validated.")
        return
    account = st.selectbox("Billing account", accounts,
                           format_func=lambda a: f"{a['account_code']} (active={a['active']})")
    if not account["active"]:
        st.warning("This account is inactive. A validated session draft does not activate it "
                   "or authorize future billing.")
    key = f"mapping_review_{selected['project_id']}"
    with st.form(key):
        decision = st.selectbox("Admin decision", review.DECISIONS)
        start = st.date_input("Proposed effective_from", value=None)
        end = st.date_input("Proposed effective_to", value=None)
        open_ended = st.checkbox("I explicitly authorize an open-ended effective period.", value=False)
        identity = st.text_input("approved_by (must match signed-in admin)", value="")
        timestamp_confirmed = st.checkbox("Record my approval time in UTC when I validate this draft.")
        source = st.text_area("Source reference / business justification", value="")
        submitted = st.form_submit_button("Validate decision draft (do not save)")
    if submitted:
        st.session_state.pop("_billto_mapping_draft", None)
        timestamp = dt.datetime.now(dt.timezone.utc) if timestamp_confirmed else None
        errors = review.validate_mapping_draft(
            selected, account, data["mappings"], decision=decision, start=start, end=end,
            open_ended=open_ended, actor=identity, approved_at=timestamp, source_reference=source,
        )
        if identity.strip() != actor.strip():
            errors.append("approved_by must match the signed-in admin.")
        for error in errors:
            st.error(error)
        if not errors:
            st.session_state["_billto_mapping_draft"] = {
                "project_id": selected["project_id"], "billing_account_id": account["id"],
                "decision": decision, "effective_from": start.isoformat() if start else None,
                "effective_to": end.isoformat() if end else None,
                "open_ended_authorized": open_ended, "approved_by": identity.strip(),
                "approved_at": timestamp.isoformat(), "source_reference": source.strip(),
            }
    if (st.session_state.get("_billto_mapping_draft", {}).get("project_id") == selected["project_id"]
            and st.session_state["_billto_mapping_draft"]["billing_account_id"] == account["id"]):
        st.info("Decision draft validated for review only. No mapping was saved; this does not authorize sending.")
        st.json(st.session_state["_billto_mapping_draft"])
    with st.expander("Verify and activate billing account"):
        st.json(account)
        cfg = {}
        if account["recipient_config_key"] == "email":
            cfg = get_email_config()
            st.write({"recipient_config_key": "email", "To": cfg.get("default_recipients", []),
                      "CC": cfg.get("default_cc", [])})
        else:
            st.warning("Recipient configuration key is not supported by this review UI.")
        with st.form(f"verify_billto_{account['id']}"):
            confirmations = {
                "legal_name": st.checkbox("I confirm the legal name."),
                "address": st.checkbox("I confirm the full billing address and country."),
                "vat": st.checkbox("I confirm the VAT number."),
                "recipients": st.checkbox("I confirm the recipient configuration and resolved recipients."),
            }
            reference = st.text_area("Verification reference")
            verify = st.form_submit_button("Verify and activate billing account", disabled=False)
        if verify:
            st.session_state.pop("_billto_account_draft", None)
            errors = review.validate_account_draft(account, confirmations, reference, actor)
            if account["recipient_config_key"] != "email" or not cfg.get("default_recipients"):
                errors.append("A supported, nonempty recipient configuration is required.")
            for error in errors:
                st.error(error)
            if not errors:
                st.session_state["_billto_account_draft"] = {
                    "billing_account_id": account["id"], "verified_by": actor,
                    "verified_at": dt.datetime.now(dt.timezone.utc).isoformat(),
                    "verification_reference": reference.strip(), "requested_active": True,
                }
        if st.session_state.get("_billto_account_draft", {}).get("billing_account_id") == account["id"]:
            st.info("Verification/activation draft validated only. Account remains unchanged in the database.")
            st.json(st.session_state["_billto_account_draft"])
