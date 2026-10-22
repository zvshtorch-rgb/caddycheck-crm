import datetime as dt
import unittest
from unittest.mock import patch

from services import billing_mapping_review as review
from streamlit.testing.v1 import AppTest


ACCOUNT = {
    "id": 1, "account_code": "IRETAILCHECK_BE", "active": False,
    "legal_name": "Iretailcheck", "address_line_1": "Street", "city": "City",
    "country_code": "BE", "vat_number": "BE123", "recipient_config_key": "email",
}


def fixture():
    ids = [i for evidence in review.EVIDENCE_GROUPS.values() for group in evidence.values() for i in group]
    projects = [dict(id=i, project_name=f"AD Site {i}", status="Active", country="Bel",
                     payment_month="July", license_eop="2027-01-01") for i in ids]
    invoices = [
        dict(id=1, project_name="AD Site 43", invoice_number="8681", year=2026,
             maintenance_year="Y3", description="Maintenance Monthly-July-2026"),
        dict(id=2, project_name="AD Site 43", invoice_number="8687", year=2026,
             maintenance_year="Y2", invoice_type="Complementary"),
    ]
    return review.build_review(projects, invoices, [ACCOUNT], [])


class TestBillingMappingReview(unittest.TestCase):
    def setUp(self):
        self.data = fixture()
        self.row = self.data["groups"]["Historically evidenced"][1]
        self.args = dict(decision="Approve for future billing", start=dt.date(2027, 7, 1),
                         end=dt.date(2028, 7, 1), open_ended=False, actor="Zvika",
                         approved_at=dt.datetime.now(dt.timezone.utc), source_reference="Explicit authorization")

    def validate(self, **changes):
        return review.validate_mapping_draft(self.row, ACCOUNT, [], **(self.args | changes))

    def test_cohort_counts_unique_and_latest_recurring_not_adjustment(self):
        self.assertEqual([len(g) for g in self.data["groups"].values()], [14, 41, 10])
        ids = [r["project_id"] for g in self.data["groups"].values() for r in g]
        self.assertEqual(len(ids), len(set(ids)))
        aalter = self.data["groups"]["Historically evidenced"][0]
        self.assertEqual(aalter["latest_maintenance_year"], "Y2")
        self.assertTrue(aalter["latest_recurring_invoice"].startswith("8681"))
        self.assertEqual(aalter["mapping_status"], "Unassigned")
        self.assertTrue(aalter["approval_blocked"])

    def test_approval_requires_explicit_dates_identity_timestamp_and_source(self):
        self.assertEqual(self.validate(), [])
        for change in ({"start": None}, {"end": None}, {"actor": ""},
                       {"approved_at": None}, {"source_reference": ""},
                       {"end": dt.date(2027, 7, 1)}, {"open_ended": True}):
            with self.subTest(change=change):
                self.assertTrue(self.validate(**change))
        self.assertEqual(self.validate(end=None, open_ended=True), [])

    def test_aalter_blocked_and_existing_overlap_rejected(self):
        self.row = self.data["groups"]["Historically evidenced"][0]
        self.assertTrue(any("AD Aalter" in e for e in self.validate()))
        self.row = self.data["groups"]["Historically evidenced"][1]
        mappings = [dict(project_id=self.row["project_id"], effective_from="2027-01-01", effective_to=None)]
        errors = review.validate_mapping_draft(self.row, ACCOUNT, mappings, **self.args)
        self.assertTrue(any("overlap" in e for e in errors))

    def test_nonapproval_decisions_do_not_require_dates(self):
        for decision in ("Historical only", "Needs confirmation", "Reject mapping"):
            self.assertEqual(self.validate(decision=decision, start=None, end=None), [])

    def test_adjacent_intervals_allowed_and_malformed_dates_rejected(self):
        mappings = [dict(project_id=self.row["project_id"], effective_from="2026-07-01",
                         effective_to="2027-07-01")]
        self.assertEqual(review.validate_mapping_draft(self.row, ACCOUNT, mappings, **self.args), [])
        self.assertTrue(self.validate(start="2027-07-01"))
        self.assertTrue(self.validate(start=dt.datetime(2027, 7, 1)))

    def test_missing_evidence_project_surfaces_explicit_error(self):
        with self.assertRaisesRegex(ValueError, "Evidence project 43 is missing"):
            review.build_review([], [], [ACCOUNT], [])

    def test_verification_requires_all_confirmations_and_reference(self):
        confirmed = dict.fromkeys(("legal_name", "address", "vat", "recipients"), True)
        self.assertEqual(review.validate_account_draft(ACCOUNT, confirmed, "Evidence", "Zvika"), [])
        for key in confirmed:
            self.assertTrue(review.validate_account_draft(ACCOUNT, confirmed | {key: False}, "Evidence", "Zvika"))
        self.assertTrue(review.validate_account_draft(ACCOUNT, confirmed, "", "Zvika"))

    def app(self, role="admin"):
        loader = patch.object(review, "load_review", return_value=self.data)
        load = loader.start()
        self.addCleanup(loader.stop)
        config = patch("ui.billing_mapping_review.get_email_config",
                       return_value={"default_recipients": ["billing@example.com"], "default_cc": []})
        config.start()
        self.addCleanup(config.stop)
        script = (
            "from ui.billing_mapping_review import render_billing_mapping_review\n"
            f"render_billing_mapping_review({role!r}, 'Zvika')"
        )
        app = AppTest.from_string(script).run()
        return app, load

    def test_admin_loads_all_sections_without_write_client(self):
        with patch("services.supabase_service._get_client", side_effect=AssertionError("No write client")):
            app, load = self.app()
            load.assert_not_called()
            app.button(key="load_billto_review").click().run()
            self.assertFalse(app.exception, [e.value for e in app.exception])
            self.assertEqual(len(app.dataframe), 3)
            self.assertTrue(any("inconsistent" in w.value for w in app.warning))
            self.assertTrue(all(d.value is None for d in app.date_input))
            self.assertFalse(any(c.value for c in app.checkbox))
            load.assert_called_once()

    def test_nonadmin_cannot_load_review(self):
        for role in ("viewer", "editor"):
            app, load = self.app(role)
            self.assertEqual(len(app.button), 0)
            load.assert_not_called()

    def test_invalid_submit_reports_errors_and_creates_no_draft(self):
        app, _ = self.app()
        app.button(key="load_billto_review").click().run()
        app.selectbox[0].set_value(self.row).run()
        next(s for s in app.selectbox if s.label == "Admin decision").set_value("Approve for future billing")
        next(b for b in app.button if b.label == "Validate decision draft (do not save)").click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertTrue(any("effective_from" in e.value for e in app.error), [e.value for e in app.error])
        self.assertNotIn("_billto_mapping_draft", app.session_state)

    def test_valid_mapping_and_activation_are_session_drafts_only(self):
        with patch("services.supabase_service._get_client", side_effect=AssertionError("No persistence")):
            app, _ = self.app()
            app.button(key="load_billto_review").click().run()
            app.selectbox[0].set_value(self.row).run()
            next(s for s in app.selectbox if s.label == "Admin decision").set_value("Approve for future billing")
            app.date_input[0].set_value(dt.date(2027, 7, 1))
            app.date_input[1].set_value(dt.date(2028, 7, 1))
            app.text_input[0].set_value("Zvika")
            next(c for c in app.checkbox if c.label.startswith("Record my approval")).check()
            next(t for t in app.text_area if t.label.startswith("Source reference")).set_value("Business authorization")
            next(b for b in app.button if b.label == "Validate decision draft (do not save)").click().run()
            self.assertFalse(app.exception, [e.value for e in app.exception])
            self.assertEqual(app.session_state["_billto_mapping_draft"]["effective_to"], "2028-07-01")
            for c in app.checkbox:
                if c.label.startswith("I confirm"):
                    c.check()
            next(t for t in app.text_area if t.label == "Verification reference").set_value("Verified legal documents")
            next(b for b in app.button if b.label == "Verify and activate billing account").click().run()
            self.assertFalse(app.exception, [e.value for e in app.exception])
            self.assertTrue(app.session_state["_billto_account_draft"]["requested_active"])
            self.assertFalse(self.data["accounts"][0]["active"])
            self.assertEqual(self.data["mappings"], [])
            app.selectbox[0].set_value(self.data["groups"]["Historically evidenced"][2]).run()
            self.assertFalse(any("Decision draft validated" in i.value for i in app.info))
            self.assertTrue(all(d.value is None for d in app.date_input))

    def test_draft_approver_must_match_authenticated_admin(self):
        app, _ = self.app()
        app.button(key="load_billto_review").click().run()
        app.text_input[0].set_value("Someone else")
        next(c for c in app.checkbox if c.label.startswith("Record my approval")).check()
        next(t for t in app.text_area if t.label.startswith("Source reference")).set_value("Evidence")
        next(b for b in app.button if b.label == "Validate decision draft (do not save)").click().run()
        self.assertTrue(any("signed-in admin" in e.value for e in app.error))
        self.assertNotIn("_billto_mapping_draft", app.session_state)
