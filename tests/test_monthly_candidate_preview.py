import copy
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from streamlit.testing.v1 import AppTest
from services import monthly_candidate_preview as preview


def inputs():
    return dict(
        projects=[{"id": 93, "project_name": "AD Aartselar", "installation_year": 2021,
                   "payment_month": "July", "status": "Active", "num_cameras": 10,
                   "rate_y2_override": None, "m7y": None}],
        accounts=[{"id": 1, "account_code": "IRETAILCHECK_BE", "legal_name": "Iretailcheck",
                   "address_line_1": "Bijkhoevelaan 11", "address_line_2": None,
                   "postal_code": "2110", "city": "WIJNEGEM", "country_code": "BE",
                   "vat_number": "BE 0537.905.085", "recipient_config_key": "email",
                   "active": True, "verified_by": "Zvika", "verified_at": "2026-10-22T19:00:00Z",
                   "verification_reference": "Approved"}],
        mappings=[{"id": 1, "project_id": 93, "billing_account_id": 1,
                   "effective_from": "2027-07-01", "effective_to": None,
                   "approved_by": "Zvika", "approved_at": "2026-10-22T19:00:00Z",
                   "source_reference": "Approved future billing"}],
        invoices=[{"id": 3054, "invoice_number": "8681", "project_name": "AD Aartselar",
                   "maintenance_year": "Y6", "year": 2026, "description": "Maintenance Monthly-July-2026"}],
        sent_history=[{"id": 5, "invoice_number": 8681, "month": "July", "year": 2026}],
        overrides={}, email_config={"default_recipients": ["billing@example.com"], "default_cc": []},
    )


class TestMonthlyCandidatePreview(unittest.TestCase):
    def setUp(self):
        self.data = inputs()

    def build(self):
        return preview.build_candidate(**self.data)

    def assert_blocked(self):
        result = self.build()
        self.assertEqual(result["validation"]["state"], "blocked", result)
        return result

    def test_exact_canonical_payload_and_inputs_unchanged(self):
        before = copy.deepcopy(self.data)
        result = self.build()
        self.assertEqual(result["validation"], {"state": "eligible_preview", "errors": []})
        expected = {
            "project_id": 93, "project_name": "AD Aartselar", "billing_account_id": 1,
            "billing_period": "2027-07", "planned_send_date": "2027-06-01",
            "planned_invoice_date": "2027-06-01", "maintenance_year": "Y7",
            "camera_quantity": 10, "unit_price": "228.00", "total_amount": "2280.00",
            "currency": "EUR", "pricing_source": "config.settings.RATE_Y2_PLUS_PER_CAM",
        }
        for key, value in expected.items():
            self.assertEqual(result[key], value, key)
        self.assertEqual(result["billing_account"]["legal_name"], "Iretailcheck")
        self.assertEqual(result["network"], "AD")
        self.assertEqual(self.data, before)
        self.assertNotIn("invoice_number", result)
        self.assertNotIn("smtp_password", result)

    def test_override_precedence_and_decimal_total(self):
        self.data["projects"][0]["rate_y2_override"] = 240.25
        self.assertEqual(self.build()["total_amount"], "2402.50")
        self.assertEqual(self.build()["pricing_source"], "projects.rate_y2_override")
        self.data["overrides"] = {"ad aartselar": {"y2_rate": 230.15}}
        self.assertEqual(self.build()["total_amount"], "2301.50")
        self.assertEqual(self.build()["pricing_source"], "project_overrides.y2_rate")
        self.data["overrides"]["ad aartselar"]["y2_rate"] = None
        self.assertEqual(self.build()["total_amount"], "2280.00")

    def test_missing_ambiguous_partial_and_boundary_mapping(self):
        for mappings in ([], [self.data["mappings"][0]] * 2,
                         [self.data["mappings"][0] | {"effective_from": "2027-07-02"}],
                         [self.data["mappings"][0] | {"effective_to": "2027-07-15"}],
                         [self.data["mappings"][0] | {"effective_to": "2027-07-01"}],
                         [self.data["mappings"][0] | {"approved_by": ""}]):
            with self.subTest(mappings=mappings):
                original = self.data["mappings"]
                self.data["mappings"] = mappings
                self.assert_blocked()
                self.data["mappings"] = original
        self.data["mappings"][0]["effective_to"] = "2027-08-01"
        self.assertEqual(self.build()["validation"]["state"], "eligible_preview")

    def test_account_must_be_active_verified_and_complete(self):
        for key, value in (("active", False), ("verified_by", None), ("verified_at", None),
                           ("verification_reference", ""), ("vat_number", None),
                           ("recipient_config_key", "unknown")):
            with self.subTest(key=key):
                original = self.data["accounts"][0][key]
                self.data["accounts"][0][key] = value
                self.assert_blocked()
                self.data["accounts"][0][key] = original

    def test_unresolved_quantity_year_and_prices(self):
        for key, values in (("num_cameras", (None, 0, -1, 1.5, "NaN")),
                            ("installation_year", (None, "invalid", 2027, 2028)),
                            ("rate_y2_override", ("invalid", -1, 0, "NaN", 12.345))):
            for value in values:
                with self.subTest(key=key, value=value):
                    original = self.data["projects"][0].get(key)
                    self.data["projects"][0][key] = value
                    self.assert_blocked()
                    self.data["projects"][0][key] = original

    def test_bad_recipient_types_empty_and_invalid_addresses(self):
        for config in ({}, {"default_recipients": []}, {"default_recipients": "a@example.com"},
                       {"default_recipients": ["not-an-address"]},
                       {"default_recipients": ["a@example.com"], "default_cc": ["bad\ncc"]}):
            with self.subTest(config=config):
                self.data["email_config"] = config
                self.assert_blocked()

    def test_ledger_period_maintenance_and_ambiguous_references_block(self):
        for changes in (
                {"description": "Maintenance Monthly-July-2027"},
                {"year": 2027, "for_month": "July", "description": None},
                {"year": 2027, "for_month": "2027-07", "description": None},
                {"year": 2027, "description": None},
                {"year": None, "maintenance_year": "Y7", "description": None},
                {"year": 2028, "maintenance_year": "Y7", "description": None}):
            with self.subTest(changes=changes):
                original = self.data["invoices"]
                self.data["invoices"] = [original[0] | changes]
                self.assert_blocked()
                self.data["invoices"] = original
        self.data["projects"][0]["m7y"] = "12345"
        self.assert_blocked()

    def test_sent_without_ledger_blocks_or_is_ambiguous(self):
        for row in ({"year": 2027, "month": "July"}, {"year": 2027}, {"month": "July"}, {}):
            with self.subTest(row=row):
                self.data["sent_history"] = [row]
                self.assert_blocked()

    def test_unrelated_project_ledger_does_not_match_fuzzily(self):
        self.data["projects"].append({"id": 94, "project_name": "AD Other"})
        self.data["invoices"].append({"id": 9, "project_name": "AD Other", "year": 2027,
                                     "description": "Maintenance Monthly-July-2027"})
        self.assertEqual(self.build()["validation"]["state"], "eligible_preview")

    def test_unknown_target_period_ledger_identity_is_ambiguous(self):
        self.data["invoices"].append({"id": 8, "project_name": "Unknown misspelling",
                                     "year": 2027, "description": "Maintenance Monthly-July-2027"})
        self.assert_blocked()

    def test_duplicate_project_name_is_blocked(self):
        self.data["projects"].append(self.data["projects"][0] | {"id": 99})
        self.assert_blocked()

    def test_select_only_loader_and_strict_local_history(self):
        tables = {
            "projects": self.data["projects"], "billing_accounts": self.data["accounts"],
            "project_billing_accounts": self.data["mappings"], "invoices": self.data["invoices"],
            "sent_invoices": self.data["sent_history"],
        }
        with tempfile.TemporaryDirectory() as root:
            history = Path(root) / "history.json"
            history.write_text('[{"year":2027,"month":"July","invoice_number":999}]', encoding="utf-8")
            with patch.object(preview, "_select_all_rows", side_effect=lambda table: copy.deepcopy(tables[table])) as select, \
                    patch.object(preview.settings, "OVERRIDES_FILE", Path(root) / "overrides.json"), \
                    patch.object(preview.settings, "SENT_INVOICES_LOG_FILE", history), \
                    patch.object(preview.settings, "get_email_config", return_value=self.data["email_config"]):
                result = preview.load_candidate()
                self.assertEqual(result["validation"]["state"], "blocked")
                self.assertEqual({call.args[0] for call in select.call_args_list}, set(tables))
                history.write_text("invalid json", encoding="utf-8")
                with self.assertRaises(ValueError):
                    preview.load_candidate()

    def test_admin_ui_is_explicitly_loaded_and_nonadmin_has_no_access(self):
        with patch("ui.monthly_candidate_preview.load_candidate", return_value=self.build()) as load:
            app = AppTest.from_string(
                "from ui.monthly_candidate_preview import render_monthly_candidate_preview\n"
                "render_monthly_candidate_preview('admin')"
            ).run()
            load.assert_not_called()
            app.button(key="load_monthly_candidate_preview").click().run()
            self.assertFalse(app.exception)
            self.assertEqual(len(app.json), 1)
            self.assertEqual(len(app.button), 1)
            self.assertEqual(app.session_state["_monthly_candidate_preview"]["total_amount"], "2280.00")
            nonadmin = AppTest.from_string(
                "from ui.monthly_candidate_preview import render_monthly_candidate_preview\n"
                "render_monthly_candidate_preview('viewer')"
            ).run()
            self.assertEqual(len(nonadmin.button), 0)
            load.assert_called_once()

    def test_ui_refresh_failure_clears_old_eligibility(self):
        with patch("ui.monthly_candidate_preview.load_candidate", side_effect=[self.build(), RuntimeError("Read failed")]):
            app = AppTest.from_string(
                "from ui.monthly_candidate_preview import render_monthly_candidate_preview\n"
                "render_monthly_candidate_preview('admin')"
            ).run()
            app.button[0].click().run()
            app.button[0].click().run()
            self.assertEqual(len(app.json), 0)
            self.assertIn("eligibility was not established", app.error[0].value)
