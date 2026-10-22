"""Tests for services/order_confirmation_service.py (Order Approvals -> Year-1 invoice + PDF +
confirmation-email extension, dry-run step).

This file covers only the dry-run preview building blocks (validation, amount computation, email
rendering, PDF bytes, Graph mailbox reachability, and the build_dry_run_preview orchestrator). No real
network calls are ever made; no real Supabase writes occur. The live orchestrator that actually creates
the invoice/PDF/email and marks the order approved is a deliberately separate, later change -- not
covered here.
"""
import os
import unittest
from unittest import mock

from services import order_confirmation_service as ocs

ORDER = {
    "id": "po-1",
    "project_name": "Edeka Steinebach Hohn",
    "customer_name": "",
    "customer_email": "tom.symons@iretailcheck.com",
    "order_reference": "",
    "ordered_cameras": 3,
    "amount": 2334.0,
}


class TestResolveOrderedCameras(unittest.TestCase):
    """purchase_orders has no ordered_cameras column -- the real count lives on the matching
    auto-provisioned Orders-page record (tagged via its 'notes' field with the exact purchase_order_id)."""

    def test_direct_value_wins_without_any_lookup(self):
        with mock.patch("services.supabase_service.load_orders") as load_orders:
            cameras = ocs.resolve_ordered_cameras({"id": "po-1", "ordered_cameras": 3})
        load_orders.assert_not_called()
        self.assertEqual(cameras, 3)

    def test_falls_back_to_the_matching_orders_table_record(self):
        rows = [{"notes": "Auto-created from approved purchase order po-1.", "ordered_cameras": 3}]
        with mock.patch("services.supabase_service.load_orders", return_value=rows):
            cameras = ocs.resolve_ordered_cameras({"id": "po-1", "ordered_cameras": 0})
        self.assertEqual(cameras, 3)

    def test_pending_order_resolves_twelve_camera_licenses_from_stored_pdf_text(self):
        order = dict(ORDER, ordered_cameras=0, incoming_email_id="incoming-1", amount=9336)
        text = (
            "Installation Type CaddyCheck 6 Checkouts TD + BT\n"
            "IRC-CC-SW-901 CaddyCheck Camera - starting license 12 EUR 550,00 EUR 6.600,00\n"
            "IRC-CC-SW-902 CaddyCheck - Maintenance & Licensing - 1Y 12 EUR 228,00 EUR 2.736,00"
        )
        with mock.patch("services.supabase_service.load_orders", return_value=[]), \
             mock.patch("services.order_approval_service.get_incoming_email",
                        return_value={"extracted_text": text}):
            cameras = ocs.resolve_ordered_cameras(order)
        self.assertEqual(cameras, 12, "count cameras, not checkouts or repeated maintenance units")

    def test_no_id_returns_zero_without_calling_load_orders(self):
        with mock.patch("services.supabase_service.load_orders") as load_orders:
            cameras = ocs.resolve_ordered_cameras({"ordered_cameras": 0})
        load_orders.assert_not_called()
        self.assertEqual(cameras, 0)

    def test_no_matching_orders_row_returns_zero(self):
        with mock.patch("services.supabase_service.load_orders", return_value=[{"notes": "unrelated"}]):
            cameras = ocs.resolve_ordered_cameras({"id": "po-1", "ordered_cameras": 0})
        self.assertEqual(cameras, 0)

    def test_load_orders_failure_is_not_silently_treated_as_missing_quantity(self):
        with mock.patch("services.supabase_service.load_orders", side_effect=RuntimeError("no creds")):
            with self.assertRaisesRegex(RuntimeError, "no creds"):
                ocs.resolve_ordered_cameras({"id": "po-1", "ordered_cameras": 0})

    def test_amount_alone_does_not_prove_camera_quantity(self):
        with mock.patch("services.supabase_service.load_orders", return_value=[]):
            self.assertEqual(ocs.resolve_ordered_cameras({"id": "po-1", "amount": 9336}), 0)

    def test_conflicting_pdf_quantities_are_reported(self):
        with mock.patch("services.supabase_service.load_orders", return_value=[]), \
             mock.patch("services.order_approval_service.get_incoming_email",
                        return_value={"extracted_text": "Cameras: 12\nCamera count: 6"}):
            with self.assertRaisesRegex(ValueError, "Conflicting camera quantities"):
                ocs.resolve_ordered_cameras({"id": "po-1", "incoming_email_id": "incoming-1"})


class TestComputeYear1Amount(unittest.TestCase):
    def test_uses_explicit_amount_when_positive(self):
        cameras, amount = ocs.compute_year1_amount({"ordered_cameras": 3, "amount": 2334.0})
        self.assertEqual(cameras, 3)
        self.assertEqual(amount, 2334.0)

    def test_falls_back_to_camera_rate_when_amount_missing(self):
        cameras, amount = ocs.compute_year1_amount({"ordered_cameras": 4, "amount": 0})
        self.assertEqual(cameras, 4)
        self.assertEqual(amount, 4 * 778.0)

    def test_zero_cameras_and_zero_amount_yields_zero(self):
        cameras, amount = ocs.compute_year1_amount({"ordered_cameras": 0, "amount": 0})
        self.assertEqual((cameras, amount), (0, 0.0))

    def test_accepts_payment_amount_key_too(self):
        cameras, amount = ocs.compute_year1_amount({"ordered_cameras": 2, "payment_amount": 1000.0})
        self.assertEqual((cameras, amount), (2, 1000.0))


class TestValidateOrderForYear1Invoicing(unittest.TestCase):
    def test_a_fully_populated_order_has_no_errors(self):
        with mock.patch.object(ocs, "resolve_project_country", return_value="Germany"):
            result = ocs.validate_order_for_year1_invoicing(ORDER)
        self.assertEqual(result["errors"], [])

    def test_unreliable_project_name_is_a_hard_error(self):
        bad_order = dict(ORDER, project_name="document iretailche 20260923 2026-0047-2")
        with mock.patch.object(ocs, "resolve_project_country", return_value=""):
            result = ocs.validate_order_for_year1_invoicing(bad_order)
        self.assertTrue(any("unreliable" in e for e in result["errors"]), result["errors"])

    def test_missing_camera_count_and_amount_is_a_hard_error(self):
        bad_order = dict(ORDER, ordered_cameras=0, amount=0)
        with mock.patch.object(ocs, "resolve_project_country", return_value=""), \
             mock.patch("services.supabase_service.load_orders", return_value=[]):
            result = ocs.validate_order_for_year1_invoicing(bad_order)
        self.assertTrue(any("camera count nor an explicit amount" in e for e in result["errors"]), result["errors"])

    def test_missing_country_is_a_soft_warning_not_an_error(self):
        with mock.patch.object(ocs, "resolve_project_country", return_value=""):
            result = ocs.validate_order_for_year1_invoicing(ORDER)
        self.assertEqual(result["errors"], [])
        self.assertTrue(any("Country" in w for w in result["warnings"]), result["warnings"])

    def test_missing_customer_details_is_a_soft_warning_not_an_error(self):
        bad_order = dict(ORDER, project_name="Unknown Site", customer_name="", customer_email="")
        with mock.patch.object(ocs, "resolve_project_country", return_value="Germany"):
            result = ocs.validate_order_for_year1_invoicing(bad_order)
        self.assertEqual(result["errors"], [])
        self.assertTrue(any("customer" in w.lower() for w in result["warnings"]), result["warnings"])


class TestRenderConfirmationEmail(unittest.TestCase):
    def test_email_address_can_never_render_as_customer_network(self):
        for value in ("Karen.Duhoux@profit-protection.com", "Karen <karen@example.com>"):
            with self.subTest(value=value):
                _, body = ocs.render_confirmation_email(
                    project_name="Unknown Site", customer_display=value, country="",
                    po_reference="PO-1", cameras=12, amount=9336, invoice_number=8719,
                    sender="sender@example.com",
                )
                self.assertNotIn("Customer/Network:", body)

    def test_subject_matches_the_required_format(self):
        subject, _ = ocs.render_confirmation_email(
            project_name="Edeka Steinebach Hohn", customer_display="iRetailCheck", country="Germany",
            po_reference="PO-301737ff", cameras=3, amount=2334.0, invoice_number=8719,
            sender="orders.chaddycheck@video-inform.com",
        )
        self.assertEqual(subject, "Order Confirmation \u2013 CaddyCheck \u2013 Edeka Steinebach Hohn \u2013 Invoice 8719")

    def test_body_includes_all_available_fields(self):
        _, body = ocs.render_confirmation_email(
            project_name="Edeka Steinebach Hohn", customer_display="iRetailCheck", country="Germany",
            po_reference="PO-301737ff", cameras=3, amount=2334.0, invoice_number=8719,
            sender="orders.chaddycheck@video-inform.com",
        )
        self.assertIn("Dear Tom and Karen,", body)
        self.assertIn("- Project/Site: Edeka Steinebach Hohn", body)
        self.assertIn("- Customer/Network: iRetailCheck", body)
        self.assertIn("- Country: Germany", body)
        self.assertIn("- Order / PO reference: PO-301737ff", body)
        self.assertIn("- Number of cameras: 3", body)
        self.assertIn("- Year-1 amount: \u20ac2,334.00", body)
        self.assertIn("- Proposed invoice number: 8719", body)
        self.assertIn("orders.chaddycheck@video-inform.com", body)

    def test_unavailable_fields_are_omitted_not_shown_blank(self):
        _, body = ocs.render_confirmation_email(
            project_name="Edeka Steinebach Hohn", customer_display="", country="", po_reference="",
            cameras=0, amount=2334.0, invoice_number=8719, sender="orders.chaddycheck@video-inform.com",
        )
        self.assertNotIn("Customer/Network", body)
        self.assertNotIn("Country:", body)
        self.assertNotIn("PO reference", body)
        self.assertNotIn("Number of cameras", body)
        self.assertNotIn(": \n", body)   # no dangling "Label: " blank-placeholder line


class TestBuildYear1InvoiceRows(unittest.TestCase):
    def test_row_shape_matches_the_orders_page_flow(self):
        rows = ocs.build_year1_invoice_rows(ORDER, invoice_number=8719)
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["invoice_number"], "8719")
        self.assertEqual(row["project_name"], "Edeka Steinebach Hohn")
        self.assertEqual(row["maintenance_year"], "Y1")
        self.assertEqual(row["payment_amount"], 2334.0)
        self.assertEqual(row["cameras_number"], 3)
        self.assertEqual(row["paid"], "No")
        self.assertEqual(row["invoice_type"], "New installation")

    def test_falls_back_to_a_synthetic_po_reference_when_missing(self):
        rows = ocs.build_year1_invoice_rows(ORDER, invoice_number=8719)
        self.assertIn("PO-po-1", rows[0]["description"])


class TestGenerateYear1InvoicePdfBytes(unittest.TestCase):
    def test_returns_nonempty_pdf_bytes(self):
        rows = ocs.build_year1_invoice_rows(ORDER, invoice_number=8719)
        pdf_bytes = ocs.generate_year1_invoice_pdf_bytes(8719, rows)
        self.assertIsInstance(pdf_bytes, bytes)
        self.assertGreater(len(pdf_bytes), 0)
        self.assertTrue(pdf_bytes.startswith(b"%PDF"))


class TestCheckGraphMailboxReachable(unittest.TestCase):
    SECRET = "s3cr3t"

    class _Ok:
        def raise_for_status(self):
            return None

        def json(self):
            return {"access_token": "tok"}

    class _Resp:
        def __init__(self, status_code):
            self.status_code = status_code

    def test_no_mailbox_given(self):
        result = ocs.check_graph_mailbox_reachable("")
        self.assertFalse(result["ok"])

    def test_missing_graph_credentials_is_reported_not_raised(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            result = ocs.check_graph_mailbox_reachable("orders.chaddycheck@video-inform.com")
        self.assertFalse(result["ok"])
        self.assertIn("not configured", result["message"])

    def test_mailbox_found_returns_ok(self):
        env = {"ORDER_INTAKE_GRAPH_TENANT_ID": "t", "ORDER_INTAKE_GRAPH_CLIENT_ID": "c",
               "ORDER_INTAKE_GRAPH_CLIENT_SECRET": self.SECRET}
        with mock.patch.dict(os.environ, env), \
             mock.patch("requests.post", return_value=self._Ok()), \
             mock.patch("requests.get", return_value=self._Resp(200)):
            result = ocs.check_graph_mailbox_reachable("orders.chaddycheck@video-inform.com")
        self.assertTrue(result["ok"], result["message"])

    def test_mailbox_not_found_returns_not_ok(self):
        env = {"ORDER_INTAKE_GRAPH_TENANT_ID": "t", "ORDER_INTAKE_GRAPH_CLIENT_ID": "c",
               "ORDER_INTAKE_GRAPH_CLIENT_SECRET": self.SECRET}
        with mock.patch.dict(os.environ, env), \
             mock.patch("requests.post", return_value=self._Ok()), \
             mock.patch("requests.get", return_value=self._Resp(404)):
            result = ocs.check_graph_mailbox_reachable("nosuchmailbox@video-inform.com")
        self.assertFalse(result["ok"])

    def test_never_raises_on_an_unexpected_exception(self):
        env = {"ORDER_INTAKE_GRAPH_TENANT_ID": "t", "ORDER_INTAKE_GRAPH_CLIENT_ID": "c",
               "ORDER_INTAKE_GRAPH_CLIENT_SECRET": self.SECRET}
        with mock.patch.dict(os.environ, env), \
             mock.patch("requests.post", side_effect=RuntimeError("network down")):
            result = ocs.check_graph_mailbox_reachable("orders.chaddycheck@video-inform.com")
        self.assertFalse(result["ok"])
        self.assertIn("network down", result["message"])


class TestBuildDryRunPreview(unittest.TestCase):
    """The dry-run orchestrator must never write anything -- no invoice row, no purchase-order update, no
    email send, no audit log."""

    def _patches(self, order=None, cfg=None, mailbox_ok=True):
        return [
            mock.patch("services.order_confirmation_service.get_purchase_order", return_value=order or ORDER),
            mock.patch("services.order_confirmation_service.get_order_confirmation_config", return_value=cfg or {
                "sender": "orders.chaddycheck@video-inform.com",
                "to": ["tom@example.com", "karen@example.com"], "cc": ["yoram@example.com"],
            }),
            mock.patch.object(ocs, "resolve_project_country", return_value="Germany"),
            mock.patch.object(ocs, "_next_invoice_number_preview", return_value=8719),
            mock.patch.object(ocs, "check_graph_mailbox_reachable", return_value={"ok": mailbox_ok, "message": "x"}),
        ]

    def test_happy_path_returns_everything_needed_for_display(self):
        patches = self._patches()
        with patches[0], patches[1], patches[2], patches[3], patches[4]:
            result = ocs.build_dry_run_preview("po-1")
        self.assertTrue(result["ok"], result["errors"])
        self.assertEqual(result["invoice_number_preview"], 8719)
        self.assertEqual(result["amount"], 2334.0)
        self.assertIn("Edeka Steinebach Hohn", result["subject"])
        self.assertTrue(result["pdf_bytes"].startswith(b"%PDF"))
        self.assertEqual(result["to"], ["tom@example.com", "karen@example.com"])
        self.assertEqual(result["cc"], ["yoram@example.com"])

    def test_ad_network_and_twelve_cameras_are_shown_in_preview_and_body(self):
        order = dict(ORDER, project_name="AD DELHAIZE Roeselare (Mandel)",
                     customer_name="", customer_email="karen@example.com",
                     ordered_cameras=12, amount=9336)
        patches = self._patches(order=order)
        with patches[0], patches[1], patches[2], patches[3], patches[4]:
            result = ocs.build_dry_run_preview("po-1")
        self.assertEqual(result["customer_display"], "AD")
        self.assertEqual(result["cameras"], 12)
        self.assertEqual(result["amount"], 9336)
        self.assertIn("- Customer/Network: AD\n", result["body"])
        self.assertIn("- Number of cameras: 12", result["body"])
        self.assertIn("- Proposed invoice number: 8719", result["body"])

    def test_missing_recipient_config_blocks_before_building_anything(self):
        patches = self._patches(cfg={"sender": "", "to": [], "cc": []})
        with patches[0], patches[1], patches[2], patches[3], patches[4]:
            result = ocs.build_dry_run_preview("po-1")
        self.assertFalse(result["ok"])
        self.assertTrue(any("sender mailbox" in e.lower() for e in result["errors"]), result["errors"])
        self.assertTrue(any("'to' recipients" in e.lower() for e in result["errors"]), result["errors"])
        self.assertNotIn("pdf_bytes", result)

    def test_validation_failure_blocks_before_building_anything(self):
        bad_order = dict(ORDER, ordered_cameras=0, amount=0)
        patches = self._patches(order=bad_order)
        with patches[0], patches[1], patches[2], patches[3], patches[4], \
             mock.patch("services.supabase_service.load_orders", return_value=[]):
            result = ocs.build_dry_run_preview("po-1")
        self.assertFalse(result["ok"])
        self.assertNotIn("pdf_bytes", result)

    def test_order_not_found(self):
        with mock.patch("services.order_confirmation_service.get_purchase_order", return_value=None):
            result = ocs.build_dry_run_preview("missing")
        self.assertFalse(result["ok"])
        self.assertIn("not found", result["errors"][0].lower())

    def test_never_writes_to_the_invoice_ledger_or_sends_anything(self):
        patches = self._patches()
        with patches[0], patches[1], patches[2], patches[3], patches[4], \
             mock.patch("services.supabase_service.replace_invoice_rows") as replace_rows, \
             mock.patch("services.order_approval_service.update_purchase_order") as update_po, \
             mock.patch("services.email_service.send_graph_email") as send_email:
            ocs.build_dry_run_preview("po-1")
        replace_rows.assert_not_called()
        update_po.assert_not_called()
        send_email.assert_not_called()

    def test_repeated_previews_only_select_candidate_without_reserving_or_consuming_it(self):
        class ReadOnlyClient:
            def __init__(self):
                self.reads = []
                self.invoices = [{"invoice_number": "8718"}]
                self.registry = []

            def table(self, name):
                if name != "invoices":
                    raise AssertionError(f"Unexpected table access: {name}")
                return self

            def select(self, columns):
                if columns != "invoice_number":
                    raise AssertionError(f"Unexpected projection: {columns}")
                self.reads.append(("invoices", columns))
                return self

            def execute(self):
                return mock.Mock(data=self.invoices.copy())

            def __getattr__(self, name):
                raise AssertionError(f"Forbidden client operation: {name}")

        client = ReadOnlyClient()
        patches = self._patches()
        with patches[0], patches[1], patches[2], patches[4], \
             mock.patch("services.order_approval_service._get_client", return_value=client), \
             mock.patch("services.supabase_service.replace_invoice_rows") as replace_rows, \
             mock.patch("services.order_approval_service.update_purchase_order") as update_po, \
             mock.patch("services.email_service.send_graph_email") as send_email:
            first = ocs.build_dry_run_preview("po-1")
            second = ocs.build_dry_run_preview("po-1")
        self.assertEqual(first["invoice_number_preview"], 8719)
        self.assertEqual(second["invoice_number_preview"], 8719)
        self.assertEqual(client.invoices, [{"invoice_number": "8718"}])
        self.assertEqual(client.registry, [])
        self.assertEqual(client.reads, [("invoices", "invoice_number")] * 2)
        replace_rows.assert_not_called()
        update_po.assert_not_called()
        send_email.assert_not_called()


class TestResolveCustomerNetwork(unittest.TestCase):
    def test_canonical_network_precedes_contact_or_customer_email(self):
        for customer in ("", "Karen <karen@example.com>", "karen@example.com"):
            with self.subTest(customer=customer):
                order = dict(ORDER, project_name="AD DELHAIZE Roeselare (Mandel)",
                             customer_name=customer)
                self.assertEqual(ocs.resolve_customer_network(order), "AD")

    def test_unknown_network_uses_real_customer_name(self):
        self.assertEqual(ocs.resolve_customer_network(
            {"project_name": "Unknown Site", "customer_name": "Customer Ltd"}), "Customer Ltd")

    def test_contact_and_sender_emails_are_never_customer_network(self):
        order = {"project_name": "Unknown Site", "customer_name": "name@example.com",
                 "customer_email": "customer@example.com", "contact_email": "contact@example.com",
                 "sender_email": "sender@example.com"}
        self.assertEqual(ocs.resolve_customer_network(order), "")


if __name__ == "__main__":
    unittest.main()
