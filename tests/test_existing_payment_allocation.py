"""Offline validation and service tests. All database clients are fake."""
from copy import deepcopy
import os
import unittest
from unittest.mock import Mock, patch

from streamlit.errors import StreamlitSecretNotFoundError

from services.existing_payment_allocation import (
    AllocationBlocked, choose_invoice, invoice_number, preview_allocation, reconciliation,
    MANUAL_ALLOCATION_GATE, manual_allocation_gate_state,
)
from services.supabase_service import add_invoice_to_existing_payment, preview_existing_payment_allocation
from services.supabase_service import preview_manual_allocation_reversal, reverse_manual_payment_allocation


class TestManualAllocationGateSecretsFallback(unittest.TestCase):
    def test_environment_values_take_precedence_without_reading_secrets(self):
        for value, expected in ((" ON ", "ON"), ("off", "OFF"), ("", "INVALID"), ("bad", "INVALID")):
            with self.subTest(value=value), patch.dict(os.environ, {MANUAL_ALLOCATION_GATE: value}), \
                    patch("streamlit.secrets") as secrets:
                secrets.get.return_value = "ON" if expected != "ON" else "OFF"
                self.assertEqual(manual_allocation_gate_state(), {
                    "state": expected, "effective_state": "ON" if expected == "ON" else "OFF",
                })
                secrets.get.assert_not_called()

    def test_root_level_secrets_on_and_off_fallback(self):
        for value, expected in ((" ON ", "ON"), ("off", "OFF")):
            with self.subTest(value=value), patch.dict(os.environ, {}, clear=True), \
                    patch("streamlit.secrets", {MANUAL_ALLOCATION_GATE: value}):
                self.assertEqual(manual_allocation_gate_state(),
                                 {"state": expected, "effective_state": expected})

    def test_missing_and_nested_secrets_fail_closed(self):
        for secrets in ({}, {"order_confirmation": {MANUAL_ALLOCATION_GATE: "ON"}}):
            with self.subTest(secrets=secrets), patch.dict(os.environ, {}, clear=True), \
                    patch("streamlit.secrets", secrets):
                self.assertEqual(manual_allocation_gate_state(),
                                 {"state": "UNSET", "effective_state": "OFF"})

    def test_invalid_secret_values_fail_closed_without_exposing_values(self):
        for value in ("", "NEVER_DISPLAY_THIS_SECRET", True, 1, [], {}):
            with self.subTest(value=value), patch.dict(os.environ, {}, clear=True), \
                    patch("streamlit.secrets", {MANUAL_ALLOCATION_GATE: value}), \
                    self.assertLogs("services.existing_payment_allocation", level="WARNING") as logs:
                self.assertEqual(manual_allocation_gate_state(),
                                 {"state": "INVALID", "effective_state": "OFF"})
                self.assertNotIn("NEVER_DISPLAY_THIS_SECRET", str(logs.output))

    def test_unavailable_secrets_fail_closed_with_sanitized_warning(self):
        with patch.dict(os.environ, {}, clear=True), patch("streamlit.secrets") as secrets, \
                self.assertLogs("services.existing_payment_allocation", level="WARNING") as logs:
            secrets.get.side_effect = StreamlitSecretNotFoundError("NEVER_DISPLAY_THIS_SECRET")
            self.assertEqual(manual_allocation_gate_state(),
                             {"state": "UNSET", "effective_state": "OFF"})
            self.assertNotIn("NEVER_DISPLAY_THIS_SECRET", str(logs.output))


def fixture(limit=46680):
    payment = {
        "id": 14, "payment_date": "2026-10-20", "invoice_number": "8596,8597,8700,8709",
        "source_name": "Transaction_21_Oct_2026.pdf", "source_kind": "pdf", "currency": "EUR",
        "instructed_amount": limit, "received_amount": limit - 63.68, "fee_amount": 63.68,
        "applied_amount": 38900, "payment_fingerprint": "original",
        "pdf_storage_bucket": "bank-payments", "pdf_storage_path": "original.pdf",
        "raw_text": "original remittance", "parsed_payload": {"unchanged": True}, "notes": "original",
    }
    allocations = [{
        "id": index, "payment_id": 14, "invoice_row_id": row_id, "invoice_number": number,
        "project_name": project, "amount_applied": amount, "maintenance_year": "Y1", "year": None,
    } for index, (row_id, number, project, amount) in enumerate([
        (500, 8596, "AD Wondelgem", 9336), (501, 8597, "AD Tienen", 15560),
        (3080, 8700, "AD Pallieter", 6224), (3089, 8709, "Proxy Lovenjoel", 7780),
    ], 73)]
    invoice = {
        "id": 3087, "invoice_number": "8707", "project_name": "Intermarche Heusy",
        "maintenance_year": "Y1", "payment_amount": 7780, "paid": "No", "payment_date": None,
        "year": None,
    }
    return payment, allocations, invoice


class TestManualReversalService(unittest.TestCase):
    def setUp(self):
        p, allocations, invoice = fixture()
        self.original = deepcopy(invoice)
        self.invoice = {**invoice, "paid": "Yes", "payment_date": p["payment_date"]}
        self.payment = {**p, "applied_amount": 46680, "invoice_number": "8596,8597,8700,8707,8709"}
        self.manual = {
            "id": 77, "payment_id": 14, "invoice_row_id": 3087, "invoice_number": 8707,
            "project_name": invoice["project_name"], "amount_applied": 7780,
            "manual_invoice_before": invoice,
        }
        self.preview = {"allocation_id": 77, "expected": {
            "actor": "admin", "payment": self.payment, "invoice": self.invoice, "allocations": allocations + [self.manual],
        }}
        self.result = {"status": "reversed", "allocation_id": 77,
                       "payment": p, "invoice": invoice, "allocations": allocations}
        self.client = Mock()
        self.client.rpc.return_value.execute.return_value.data = self.result
        self.client_patch = patch("services.supabase_service._get_client", return_value=self.client)
        self.gate_patch = patch("services.financial_writer_adapter.enabled", return_value=False)
        self.client_patch.start()
        self.gate = self.gate_patch.start()
        self.addCleanup(self.client_patch.stop)
        self.addCleanup(self.gate_patch.stop)

    def test_admin_confirmation_actor_reason_and_gate_are_checked_before_write(self):
        for changes in ({"role": "viewer"}, {"confirmed": False}, {"actor": ""}, {"actor": "wrong-admin"}, {"reason": ""}):
            kwargs = {"role": "admin", "confirmed": True, "actor": "admin", "reason": "Correction"}
            kwargs.update(changes)
            with self.subTest(changes=changes), self.assertRaises(AllocationBlocked):
                reverse_manual_payment_allocation(self.preview, **kwargs)
        self.gate.return_value = True
        with self.assertRaises(AllocationBlocked):
            reverse_manual_payment_allocation(self.preview, role="admin", confirmed=True, actor="admin", reason="Correction")
        self.client.rpc.assert_not_called()

    def test_one_rpc_reverses_and_verifies_without_side_effects(self):
        result = reverse_manual_payment_allocation(
            self.preview, role="admin", confirmed=True, actor="admin", reason="Correction",
        )
        self.assertEqual(result, self.result)
        self.client.rpc.assert_called_once_with("reverse_manual_payment_allocation", {
            "p_allocation_id": 77, "p_expected": self.preview["expected"],
            "p_actor": "admin", "p_reason": "Correction", "p_confirmed": True,
        })
        self.client.table.assert_not_called()

    def test_uncertain_or_altered_correction_readback_is_explicit_error(self):
        for field, value in (("status", "unknown"), ("allocations", [])):
            self.client.rpc.return_value.execute.return_value.data = {**self.result, field: value}
            with self.subTest(field=field), self.assertLogs("services.supabase_service", level="ERROR"), \
                    self.assertRaises(AllocationBlocked):
                reverse_manual_payment_allocation(
                    self.preview, role="admin", confirmed=True, actor="admin", reason="Correction",
                )

    def test_sql_error_is_logged_and_propagated_without_fallback(self):
        self.client.rpc.return_value.execute.side_effect = RuntimeError("stale correction")
        with self.assertLogs("services.supabase_service", level="ERROR"), \
                self.assertRaisesRegex(RuntimeError, "stale correction"):
            reverse_manual_payment_allocation(
                self.preview, role="admin", confirmed=True, actor="admin", reason="Correction",
            )
        self.client.table.assert_not_called()

    def test_preview_reads_only_and_requires_existing_manual_allocation(self):
        query = self.client.table.return_value
        query.select.return_value = query
        query.eq.return_value = query
        query.execute.side_effect = [
            Mock(data=[self.manual]), Mock(data=[self.payment]), Mock(data=[self.invoice]),
        ]
        with patch("services.supabase_service.load_bank_payment_allocations",
                   return_value=self.preview["expected"]["allocations"]):
            self.assertEqual(preview_manual_allocation_reversal(77, role="admin", actor="admin"), self.preview)
        self.client.rpc.assert_not_called()

    def test_invalid_correction_id_or_nonmanual_preview_is_blocked(self):
        for allocation_id in (None, True, 0):
            with self.subTest(allocation_id=allocation_id), self.assertRaises(AllocationBlocked):
                preview_manual_allocation_reversal(allocation_id, role="admin", actor="admin")
        preview = deepcopy(self.preview)
        preview["expected"]["allocations"][-1]["manual_invoice_before"] = None
        with self.assertRaises(AllocationBlocked):
            reverse_manual_payment_allocation(preview, role="admin", confirmed=True, actor="admin", reason="Correction")
        self.client.rpc.assert_not_called()


class TestPreview(unittest.TestCase):
    def test_exact_reconciliation_without_mutating_inputs(self):
        p, a, i = fixture()
        original = deepcopy((p, a, i))
        result = preview_allocation(p, a, i, [])
        self.assertEqual(result["remaining_before"], "7780.00")
        self.assertEqual(result["remaining_after"], "0.00")
        self.assertEqual(result["payment_date"], "2026-10-20")
        self.assertEqual((p, a, i), original)
        result["expected"]["payment"]["source_name"] = "changed"
        self.assertEqual(p["source_name"], original[0]["source_name"])

    def test_remaining_balance(self):
        p, a, i = fixture(50000)
        self.assertEqual(preview_allocation(p, a, i, [])["remaining_after"], "3320.00")

    def test_duplicate_and_allocation_elsewhere(self):
        for payment_id in (14, 99):
            with self.subTest(payment_id=payment_id):
                p, a, i = fixture()
                match = {"invoice_row_id": i["id"], "payment_id": payment_id}
                with self.assertRaisesRegex(AllocationBlocked, "already allocated"):
                    preview_allocation(p, a, i, [match])

    def test_legacy_allocation_elsewhere(self):
        p, a, i = fixture()
        legacy = {"payment_id": 99, "invoice_row_id": None, "invoice_number": 8707,
                  "project_name": i["project_name"]}
        with self.assertRaisesRegex(AllocationBlocked, "elsewhere"):
            preview_allocation(p, a, i, [legacy])

    def test_non_payable_and_positive_full_amount_only(self):
        for status in ("Yes", "cancelled", "Cancelled", "unknown"):
            p, a, i = fixture()
            i["paid"] = status
            with self.subTest(status=status), self.assertRaisesRegex(AllocationBlocked, "non-payable"):
                preview_allocation(p, a, i, [])
        for amount in (0, -1, None, "NaN", "Infinity", "7780.001"):
            p, a, i = fixture()
            i["payment_amount"] = amount
            with self.subTest(amount=amount), self.assertRaises(AllocationBlocked):
                preview_allocation(p, a, i, [])
        for status in ("", None, " no "):
            p, a, i = fixture()
            i["paid"] = status
            self.assertEqual(preview_allocation(p, a, i, [])["amount"], "7780")

    def test_missing_and_ambiguous_invoice_require_selection(self):
        _, _, i = fixture()
        with self.assertRaisesRegex(AllocationBlocked, "No existing"):
            choose_invoice([])
        with self.assertRaisesRegex(AllocationBlocked, "multiple rows"):
            choose_invoice([i, {**i, "id": 2}])
        self.assertIs(choose_invoice([i, {**i, "id": 2}], 3087), i)
        with self.assertRaises(AllocationBlocked):
            choose_invoice([i], 1)

    def test_over_allocation_invalid_date_currency_and_limit(self):
        cases = [
            ("instructed_amount", 46679), ("instructed_amount", None),
            ("instructed_amount", 0), ("instructed_amount", "NaN"),
            ("payment_date", None), ("payment_date", "2026-02-30"),
            ("payment_date", "20261020"), ("currency", "USD"),
        ]
        for field, value in cases:
            p, a, i = fixture()
            p[field] = value
            with self.subTest(field=field, value=value), self.assertRaises(AllocationBlocked):
                preview_allocation(p, a, i, [])
        for value in ("USD", None, ""):
            p, a, i = fixture()
            i["currency"] = value
            with self.subTest(currency=value), self.assertRaises(AllocationBlocked):
                preview_allocation(p, a, i, [])

    def test_summary_mismatch_and_duplicate_existing_rows(self):
        p, a, _ = fixture()
        for field, value in (("applied_amount", 0), ("invoice_number", "8596"), ("invoice_number", "8596,8596")):
            with self.subTest(field=field), self.assertRaises(AllocationBlocked):
                reconciliation({**p, field: value}, a)
        a[-1]["invoice_row_id"] = a[0]["invoice_row_id"]
        with self.assertRaisesRegex(AllocationBlocked, "duplicate"):
            reconciliation(p, a)

    def test_invoice_number_input_is_not_a_multi_invoice_or_injection_input(self):
        for text in ("8707,8709", "8707.0", "0", "-1", "foo", "8707 or 1=1"):
            with self.subTest(text=text), self.assertRaises(AllocationBlocked):
                invoice_number(text)


class FakeQuery:
    def __init__(self, client, table):
        self.client, self.table = client, table
        self.filters = []

    def select(self, *args):
        return self

    def eq(self, key, value):
        self.filters.append((key, str(value)))
        return self

    def or_(self, value):
        self.client.or_queries.append(value)
        return self

    def limit(self, value):
        self.limit_value = value
        return self

    def execute(self):
        data = self.client.rows[self.table]
        for key, value in self.filters:
            data = [row for row in data if str(row.get(key)) == value]
        return Mock(data=deepcopy(data))


class FakeClient:
    def __init__(self, p, a, i):
        self.rows = {"bank_payments": [p], "bank_payment_allocations": [], "invoices": [i],
                     "manual_allocation_release_policy": [{"singleton": True, "enabled": True}]}
        self.or_queries = []
        self.rpc = Mock()

    def table(self, name):
        return FakeQuery(self, name)


class TestService(unittest.TestCase):
    def setUp(self):
        release = patch.dict(os.environ, {MANUAL_ALLOCATION_GATE: "on"})
        release.start()
        self.addCleanup(release.stop)
        self.p, self.a, self.i = fixture()
        self.preview = preview_allocation(self.p, self.a, self.i, [])
        self.client = FakeClient(self.p, self.a, self.i)
        self.client_patch = patch("services.supabase_service._get_client", return_value=self.client)
        self.client_patch.start()
        self.addCleanup(self.client_patch.stop)
        self.gate_patch = patch("services.supabase_service.financial_writer.enabled", return_value=False)
        self.gate_patch.start()
        self.addCleanup(self.gate_patch.stop)

    def test_release_gate_blocks_preview_and_direct_backend_before_client_access(self):
        for value in (None, "off", "", "invalid-secret-value"):
            with self.subTest(value=value), patch.dict(os.environ, {}, clear=True):
                if value is not None:
                    os.environ[MANUAL_ALLOCATION_GATE] = value
                with patch("services.supabase_service._get_client") as client:
                    with self.assertRaisesRegex(AllocationBlocked, "release gate"):
                        preview_existing_payment_allocation(14, 8707, role="admin")
                    with self.assertRaisesRegex(AllocationBlocked, "release gate"):
                        add_invoice_to_existing_payment(self.preview, role="admin", confirmed=True)
                    client.assert_not_called()

    def test_database_release_off_or_invalid_blocks_preview_and_direct_backend(self):
        for rows in (
            [{"singleton": True, "enabled": False}], [], None,
            [{"singleton": True, "enabled": "on"}],
            [{"singleton": False, "enabled": True}],
            [{"singleton": True, "enabled": True}] * 2,
        ):
            self.client.rows["manual_allocation_release_policy"] = rows
            with self.subTest(rows=rows):
                with self.assertLogs("services.supabase_service", level="WARNING"):
                    with self.assertRaises(AllocationBlocked):
                        preview_existing_payment_allocation(14, 8707, role="admin")
                with self.assertLogs("services.supabase_service", level="WARNING"):
                    with self.assertRaises(AllocationBlocked):
                        add_invoice_to_existing_payment(self.preview, role="admin", confirmed=True)
        self.client.rpc.assert_not_called()

    def test_database_release_read_failure_blocks_without_rpc(self):
        with patch.object(self.client, "table", side_effect=RuntimeError("read unavailable")):
            with self.assertLogs("services.supabase_service", level="WARNING"):
                with self.assertRaisesRegex(RuntimeError, "read unavailable"):
                    add_invoice_to_existing_payment(self.preview, role="admin", confirmed=True)
        self.client.rpc.assert_not_called()

    def outcome(self):
        allocation = {**self.a[0], "id": 77, "invoice_row_id": self.i["id"],
                      "invoice_number": 8707, "project_name": self.i["project_name"], "amount_applied": 7780}
        return {"status": "added", "payment": {**self.p, "applied_amount": 46680,
                "invoice_number": "8596,8597,8700,8707,8709"},
                "allocations": [*self.a, allocation],
                "invoice": {**self.i, "paid": "Yes", "payment_date": "2026-10-20"}, "allocation_id": 77}

    def test_preview_is_select_only(self):
        with patch("services.supabase_service.load_bank_payment_allocations", return_value=self.a):
            result = preview_existing_payment_allocation(14, 8707, role="admin")
        self.assertEqual(result, self.preview)
        self.client.rpc.assert_not_called()
        self.assertEqual(self.client.or_queries, ["invoice_row_id.eq.3087,invoice_number.eq.8707"])

    def test_admin_and_confirmation_required_before_client_access(self):
        for role, confirmed in (("viewer", True), ("admin", False), ("admin", 1)):
            with self.subTest(role=role, confirmed=confirmed), self.assertRaises(AllocationBlocked):
                add_invoice_to_existing_payment(self.preview, role=role, confirmed=confirmed)
        self.client.rpc.assert_not_called()
        with self.assertRaises(AllocationBlocked):
            preview_existing_payment_allocation(14, 8707, role="viewer")

    def test_financial_writer_mode_has_no_fallback(self):
        with patch("services.supabase_service.financial_writer.enabled", return_value=True):
            with self.assertRaisesRegex(AllocationBlocked, "disabled"):
                add_invoice_to_existing_payment(self.preview, role="admin", confirmed=True)
            with self.assertRaisesRegex(AllocationBlocked, "disabled"):
                preview_existing_payment_allocation(14, 8707, role="admin")
        self.client.rpc.assert_not_called()

    def test_one_narrow_rpc_and_readback(self):
        self.client.rpc.return_value.execute.return_value = Mock(data=self.outcome())
        result = add_invoice_to_existing_payment(self.preview, role="admin", confirmed=True)
        self.assertEqual(result["allocation_id"], 77)
        self.client.rpc.assert_called_once_with("add_invoice_to_existing_payment", {
            "p_payment_id": 14, "p_invoice_row_id": 3087, "p_expected": self.preview["expected"],
        })

    def test_tampered_preview_blocked_before_rpc(self):
        self.preview["amount"] = "1.00"
        with self.assertRaisesRegex(AllocationBlocked, "invalid"):
            add_invoice_to_existing_payment(self.preview, role="admin", confirmed=True)
        self.client.rpc.assert_not_called()

    def test_preservation_and_single_allocation_readback_are_checked(self):
        for field in ("pdf_storage_path", "received_amount", "payment_fingerprint"):
            outcome = self.outcome()
            outcome["payment"][field] = "changed"
            self.client.rpc.return_value.execute.return_value = Mock(data=outcome)
            with self.subTest(field=field), self.assertLogs("services.supabase_service", level="ERROR"):
                with self.assertRaisesRegex(AllocationBlocked, "metadata changed"):
                    add_invoice_to_existing_payment(self.preview, role="admin", confirmed=True)
        outcome = self.outcome()
        outcome["allocations"][0]["project_name"] = "rewritten"
        self.client.rpc.return_value.execute.return_value = Mock(data=outcome)
        with self.assertLogs("services.supabase_service", level="ERROR"):
            with self.assertRaisesRegex(AllocationBlocked, "allocations changed"):
                add_invoice_to_existing_payment(self.preview, role="admin", confirmed=True)

    def test_rpc_failure_or_uncertain_response_is_explicit(self):
        self.client.rpc.return_value.execute.side_effect = RuntimeError("Preview is stale")
        with self.assertLogs("services.supabase_service", level="ERROR"), self.assertRaisesRegex(RuntimeError, "stale"):
            add_invoice_to_existing_payment(self.preview, role="admin", confirmed=True)
        self.client.rpc.return_value.execute.side_effect = None
        self.client.rpc.return_value.execute.return_value = Mock(data=None)
        with self.assertLogs("services.supabase_service", level="ERROR"), self.assertRaisesRegex(AllocationBlocked, "unconfirmed"):
            add_invoice_to_existing_payment(self.preview, role="admin", confirmed=True)


class TestManualAllocationGate(unittest.TestCase):
    def test_states_are_independent_and_sanitized(self):
        for value, state, effective in (
            (None, "UNSET", "OFF"), ("off", "OFF", "OFF"), ("on", "ON", "ON"),
            ("", "INVALID", "OFF"), ("secret-invalid", "INVALID", "OFF"),
        ):
            for writer in ("on", "off"):
                with self.subTest(value=value, writer=writer), patch.dict(
                    os.environ, {"CADDYCHECK_COVERAGE_WRITERS": writer}, clear=True,
                ):
                    if value is not None:
                        os.environ[MANUAL_ALLOCATION_GATE] = value
                    if state == "INVALID":
                        with self.assertLogs("services.existing_payment_allocation", level="WARNING") as logs:
                            result = manual_allocation_gate_state()
                        self.assertNotIn("secret-invalid", str(logs.output))
                    else:
                        result = manual_allocation_gate_state()
                    self.assertEqual(result, {"state": state, "effective_state": effective})
