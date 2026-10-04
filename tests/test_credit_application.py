"""Failure-injection tests for the atomic "Apply Credit" operation (services/credit_application.py).

Every simulated failure must leave the ledger EXACTLY as it was: credit restored, target rows restored,
no orphan usage/remainder rows and no (misleading) successful credit-application payment record.
Nothing here touches production data -- the "database" is an in-memory fake.
"""
import json
import unittest

try:
    from tests.test_bank_payment_save import PAY_DATE, FakeLedger, _credit, _target
except ImportError:  # `unittest discover -s tests` runs from inside the tests folder
    from test_bank_payment_save import PAY_DATE, FakeLedger, _credit, _target

from services.credit_application import (
    CreditPlan,
    apply_credit_atomically,
    execute_with_compensation,
    plan_credit_application_ops,
)

LOGGER = "services.credit_application"


class FakeApiError(Exception):
    """Looks like postgrest.APIError: the database answered with an error code."""

    def __init__(self, message, code):
        super().__init__(message)
        self.code = code


class TestCreditPlan(unittest.TestCase):
    def test_operations_are_ordered_and_the_payment_record_is_last(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        plan = ledger.plan(8700)
        self.assertEqual([op["op"] for op in plan.ops], ["credit_consume", "insert_row", "target_update", "save_payment"])
        self.assertEqual((plan.consumed, plan.applied), (130.0, 130.0))
        credit_op, usage_op = plan.ops[0], plan.ops[1]
        self.assertEqual((credit_op["new_amount"], credit_op["mark_paid"]), (-870.0, False))
        self.assertEqual((usage_op["row"]["payment_amount"], usage_op["row"]["paid"]), (-130.0, "Yes"))

    def test_payment_entry_is_complete_and_consistent_with_the_allocations(self):
        plan = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)]).plan(8700)
        entry = plan.payment_entry
        self.assertEqual((entry["source_kind"], entry["invoice_number"], entry["applied_amount"]), ("credit-apply", "8700", 130.0))
        self.assertEqual(entry["applied_amount"], sum(a["amount_applied"] for a in plan.allocations))
        self.assertEqual(entry["parsed_payload"]["credit_sources"][0]["credit_row_id"], 1)
        self.assertEqual(len(entry["payment_fingerprint"]), 64)

    def test_nothing_to_apply_gives_an_empty_plan(self):
        self.assertEqual(FakeLedger([_target(2, 130.0)]).plan(8700).ops, [])  # no credit
        self.assertEqual(FakeLedger([_credit(1, -50.0)]).plan(8700).ops, [])  # no target rows
        self.assertEqual(FakeLedger([_credit(1, -50.0, paid="cancelled"), _target(2, 9.0)]).plan(8700).ops, [])

    def test_plan_is_json_serialisable_for_the_database_function(self):
        payload = FakeLedger([_credit(1, -100.0), _target(2, 130.0)]).plan(8700).rpc_payload()
        self.assertEqual(json.loads(json.dumps(payload)), payload)


class TestCompensatingRollbackFailureInjection(unittest.TestCase):
    """Rollback mode (used until the SQL function is installed): undo applied steps in reverse order."""

    def _run_failing(self, ledger, invoice=8700):
        initial = ledger.state()
        with self.assertLogs(LOGGER, level="ERROR"):
            result = ledger.apply(invoice)
        return initial, result

    def _assert_fully_restored(self, ledger, initial, result):
        self.assertFalse(result.ok)
        self.assertTrue(result.rolled_back, result.rollback_errors)
        self.assertEqual(ledger.state(), initial, "ledger must be byte-for-byte what it was before")
        self.assertEqual(ledger.payments, {}, "no payment/credit-application record may survive a failure")
        self.assertFalse([i for i in ledger.rows if i > 1000], "no orphan usage/remainder rows")
        self.assertEqual((result.consumed, result.applied), (0.0, 0.0))
        self.assertTrue(result.errors and "rolled back" in result.errors[0])

    def test_failure_after_consuming_the_credit_but_before_updating_the_target(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        ledger.fail_calls = {"update_target_row": {1}}
        initial, result = self._run_failing(ledger)
        self._assert_fully_restored(ledger, initial, result)
        self.assertEqual(ledger.available(), 1000.0)  # original available credit restored
        self.assertEqual(ledger.rows[2]["paid"], "No")

    def test_failure_while_creating_the_remainder_row(self):
        ledger = FakeLedger([_credit(1, -100.0), _target(2, 130.0)])  # target only partly funded -> remainder row
        ledger.fail_calls = {"insert_row": {1}}
        initial, result = self._run_failing(ledger)
        self._assert_fully_restored(ledger, initial, result)
        self.assertEqual(ledger.rows[2]["payment_amount"], 130.0)  # partly-paid split was undone too

    def test_failure_while_creating_the_credit_usage_row(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        ledger.fail_calls = {"insert_row": {1}}  # first insert is the usage row for the partly used credit
        initial, result = self._run_failing(ledger)
        self._assert_fully_restored(ledger, initial, result)
        self.assertEqual(ledger.available(), 1000.0)

    def test_failure_while_creating_the_audit_payment_record(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        ledger.fail_calls = {"save_payment": {1}}
        initial, result = self._run_failing(ledger)
        self._assert_fully_restored(ledger, initial, result)
        self.assertEqual(ledger.available(), 1000.0)
        self.assertEqual((ledger.rows[2]["paid"], ledger.rows[2]["payment_amount"]), ("No", 130.0))

    def test_multiple_targets_where_one_succeeds_and_a_later_one_fails(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0), _target(3, 70.0)])
        ledger.fail_calls = {"update_target_row": {2}}  # the first target is already paid when the second fails
        initial, result = self._run_failing(ledger)
        self._assert_fully_restored(ledger, initial, result)
        self.assertEqual((ledger.rows[2]["paid"], ledger.rows[3]["paid"]), ("No", "No"))

    def test_failure_across_several_credit_rows_restores_all_of_them(self):
        ledger = FakeLedger([_credit(1, -100.0), _credit(2, -50.0), _credit(3, -200.0), _target(10, 220.0)])
        ledger.fail_calls = {"save_payment": {1}}
        initial, result = self._run_failing(ledger, invoice=8700)
        self._assert_fully_restored(ledger, initial, result)
        self.assertEqual(ledger.available(), 350.0)

    def test_a_row_changed_by_someone_else_aborts_and_rolls_back_everything(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0), _target(3, 70.0)])
        plan = ledger.plan(8700)
        ledger.rows[3]["payment_amount"] = 75.0  # edited elsewhere after the page was loaded
        initial = ledger.state()
        with self.assertLogs(LOGGER, level="ERROR"):
            result = execute_with_compensation(plan, **ledger.executors())
        self.assertFalse(result.ok)
        self.assertTrue(result.rolled_back)
        self.assertEqual(ledger.state(), initial)
        self.assertIn("changed since it was loaded", result.errors[0])

    def test_failure_to_undo_is_reported_loudly_not_swallowed(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        ledger.fail_calls = {"update_target_row": {1}, "restore_row": {1}}
        with self.assertLogs(LOGGER, level="ERROR"):
            result = ledger.apply(8700)
        self.assertFalse(result.ok)
        self.assertFalse(result.rolled_back)
        self.assertTrue(result.rollback_errors)
        self.assertTrue(any("ROLLBACK INCOMPLETE" in e for e in result.errors))

    def test_retry_after_a_rolled_back_failure_succeeds_cleanly(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        ledger.fail_calls = {"save_payment": {1}}
        with self.assertLogs(LOGGER, level="ERROR"):
            self.assertFalse(ledger.apply(8700).ok)
        ledger.fail_calls = {}
        result = ledger.apply(8700)
        self.assertTrue(result.ok, result.errors)
        self.assertEqual(ledger.available(), 870.0)
        self.assertEqual(len(ledger.payments), 1)
        self.assertEqual(ledger.rows[2]["paid"], "Yes")

    def test_success_leaves_exactly_one_payment_record_matching_the_allocations(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        result = ledger.apply(8700)
        self.assertTrue(result.ok)
        self.assertEqual(result.mode, "compensation")
        (payment,) = ledger.payments.values()
        self.assertEqual(payment["entry"]["applied_amount"], sum(a["amount_applied"] for a in payment["allocations"]))


class TestAtomicOrchestrator(unittest.TestCase):
    """apply_credit_atomically(): one DB transaction (RPC) first; compensation only if the function is missing."""

    def _fallback(self, ledger, calls):
        def run(plan):
            calls.append("fallback")
            return execute_with_compensation(plan, **ledger.executors())
        return run

    def test_rpc_success_is_a_single_transaction_and_needs_no_fallback(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        plan, sent, calls = ledger.plan(8700), [], []
        result = apply_credit_atomically(plan, rpc=sent.append, fallback=self._fallback(ledger, calls))
        self.assertTrue(result.ok)
        self.assertEqual(result.mode, "transaction")
        self.assertEqual((result.consumed, result.applied), (130.0, 130.0))
        self.assertEqual(sent, [plan.rpc_payload()])
        self.assertEqual(calls, [])

    def test_database_rejection_means_nothing_changed_and_no_fallback_is_attempted(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        initial, calls = ledger.state(), []

        def rejecting_rpc(payload):
            raise FakeApiError("credit row 1 changed since it was loaded", code="P0001")

        with self.assertLogs(LOGGER, level="ERROR"):
            result = apply_credit_atomically(ledger.plan(8700), rpc=rejecting_rpc, fallback=self._fallback(ledger, calls))
        self.assertFalse(result.ok)
        self.assertTrue(result.rolled_back)
        self.assertFalse(result.outcome_unknown)
        self.assertEqual((result.consumed, result.applied), (0.0, 0.0))
        self.assertEqual(calls, [])  # never retried non-atomically after a real rejection
        self.assertEqual(ledger.state(), initial)

    def test_transport_failure_is_reported_as_unknown_outcome(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        calls = []

        def timing_out_rpc(payload):
            raise TimeoutError("read timed out")

        with self.assertLogs(LOGGER, level="ERROR"):
            result = apply_credit_atomically(ledger.plan(8700), rpc=timing_out_rpc, fallback=self._fallback(ledger, calls))
        self.assertFalse(result.ok)
        self.assertTrue(result.outcome_unknown)
        self.assertFalse(result.rolled_back)
        self.assertIn("UNKNOWN", result.errors[0])
        self.assertEqual(calls, [])

    def test_missing_function_falls_back_to_compensating_rollback(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        calls = []

        def missing_rpc(payload):
            raise FakeApiError("Could not find the function public.apply_credit_atomic", code="PGRST202")

        with self.assertLogs(LOGGER, level="WARNING"):
            result = apply_credit_atomically(ledger.plan(8700), rpc=missing_rpc, fallback=self._fallback(ledger, calls))
        self.assertTrue(result.ok)
        self.assertEqual(result.mode, "compensation")
        self.assertEqual(calls, ["fallback"])
        self.assertEqual(ledger.available(), 870.0)

    def test_fallback_failure_is_also_rolled_back(self):
        ledger = FakeLedger([_credit(1, -1000.0), _target(2, 130.0)])
        ledger.fail_calls = {"save_payment": {1}}
        initial, calls = ledger.state(), []

        def missing_rpc(payload):
            raise FakeApiError("Could not find the function", code="PGRST202")

        with self.assertLogs(LOGGER, level="WARNING"):
            result = apply_credit_atomically(ledger.plan(8700), rpc=missing_rpc, fallback=self._fallback(ledger, calls))
        self.assertFalse(result.ok)
        self.assertTrue(result.rolled_back)
        self.assertEqual(ledger.state(), initial)

    def test_empty_plan_does_nothing(self):
        called = []
        result = apply_credit_atomically(CreditPlan(), rpc=called.append, fallback=lambda p: called.append("fb"))
        self.assertTrue(result.ok)
        self.assertEqual(called, [])


if __name__ == "__main__":
    unittest.main()
