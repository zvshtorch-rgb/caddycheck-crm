"""Ask Data Phase 3: ToolResult, Decimal money, data freshness, grounding, debt_by_country, drill-down."""
import ast
import datetime
import json
import unittest
from decimal import Decimal
from pathlib import Path
from unittest import mock

from models.invoice import Invoice
from services.ask_data import data_clock, router as router_module
from services.ask_data.executor import execute_tool
from services.ask_data.money import amounts_in_text, fmt_eur, fmt_percent, percent, sum_money, to_money
from services.ask_data.tool_result import (
    Component, ResultStore, ToolResult, UnknownResultError, get_result_rows,
)
from services.ask_data.tool_specs import TOOL_SPECS
from services.ask_data.validation import validate_tool_call
from tests._ask_data_harness import (
    BASE_CFG, FIXED_LOADED_AT, NewImpl, OracleImpl, TOOL_CASES, _gemini_ok, _make_post, build_fixture, fake_streamlit,
)
from tests.test_ask_data_resolver import _inv, _project, make_data

ROOT = Path(__file__).resolve().parent.parent
D = Decimal


def run(tool, args=None, data=None):
    projects, invoices = data or make_data()
    with mock.patch.dict(data_clock._state, {"loaded_at": FIXED_LOADED_AT, "source": "test"}):
        return execute_tool(tool, dict(args or {}), projects, invoices, [])


class TestMoney(unittest.TestCase):
    def test_float_artifacts_never_reach_amounts(self):
        self.assertEqual(to_money(1000.5), D("1000.50"))
        self.assertEqual(to_money(0.1 + 0.2), D("0.30"))                       # 0.30000000000000004 as a float
        self.assertEqual(sum_money([0.1] * 10), D("1.00"))
        drifting = 0.0
        for _ in range(10):
            drifting += 0.1
        self.assertNotEqual(drifting, 1.0)                                      # naive float accumulation gives 0.9999999999999999
        self.assertEqual(sum_money(["1.005", D("2.005")]), D("1.01") + D("2.01"))
        self.assertEqual((to_money(None), to_money(""), to_money(float("nan"))), (D("0.00"),) * 3)

    def test_rounding_is_half_up_to_cents(self):
        self.assertEqual(to_money("2.675"), D("2.68"))
        self.assertEqual(to_money("-2.675"), D("-2.68"))
        self.assertEqual(to_money(5016), D("5016.00"))

    def test_formatting_shows_cents_and_signs(self):
        self.assertEqual(fmt_eur(D("189998")), "€189,998.00")
        self.assertEqual(fmt_eur(-2723), "-€2,723.00")
        self.assertEqual(fmt_eur(D("0")), "€0.00")

    def test_percentages_are_decimal_and_safe_on_zero(self):
        self.assertEqual(percent(D("168590"), D("189998")), D("88.7"))
        self.assertIsInstance(percent(D("1"), D("3")), Decimal)
        self.assertIsNone(percent(D("5"), D("0")))
        self.assertEqual(fmt_percent(None), "n/a")

    def test_amounts_in_text(self):
        self.assertEqual(amounts_in_text("a €1,234.50 and -€2,723.00, old €-5"), [D("1234.50"), D("-2723.00"), D("-5")])


class TestNoFloatsInFinancialResults(unittest.TestCase):
    FINANCIAL = [(t, a) for _id, t, a in TOOL_CASES if t in (
        "get_invoice", "get_invoices", "get_unpaid_invoices", "get_debt", "get_paid_amount", "get_outstanding_amount",
        "get_top_debt_projects", "debt_by_country")]

    def test_values_components_derived_and_tables_are_decimal(self):
        projects, invoices = build_fixture()
        checked = 0
        with mock.patch.dict(data_clock._state, {"loaded_at": FIXED_LOADED_AT, "source": "test"}):
            for tool, args in self.FINANCIAL:
                result = execute_tool(tool, dict(args), projects, invoices, [])
                if not result.financial:
                    continue
                checked += 1
                for amount in [result.value, *[c.amount for c in result.components], *result.derived_amounts]:
                    self.assertIsInstance(amount, Decimal, f"{tool} {args}")
                if result.dataframe is not None:
                    for column in result.dataframe.columns:
                        if "€" in str(column):
                            self.assertTrue(all(isinstance(v, Decimal) for v in result.dataframe[column]), f"{tool} {args} {column}")
        self.assertGreater(checked, 30)

    def test_exact_total_where_a_float_sum_would_drift(self):
        projects = [_project("AD Alpha", "Bel")]
        invoices = [_inv(9000 + i, "AD Alpha", 0.1) for i in range(10)]
        result = run("get_debt", data=(projects, invoices))
        self.assertEqual(result.value, D("1.00"))
        self.assertIn("totaling €1.00.", result.text)
        self.assertEqual(result.display_text().split("\n\n")[0], "I found 10 unpaid invoice row(s) totaling €1.00.")

    def test_cents_are_displayed_and_kept(self):
        projects, invoices = make_data()
        result = run("get_outstanding_amount", {"country": "Hol"}, (projects, invoices))
        self.assertEqual(result.value, D("200.50"))
        self.assertIn("€200.50", result.text)

    def test_the_app_legacy_path_uses_decimal_and_shows_freshness(self):
        projects, invoices = make_data()
        app = OracleImpl((ROOT / "streamlit_app.py").read_text(encoding="utf-8"), "app", delegate=NewImpl())
        with fake_streamlit(), app.env(cfg={"api_key": "", "model": "m"}):
            text, df = app.ns["_answer_data_question_smart"]("show unpaid debt in belgium", projects, invoices, [])
        self.assertIn("I found 5 unpaid invoice row(s) totaling €11,942.00.", text)      # canonical country matching, like the tools
        self.assertIn("Data as of 2026-10-20 09:05 UTC.", text)
        self.assertTrue(all(isinstance(v, Decimal) for v in df["Amount (€)"]))


class TestToolResultSchema(unittest.TestCase):
    def test_a_financial_result_carries_every_field(self):
        result = run("get_unpaid_invoices", {"country": "Bel", "year": 2026})
        self.assertTrue(result.financial)
        self.assertEqual(result.currency, "EUR")
        self.assertEqual(result.value, D("11942.00"))
        self.assertEqual(result.row_count, 5)
        self.assertEqual(len(result.rows), 5)
        self.assertTrue(all(isinstance(r, Invoice) for r in result.rows))                 # the source rows themselves
        self.assertEqual(sum_money(r.payment_amount for r in result.rows), result.value)
        self.assertEqual(sum((c.amount for c in result.components), D("0")), result.value)
        self.assertEqual({c.label for c in result.components}, {"Charges", "Credits"})
        self.assertIn("charges (4 row(s))", result.calculation)
        self.assertEqual(result.filters, {"country": "Bel", "year": 2026})
        self.assertEqual(result.data_as_of, FIXED_LOADED_AT)
        self.assertTrue(any("credit row(s)" in w for w in result.warnings))
        self.assertRegex(result.result_id, r"^[0-9a-f]{12}$")
        self.assertIsNotNone(result.dataframe)
        json.dumps(result.summary())                                                        # JSON-safe summary

    def test_result_ids_are_unique_per_execution(self):
        self.assertNotEqual(run("get_debt").result_id, run("get_debt").result_id)

    def test_display_text_has_headline_calculation_scope_freshness_and_notes_in_order(self):
        text = run("get_debt", {"country": "Bel"}).display_text()
        parts = text.split("\n\n")
        self.assertTrue(parts[0].startswith("I found 5 unpaid invoice row(s) totaling €11,942.00"))
        self.assertTrue(parts[1].startswith("Calculation: €11,992.00 charges (4 row(s)) − €50.00 credits (1 row(s)) = €11,942.00"))
        self.assertEqual(parts[2], "Scope: country=Bel.")
        self.assertEqual(parts[3], "Data as of 2026-10-20 09:05 UTC.")
        self.assertTrue(parts[4].startswith("Note: 1 credit row(s)"))

    def test_non_financial_results_are_displayed_unchanged(self):
        result = run("count_projects", {"country": "Bel"})
        self.assertFalse(result.financial)
        self.assertEqual(result.display_text(), result.text)
        self.assertNotIn("Data as of", result.display_text())

    def test_empty_results_are_plain_messages(self):
        result = run("get_debt", {"project": "Spar Eta"})
        self.assertEqual(result.text, "No unpaid invoice rows match that question.")
        self.assertFalse(result.financial)


class TestDataFreshness(unittest.TestCase):
    def test_every_financial_answer_shows_the_load_timestamp(self):
        for tool, args in (("get_debt", {}), ("get_outstanding_amount", {}), ("get_paid_amount", {}), ("get_invoice", {"invoice_number": 8700}),
                           ("get_top_debt_projects", {}), ("debt_by_country", {}), ("get_unpaid_invoices", {})):
            with self.subTest(tool=tool):
                self.assertIn("Data as of 2026-10-20 09:05 UTC.", run(tool, args).display_text())

    def test_unknown_load_time_is_stated_not_invented(self):
        projects, invoices = make_data()
        with mock.patch.dict(data_clock._state, {"loaded_at": None, "source": None}):
            result = execute_tool("get_debt", {}, projects, invoices, [])
        self.assertIsNone(result.data_as_of)
        self.assertIn("Data timestamp unavailable.", result.display_text())
        self.assertTrue(any("load time is unknown" in w for w in result.warnings))

    def test_format_handles_naive_and_offset_datetimes(self):
        self.assertEqual(data_clock.format_as_of(datetime.datetime(2026, 1, 2, 3, 4)), "Data as of 2026-01-02 03:04 UTC")
        plus3 = datetime.timezone(datetime.timedelta(hours=3))
        self.assertEqual(data_clock.format_as_of(datetime.datetime(2026, 1, 2, 6, 4, tzinfo=plus3)), "Data as of 2026-01-02 03:04 UTC")

    def test_mark_loaded_records_the_time(self):
        with mock.patch.dict(data_clock._state, {"loaded_at": None, "source": None}):
            stamp = data_clock.mark_loaded("Supabase")
            self.assertEqual(data_clock.loaded_at(), stamp)
            self.assertEqual(data_clock.source(), "Supabase")
            self.assertIsNotNone(stamp.tzinfo)

    def test_load_data_records_its_load_time_and_keeps_its_return_tuple(self):
        tree = ast.parse((ROOT / "streamlit_app.py").read_text(encoding="utf-8"))
        load_data = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "load_data")
        returns = [n for n in ast.walk(load_data) if isinstance(n, ast.Return)]
        self.assertEqual(len(returns), 1)
        self.assertEqual([e.id for e in returns[0].value.elts], ["projects", "invoices", "debt_summaries", "yearly_summary", "source_name"])
        calls = [ast.unparse(n) for n in ast.walk(load_data) if isinstance(n, ast.Call)]
        self.assertIn("data_clock.mark_loaded(source_name)", calls)


class TestGrounding(unittest.TestCase):
    def test_every_amount_in_every_financial_answer_comes_from_its_toolresult(self):
        projects, invoices = build_fixture()
        checked = 0
        with mock.patch.dict(data_clock._state, {"loaded_at": FIXED_LOADED_AT, "source": "test"}):
            for case_id, tool, args in TOOL_CASES:
                result = execute_tool(tool, dict(args), projects, invoices, [], next_invoice_number=lambda: 8042)
                if result.financial:
                    checked += 1
                    self.assertEqual(result.ungrounded_amounts(), [], f"{case_id}: {result.display_text()}")
        self.assertGreater(checked, 30)

    def test_an_invented_number_is_detected(self):
        result = run("get_debt", {"country": "Bel"})
        self.assertEqual(result.ungrounded_amounts(), [])
        self.assertEqual(result.ungrounded_amounts(result.display_text() + "\n\nAlso €999,999.00."), ["€999,999.00"])
        self.assertEqual(result.ungrounded_amounts("totaling €11,942.01"), ["€11,942.01"])        # one cent off is still caught

    def test_router_refuses_to_show_an_answer_whose_figures_do_not_ground(self):
        projects, invoices = make_data()
        real = router_module.execute_tool

        def tampered(*a, **k):
            result = real(*a, **k)
            result.text = result.text.replace("€11,942.00", "€999,999.00")
            return result

        payload = '{"tool": "get_debt", "arguments": {"country": "Bel"}}'
        with fake_streamlit() as state, NewImpl().env(cfg=BASE_CFG), mock.patch("requests.post", _make_post([_gemini_ok(payload)], [])), \
                mock.patch.object(router_module, "execute_tool", tampered):
            text, df = NewImpl().smart("Belgium debt", projects, invoices, lambda q: None, lambda q, p, i, d: ("LEGACY", None))
        self.assertEqual(text, router_module.UNVERIFIED_TEXT)
        self.assertNotIn("999,999", text)
        self.assertIn("999,999.00", state["_ask_data_grounding_error"])
        self.assertIsNotNone(df)                                                                 # the table is still shown

    def test_model_supplied_numbers_never_reach_the_answer(self):
        projects, invoices = make_data()
        payload = ('{"tool": "get_debt", "arguments": {"country": "Bel", "total": 999999, "amount": "€1,000,000.00", '
                   '"value": 123456.78, "sql": "SELECT 1"}, "answer": "Belgium owes €5,000,000.00", "total": 42}')
        with fake_streamlit() as state, NewImpl().env(cfg=BASE_CFG), mock.patch("requests.post", _make_post([_gemini_ok(payload)], [])):
            text, _ = NewImpl().smart("Belgium debt", projects, invoices, lambda q: None, lambda q, p, i, d: ("LEGACY", None))
        self.assertNotIn("999999", text.replace(",", ""))
        self.assertNotIn("1,000,000", text)
        self.assertNotIn("5,000,000", text)
        self.assertNotIn("123,456", text)
        self.assertEqual(state["_ask_data_last_result"].ungrounded_amounts(text), [])
        self.assertEqual(state["_ask_data_last_result"].value, D("11942.00"))
        self.assertEqual(state["_ask_data_llm_tool_call"]["arguments"], {"project": None, "network": None, "year": None, "country": "Bel", "category": None})


class TestDebtByCountryTool(unittest.TestCase):
    def setUp(self):
        self.data = make_data()
        self.total = run("get_outstanding_amount", data=self.data)

    def test_the_tool_is_whitelisted_with_exactly_these_arguments(self):
        self.assertEqual(TOOL_SPECS["debt_by_country"], {
            "year": ("int", False), "category": ("enum:trial,y1,y2plus", False),
            "country": ("country", False), "compare_country": ("country", False)})
        projects, _ = self.data
        ok = validate_tool_call({"tool": "debt_by_country", "arguments": {"country": "Belgium", "compare_country": "Germany"}}, projects)
        self.assertEqual((ok.tool, ok.arguments["country"], ok.arguments["compare_country"]), ("debt_by_country", "Bel", "Ger"))
        he = validate_tool_call({"tool": "debt_by_country", "arguments": {"country": "בלגיה", "compare_country": "גרמניה"}}, projects)
        self.assertEqual((he.arguments["country"], he.arguments["compare_country"]), ("Bel", "Ger"))
        self.assertIsNone(validate_tool_call({"tool": "debt_by_country", "arguments": {"country": "Atlantis"}}, projects).tool)

    def test_how_much_debt_do_we_have_by_country(self):
        result = run("debt_by_country", data=self.data)
        self.assertEqual(result.value, self.total.value)
        self.assertEqual(result.row_count, self.total.row_count)
        self.assertEqual({c.label: c.amount for c in result.components}, {
            "Bel": D("11942.00"), "Ger": D("500.00"), "Hol": D("200.50"), "Unassigned (no matching project)": D("111.00")})
        self.assertEqual(sum((c.amount for c in result.components), D("0")), self.total.value)
        self.assertEqual(sum(c.row_count for c in result.components), self.total.row_count)
        self.assertIn("Total outstanding debt is €12,753.50 across 8 unpaid row(s).", result.text)
        self.assertEqual(list(result.dataframe["Country"])[-1], "Total")
        self.assertEqual(result.dataframe.iloc[-1]["Amount (€)"], self.total.value)

    def test_which_country_has_the_most_debt(self):
        result = run("debt_by_country", data=self.data)
        self.assertIn("Belgium (Bel) has the most: €11,942.00 (93.6% of the total).", result.text)
        self.assertEqual(result.components[0].label, "Bel")                                      # ranked by amount, descending

    def test_compare_belgium_and_germany(self):
        result = run("debt_by_country", {"country": "Bel", "compare_country": "Ger"}, self.data)
        self.assertEqual(result.text, "Belgium (Bel) has €11,942.00 and Germany (Ger) has €500.00 of outstanding debt; "
                                      "Belgium (Bel) has €11,442.00 more (2288.4% higher).")
        self.assertEqual([c.label for c in result.components], ["Bel", "Ger", "All other countries and unassigned"])
        self.assertEqual(sum((c.amount for c in result.components), D("0")), self.total.value)   # reconciles exactly
        self.assertEqual(result.value, self.total.value)
        self.assertEqual(result.derived_amounts[0], D("11442.00"))
        self.assertEqual(result.ungrounded_amounts(), [])

    def test_comparison_edge_cases(self):
        same = run("debt_by_country", {"country": "Bel", "compare_country": "Bel"}, self.data)
        self.assertIn("Belgium (Bel) has €11,942.00 of outstanding debt", same.text)             # same country twice = single country
        zero = run("debt_by_country", {"country": "Ger", "compare_country": "Lux"}, self.data)
        self.assertIn("Luxembourg (Lux) has none", zero.text)
        self.assertEqual(sum((c.amount for c in zero.components), D("0")), self.total.value)

    def test_single_country_shows_its_share(self):
        result = run("debt_by_country", {"country": "Ger"}, self.data)
        self.assertEqual(result.text, "Germany (Ger) has €500.00 of outstanding debt (3.9% of the €12,753.50 total) across 1 row(s).")

    def test_unmatched_rows_stay_visible_as_their_own_bucket(self):
        result = run("debt_by_country", data=self.data)
        self.assertTrue(any("Unknown Shop" in w and "€111.00" in w for w in result.warnings))
        self.assertIn("Unassigned (no matching project)", list(result.dataframe["Country"]))
        self.assertEqual(sum(c.row_count for c in result.components if c.label.startswith("Unassigned")), 1)

    def test_year_and_category_filters_still_reconcile_to_the_filtered_total(self):
        projects, invoices = self.data
        for args in ({"year": 2026}, {"category": "y1"}, {"year": 2026, "category": "y1"}):
            result = run("debt_by_country", args, self.data)
            filtered = [i for i in invoices if i.is_unpaid() and (args.get("year") is None or i.year == args["year"])
                        and (args.get("category") is None or i.maintenance_year in ("Y1", "Credit"))]   # credits count as Y1 by design
            self.assertEqual(result.value, sum_money(i.payment_amount for i in filtered), args)
            self.assertEqual(sum((c.amount for c in result.components), D("0")), result.value, args)

    def test_no_debt(self):
        result = run("debt_by_country", {"year": 1999}, self.data)
        self.assertEqual((result.text, result.value, result.financial), ("There is no outstanding debt for the selected filters.", D("0.00"), True))

    def test_total_on_the_live_shaped_data_is_exact(self):
        projects, invoices = self.data
        result = run("debt_by_country", data=self.data)
        self.assertEqual(result.value, sum_money(i.payment_amount for i in invoices if i.is_unpaid()))

    def test_router_answers_the_three_example_questions(self):
        projects, invoices = self.data
        cases = {
            "How much debt do we have by country?": '{"tool": "debt_by_country", "arguments": {}}',
            "Which country has the most debt?": '{"tool": "debt_by_country", "arguments": {}}',
            "Compare Belgium and Germany debt": '{"tool": "debt_by_country", "arguments": {"country": "Belgium", "compare_country": "Germany"}}',
        }
        for question, payload in cases.items():
            with fake_streamlit() as state, NewImpl().env(cfg=BASE_CFG), mock.patch("requests.post", _make_post([_gemini_ok(payload)], [])):
                text, df = NewImpl().smart(question, projects, invoices, lambda q: None, lambda q, p, i, d: self.fail("legacy"))
            self.assertIn("Data as of 2026-10-20 09:05 UTC.", text, question)
            self.assertIsNotNone(df)
            self.assertEqual(state["_ask_data_last_result"].tool, "debt_by_country")


class TestOutstandingAndPaidReturnTheirRows(unittest.TestCase):
    def test_outstanding_amount_returns_the_contributing_rows(self):
        result = run("get_outstanding_amount", {"country": "Bel"})
        self.assertEqual(result.row_count, 5)
        self.assertEqual(len(result.dataframe), 5)
        self.assertEqual(sum((r for r in result.dataframe["Amount (€)"]), D("0")), result.value)
        self.assertEqual({r.invoice_number for r in result.rows}, {8595, 8700, 8701, 8703, 8704})

    def test_paid_amount_returns_paid_rows_and_keeps_the_outstanding_rows_as_a_component(self):
        result = run("get_paid_amount", {"project": "CM Essen"})
        self.assertEqual(result.value, D("4668.00"))
        self.assertEqual(len(result.dataframe), 1)
        self.assertEqual({c.label: (c.amount, c.row_count) for c in result.components},
                         {"Paid": (D("4668.00"), 1), "Still outstanding": (D("3112.00"), 1)})
        outstanding_rows = next(c for c in result.components if c.label == "Still outstanding").rows
        self.assertEqual([r.invoice_number for r in outstanding_rows], [8595])
        self.assertIn("€3,112.00 across 1 invoice row(s) is still outstanding", result.text)

    def test_zero_row_scope_still_answers_zero(self):
        result = run("get_outstanding_amount", {"project": "Spar Eta"})
        self.assertEqual((result.value, result.row_count, result.dataframe), (D("0.00"), 0, None))
        self.assertIn("€0.00 across 0 invoice row(s)", result.text)


class TestDrillDown(unittest.TestCase):
    def test_result_rows_add_up_to_the_figure_they_explain(self):
        store = ResultStore()
        result = run("get_outstanding_amount")
        result_id = store.add(result)
        self.assertEqual(result_id, result.result_id)
        rows = store.get_result_rows(result_id)
        self.assertEqual((rows.result_id, rows.tool, rows.row_count), (result_id, "get_outstanding_amount", result.row_count))
        self.assertEqual(rows.total, result.value)
        self.assertEqual(len(rows.dataframe), result.row_count)
        self.assertEqual(sum(rows.dataframe["Amount (€)"], D("0")), result.value)
        self.assertEqual({r.invoice_number for r in rows.rows}, {r.invoice_number for r in result.rows})

    def test_drill_into_one_component_of_a_breakdown(self):
        store = ResultStore()
        result = run("debt_by_country")
        store.add(result)
        belgium = store.get_result_rows(result.result_id, component="Bel")
        self.assertEqual(belgium.row_count, 5)
        self.assertEqual(belgium.total, D("11942.00"))
        self.assertEqual(sum(belgium.dataframe["Country"].map(lambda c: c == "Bel")), 5)
        unassigned = store.get_result_rows(result.result_id, component="Unassigned (no matching project)")
        self.assertEqual([r.project_name for r in unassigned.rows], ["Unknown Shop"])
        with self.assertRaises(KeyError):
            store.get_result_rows(result.result_id, component="Atlantis")

    def test_drill_into_credits_and_charges(self):
        store = ResultStore()
        result = run("get_debt", {"country": "Bel"})
        store.add(result)
        self.assertEqual(store.get_result_rows(result.result_id, "Credits").total, D("-50.00"))
        self.assertEqual(store.get_result_rows(result.result_id, "Charges").total, D("11992.00"))

    def test_bank_payment_allocations_can_be_drilled_into(self):
        from tests._ask_data_harness import BANK_ALLOCATIONS, BANK_PAYMENTS
        projects, invoices = make_data()
        with NewImpl().env(bank={"payments": BANK_PAYMENTS, "allocations": BANK_ALLOCATIONS}), \
                mock.patch.dict(data_clock._state, {"loaded_at": FIXED_LOADED_AT, "source": "test"}):
            result = execute_tool("get_last_bank_payment", {}, projects, invoices, [])
        self.assertEqual(result.value, D("11670.00"))
        store = ResultStore()
        store.add(result)
        rows = store.get_result_rows(result.result_id)
        self.assertEqual((rows.row_count, rows.total), (2, D("11670.00")))

    def test_unknown_id_and_bounded_store(self):
        store = ResultStore(max_results=3)
        with self.assertRaises(UnknownResultError):
            store.get_result_rows("does-not-exist")
        ids = [store.add(run("get_debt")) for _ in range(5)]
        self.assertEqual(len(store._results), 3)
        with self.assertRaises(UnknownResultError):
            store.get_result_rows(ids[0])                                                  # evicted
        self.assertEqual(store.get_result_rows(ids[-1]).row_count, 5 if False else store.get(ids[-1]).row_count)
        self.assertIs(store.latest(), store.get(ids[-1]))

    def test_module_level_get_result_rows_uses_the_given_store(self):
        store = ResultStore()
        result = run("get_debt")
        store.add(result)
        self.assertEqual(get_result_rows(result.result_id, store=store).total, result.value)

    def test_drill_down_is_read_only(self):
        projects, invoices = make_data()
        before = [(i.invoice_number, i.project_name, i.payment_amount, i.paid, i.payment_date) for i in invoices]
        with mock.patch.dict(data_clock._state, {"loaded_at": FIXED_LOADED_AT, "source": "test"}):
            result = execute_tool("debt_by_country", {}, projects, invoices, [])
        store = ResultStore()
        store.add(result)
        for label in [None] + [c.label for c in result.components]:
            store.get_result_rows(result.result_id, label)
        self.assertEqual(before, [(i.invoice_number, i.project_name, i.payment_amount, i.paid, i.payment_date) for i in invoices])

    def test_the_router_keeps_each_result_in_a_per_session_store(self):
        projects, invoices = make_data()
        payload = '{"tool": "get_outstanding_amount", "arguments": {"country": "Bel"}}'
        with fake_streamlit() as state, NewImpl().env(cfg=BASE_CFG), mock.patch("requests.post", _make_post([_gemini_ok(payload)], [])):
            NewImpl().smart("Belgium outstanding", projects, invoices, lambda q: None, lambda q, p, i, d: ("LEGACY", None))
        last = state["_ask_data_last_result"]
        rows = state["_ask_data_result_store"].get_result_rows(last.result_id)
        self.assertEqual((rows.row_count, rows.total), (5, D("11942.00")))

    def test_an_explicit_store_can_be_injected_and_sessions_do_not_share_results(self):
        projects, invoices = make_data()
        payload = '{"tool": "get_debt", "arguments": {}}'
        shared = ResultStore()
        with fake_streamlit() as state_a, NewImpl().env(cfg=BASE_CFG), mock.patch("requests.post", _make_post([_gemini_ok(payload)], [])):
            router_module._answer_data_question_smart("q", projects, invoices, [], legacy_normalize_question=lambda q: None,
                                                      legacy_answer=lambda q, p, i, d: ("LEGACY", None), result_store=shared)
        self.assertIsNotNone(shared.latest())
        self.assertNotIn("_ask_data_result_store", state_a)
        with fake_streamlit() as state_b, NewImpl().env(cfg=BASE_CFG), mock.patch("requests.post", _make_post([_gemini_ok(payload)], [])):
            NewImpl().smart("q", projects, invoices, lambda q: None, lambda q, p, i, d: ("LEGACY", None))
        self.assertIsNot(state_b["_ask_data_result_store"], shared)
        with self.assertRaises(UnknownResultError):
            state_b["_ask_data_result_store"].get_result_rows(shared.latest().result_id)


class TestOutOfScopeStaysOutOfScope(unittest.TestCase):
    def test_no_chat_memory_party_network_trend_or_snapshot_tools_were_added(self):
        self.assertEqual(sorted(set(TOOL_SPECS) - {
            "get_invoice", "get_invoices", "get_unpaid_invoices", "get_project", "get_projects", "count_projects", "get_debt",
            "get_paid_amount", "get_outstanding_amount", "get_top_debt_projects", "get_sent_invoices", "get_next_invoice_number",
            "get_camera_statistics", "get_last_bank_payment", "get_licenses", "get_license_payment_status",
            "get_license_eop_reconciliation"}), ["debt_by_country"])
        sources = "".join(p.read_text(encoding="utf-8").lower() for p in (ROOT / "services" / "ask_data").glob("*.py"))
        for word in ("invoice_snapshots", "conversation_memory", "pps_accepted"):
            self.assertNotIn(word, sources)

    def test_drill_down_is_not_exposed_to_the_model(self):
        self.assertNotIn("get_result_rows", TOOL_SPECS)


if __name__ == "__main__":
    unittest.main()
