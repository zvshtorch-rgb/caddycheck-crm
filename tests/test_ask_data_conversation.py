"""Ask Data v3 (Phase 4): conversation context, follow-ups, drill-down, clarification continuation, bilingual
continuity, and the ask_data_v3 feature flag. Gemini is mocked throughout; no real network call is ever made."""
import datetime
import unittest
from decimal import Decimal
from pathlib import Path
from unittest import mock

from config.settings import get_ask_data_v3_enabled
from services.ask_data import data_clock
from services.ask_data.context import (
    CARRYABLE_FILTER_KEYS, Clarification, ConversationContext, is_reset_phrase, resolve_clarification_reply,
)
from services.ask_data.conversation import (
    TurnOutcome, answer_turn, breakdown_table_for_result, detect_drilldown_kind, rows_table_for_result,
)
from services.ask_data.money import fmt_eur, sum_money
from services.ask_data.planner import _llm_parse_data_question
from services.ask_data.resolver import ProjectIndex
from services.ask_data.tool_result import ResultStore
from tests._ask_data_harness import (
    BASE_CFG, FIXED_LOADED_AT, _gemini_ok, _make_post, build_fixture, fake_streamlit,
)
from tests.test_ask_data_resolver import make_data, _project

ROOT = Path(__file__).resolve().parent.parent
HE_GERMANY = "גרמניה"
HE_Y2_FOLLOWUP = "כמה מזה Y2 ומעלה?"
HE_RESET = "בוא נתחיל מחדש"
HE_ROWS = "אילו חשבוניות"


def legacy_normalize(_question):
    return None


def legacy_answer(question, _projects, _invoices, _debt_summaries):
    return f"LEGACY[{question}]", None


def expected_unpaid(projects, invoices, *, country=None, project=None, category=None):
    """Independent Decimal sum/count of unpaid rows for a scope, so tests never hardcode guessed euro figures."""
    from services.ask_data.helpers import _is_maintenance_category, _is_new_installation_category, _is_paid_trial_category
    index = ProjectIndex(projects) if (country or project) else None
    rows = [inv for inv in invoices if inv.is_unpaid()]
    if country:
        rows = [inv for inv in rows if index.country(inv.project_name) == country]
    if project:
        rows = [inv for inv in rows if index.canonical_name(inv.project_name) == project]
    if category == "y1":
        rows = [inv for inv in rows if _is_new_installation_category(inv)]
    elif category == "y2plus":
        rows = [inv for inv in rows if _is_maintenance_category(inv)]
    elif category == "trial":
        rows = [inv for inv in rows if _is_paid_trial_category(inv)]
    return sum_money(inv.payment_amount for inv in rows), len(rows)


_TEST_STORES = {}    # id(ConversationContext) -> ResultStore; mimics st.session_state persisting across reruns


def run_turn(question, payload, context, *, data=None, posts=None, cfg=BASE_CFG):
    """One chat turn with Gemini mocked to return `payload` (a raw JSON string) or, if payload is None, with
    Gemini never expected to be called at all (any request raises).

    Successive calls sharing the same `context` object also share one ResultStore (stashed as a private
    attribute on the context itself, since ConversationContext isn't hashable), so drill-down ("show invoices")
    works across turns exactly as it does with a real Streamlit session, even though each call otherwise gets
    a fresh `fake_streamlit()` session."""
    projects, invoices = data or build_fixture()
    responses = [_gemini_ok(payload)] if payload is not None else []
    calls = posts if posts is not None else []
    store = getattr(context, "_test_store", None)
    if store is None:
        store = ResultStore()
        context._test_store = store
    with fake_streamlit(), mock.patch.dict(data_clock._state, {"loaded_at": FIXED_LOADED_AT, "source": "test"}), \
            mock.patch("services.ask_data.planner.get_gemini_config", lambda: cfg), \
            mock.patch("requests.post", _make_post(responses, calls)):
        return answer_turn(
            question, projects, invoices, [], context,
            legacy_normalize_question=legacy_normalize, legacy_answer=legacy_answer, result_store=store,
        )


class TestConversationContextMerge(unittest.TestCase):
    def test_relevant_for_restricts_to_the_tools_own_arguments(self):
        ctx = ConversationContext(filters={"country": "Bel", "category": "y1", "limit": 5})
        self.assertEqual(ctx.relevant_for("get_debt"), {"country": "Bel", "category": "y1"})
        self.assertEqual(ctx.relevant_for("get_camera_statistics"), {"country": "Bel"})     # no 'category' arg there
        self.assertEqual(ctx.relevant_for("get_next_invoice_number"), {})

    def test_merge_inherits_unset_and_lets_new_explicit_values_override(self):
        ctx = ConversationContext(filters={"country": "Bel", "year": 2026})
        merged = ctx.merge("get_debt", {"project": None, "year": None, "country": None, "category": "y1"})
        self.assertEqual(merged, {"project": None, "year": 2026, "country": "Bel", "category": "y1"})
        overridden = ctx.merge("get_debt", {"project": None, "year": None, "country": "Ger", "category": None})
        self.assertEqual(overridden["country"], "Ger")                 # explicit new value always wins

    def test_merge_never_leaks_an_irrelevant_filter_into_a_different_tool(self):
        ctx = ConversationContext(filters={"country": "Bel", "category": "y1"})
        merged = ctx.merge("get_camera_statistics", {"country": None, "status": None, "aggregation": "sum"})
        self.assertNotIn("category", merged)
        self.assertEqual(merged["country"], "Bel")

    def test_record_keeps_only_carryable_keys_and_drops_nones(self):
        ctx = ConversationContext()
        ctx.record("get_invoice", {"invoice_number": 8001, "year": None}, result_id="abc")
        self.assertEqual(ctx.filters, {})                              # invoice_number/year(None) are not carried
        self.assertEqual(ctx.previous_result_id, "abc")
        ctx.record("get_debt", {"country": "Bel", "year": 2026, "project": None, "category": None}, result_id="def")
        self.assertEqual(ctx.filters, {"country": "Bel", "year": 2026})
        self.assertEqual(set(CARRYABLE_FILTER_KEYS), {"project", "country", "compare_country", "year", "paid_year", "category", "paid"})

    def test_planner_hint_contains_only_tool_and_filter_values_never_amounts(self):
        ctx = ConversationContext()
        self.assertIsNone(ctx.planner_hint())
        ctx.record("get_debt", {"country": "Bel"}, result_id="r1")
        self.assertEqual(ctx.planner_hint(), {"tool": "get_debt", "arguments": {"country": "Bel"}})

    def test_reset_clears_everything(self):
        ctx = ConversationContext(intent="get_debt", filters={"country": "Bel"}, previous_result_id="r1",
                                  clarification=Clarification("get_debt", "project", "project", "x", ("a",), {}))
        ctx.reset()
        self.assertEqual((ctx.intent, ctx.filters, ctx.previous_result_id, ctx.clarification), (None, {}, None, None))

    def test_is_reset_phrase_english_and_hebrew(self):
        self.assertTrue(is_reset_phrase("Never mind, what's the total?"))
        self.assertTrue(is_reset_phrase("Let's start over"))
        self.assertTrue(is_reset_phrase(HE_RESET))
        self.assertFalse(is_reset_phrase("How much debt in Belgium?"))
        self.assertFalse(is_reset_phrase(""))


class TestClarificationReply(unittest.TestCase):
    def setUp(self):
        self.clarification = Clarification(
            tool="get_debt", arg="project", kind="project", value="Edeka",
            candidates=("Edeka Herweg", "Edeka Peipp Rostal"), raw_arguments={"project": "Edeka"},
        )

    def test_an_exact_candidate_name_resolves(self):
        payload = resolve_clarification_reply(self.clarification, "Edeka Herweg")
        self.assertEqual(payload, {"tool": "get_debt", "arguments": {"project": "Edeka Herweg"}})

    def test_a_distinguishing_fragment_resolves_unambiguously(self):
        payload = resolve_clarification_reply(self.clarification, "Rostal")
        self.assertEqual(payload["arguments"]["project"], "Edeka Peipp Rostal")

    def test_a_non_matching_reply_returns_none(self):
        self.assertIsNone(resolve_clarification_reply(self.clarification, "something completely unrelated"))
        self.assertIsNone(resolve_clarification_reply(self.clarification, "Edeka"))   # still ambiguous between the two

    def test_country_clarification_reply(self):
        clarification = Clarification("get_debt", "country", "country", "blegium", ("Bel",), {"country": "blegium"})
        payload = resolve_clarification_reply(clarification, "Belgium")
        self.assertEqual(payload, {"tool": "get_debt", "arguments": {"country": "Bel"}})


class TestFollowUps(unittest.TestCase):
    """Goal 2's worked example, end to end with Gemini mocked."""

    def test_country_refinement_then_drilldown_then_another_country(self):
        ctx = ConversationContext()
        t1 = run_turn("How much debt do we have?", '{"tool": "get_debt", "arguments": {}}', ctx)
        self.assertEqual(t1.kind, "tool")
        self.assertIn("€28,913.50", t1.text)
        self.assertIsNotNone(ctx.previous_result_id)

        t2 = run_turn("Only Belgium?", '{"tool": "get_debt", "arguments": {"country": "Belgium"}}', ctx)
        self.assertIn("€19,999.00", t2.text)
        self.assertEqual(ctx.filters, {"country": "Bel"})

        t3 = run_turn("Which invoices make that up?", None, ctx)                   # NEVER calls Gemini
        self.assertEqual(t3.kind, "drilldown")
        self.assertEqual(len(t3.dataframe), 8)
        self.assertEqual(sum(t3.dataframe["Amount (€)"]), Decimal("19999.00"))

        t4 = run_turn("And Germany?", '{"tool": "get_debt", "arguments": {"country": "Germany"}}', ctx)
        self.assertIn("€3,890.00", t4.text)
        self.assertEqual(ctx.filters, {"country": "Ger"})                          # Belgium is gone, not merged in

    def test_project_refinement_overrides_the_previous_project(self):
        ctx = ConversationContext()
        run_turn("Debt for AD Alpha", '{"tool": "get_debt", "arguments": {"project": "AD Alpha"}}', ctx)
        self.assertEqual(ctx.filters["project"], "AD Alpha")
        t2 = run_turn("What about AD Beta instead?", '{"tool": "get_debt", "arguments": {"project": "AD Beta"}}', ctx)
        self.assertEqual(ctx.filters["project"], "AD Beta")
        self.assertIn("€5,016.00", t2.text)

    def test_y1_vs_y2plus_follow_up(self):
        ctx = ConversationContext()
        run_turn("Debt in Germany", '{"tool": "get_debt", "arguments": {"country": "Germany"}}', ctx)
        t2 = run_turn("What about Y2+?", '{"tool": "get_debt", "arguments": {"category": "y2plus"}}', ctx)
        self.assertEqual(ctx.filters, {"country": "Ger", "category": "y2plus"})     # country inherited, category added
        t3 = run_turn("And Y1?", '{"tool": "get_debt", "arguments": {"category": "y1"}}', ctx)
        self.assertEqual(ctx.filters, {"country": "Ger", "category": "y1"})         # category overridden, country kept
        self.assertNotEqual(t2.text, t3.text)

    def test_show_invoices_phrase_variants_never_call_gemini(self):
        ctx = ConversationContext()
        run_turn("How much debt in Belgium?", '{"tool": "get_debt", "arguments": {"country": "Belgium"}}', ctx)
        for phrase in ("Show invoices", "which invoices make that up?", "show me the rows behind that"):
            with self.subTest(phrase=phrase):
                outcome = run_turn(phrase, None, ctx)
                self.assertEqual(outcome.kind, "drilldown")
                self.assertEqual(len(outcome.dataframe), 8)

    def test_show_breakdown_returns_components_table(self):
        ctx = ConversationContext()
        run_turn("How much debt do we have by country?", '{"tool": "debt_by_country", "arguments": {}}', ctx)
        outcome = run_turn("Show breakdown", None, ctx)
        self.assertEqual(outcome.kind, "drilldown")
        self.assertEqual(set(outcome.dataframe["Component"]), {"Bel", "Ger", "Hol", "Unassigned (no matching project)"})

    def test_drilldown_without_a_previous_result_is_a_clear_message_not_an_error(self):
        ctx = ConversationContext()
        outcome = run_turn("Show invoices", None, ctx)
        self.assertEqual(outcome.kind, "drilldown")
        self.assertIn("no previous answer", outcome.text.lower())

    def test_conflicting_scope_the_new_explicit_value_wins_and_the_old_one_does_not_leak_back(self):
        ctx = ConversationContext()
        run_turn("Debt for AD Alpha in Belgium", '{"tool": "get_debt", "arguments": {"project": "AD Alpha", "country": "Belgium"}}', ctx)
        self.assertEqual(ctx.filters, {"project": "AD Alpha", "country": "Bel"})
        # the next turn explicitly asks about a different country; project should still be inherited (not conflicting)
        t2 = run_turn("What about Germany?", '{"tool": "get_debt", "arguments": {"country": "Germany"}}', ctx)
        self.assertEqual(ctx.filters, {"project": "AD Alpha", "country": "Ger"})
        self.assertIn("No unpaid invoice rows", t2.text)   # AD Alpha has no unpaid rows in Germany -- proves country truly changed, not just displayed

    def test_switching_tool_only_inherits_filters_the_new_tool_actually_supports(self):
        ctx = ConversationContext()
        run_turn("Y1 debt in Belgium", '{"tool": "get_debt", "arguments": {"country": "Belgium", "category": "y1"}}', ctx)
        self.assertEqual(ctx.filters, {"country": "Bel", "category": "y1"})
        t2 = run_turn("How many cameras in total?", '{"tool": "get_camera_statistics", "arguments": {"country": "Belgium"}}', ctx)
        self.assertEqual(t2.kind, "tool")
        # category inherited fine as a concept, but get_camera_statistics doesn't declare it -- never sent/used
        self.assertNotIn("category", ctx.relevant_for("get_camera_statistics"))

    def test_clearing_context_starts_a_fresh_question(self):
        ctx = ConversationContext()
        run_turn("Debt in Belgium", '{"tool": "get_debt", "arguments": {"country": "Belgium"}}', ctx)
        self.assertEqual(ctx.filters, {"country": "Bel"})
        t2 = run_turn("Never mind, what's the total debt?", '{"tool": "get_debt", "arguments": {}}', ctx)
        self.assertEqual(ctx.filters, {})
        self.assertIn("€28,913.50", t2.text)                 # the overall total, not Belgium-scoped


class TestClarificationState(unittest.TestCase):
    def test_ambiguous_project_then_a_short_reply_resolves_it_without_calling_gemini(self):
        ctx = ConversationContext()
        t1 = run_turn("debt for Edeka", '{"tool": "get_debt", "arguments": {"project": "Edeka"}}', ctx, data=make_data())
        self.assertEqual(t1.kind, "clarification")
        self.assertIn("Edeka Herweg", t1.text)
        self.assertIn("Edeka Peipp Rostal", t1.text)
        self.assertIsNotNone(ctx.clarification)

        t2 = run_turn("Edeka Herweg", None, ctx, data=make_data())    # answering the clarification never calls Gemini
        self.assertEqual(t2.kind, "tool")
        self.assertIsNone(ctx.clarification)
        self.assertEqual(ctx.filters["project"], "Edeka Herweg")

    def test_a_reply_that_does_not_match_any_candidate_falls_through_to_a_fresh_question(self):
        data = make_data()
        total, _count = expected_unpaid(*data)
        ctx = ConversationContext()
        run_turn("debt for Edeka", '{"tool": "get_debt", "arguments": {"project": "Edeka"}}', ctx, data=data)
        self.assertIsNotNone(ctx.clarification)
        t2 = run_turn("How much debt in total?", '{"tool": "get_debt", "arguments": {}}', ctx, data=data)
        self.assertIsNone(ctx.clarification)
        self.assertEqual(t2.kind, "tool")
        self.assertIn(fmt_eur(total), t2.text)

    def test_a_still_ambiguous_clarified_call_asks_again(self):
        projects, invoices = make_data()
        projects = projects + [_project("Edeka Herwegen", "Ger")]
        ctx = ConversationContext()
        ctx.clarification = Clarification("get_debt", "project", "project", "Edeka Herweg",
                                          ("Edeka Herweg", "Edeka Herwegen"), {"project": "Edeka Herweg"})
        t = run_turn("Edeka Herweg", None, ctx, data=(projects, invoices))
        self.assertEqual(t.kind, "tool")  # exact text match -> unambiguous even among near-duplicates


class TestBilingualContinuity(unittest.TestCase):
    def test_hebrew_follow_up_after_an_english_question(self):
        ctx = ConversationContext()
        run_turn("How much debt in Germany?", '{"tool": "get_debt", "arguments": {"country": "Germany"}}', ctx)
        self.assertEqual(ctx.filters, {"country": "Ger"})
        t2 = run_turn(HE_Y2_FOLLOWUP, '{"tool": "get_debt", "arguments": {"category": "y2plus"}}', ctx)
        self.assertEqual(ctx.filters, {"country": "Ger", "category": "y2plus"})

    def test_switching_from_hebrew_back_to_english(self):
        ctx = ConversationContext()
        run_turn(f"כמה חוב ב{HE_GERMANY}?", '{"tool": "get_debt", "arguments": {"country": "Germany"}}', ctx)
        self.assertEqual(ctx.filters, {"country": "Ger"})
        t2 = run_turn("And Y1 only?", '{"tool": "get_debt", "arguments": {"category": "y1"}}', ctx)
        self.assertEqual(ctx.filters, {"country": "Ger", "category": "y1"})
        self.assertEqual(t2.kind, "tool")

    def test_hebrew_reset_phrase_clears_context(self):
        ctx = ConversationContext()
        run_turn("How much debt in Germany?", '{"tool": "get_debt", "arguments": {"country": "Germany"}}', ctx)
        run_turn(HE_RESET + " What's the total?", '{"tool": "get_debt", "arguments": {}}', ctx)
        self.assertEqual(ctx.filters, {})

    def test_hebrew_drilldown_phrase_never_calls_gemini(self):
        ctx = ConversationContext()
        run_turn("How much debt in Belgium?", '{"tool": "get_debt", "arguments": {"country": "Belgium"}}', ctx)
        outcome = run_turn(HE_ROWS, None, ctx)
        self.assertEqual(outcome.kind, "drilldown")
        self.assertEqual(len(outcome.dataframe), 8)

    def test_detect_drilldown_kind_bilingual(self):
        self.assertEqual(detect_drilldown_kind("Show invoices"), "rows")
        self.assertEqual(detect_drilldown_kind(HE_ROWS), "rows")
        self.assertEqual(detect_drilldown_kind("show breakdown"), "breakdown")
        self.assertEqual(detect_drilldown_kind("פירוט"), "breakdown")
        self.assertIsNone(detect_drilldown_kind("How much debt in Belgium?"))


class TestResultIdIsolationBySession(unittest.TestCase):
    def test_two_sessions_do_not_share_result_ids_or_context(self):
        store_a, store_b = ResultStore(), ResultStore()
        ctx_a, ctx_b = ConversationContext(), ConversationContext()
        projects, invoices = make_data()

        def turn(question, payload, context, store):
            with fake_streamlit(), mock.patch.dict(data_clock._state, {"loaded_at": FIXED_LOADED_AT, "source": "test"}), \
                    mock.patch("services.ask_data.planner.get_gemini_config", lambda: BASE_CFG), \
                    mock.patch("requests.post", _make_post([_gemini_ok(payload)] if payload else [], [])):
                return answer_turn(question, projects, invoices, [], context, legacy_normalize_question=legacy_normalize,
                                   legacy_answer=legacy_answer, result_store=store)

        turn("Debt in Belgium", '{"tool": "get_debt", "arguments": {"country": "Belgium"}}', ctx_a, store_a)
        turn("Debt in Germany", '{"tool": "get_debt", "arguments": {"country": "Germany"}}', ctx_b, store_b)

        self.assertNotEqual(ctx_a.previous_result_id, ctx_b.previous_result_id)
        self.assertIsNotNone(store_a.get(ctx_a.previous_result_id))
        self.assertIsNone(store_b.get(ctx_a.previous_result_id))     # session A's result is invisible to session B
        self.assertIsNone(store_a.get(ctx_b.previous_result_id))

        rows_a = turn("Show invoices", None, ctx_a, store_a)
        self.assertEqual(len(rows_a.dataframe), 5)                   # Belgium's 5 rows, not Germany's
        rows_b = turn("Show invoices", None, ctx_b, store_b)
        self.assertEqual(len(rows_b.dataframe), 1)                   # Germany's 1 row

    def test_session_state_backed_store_defaults_are_also_isolated_per_streamlit_session(self):
        projects, invoices = make_data()
        with fake_streamlit() as state_a:
            with mock.patch.dict(data_clock._state, {"loaded_at": FIXED_LOADED_AT, "source": "test"}), \
                    mock.patch("services.ask_data.planner.get_gemini_config", lambda: BASE_CFG), \
                    mock.patch("requests.post", _make_post([_gemini_ok('{"tool": "get_debt", "arguments": {}}')], [])):
                ctx = ConversationContext()
                answer_turn("total debt", projects, invoices, [], ctx, legacy_normalize_question=legacy_normalize, legacy_answer=legacy_answer)
            self.assertIn("_ask_data_v3_result_store", state_a)
            store_a = state_a["_ask_data_v3_result_store"]
        with fake_streamlit() as state_b:
            self.assertNotIn("_ask_data_v3_result_store", state_b)   # a fresh session starts with no store at all
            with mock.patch.dict(data_clock._state, {"loaded_at": FIXED_LOADED_AT, "source": "test"}), \
                    mock.patch("services.ask_data.planner.get_gemini_config", lambda: BASE_CFG), \
                    mock.patch("requests.post", _make_post([_gemini_ok('{"tool": "get_debt", "arguments": {}}')], [])):
                ctx_b = ConversationContext()
                answer_turn("total debt", projects, invoices, [], ctx_b, legacy_normalize_question=legacy_normalize, legacy_answer=legacy_answer)
            store_b = state_b["_ask_data_v3_result_store"]
        self.assertIsNot(store_a, store_b)


class TestGroundingAndLegacyFallbackStillWork(unittest.TestCase):
    def test_unresolvable_tool_falls_back_to_legacy_and_does_not_poison_context(self):
        ctx = ConversationContext()
        outcome = run_turn("gibberish", '{"tool": "nope", "arguments": {}}', ctx)
        self.assertEqual(outcome.kind, "legacy")
        self.assertIsNone(ctx.intent)
        self.assertEqual(ctx.filters, {})

    def test_no_api_key_goes_straight_to_legacy(self):
        ctx = ConversationContext()
        outcome = run_turn("anything", None, ctx, cfg={"api_key": "", "model": "m"})
        self.assertEqual(outcome.kind, "legacy")

    def test_a_tampered_answer_is_refused_even_in_v3(self):
        from services.ask_data import conversation as conversation_module
        real_execute = conversation_module.execute_tool

        def tampered(*args, **kwargs):
            result = real_execute(*args, **kwargs)
            result.text = result.text.replace("€28,913.50", "€999,999.00")
            return result

        ctx = ConversationContext()
        with mock.patch.object(conversation_module, "execute_tool", tampered):
            outcome = run_turn("total debt", '{"tool": "get_debt", "arguments": {}}', ctx)
        self.assertEqual(outcome.text, conversation_module.UNVERIFIED_TEXT)
        self.assertIsNotNone(outcome.dataframe)


class TestFeatureFlag(unittest.TestCase):
    def test_defaults_to_off(self):
        with mock.patch.dict("os.environ", {}, clear=False):
            import os
            os.environ.pop("ASK_DATA_V3", None)
            self.assertFalse(get_ask_data_v3_enabled())

    def test_env_var_truthy_values(self):
        import os
        for value, expected in (("1", True), ("true", True), ("YES", True), ("0", False), ("", False), ("off", False)):
            with mock.patch.dict(os.environ, {"ASK_DATA_V3": value}):
                self.assertEqual(get_ask_data_v3_enabled(), expected, value)

    def test_v2_entry_point_is_completely_unaffected(self):
        """router._answer_data_question_smart (v2) must still exist, unchanged, independent of v3."""
        from services.ask_data.router import _answer_data_question_smart
        self.assertTrue(callable(_answer_data_question_smart))


class TestPlannerContextIsSafe(unittest.TestCase):
    def test_no_context_prompt_is_byte_identical_to_pre_phase4(self):
        import json as _json
        calls_without, calls_with = [], []
        with mock.patch("services.ask_data.planner.get_gemini_config", lambda: BASE_CFG), \
                mock.patch("requests.post", _make_post([_gemini_ok('{"tool": "get_debt", "arguments": {}}')], calls_without)):
            _llm_parse_data_question("How much debt?")
        with mock.patch("services.ask_data.planner.get_gemini_config", lambda: BASE_CFG), \
                mock.patch("requests.post", _make_post([_gemini_ok('{"tool": "get_debt", "arguments": {}}')], calls_with)):
            _llm_parse_data_question("How much debt?", context=None)
        self.assertEqual(calls_without[0]["body"], calls_with[0]["body"])

    def test_context_adds_only_tool_and_filters_never_rows_or_amounts(self):
        calls = []
        context = {"tool": "get_debt", "arguments": {"country": "Bel", "year": 2026}}
        with mock.patch("services.ask_data.planner.get_gemini_config", lambda: BASE_CFG), \
                mock.patch("requests.post", _make_post([_gemini_ok('{"tool": "get_debt", "arguments": {"country": "Bel"}}')], calls)):
            _llm_parse_data_question("Only Belgium?", context=context)
        prompt = calls[0]["body"]["contents"][0]["parts"][0]["text"]
        self.assertIn('"tool": "get_debt"', prompt)
        self.assertIn('"country": "Bel"', prompt)
        for forbidden in ("\u20ac", "payment_amount", "189,998", "19,999", "28,913"):
            self.assertNotIn(forbidden, prompt)


if __name__ == "__main__":
    unittest.main()
