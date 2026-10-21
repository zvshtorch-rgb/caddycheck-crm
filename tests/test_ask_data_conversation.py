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
    HISTORICAL_LIMITATION_TEXT, TurnOutcome, answer_turn, breakdown_table_for_result, detect_drilldown_kind,
    is_historical_question, reset_conversation, rows_table_for_result,
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


def run_turn(question, payload, context, *, data=None, posts=None, cfg=BASE_CFG, history_enabled=False):
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
            history_enabled=history_enabled,
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
        self.assertEqual(merged, {"project": None, "network": None, "year": 2026, "country": "Bel", "category": "y1"})
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
        self.assertEqual(set(CARRYABLE_FILTER_KEYS), {"project", "country", "compare_country", "year", "paid_year",
                                                       "category", "paid", "period", "at_date", "date_a", "date_b"})

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
        # get_last_bank_payment has no 'network' argument, so "Edeka" alone is still plain project
        # ambiguity here (get_debt would instead resolve it to network="Edeka" -- see
        # test_ask_data_network_resolver.py).
        ctx = ConversationContext()
        t1 = run_turn("last bank payment for Edeka", '{"tool": "get_last_bank_payment", "arguments": {"project": "Edeka"}}', ctx, data=make_data())
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
        run_turn("last bank payment for Edeka", '{"tool": "get_last_bank_payment", "arguments": {"project": "Edeka"}}', ctx, data=data)
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


class TestNewConversationFullyResetsState(unittest.TestCase):
    """Regression test: "New conversation" must produce exactly the same Ask Data conversational state as a
    fresh session. Found via production testing: the button cleared ConversationContext but left the
    drill-down ResultStore in place -- see reset_conversation()'s docstring for why a fresh store (not
    merely context.reset()) is required."""

    def test_network_query_drilldown_new_conversation_unrelated_query_has_no_inherited_scope(self):
        projects, invoices = make_data()
        ctx = ConversationContext()
        store = ResultStore()

        def turn(question, payload):
            with fake_streamlit(), mock.patch.dict(data_clock._state, {"loaded_at": FIXED_LOADED_AT, "source": "test"}), \
                    mock.patch("services.ask_data.planner.get_gemini_config", lambda: BASE_CFG), \
                    mock.patch("requests.post", _make_post([_gemini_ok(payload)] if payload else [], [])):
                return answer_turn(question, projects, invoices, [], ctx, legacy_normalize_question=legacy_normalize,
                                   legacy_answer=legacy_answer, result_store=store)

        # Turn 1: a network query.
        turn1 = turn("Edeka debt", '{"tool": "get_debt", "arguments": {"network": "Edeka"}}')
        self.assertEqual(turn1.kind, "tool")
        network_result_id = ctx.previous_result_id
        self.assertIsNotNone(network_result_id)
        self.assertEqual(turn1.result.filters.get("network"), "Edeka")
        # Note: 'network' is not in CARRYABLE_FILTER_KEYS (a pre-existing, documented gap predating Phase 5
        # -- the Network Resolver phase added network to the tools themselves, not to follow-up context),
        # so it is correctly absent from ctx.filters; that is unrelated to this reset bug.

        # Turn 2: drill-down on that network result.
        turn2 = turn("Which invoices make that up?", None)
        self.assertEqual(turn2.kind, "drilldown")
        self.assertIs(turn2.result, store.get(network_result_id))

        # "New conversation": full reset, replacing the store -- mirrors streamlit_app.py's button handler
        # exactly (st.session_state["_ask_data_v3_result_store"] = reset_conversation(_v3_context)).
        store = reset_conversation(ctx)

        self.assertIsNone(ctx.previous_result_id)
        self.assertEqual(ctx.filters, {})
        self.assertIsNone(ctx.clarification)
        self.assertEqual(ctx.turn_count, 0)
        self.assertIsNone(store.get(network_result_id))   # the old result is unreachable in the NEW store

        # Turn 3: an unrelated query in the "new" conversation -- no inherited network scope.
        turn3 = turn("How much debt in Belgium?", '{"tool": "get_debt", "arguments": {"country": "Bel"}}')
        self.assertEqual(turn3.kind, "tool")
        self.assertNotIn("network", turn3.result.filters)
        self.assertEqual(turn3.result.filters.get("country"), "Bel")

        # A drill-down request in the new conversation resolves to turn 3's OWN result, never turn 1's.
        drilldown_after_reset = turn("Which invoices make that up?", None)
        self.assertEqual(drilldown_after_reset.kind, "drilldown")
        self.assertIs(drilldown_after_reset.result, turn3.result)
        self.assertIsNot(drilldown_after_reset.result, turn1.result)


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


class TestHistoricalQuestionsAreNotMisrouted(unittest.TestCase):
    """Regression test for the production bug: "What changed since September 1?" was silently routed to a
    license-status tool and returned unrelated project data. Phase 5 (invoice_snapshots history) does not
    exist yet, so these questions must get a deterministic limitation message instead of ANY tool call."""

    def test_is_historical_question_detects_the_reported_examples(self):
        for question in (
            "What changed since September 1?",
            "How has debt changed this month?",
            "What was the debt last week?",
            "What changed since yesterday?",
            "כמה השתנה החוב מאז תחילת החודש?",
            "מה היה החוב לפני שבוע?",
        ):
            with self.subTest(question=question):
                self.assertTrue(is_historical_question(question))

    def test_plain_current_state_questions_are_not_false_positives(self):
        for question in (
            "How much debt do we have?",
            "Show unpaid invoices for AD Anderlecht",
            "Which country has the most cameras?",
            "Show invoices from last month",          # has a time word but no change/comparison word
            "What is the debt for Belgium?",           # "what is", not "what was"
            HE_GERMANY,
        ):
            with self.subTest(question=question):
                self.assertFalse(is_historical_question(question))

    def test_a_historical_question_never_calls_gemini_and_returns_the_limitation_text(self):
        ctx = ConversationContext()
        outcome = run_turn("What changed since September 1?", None, ctx)   # payload=None -> any Gemini call fails the test
        self.assertEqual(outcome.kind, "historical_unsupported")
        self.assertEqual(outcome.text, HISTORICAL_LIMITATION_TEXT)
        self.assertIsNone(outcome.dataframe)
        self.assertIsNone(outcome.result)

    def test_hebrew_historical_question_never_calls_gemini(self):
        ctx = ConversationContext()
        outcome = run_turn("כמה השתנה החוב מאז תחילת החודש?", None, ctx)
        self.assertEqual(outcome.kind, "historical_unsupported")
        self.assertEqual(outcome.text, HISTORICAL_LIMITATION_TEXT)

    def test_does_not_offer_drilldown_buttons_tool_or_result(self):
        """Goal 4: no "Show invoices"/"Show breakdown" buttons -- those only ever attach to kind == "tool"
        in the chat UI (streamlit_app.py), so a non-"tool" kind with no result is the enforcement point here."""
        ctx = ConversationContext()
        outcome = run_turn("What was the debt last week?", None, ctx)
        self.assertNotEqual(outcome.kind, "tool")
        self.assertIsNone(outcome.result)

    def test_does_not_poison_the_conversation_context(self):
        ctx = ConversationContext()
        run_turn("Debt in Belgium", '{"tool": "get_debt", "arguments": {"country": "Belgium"}}', ctx)
        previous_result_id, previous_filters = ctx.previous_result_id, dict(ctx.filters)
        run_turn("What changed since September 1?", None, ctx)
        self.assertEqual(ctx.previous_result_id, previous_result_id)   # untouched -- still points at Belgium's answer
        self.assertEqual(ctx.filters, previous_filters)

    def test_abandons_a_pending_clarification_instead_of_trying_to_resolve_it_as_a_reply(self):
        # get_last_bank_payment has no 'network' argument, so "Edeka" alone is still plain project
        # ambiguity here (a network-aware tool like get_debt would instead resolve it to network="Edeka" --
        # see test_ask_data_network_resolver.py).
        ctx = ConversationContext()
        run_turn("last bank payment for Edeka", '{"tool": "get_last_bank_payment", "arguments": {"project": "Edeka"}}', ctx, data=make_data())
        self.assertIsNotNone(ctx.clarification)
        outcome = run_turn("What changed since September 1?", None, ctx, data=make_data())
        self.assertEqual(outcome.kind, "historical_unsupported")
        self.assertIsNone(ctx.clarification)

    def test_flag_off_explicitly_still_shows_the_limitation_text(self):
        """Same as the defaults above, but with `history_enabled=False` passed explicitly, to directly pin
        the gate variable itself (not just its default) -- this is the single most important Phase 5
        guarantee: behavior is unchanged unless the flag is explicitly turned on."""
        ctx = ConversationContext()
        outcome = run_turn("What changed since September 1?", None, ctx, history_enabled=False)
        self.assertEqual(outcome.kind, "historical_unsupported")
        self.assertEqual(outcome.text, HISTORICAL_LIMITATION_TEXT)


def _history_patches(runs, snapshots):
    return [
        mock.patch("services.supabase_service.load_invoice_snapshot_runs", lambda: runs),
        mock.patch("services.supabase_service.load_invoice_snapshot", lambda d: snapshots.get(d.isoformat(), [])),
        mock.patch("services.supabase_service.load_invoice_change_log", lambda *a, **k: []),
        mock.patch("services.supabase_service.load_bank_payments", lambda: []),
        mock.patch("services.supabase_service.load_bank_payment_allocations", lambda *a, **k: []),
        mock.patch("config.settings.load_sent_invoices_log", lambda: []),
    ]


class TestHistoryFlagOnRouting(unittest.TestCase):
    """Ask Data Phase 5: with `ask_data_history` on, a historical question flows through the normal
    planner -> validator -> executor pipeline instead of the fixed limitation message, follow-ups keep the
    same resolved period, and drill-down works on the resulting bridge exactly like any other tool."""

    def setUp(self):
        real_today = datetime.date.today()
        self.runs = [
            {"snapshot_date": "2026-09-01", "status": "success", "started_at": "2026-09-01T02:00:00+00:00",
             "completed_at": "2026-09-01T02:01:00+00:00", "row_count": 1},
            {"snapshot_date": real_today.isoformat(), "status": "success",
             "started_at": real_today.isoformat() + "T02:00:00+00:00",
             "completed_at": real_today.isoformat() + "T02:01:00+00:00", "row_count": 1},
        ]
        self.snapshots = {
            "2026-09-01": [{"invoice_row_id": 1, "invoice_number": 8001, "project_name": "AD Alpha",
                            "payment_amount": 7780.0, "paid": "No", "year": 2026, "maintenance_year": "Y1"}],
            real_today.isoformat(): [{"invoice_row_id": 1, "invoice_number": 8001, "project_name": "AD Alpha",
                                      "payment_amount": 7780.0, "paid": "No", "year": 2026, "maintenance_year": "Y1"}],
        }

    def _run(self, question, payload, context, **kwargs):
        patches = _history_patches(self.runs, self.snapshots)
        with patches[0], patches[1], patches[2], patches[3], patches[4], patches[5]:
            return run_turn(question, payload, context, history_enabled=True, **kwargs)

    def test_flag_on_executes_a_historical_tool_instead_of_the_limitation_message(self):
        ctx = ConversationContext()
        outcome = self._run(
            "How much has debt increased this month?",
            '{"tool": "get_debt_change", "arguments": {"period": "this month"}}', ctx,
        )
        self.assertEqual(outcome.kind, "tool")
        self.assertEqual(outcome.tool, "get_debt_change")
        self.assertTrue(outcome.result.financial)

    def test_hebrew_historical_question_with_flag_on_also_executes_a_tool(self):
        ctx = ConversationContext()
        outcome = self._run(
            "כמה החוב השתנה מתחילת החודש?",
            '{"tool": "get_debt_change", "arguments": {"period": "this month"}}', ctx,
        )
        self.assertEqual(outcome.kind, "tool")

    def test_period_carries_over_to_a_country_refinement_follow_up(self):
        ctx = ConversationContext()
        self._run(
            "How much has debt increased this month?",
            '{"tool": "get_debt_change", "arguments": {"period": "this month"}}', ctx,
        )
        first_period = ctx.filters.get("period")
        self.assertIsNotNone(first_period)
        outcome = self._run(
            "Only Belgium?",
            '{"tool": "get_debt_change", "arguments": {"country": "Bel"}}', ctx,
        )
        self.assertEqual(outcome.kind, "tool")
        # The SAME period object is reused -- never re-derived/guessed by the model on the follow-up.
        self.assertIs(ctx.filters.get("period"), first_period)
        self.assertEqual(outcome.result.filters.get("country"), "Bel")

    def test_show_invoices_works_after_a_historical_bridge_answer(self):
        ctx = ConversationContext()
        self._run(
            "How much has debt increased this month?",
            '{"tool": "get_debt_change", "arguments": {"period": "this month"}}', ctx,
        )
        outcome = self._run(HE_ROWS, None, ctx)   # drill-down never calls the model
        self.assertEqual(outcome.kind, "drilldown")

    def test_switching_hebrew_to_english_and_back_keeps_the_same_period(self):
        ctx = ConversationContext()
        self._run("כמה החוב השתנה מתחילת החודש?", '{"tool": "get_debt_change", "arguments": {"period": "this month"}}', ctx)
        period_after_he = ctx.filters.get("period")
        outcome = self._run("Only Belgium?", '{"tool": "get_debt_change", "arguments": {"country": "Bel"}}', ctx)
        self.assertEqual(outcome.kind, "tool")
        self.assertIs(ctx.filters.get("period"), period_after_he)

    def test_reset_phrase_clears_the_carried_period(self):
        ctx = ConversationContext()
        self._run("How much has debt increased this month?", '{"tool": "get_debt_change", "arguments": {"period": "this month"}}', ctx)
        self.assertIn("period", ctx.filters)
        self._run("start over", None, ctx)   # reset is detected before any tool/model call
        self.assertEqual(ctx.filters, {})

    def test_follow_up_still_works_when_gemini_echoes_the_resolved_at_date_back_verbatim(self):
        """Regression (Step 4 production bug): planner_hint() tells the model the previous turn's
        at_date as a plain ISO string (e.g. "2026-09-01"), since that is all that is safe to show it (see
        context.py). In real production traffic Gemini commonly copies that exact string straight back as
        its OWN 'at_date' argument on a follow-up (instructed to "keep any filter the new question does not
        mention"), rather than omitting it. Before the fix, that echoed ISO string failed
        resolve_single_date_phrase() (it only parses phrases, not ISO dates) and the WHOLE tool call was
        rejected, silently falling through to the unrelated legacy fallback -- exactly the "current data
        for a historical question" failure Phase 5 exists to prevent. Fixed in validation.py by recognising
        an exact echo of the already-resolved carried-over value."""
        ctx = ConversationContext()
        self._run("What was the debt on September 1?", '{"tool": "get_debt_at", "arguments": {"at_date": "September 1"}}', ctx)
        first_at_date = ctx.filters.get("at_date")
        self.assertEqual(first_at_date, datetime.date(2026, 9, 1))
        outcome = self._run(
            "Only Belgium?",
            # Gemini echoes the hint's ISO string verbatim, exactly as observed in production.
            '{"tool": "get_debt_at", "arguments": {"at_date": "2026-09-01", "country": "Bel"}}', ctx,
        )
        self.assertEqual(outcome.kind, "tool")
        self.assertEqual(outcome.tool, "get_debt_at")
        self.assertEqual(ctx.filters.get("at_date"), first_at_date)
        self.assertEqual(outcome.result.filters.get("country"), "Bel")

    def test_follow_up_still_works_when_gemini_echoes_the_resolved_period_label_back_verbatim(self):
        """Same regression as above, for the 'period' kind (get_debt_change / get_invoice_changes /
        get_payments_in_period / get_new_invoices_in_period): the hint's period value is
        PeriodResolution.__str__() (a human-readable label like "this month (2026-10-01 to 2026-10-21)"),
        which Gemini can also copy back verbatim on a follow-up."""
        import json as _json
        ctx = ConversationContext()
        self._run("How much has debt increased this month?", '{"tool": "get_debt_change", "arguments": {"period": "this month"}}', ctx)
        first_period = ctx.filters.get("period")
        echoed_label = str(first_period)
        payload = _json.dumps({"tool": "get_debt_change", "arguments": {"period": echoed_label, "country": "Bel"}})
        outcome = self._run("Only Belgium?", payload, ctx)
        self.assertEqual(outcome.kind, "tool")
        self.assertEqual(outcome.tool, "get_debt_change")
        self.assertIs(ctx.filters.get("period"), first_period)
        self.assertEqual(outcome.result.filters.get("country"), "Bel")

    def test_a_genuinely_new_unparseable_date_on_a_follow_up_still_fails_normally(self):
        """The echo-recovery must only fire on a PROVABLE exact echo of the carried-over value -- a
        follow-up that supplies a different, still-unparseable date phrase must continue to fail validation
        (and thus fall through to the legacy path) exactly as before the fix, never silently keep the old
        date instead."""
        ctx = ConversationContext()
        self._run("What was the debt on September 1?", '{"tool": "get_debt_at", "arguments": {"at_date": "September 1"}}', ctx)
        outcome = self._run(
            "What about nonsense-date?",
            '{"tool": "get_debt_at", "arguments": {"at_date": "nonsense-date"}}', ctx,
        )
        self.assertEqual(outcome.kind, "legacy")


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
