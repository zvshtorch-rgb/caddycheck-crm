"""Ask Data Phase 2 unit tests: canonical resolver, aliases, thresholds, country reconciliation, month fix."""
import datetime
import unittest
from decimal import Decimal
from pathlib import Path
from unittest import mock

from config.ask_data_aliases import PROJECT_ALIASES
from config.settings import MONTH_ORDER
from models.invoice import Invoice
from models.project import Project
from services.ask_data import breakdown, resolver
from services.ask_data.executor import _execute_ask_data_tool, _filter_invoices_for_debt_tools
from services.ask_data.helpers import _build_invoice_answer_df, _normalize_country
from services.ask_data.resolver import (
    ACCEPTED, AMBIGUOUS, NOT_FOUND, ProjectIndex, classify_score, compact_key, normalize_text,
    resolve_country, resolve_month, resolve_project,
)
from services.ask_data.validation import _validate_ask_data_tool_call, clarification_question, validate_tool_call
from tests._ask_data_harness import BASE_CFG, NewImpl, OracleImpl, _gemini_ok, _make_post, fake_streamlit

ROOT = Path(__file__).resolve().parent.parent
HE_GERMANY, HE_IN_GERMANY, HE_HOLLAND, HE_FROM_HOLLAND, HE_BELGIUM = "גרמניה", "בגרמניה", "הולנד", "מהולנד", "בלגיה"
HE_APRIL, HE_IN_APRIL, HE_MARCH = "אפריל", "באפריל", "מרץ"
HE_QUESTION_PROJECTS_BILLED_IN_APRIL = "אילו פרויקטים מחויבים באפריל?"


def _project(name, country, month="", cams=4, status="Active"):
    return Project(name, country=country, num_cams=cams, payment_month=month, installation_year=2022, status=status)


def _inv(number, name, amount, paid="No", year=2026, kind="Y1"):
    return Invoice(number, name, kind, float(amount), 4, None, paid, year)


def make_data():
    projects = [
        _project("CM Essen", "Bel", "June", 10),
        _project("Intermarché Heusy", "Bel", "April", 6),
        _project("AD Alpha", "Bel", "April", 10),
        _project("AD Beta", "Bel", "May", 8),
        _project("Edeka Herweg", "Ger", "April", 5),
        _project("Edeka Peipp Rostal", "Ger", "March", 4),
        _project("Plus Zeta", "Hol", "April", 4),
        _project("Spar Eta", "Lux", "July", 2),
        _project("Old Shop", "Bel", "April", 3, status="Cancelled"),
    ]
    invoices = [
        _inv(8595, "CM Essen (coplementary)", 3112),          # approved alias -> CM Essen (Bel)
        _inv(8562, "CM Essen", 4668, paid="Yes", year=2024),
        _inv(8700, "intermarché Heusy", 7780),                # case/accent variant -> Intermarché Heusy (Bel)
        _inv(8701, "INTERMARCHE HEUSY", 100),
        _inv(8702, "Intermarché Heusy", 50, paid="Yes"),
        _inv(8703, "AD Alpha", 1000),
        _inv(8704, "AD Beta", -50, kind="Credit"),            # unpaid credit nets against debt
        _inv(8705, "Edeka Herweg", 500),
        _inv(8706, "Plus Zeta", 200.5),
        _inv(8707, "Unknown Shop", 111),                       # no matching project at all
        _inv(8708, "Spar Eta", 75, paid="cancelled"),
    ]
    return projects, invoices


class TestNormalization(unittest.TestCase):
    def test_case_accent_and_punctuation_are_folded(self):
        for variant in ("intermarché Heusy", "Intermarché Heusy", "INTERMARCHE HEUSY", "intermarche  heusy!", "Intermarché-Heusy"):
            self.assertEqual(normalize_text(variant), "intermarche heusy", variant)
        self.assertEqual(compact_key("AD Alpha"), compact_key("ad-alpha"))
        self.assertEqual(compact_key("AD Alpha"), compact_key("ADAlpha"))
        self.assertEqual(normalize_text(None), "")

    def test_spelling_variants_are_unified(self):
        self.assertEqual(normalize_text("CM Essen (coplementary)"), normalize_text("CM Essen (Complementary)"))
        self.assertEqual(normalize_text("CM Essen (complimentary)"), "cm essen complementary")

    def test_hebrew_terms_points_final_letters_and_prefixes(self):
        self.assertEqual(normalize_text(HE_GERMANY), "germany")
        self.assertEqual(normalize_text(HE_IN_GERMANY), "germany")          # prefix ב ("in")
        self.assertEqual(normalize_text(HE_FROM_HOLLAND), "holland")        # prefix מ ("from")
        self.assertEqual(normalize_text(HE_IN_APRIL), "april")
        self.assertEqual(normalize_text(HE_MARCH), "march")                 # final letter ץ folded
        self.assertEqual(normalize_text("בלגיָה"), "belgium")               # niqqud stripped
        self.assertEqual(normalize_text("שלום"), "שלומ")                    # unknown Hebrew is only folded, never translated


class TestAliases(unittest.TestCase):
    def test_exactly_the_four_approved_aliases_are_configured(self):
        self.assertEqual(PROJECT_ALIASES, {
            "CM Essen (coplementary)": "CM Essen",
            "intermarché Heusy": "Intermarché Heusy",
            "Plus Landgraaf Arts": "Plus langdraaf Aarts",
            "Plus Landgraaf Aarts": "Plus langdraaf Aarts",
        })

    def test_cm_essen_coplementary_resolves_to_cm_essen(self):
        projects, _ = make_data()
        names = [p.project_name for p in projects]
        for text in ("CM Essen (coplementary)", "cm essen (complementary)", "CM ESSEN coplementary"):
            r = resolve_project(text, names)
            self.assertEqual((r.status, r.value, r.reason), (ACCEPTED, "CM Essen", "alias"), text)

    def test_intermarche_heusy_resolves_to_accented_project(self):
        projects, _ = make_data()
        names = [p.project_name for p in projects]
        self.assertEqual(resolve_project("intermarché Heusy", names).value, "Intermarché Heusy")
        # also by plain normalisation, with no alias table at all
        r = resolve_project("intermarche HEUSY", names, aliases={})
        self.assertEqual((r.status, r.value, r.reason), (ACCEPTED, "Intermarché Heusy", "exact"))

    def test_alias_with_a_missing_target_project_is_ignored(self):
        self.assertEqual(resolver.project_alias_map(["AD Alpha"]), {})
        r = resolve_project("CM Essen (coplementary)", ["AD Alpha", "AD Beta"])
        self.assertNotEqual(r.status, ACCEPTED)

    def test_ledger_rows_map_through_the_index_for_both_aliases(self):
        projects, _ = make_data()
        index = ProjectIndex(projects)
        self.assertEqual(index.canonical_name("CM Essen (coplementary)"), "CM Essen")
        self.assertEqual(index.country("CM Essen (coplementary)"), "Bel")
        self.assertEqual(index.canonical_name("intermarché Heusy"), "Intermarché Heusy")
        self.assertEqual(index.canonical_name("INTERMARCHE HEUSY"), "Intermarché Heusy")
        self.assertIsNone(index.canonical_name("Unknown Shop"))
        self.assertIsNone(index.country(""))

    def test_aliases_never_touch_ledger_rows(self):
        projects, invoices = make_data()
        before = [(i.invoice_number, i.project_name, i.payment_amount, i.paid) for i in invoices]
        ProjectIndex(projects).canonical_name("CM Essen (coplementary)")
        breakdown.debt_by_country(invoices, projects)
        _filter_invoices_for_debt_tools(invoices, projects, {"country": "Bel"}, lambda inv: inv.is_unpaid())
        self.assertEqual(before, [(i.invoice_number, i.project_name, i.payment_amount, i.paid) for i in invoices])

    def test_ledger_matching_is_never_fuzzy(self):
        projects, _ = make_data()
        self.assertIsNone(ProjectIndex(projects).project("AD Alfa"))        # typo: fine for user input, never for ledger rows

    LANDGRAAF_PROJECTS = [_project("Plus langdraaf Aarts", "Hol"), _project("Plus Landgraaf Wetzels", "Hol")]

    def test_landgraaf_aliases_resolve_to_the_misspelled_project(self):
        names = [p.project_name for p in self.LANDGRAAF_PROJECTS]
        for text in ("Plus Landgraaf Arts", "Plus Landgraaf Aarts", "plus landgraaf arts"):
            r = resolve_project(text, names)
            self.assertEqual((r.status, r.value), (ACCEPTED, "Plus langdraaf Aarts"), text)
        index = ProjectIndex(self.LANDGRAAF_PROJECTS)
        self.assertEqual(index.canonical_name("Plus Landgraaf Arts"), "Plus langdraaf Aarts")
        self.assertEqual(index.canonical_name("Plus Landgraaf Aarts"), "Plus langdraaf Aarts")

    def test_landgraaf_alone_stays_ambiguous_between_the_two_stores(self):
        """'Landgraaf' alone must NOT silently resolve to Wetzels just because that project's real name
        happens to contain the word -- the misspelled project's own alias sources also contain it."""
        names = [p.project_name for p in self.LANDGRAAF_PROJECTS]
        r = resolve_project("Landgraaf", names)
        self.assertEqual(r.status, AMBIGUOUS)
        self.assertEqual(set(r.candidates), {"Plus langdraaf Aarts", "Plus Landgraaf Wetzels"})
        r = resolve_project("Plus Landgraaf", names)
        self.assertEqual(r.status, AMBIGUOUS)
        self.assertEqual(set(r.candidates), {"Plus langdraaf Aarts", "Plus Landgraaf Wetzels"})

    def test_wetzels_and_aarts_remain_individually_unambiguous(self):
        names = [p.project_name for p in self.LANDGRAAF_PROJECTS]
        self.assertEqual(resolve_project("Wetzels", names).value, "Plus Landgraaf Wetzels")
        self.assertEqual(resolve_project("Aarts", names).value, "Plus langdraaf Aarts")
        # ledger rows: never fuzzy, so the bare "Landgraaf" ambiguity never corrupts country matching
        index = ProjectIndex(self.LANDGRAAF_PROJECTS)
        self.assertIsNone(index.canonical_name("Landgraaf"))

    def test_landgraaf_aliases_do_not_affect_the_usual_fixture(self):
        """The alias table is global config; confirm it's inert when neither Landgraaf project exists."""
        projects, _ = make_data()
        self.assertIsNone(resolver.project_alias_map([p.project_name for p in projects]).get(compact_key("Plus Landgraaf Arts")))


class TestThresholds(unittest.TestCase):
    def test_classify_score_boundaries(self):
        self.assertEqual(classify_score(0.90), ACCEPTED)
        self.assertEqual(classify_score(0.8999), AMBIGUOUS)
        self.assertEqual(classify_score(0.60), AMBIGUOUS)
        self.assertEqual(classify_score(0.5999), NOT_FOUND)
        self.assertEqual(classify_score(1.0), ACCEPTED)
        self.assertEqual(classify_score(0.0), NOT_FOUND)

    def _resolve_with_score(self, score):
        with mock.patch.object(resolver, "_project_score", return_value=score):
            return resolve_project("anything", ["Alpha Store"])

    def test_resolve_project_applies_the_thresholds_exactly(self):
        self.assertEqual(self._resolve_with_score(0.90).status, ACCEPTED)
        self.assertEqual(self._resolve_with_score(0.8999).status, AMBIGUOUS)
        self.assertEqual(self._resolve_with_score(0.60).status, AMBIGUOUS)
        self.assertEqual(self._resolve_with_score(0.5999).status, NOT_FOUND)
        low = self._resolve_with_score(0.75)
        self.assertEqual((low.value, low.candidates, low.reason), (None, ("Alpha Store",), "low_confidence"))

    def test_ambiguous_entity_asks_for_clarification_with_real_candidates(self):
        projects, _ = make_data()
        r = resolve_project("Edeka", [p.project_name for p in projects])
        self.assertEqual((r.status, r.value), (AMBIGUOUS, None))
        self.assertEqual(set(r.candidates), {"Edeka Herweg", "Edeka Peipp Rostal"})
        self.assertLessEqual(len(r.candidates), resolver.MAX_CANDIDATES)

    def test_duplicate_project_names_are_ambiguous_not_guessed(self):
        r = resolve_project("AD Alpha", ["AD Alpha", "ad-alpha"])
        self.assertEqual((r.status, r.reason), (AMBIGUOUS, "duplicate_names"))

    def test_typo_scoring_between_60_and_90_asks_instead_of_guessing(self):
        r = resolve_project("Plus Zefa", ["Plus Zeta Market"])          # similar but not close enough to accept
        self.assertIn(r.status, (AMBIGUOUS, NOT_FOUND))
        r = resolve_project("AD Alfa", ["AD Alpha", "Plus Zeta"])
        self.assertEqual((r.status, r.candidates), (AMBIGUOUS, ("AD Alpha",)))

    def test_not_found(self):
        projects, _ = make_data()
        names = [p.project_name for p in projects]
        for text in ("Zzzzzz Qqqq", "", "   ", None):
            r = resolve_project(text, names)
            self.assertEqual((r.status, r.value, r.candidates), (NOT_FOUND, None, ()), repr(text))
        self.assertEqual(resolve_project("AD Alpha", []).status, NOT_FOUND)


class TestCountries(unittest.TestCase):
    CODES = ["Bel", "Ger", "Hol", "Lux"]

    def r(self, text):
        return resolve_country(text, self.CODES, _normalize_country)

    def test_codes_names_variants_and_hebrew(self):
        cases = {
            "Bel": "Bel", "belgium": "Bel", "BELGIË": "Bel", "Germany": "Ger", "Deutschland": "Ger", "Holland": "Hol",
            "Netherlands": "Hol", "Nederland": "Hol", "Luxembourg": "Lux", HE_GERMANY: "Ger", HE_IN_GERMANY: "Ger",
            HE_HOLLAND: "Hol", HE_FROM_HOLLAND: "Hol", HE_BELGIUM: "Bel",
        }
        for text, code in cases.items():
            r = self.r(text)
            self.assertEqual((r.status, r.value), (ACCEPTED, code), text)

    def test_small_typo_is_accepted_but_a_weak_one_asks(self):
        self.assertEqual(self.r("belgum").value, "Bel")                 # 0.923
        weak = self.r("blegium")                                         # 0.857: ask, do not guess
        self.assertEqual((weak.status, weak.candidates), (AMBIGUOUS, ("Bel",)))

    def test_country_not_found(self):
        r = self.r("Atlantis")
        self.assertEqual((r.status, r.value), (NOT_FOUND, None))
        self.assertEqual(self.r("").status, NOT_FOUND)


class TestValidatorClarification(unittest.TestCase):
    def setUp(self):
        self.projects, self.invoices = make_data()

    def test_ambiguous_project_returns_a_clarification_and_no_tool(self):
        result = validate_tool_call({"tool": "get_debt", "arguments": {"project": "Edeka"}}, self.projects)
        self.assertIsNone(result.tool)
        self.assertEqual(result.clarification["arg"], "project")
        self.assertEqual(set(result.clarification["candidates"]), {"Edeka Herweg", "Edeka Peipp Rostal"})
        question = clarification_question(result.clarification)
        self.assertTrue(question.endswith("?"))
        self.assertIn("Edeka Herweg", question)
        self.assertEqual(question.count("?"), 1)                          # exactly one short question

    def test_weak_country_typo_clarifies(self):
        result = validate_tool_call({"tool": "get_debt", "arguments": {"country": "blegium"}}, self.projects)
        self.assertEqual((result.tool, result.clarification["kind"], result.clarification["candidates"]), (None, "country", ["Bel"]))
        self.assertEqual(clarification_question(result.clarification), "Did you mean Bel (for 'blegium')?")

    def test_not_found_keeps_the_original_validation_error_and_has_no_clarification(self):
        result = validate_tool_call({"tool": "get_debt", "arguments": {"project": "Zzzzzz Qqqq"}}, self.projects)
        self.assertIsNone(result.clarification)
        self.assertEqual(result.error, "Could not validate argument 'project'='Zzzzzz Qqqq' for tool 'get_debt'.")
        self.assertEqual(_validate_ask_data_tool_call({"tool": "get_debt", "arguments": {"country": "Atlantis"}}, self.projects)[0], None)

    def test_aliases_and_hebrew_arguments_are_resolved(self):
        result = validate_tool_call({"tool": "get_debt", "arguments": {"project": "CM Essen (coplementary)", "country": HE_IN_GERMANY}}, self.projects)
        self.assertEqual((result.tool, result.arguments["project"], result.arguments["country"]), ("get_debt", "CM Essen", "Ger"))
        result = validate_tool_call({"tool": "get_projects", "arguments": {"month": HE_APRIL}}, self.projects)
        self.assertEqual(result.arguments["month"], "April")

    def test_router_asks_one_question_and_neither_executes_nor_falls_back(self):
        projects, invoices = self.projects, self.invoices
        legacy_calls, posts = [], []
        payload = '{"tool": "get_debt", "arguments": {"project": "Edeka"}}'
        with fake_streamlit() as state, NewImpl().env(cfg=BASE_CFG), mock.patch("requests.post", _make_post([_gemini_ok(payload)], posts)):
            text, df = NewImpl().smart(
                "debt of edeka", projects, invoices,
                lambda q: legacy_calls.append(("normalize", q)), lambda q, p, i, d: legacy_calls.append(("answer", q)),
            )
        self.assertIsNone(df)
        self.assertIn("Which project did you mean", text)
        self.assertEqual(legacy_calls, [])
        self.assertIn("_ask_data_clarification", state)
        self.assertNotIn("_ask_data_llm_tool_call", state)
        self.assertNotIn("_ask_data_llm_error", state)


class TestCountryTotalsReconcile(unittest.TestCase):
    def setUp(self):
        self.projects, self.invoices = make_data()

    def test_variant_ledger_names_now_belong_to_their_country(self):
        rows = _filter_invoices_for_debt_tools(self.invoices, self.projects, {"country": "Bel"}, lambda inv: inv.is_unpaid())
        self.assertEqual(sorted(r.project_name for r in rows),
                         sorted(["CM Essen (coplementary)", "intermarché Heusy", "INTERMARCHE HEUSY", "AD Alpha", "AD Beta"]))

    def test_country_buckets_plus_unassigned_equal_the_total_debt_exactly(self):
        b = breakdown.debt_by_country(self.invoices, self.projects)
        all_unpaid = [i for i in self.invoices if i.is_unpaid()]
        total = sum((Decimal(str(i.payment_amount)) for i in all_unpaid), Decimal("0"))
        self.assertEqual(b.total, total)
        self.assertEqual(b.row_count, len(all_unpaid))
        self.assertTrue(b.reconciles)
        self.assertEqual(sum((x.amount for x in b.buckets), Decimal("0")), total)
        self.assertEqual({x.label: x.amount for x in b.buckets}, {
            "Bel": Decimal("3112") + Decimal("7780") + Decimal("100") + Decimal("1000") + Decimal("-50"),
            "Ger": Decimal("500"), "Hol": Decimal("200.5"), breakdown.UNASSIGNED_LABEL: Decimal("111"),
        })

    def test_unmatched_rows_stay_visible_as_their_own_bucket(self):
        b = breakdown.debt_by_country(self.invoices, self.projects)
        self.assertEqual(b.unassigned_names(), ["Unknown Shop"])
        self.assertEqual((b.unassigned.row_count, b.unassigned.amount), (1, Decimal("111")))
        frame = b.to_dataframe()
        self.assertIn(breakdown.UNASSIGNED_LABEL, list(frame["Country"]))
        self.assertEqual(list(frame["Country"])[-1], "Total")
        self.assertEqual(frame.iloc[-1]["Amount (€)"], b.total)

    def test_unassigned_bucket_is_shown_at_zero_for_paid_and_cancelled_breakdowns_too(self):
        """The 'always show, even at zero' requirement applies to ANY row_filter, not just unpaid debt."""
        for predicate in (lambda inv: inv.is_paid(), lambda inv: inv.is_cancelled()):
            b = breakdown.debt_by_country(self.invoices, self.projects, row_filter=predicate)
            self.assertIn(breakdown.UNASSIGNED_LABEL, [x.label for x in b.buckets])
            self.assertTrue(b.reconciles)
        paid = breakdown.debt_by_country(self.invoices, self.projects, row_filter=lambda inv: inv.is_paid())
        self.assertEqual(paid.unassigned.amount, Decimal("0"))    # "Unknown Shop" row is unpaid, so paid-unassigned is genuinely zero
        self.assertEqual(sum((b.amount for b in paid.buckets), Decimal("0")), paid.total)
        cancelled = breakdown.debt_by_country(self.invoices, self.projects, row_filter=lambda inv: inv.is_cancelled())
        self.assertEqual(cancelled.unassigned.amount, Decimal("0"))
        self.assertEqual(sum((b.amount for b in cancelled.buckets), Decimal("0")), cancelled.total)

    def test_an_alias_without_its_target_project_leaves_the_row_visibly_unassigned(self):
        projects = [p for p in self.projects if p.project_name != "CM Essen"]
        b = breakdown.debt_by_country(self.invoices, projects)
        self.assertTrue(b.reconciles)
        self.assertIn("CM Essen (coplementary)", b.unassigned_names())
        self.assertEqual(b.unassigned.amount, Decimal("111") + Decimal("3112"))

    def test_unassigned_bucket_is_always_shown_even_when_empty(self):
        invoices = [i for i in self.invoices if i.project_name != "Unknown Shop"]
        b = breakdown.debt_by_country(invoices, self.projects)
        self.assertEqual((b.unassigned.row_count, b.unassigned.amount), (0, Decimal("0")))
        self.assertTrue(b.reconciles)
        self.assertIn(breakdown.UNASSIGNED_LABEL, [x.label for x in b.buckets])
        self.assertEqual(b.buckets[-1].label, breakdown.UNASSIGNED_LABEL)  # always last, even at zero

    def test_a_matched_project_without_a_country_gets_its_own_visible_bucket(self):
        projects = self.projects + [Project("No Country Shop", country="", status="Active")]
        invoices = self.invoices + [_inv(8800, "No Country Shop", 9)]
        b = breakdown.debt_by_country(invoices, projects)
        self.assertTrue(b.reconciles)
        self.assertEqual({x.label: x.amount for x in b.buckets}[breakdown.NO_COUNTRY_LABEL], Decimal("9"))

    def test_tool_level_country_totals_sum_to_the_unfiltered_total_minus_unassigned(self):
        def amount(country=None):
            args = {"country": country} if country else {}
            rows = _filter_invoices_for_debt_tools(self.invoices, self.projects, args, lambda inv: inv.is_unpaid())
            return sum(Decimal(str(r.payment_amount)) for r in rows)

        self.assertEqual(sum(amount(c) for c in ("Bel", "Ger", "Hol", "Lux")) + Decimal("111"), amount())

    def test_executor_text_and_table_use_the_canonical_country(self):
        text, df = _execute_ask_data_tool("get_unpaid_invoices", {"country": "Bel"}, self.projects, self.invoices, [])
        self.assertIn("totaling €11,942", text)                          # 3112 + 7780 + 100 + 1000 - 50
        countries = {row["Project"]: row["Country"] for row in df.to_dict("records")}
        self.assertEqual(countries["CM Essen (coplementary)"], "Bel")
        self.assertEqual(countries["intermarché Heusy"], "Bel")
        unknown = next(i for i in self.invoices if i.project_name == "Unknown Shop")
        self.assertEqual(_build_invoice_answer_df([unknown], self.projects).iloc[0]["Country"], "")

    def test_project_filter_uses_the_canonical_name(self):
        text, _ = _execute_ask_data_tool("get_outstanding_amount", {"project": "Intermarché Heusy"}, self.projects, self.invoices, [])
        self.assertEqual(text.split("\n\n")[0], "Outstanding (unpaid) amount for Intermarché Heusy is €7,880.00 across 2 invoice row(s).")


class TestMonthQueries(unittest.TestCase):
    """Regression: `get_projects_for_month` was never imported, so month questions raised NameError."""

    def setUp(self):
        self.projects, self.invoices = make_data()
        self.april = {"Intermarché Heusy", "AD Alpha", "Edeka Herweg", "Plus Zeta"}       # active projects billed in April

    def test_month_resolver_english_hebrew_and_unknown(self):
        self.assertEqual(resolve_month("April", MONTH_ORDER), "April")
        self.assertEqual(resolve_month("apr", MONTH_ORDER), "April")
        self.assertEqual(resolve_month("Sept", MONTH_ORDER), "September")
        self.assertEqual(resolve_month(HE_APRIL, MONTH_ORDER), "April")
        self.assertEqual(resolve_month(HE_IN_APRIL, MONTH_ORDER), "April")
        self.assertEqual(resolve_month(HE_MARCH, MONTH_ORDER), "March")
        self.assertIsNone(resolve_month("Smarch", MONTH_ORDER))
        self.assertIsNone(resolve_month("", MONTH_ORDER))

    def test_get_projects_month_filter_works_and_no_longer_raises(self):
        text, df = _execute_ask_data_tool("get_projects", {"month": "April"}, self.projects, self.invoices, [])
        self.assertEqual(text, "Found 4 active project(s).")
        self.assertEqual({r["Project"] for r in df.to_dict("records")}, self.april)    # cancelled 'Old Shop' excluded

    def test_month_with_country_filter(self):
        _, df = _execute_ask_data_tool("get_projects", {"month": "April", "country": "Bel"}, self.projects, self.invoices, [])
        self.assertEqual({r["Project"] for r in df.to_dict("records")}, {"Intermarché Heusy", "AD Alpha"})

    def test_equivalent_hebrew_question_end_to_end_through_the_router(self):
        for month_value in ("April", HE_APRIL, HE_IN_APRIL):
            payload = '{"tool": "get_projects", "arguments": {"month": "%s"}}' % month_value
            with fake_streamlit() as state, NewImpl().env(cfg=BASE_CFG), mock.patch("requests.post", _make_post([_gemini_ok(payload)], [])):
                text, df = NewImpl().smart(
                    HE_QUESTION_PROJECTS_BILLED_IN_APRIL, self.projects, self.invoices,
                    lambda q: self.fail("legacy normaliser must not run"), lambda q, p, i, d: self.fail("legacy answer must not run"),
                )
            self.assertEqual(text, "Found 4 active project(s).", month_value)
            self.assertEqual({r["Project"] for r in df.to_dict("records")}, self.april)
            self.assertNotIn("_ask_data_llm_error", state)
            self.assertEqual(state["_ask_data_llm_tool_call"]["arguments"]["month"], "April")

    def test_unknown_month_value_is_still_a_validation_error(self):
        result = validate_tool_call({"tool": "get_projects", "arguments": {"month": "Smarch"}}, self.projects)
        self.assertEqual(result.error, "Could not validate argument 'month'='Smarch' for tool 'get_projects'.")

    def test_legacy_keyword_path_now_answers_month_questions(self):
        source = (ROOT / "streamlit_app.py").read_text(encoding="utf-8")
        app = OracleImpl(source, "app", delegate=NewImpl())
        with fake_streamlit(), app.env(cfg={"api_key": "", "model": "m"}):
            for question in ("Which projects are billed in April?", "projects billed in april"):
                text, df = app.ns["_answer_data_question_smart"](question, self.projects, self.invoices, [])
                self.assertEqual(text, "Found 4 project(s) billed in April.", question)
                self.assertEqual({r["Project"] for r in df.to_dict("records")}, self.april)

    def test_app_imports_the_month_function(self):
        import ast
        tree = ast.parse((ROOT / "streamlit_app.py").read_text(encoding="utf-8"))
        imported = {a.name for n in tree.body if isinstance(n, ast.ImportFrom) and n.module == "services.excel_service" for a in n.names}
        self.assertIn("get_projects_for_month", imported)


if __name__ == "__main__":
    unittest.main()
