"""Network Resolver for Ask Data v3/v2: free-text supermarket chain/group references ("Edeka", "AD",
"Delhaize", "Proxy", "Intermarché", "Carrefour Market"/"CM", "Albert Heijn"/"AH", "Plus", "Jumbo") resolve to
a `network` filter distinct from a specific `project` and from `country` -- see services/ask_data/resolver.py
(resolve_network, NetworkIndex) and services/ask_data/validation.py (_redirect_network_text).

Goal 2 ("never fuzzy-resolve a generic network name to a single project") is the central guarantee tested
here: a bare network name must resolve to network=<canonical>, never silently collapse onto one specific
member project, regardless of how many (or how few) projects that network currently has.
"""
import unittest
from decimal import Decimal

from models.invoice import Invoice
from models.project import Project
from services.ask_data.executor import execute_tool
from services.ask_data.resolver import (
    ACCEPTED, AMBIGUOUS, NOT_FOUND, NetworkIndex, ProjectIndex, resolve_network, resolve_project,
)
from services.ask_data.validation import validate_tool_call


def _project(name, country, cams=4, status="Active"):
    return Project(name, country=country, num_cams=cams, payment_month="April", installation_year=2022, status=status)


def _inv(number, name, amount, paid="No", kind="Y1", cams=4):
    return Invoice(number, name, kind, float(amount), cams, None, paid, 2026)


def make_network_data():
    """Two members per network (so network filtering is actually exercised, not just "the one project
    that happens to share a prefix"), plus one project per "specific store" example from the task."""
    projects = [
        _project("AD Alpha", "Bel", cams=10),
        _project("AD Anderlecht", "Bel", cams=12),
        _project("Edeka Herweg", "Ger", cams=5),
        _project("Edeka Preller-Breckerfeld", "Ger", cams=7),
        _project("Proxy Delta", "Bel", cams=3),
        _project("Proxy Epsilon", "Bel", cams=6),
        _project("Intermarché Heusy", "Bel", cams=6),
        _project("Carrefour Market Namur", "Bel", cams=8),
        _project("Albert Heijn Tilburg", "Hol", cams=9),
        _project("Plus Zeta", "Hol", cams=4),
        _project("Jumbo Breda", "Hol", cams=11),
        _project("Unrelated Shop", "Lux", cams=2),
    ]
    invoices = [
        _inv(9001, "AD Alpha", 1000),
        _inv(9002, "AD Anderlecht", 1500),
        _inv(9003, "Edeka Herweg", 500),
        _inv(9004, "Edeka Preller-Breckerfeld", 700),
        _inv(9005, "Proxy Delta", 300),
        _inv(9006, "Proxy Epsilon", 600),
        _inv(9007, "Intermarché Heusy", 800),
        _inv(9008, "Carrefour Market Namur", 900),
        _inv(9009, "Albert Heijn Tilburg", 1100),
        _inv(9010, "Plus Zeta", 400),
        _inv(9011, "Jumbo Breda", 1200),
        _inv(9012, "Unrelated Shop", 200),
    ]
    return projects, invoices


class TestResolveNetwork(unittest.TestCase):
    """resolve_network(): EXACT alias match only, never fuzzy (goal 2)."""

    def test_bare_chain_names_resolve_to_their_canonical_network(self):
        cases = {
            "Edeka": "Edeka", "edeka": "Edeka", "  EDEKA  ": "Edeka",
            "AD": "AD", "Delhaize": "AD", "AD Delhaize": "AD", "ad-delhaize": "AD",
            "Proxy": "Proxy",
            "Intermarché": "Intermarché", "Intermarche": "Intermarché",
            "Carrefour Market": "Carrefour Market", "CM": "Carrefour Market", "Carrefour": "Carrefour Market",
            "Albert Heijn": "Albert Heijn", "AH": "Albert Heijn",
            "Plus": "Plus",
            "Jumbo": "Jumbo",
        }
        for text, expected in cases.items():
            with self.subTest(text=text):
                result = resolve_network(text)
                self.assertEqual(result.status, ACCEPTED)
                self.assertEqual(result.value, expected)

    def test_hebrew_alias(self):
        result = resolve_network("עדקה")
        self.assertEqual((result.status, result.value), (ACCEPTED, "Edeka"))

    def test_a_specific_store_name_is_not_a_network_alias(self):
        for text in ("Edeka Preller-Breckerfeld", "AD Anderlecht", "Proxy Delta", "Belgium", "random text"):
            with self.subTest(text=text):
                self.assertEqual(resolve_network(text).status, NOT_FOUND)

    def test_never_fuzzy_never_mutates_input(self):
        projects, _ = make_network_data()
        before = [p.project_name for p in projects]
        resolve_network("edeca")              # a typo -- must NOT fuzzy-match "Edeka"
        resolve_network("AD ")
        self.assertEqual(resolve_network("edeca").status, NOT_FOUND)
        self.assertEqual([p.project_name for p in projects], before)   # read-only: goal 7


class TestNetworkIndex(unittest.TestCase):
    """NetworkIndex: deterministic (first-word-token) project -> network membership, never fuzzy (goal 1 + 2)."""

    def setUp(self):
        self.projects, _ = make_network_data()
        self.index = NetworkIndex(self.projects)

    def test_every_seeded_project_maps_to_its_expected_network(self):
        expected = {
            "AD Alpha": "AD", "AD Anderlecht": "AD",
            "Edeka Herweg": "Edeka", "Edeka Preller-Breckerfeld": "Edeka",
            "Proxy Delta": "Proxy", "Proxy Epsilon": "Proxy",
            "Intermarché Heusy": "Intermarché",
            "Carrefour Market Namur": "Carrefour Market",
            "Albert Heijn Tilburg": "Albert Heijn",
            "Plus Zeta": "Plus",
            "Jumbo Breda": "Jumbo",
            "Unrelated Shop": None,
        }
        for name, network in expected.items():
            with self.subTest(project=name):
                self.assertEqual(self.index.network_of(name), network)

    def test_case_and_accent_insensitive(self):
        self.assertEqual(self.index.network_of("ad alpha"), "AD")
        self.assertEqual(self.index.network_of("EDEKA HERWEG"), "Edeka")
        self.assertEqual(self.index.network_of("intermarche heusy"), "Intermarché")


class TestExactProjectVsNetworkDistinction(unittest.TestCase):
    """Goal: "Edeka" -> network=Edeka; "Edeka Preller-Breckerfeld" -> project=Edeka Preller-Breckerfeld.
    "AD" -> network=AD; "AD Anderlecht" -> project=AD Anderlecht. "Belgium" -> country=Bel."""

    def setUp(self):
        self.projects, self.invoices = make_network_data()

    def _validate(self, tool, **kwargs):
        return validate_tool_call({"tool": tool, "arguments": kwargs}, self.projects)

    def test_bare_edeka_resolves_to_network_not_project(self):
        result = self._validate("get_debt", project="Edeka")
        self.assertEqual(result.tool, "get_debt")
        self.assertEqual(result.arguments["network"], "Edeka")
        self.assertIsNone(result.arguments["project"])

    def test_specific_edeka_store_still_resolves_to_project_exactly_as_before(self):
        result = self._validate("get_debt", project="Edeka Preller-Breckerfeld")
        self.assertEqual(result.arguments["project"], "Edeka Preller-Breckerfeld")
        self.assertIsNone(result.arguments["network"])

    def test_bare_ad_resolves_to_network(self):
        result = self._validate("get_debt", project="AD")
        self.assertEqual(result.arguments["network"], "AD")
        self.assertIsNone(result.arguments["project"])

    def test_specific_ad_store_still_resolves_to_project(self):
        result = self._validate("get_debt", project="AD Anderlecht")
        self.assertEqual(result.arguments["project"], "AD Anderlecht")
        self.assertIsNone(result.arguments["network"])

    def test_belgium_resolves_to_country_not_network_or_project(self):
        result = self._validate("get_debt", country="Belgium")
        self.assertEqual(result.arguments["country"], "Bel")
        self.assertIsNone(result.arguments["network"])
        self.assertIsNone(result.arguments["project"])

    def test_planner_can_also_set_network_directly(self):
        """The redirect also works the other way: if the planner already puts the chain name straight into
        'network' (as the updated prompt now tells it to), validation accepts it directly."""
        result = self._validate("get_debt", network="Edeka")
        self.assertEqual(result.arguments["network"], "Edeka")

    def test_network_is_never_offered_for_a_tool_that_does_not_declare_it(self):
        """get_last_bank_payment has no 'network' argument -- a bare network name there is still handled as
        plain (ambiguous) project text, exactly as before this feature existed."""
        result = self._validate("get_last_bank_payment", project="Edeka")
        self.assertIsNone(result.tool)
        self.assertEqual(result.clarification["kind"], "project")
        self.assertEqual(set(result.clarification["candidates"]), {"Edeka Herweg", "Edeka Preller-Breckerfeld"})


class TestNetworkFiltersOnTools(unittest.TestCase):
    """Goal 4: network filters wired into project listing, debt, unpaid invoices, paid amount, outstanding
    amount, camera counts, licenses."""

    def setUp(self):
        self.projects, self.invoices = make_network_data()

    def _run(self, tool, **kwargs):
        return execute_tool(tool, kwargs, self.projects, self.invoices, [])

    def test_edeka_projects(self):
        result = self._run("get_projects", network="Edeka")
        names = set(result.dataframe["Project"])
        self.assertEqual(names, {"Edeka Herweg", "Edeka Preller-Breckerfeld"})

    def test_ad_projects(self):
        result = self._run("get_projects", network="AD")
        self.assertEqual(set(result.dataframe["Project"]), {"AD Alpha", "AD Anderlecht"})

    def test_proxy_projects(self):
        result = self._run("get_projects", network="Proxy")
        self.assertEqual(set(result.dataframe["Project"]), {"Proxy Delta", "Proxy Epsilon"})

    def test_count_projects_network(self):
        result = self._run("count_projects", network="Edeka")
        self.assertIn("2", result.text)

    def test_edeka_debt(self):
        result = self._run("get_debt", network="Edeka")
        self.assertEqual(result.value, Decimal("1200.00"))   # 500 + 700
        self.assertEqual(result.row_count, 2)

    def test_unpaid_invoices_network(self):
        result = self._run("get_unpaid_invoices", network="Proxy")
        self.assertEqual({str(int(r)) for r in result.dataframe["Invoice #"]}, {"9005", "9006"})

    def test_paid_amount_network(self):
        paid_invoices = list(self.invoices) + [_inv(9999, "Edeka Herweg", 50, paid="yes")]
        result = execute_tool("get_paid_amount", {"network": "Edeka"}, self.projects, paid_invoices, [])
        self.assertEqual(result.value, Decimal("50.00"))

    def test_outstanding_amount_network(self):
        result = self._run("get_outstanding_amount", network="AD")
        self.assertEqual(result.value, Decimal("2500.00"))   # 1000 + 1500

    def test_edeka_cameras(self):
        result = self._run("get_camera_statistics", network="Edeka", aggregation="sum")
        self.assertIn("12", result.text)   # 5 + 7 cameras

    def test_licenses_network(self):
        result = self._run("get_licenses", network="Proxy")
        self.assertEqual(len(result.dataframe), 2)

    def test_network_totals_match_the_sum_of_its_member_projects(self):
        """Goal 1: a network query returns ONLY projects in that network -- cross-checked against summing the
        individual member projects' own debt."""
        network_total = self._run("get_debt", network="AD").value
        alpha = self._run("get_debt", project="AD Alpha").value
        anderlecht = self._run("get_debt", project="AD Anderlecht").value
        self.assertEqual(network_total, alpha + anderlecht)


class TestHebrewNetworkQueries(unittest.TestCase):
    def setUp(self):
        self.projects, self.invoices = make_network_data()

    def test_hebrew_edeka_alias_resolves_the_same_network(self):
        result = validate_tool_call({"tool": "get_debt", "arguments": {"project": "עדקה"}}, self.projects)
        self.assertEqual(result.arguments["network"], "Edeka")
        tool_result = execute_tool("get_debt", result.arguments, self.projects, self.invoices, [])
        self.assertEqual(tool_result.value, Decimal("1200.00"))

    def test_latin_brand_name_inside_a_hebrew_sentence_still_resolves(self):
        """No alias needed: the tokenizer only translates tokens it recognises as Hebrew script, so a Latin
        brand name embedded in a Hebrew question passes through untouched."""
        result = validate_tool_call({"tool": "get_debt", "arguments": {"project": "Edeka", "country": "בגרמניה"}}, self.projects)
        self.assertEqual(result.arguments["network"], "Edeka")
        self.assertEqual(result.arguments["country"], "Ger")


if __name__ == "__main__":
    unittest.main()
