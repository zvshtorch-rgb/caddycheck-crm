"""Regression tests for the Licenses page's license-extension confirmation flow
(streamlit_app.py: `_compute_target_license_date`, `_resolve_license_project_selection`,
`_license_confirmation_still_valid`).

streamlit_app.py executes top-level Streamlit UI code on import (it's an entry-point
script, not a library module), so these pure, Streamlit-independent helper functions are
extracted via `ast` instead of `import streamlit_app` -- this avoids booting the whole app
just to unit-test isolated logic. See /memories/repo/ask-data-llm-architecture.md for the
same pattern used elsewhere this session.
"""
import ast
import datetime
import unittest
from pathlib import Path

_SRC_PATH = Path(__file__).resolve().parent.parent / "streamlit_app.py"
_WANTED_FUNCS = {
    "_compute_target_license_date",
    "_resolve_license_project_selection",
    "_license_confirmation_still_valid",
    "_add_months",
}


def _load_license_confirmation_helpers() -> dict:
    tree = ast.parse(_SRC_PATH.read_text(encoding="utf-8"))
    nodes = [
        node for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name in _WANTED_FUNCS
    ]
    found = {node.name for node in nodes}
    missing = _WANTED_FUNCS - found
    assert not missing, f"Missing extracted defs from streamlit_app.py: {missing}"
    module = ast.Module(body=nodes, type_ignores=[])
    code = compile(module, "<extracted-license-confirmation>", "exec")
    namespace = {"datetime": datetime, "calendar": __import__("calendar"), "Optional": None}
    from typing import Optional
    namespace["Optional"] = Optional
    exec(code, namespace)
    return namespace


_HELPERS = _load_license_confirmation_helpers()
_compute_target_license_date = _HELPERS["_compute_target_license_date"]
_resolve_license_project_selection = _HELPERS["_resolve_license_project_selection"]
_license_confirmation_still_valid = _HELPERS["_license_confirmation_still_valid"]


class TestComputeTargetLicenseDate(unittest.TestCase):
    """Sanity-checks the pure date computation the preview/confirm/save all share."""

    def test_set_exact_date(self):
        result = _compute_target_license_date(
            "Set exact date", datetime.date(2027, 9, 18), datetime.date(2026, 9, 23)
        )
        self.assertEqual(result, datetime.date(2027, 9, 18))

    def test_extend_by_12_months(self):
        result = _compute_target_license_date(
            "Extend by 12 months", datetime.date(2026, 1, 1), datetime.date(2026, 9, 18)
        )
        self.assertEqual(result, datetime.date(2027, 9, 18))


class TestNormalSaveWithIntendedProject(unittest.TestCase):
    """Scenario 1: reviewing then confirming the SAME project/date should stay valid."""

    def test_review_then_confirm_same_selection_is_valid(self):
        target = _compute_target_license_date(
            "Set exact date", datetime.date(2027, 9, 18), datetime.date(2026, 9, 23)
        )
        pending = {
            "project_name": "AD Sint-Kruis Brugge",
            "extend_action": "Set exact date",
            "previous_license_date": "2026-09-18",
            "target_license_date": target.isoformat(),
        }
        # Re-render with the SAME committed project + SAME recomputed target date.
        still_valid = _license_confirmation_still_valid(
            pending, "AD Sint-Kruis Brugge", target.isoformat()
        )
        self.assertTrue(still_valid)


class TestFilteredProjectDisappears(unittest.TestCase):
    """Scenario 2 + 3: the selected project falls out of the filtered list."""

    def test_selection_falls_out_of_filtered_list_triggers_warning(self):
        selected, warning = _resolve_license_project_selection(
            "AD Sint-Kruis Brugge", ["AD Aalter", "AD Brugge Sint Pieters"]
        )
        self.assertEqual(selected, "AD Aalter")
        self.assertIsNotNone(warning)
        self.assertIn("AD Aalter", warning)
        self.assertIn("no longer available", warning)

    def test_first_load_no_previous_selection_no_warning(self):
        # No previous selection yet (fresh page load) -- must NOT warn.
        selected, warning = _resolve_license_project_selection(None, ["AD Aalter", "AD Zwevegem"])
        self.assertEqual(selected, "AD Aalter")
        self.assertIsNone(warning)

    def test_selection_still_present_no_warning(self):
        selected, warning = _resolve_license_project_selection(
            "AD Sint-Kruis Brugge", ["AD Aalter", "AD Sint-Kruis Brugge"]
        )
        self.assertEqual(selected, "AD Sint-Kruis Brugge")
        self.assertIsNone(warning)


class TestConfirmationInvalidation(unittest.TestCase):
    """Scenarios 4 + 5: a pending confirmation must be invalidated by ANY drift."""

    def setUp(self):
        self.target_date = _compute_target_license_date(
            "Set exact date", datetime.date(2027, 9, 18), datetime.date(2026, 9, 23)
        )
        self.pending = {
            "project_name": "AD Sint-Kruis Brugge",
            "extend_action": "Set exact date",
            "previous_license_date": "2026-09-18",
            "target_license_date": self.target_date.isoformat(),
        }

    def test_invalidated_when_project_selection_changes(self):
        # Simulate the dropdown having actually committed to a DIFFERENT project.
        still_valid = _license_confirmation_still_valid(
            self.pending, "AD Brugge Sint Pieters", self.target_date.isoformat()
        )
        self.assertFalse(still_valid)

    def test_invalidated_when_target_date_changes(self):
        other_date = _compute_target_license_date(
            "Set exact date", datetime.date(2026, 11, 1), datetime.date(2026, 9, 23)
        )
        still_valid = _license_confirmation_still_valid(
            self.pending, "AD Sint-Kruis Brugge", other_date.isoformat()
        )
        self.assertFalse(still_valid)


class TestSaveRequiresExplicitConfirmation(unittest.TestCase):
    """Scenario 6: with no pending confirmation, there is nothing valid to save."""

    def test_no_pending_confirmation_is_never_valid(self):
        self.assertFalse(_license_confirmation_still_valid(None, "AD Sint-Kruis Brugge", "2027-09-18"))

    def test_empty_pending_confirmation_is_never_valid(self):
        self.assertFalse(_license_confirmation_still_valid({}, "AD Sint-Kruis Brugge", "2027-09-18"))


if __name__ == "__main__":
    unittest.main()
