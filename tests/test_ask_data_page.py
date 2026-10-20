"""End-to-end smoke test of the real Ask Data page (Streamlit AppTest) after the Phase 1 extraction.

Runs the real streamlit_app.py script with Supabase, Gemini and the network mocked. Skipped when Streamlit
is not installed. The expected answers are computed with the same deterministic executor, so this test
proves the *page wiring* (imports, wrappers, session state, rendering), not the arithmetic (that is pinned by
test_ask_data_characterization.py).
"""
import importlib.util
import os
import unittest
from unittest.mock import patch

from tests._ask_data_harness import BASE_CFG, NewImpl, _gemini_ok, _make_post, build_fixture

HAS_STREAMLIT = importlib.util.find_spec("streamlit") is not None
APP_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "streamlit_app.py")
NO_KEY = {"api_key": "", "model": "m"}


@unittest.skipUnless(HAS_STREAMLIT, "streamlit is not installed")
class TestAskDataPage(unittest.TestCase):
    def open_page(self, cfg, gemini_responses=()):
        import streamlit as st
        from streamlit.testing.v1 import AppTest

        st.cache_data.clear()
        self.addCleanup(st.cache_data.clear)  # load_data() is cached; never leak this fixture into other UI tests
        projects, invoices = build_fixture()
        self.posts: list = []
        patches = [
            patch("services.supabase_service.load_projects", return_value=projects),
            patch("services.supabase_service.load_invoices", return_value=invoices),
            patch("services.order_approval_service.count_unread_notifications", return_value=0),
            patch("services.order_approval_service.list_notifications", return_value=[]),
            patch("config.settings.load_project_change_log", return_value=[]),
            patch("config.settings.get_gemini_config", lambda: dict(cfg)),
            patch("services.ask_data.planner.get_gemini_config", lambda: dict(cfg)),
            patch("requests.post", _make_post(list(gemini_responses), self.posts)),
        ]
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        self.projects, self.invoices = projects, invoices
        app = AppTest.from_file(APP_PATH, default_timeout=90)
        app.session_state["role"] = "viewer"
        app.run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        app.sidebar.radio[0].set_value("❓ Ask Data").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        return app

    @staticmethod
    def ask(app, question):
        app.text_area[0].set_value(question)
        [b for b in app.button if b.label == "Ask"][0].click().run()
        return app

    def expected(self, tool, args):
        return NewImpl().execute(tool, args, self.projects, self.invoices, next_invoice=lambda: 0)

    def test_page_renders_and_legacy_keyword_path_answers_without_gemini(self):
        app = self.open_page(NO_KEY)
        self.ask(app, "How many cameras do we have in total?")
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual(self.posts, [], "no Gemini request may be made without an API key")
        self.assertEqual([s.value for s in app.success], ["There are 47 camera(s) across 6 active project(s)."])

    def test_gemini_tool_selection_path_returns_the_deterministic_answer(self):
        payload = '{"tool": "get_outstanding_amount", "arguments": {"country": "Bel"}}'
        app = self.open_page(BASE_CFG, [_gemini_ok(payload)])
        self.ask(app, "How much does Belgium owe us?")
        self.assertFalse(app.exception, [e.value for e in app.exception])
        text, _ = self.expected("get_outstanding_amount", {"country": "Bel"})
        self.assertEqual([s.value for s in app.success], [text])
        self.assertEqual(len(self.posts), 1)

    def test_tool_result_with_rows_shows_the_invoice_table(self):
        payload = '{"tool": "get_unpaid_invoices", "arguments": {"country": "Ger"}}'
        app = self.open_page(BASE_CFG, [_gemini_ok(payload)])
        self.ask(app, "Show unpaid invoices in Germany")
        self.assertFalse(app.exception, [e.value for e in app.exception])
        text, df = self.expected("get_unpaid_invoices", {"country": "Ger"})
        self.assertEqual([s.value for s in app.success], [text])
        self.assertEqual(len(app.dataframe), 1)
        self.assertEqual(len(app.dataframe[0].value), len(df))

    def test_gemini_failure_falls_back_to_keyword_matcher(self):
        app = self.open_page(BASE_CFG, [{"status": 500, "json": {}}, {"status": 500, "json": {}}, {"status": 500, "json": {}}])
        self.ask(app, "How many cameras do we have in total?")
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual([s.value for s in app.success], ["There are 47 camera(s) across 6 active project(s)."])

    def test_unresolvable_tool_call_falls_back_without_error(self):
        app = self.open_page(BASE_CFG, [_gemini_ok('{"tool": "drop_table", "arguments": {}}'), {"status": 500, "json": {}}])
        self.ask(app, "How many cameras do we have in total?")
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual([s.value for s in app.success], ["There are 47 camera(s) across 6 active project(s)."])


if __name__ == "__main__":
    unittest.main()
