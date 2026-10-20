"""End-to-end smoke test of the Ask Data v3 chat UI (Streamlit AppTest), behind the ask_data_v3 feature flag.

Mirrors tests/test_ask_data_page.py's harness (same fixture, same Supabase/Gemini mocking) but drives the
new chat interface added in streamlit_app.py's "else" branch. Flag-off behaviour (the existing v2 one-shot
form) is already fully covered by test_ask_data_page.py; this file only needs to prove: (1) the flag
correctly switches between the two UIs, and (2) the v3 chat UI itself -- history, follow-ups, and the
"Show invoices"/"Show breakdown" drill-down buttons -- renders and behaves as expected.
"""
import importlib.util
import os
import unittest
from unittest.mock import patch

from tests._ask_data_harness import BASE_CFG, _gemini_ok, _make_post, build_fixture

HAS_STREAMLIT = importlib.util.find_spec("streamlit") is not None
APP_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "streamlit_app.py")


@unittest.skipUnless(HAS_STREAMLIT, "streamlit is not installed")
class TestAskDataV3Page(unittest.TestCase):
    def open_page(self, *, v3_enabled: bool, gemini_responses=()):
        import streamlit as st
        from streamlit.testing.v1 import AppTest

        st.cache_data.clear()
        self.addCleanup(st.cache_data.clear)
        projects, invoices = build_fixture()
        self.posts: list = []
        patches = [
            patch("services.supabase_service.load_projects", return_value=projects),
            patch("services.supabase_service.load_invoices", return_value=invoices),
            patch("services.order_approval_service.count_unread_notifications", return_value=0),
            patch("services.order_approval_service.list_notifications", return_value=[]),
            patch("config.settings.load_project_change_log", return_value=[]),
            patch("config.settings.get_gemini_config", lambda: dict(BASE_CFG)),
            patch("services.ask_data.planner.get_gemini_config", lambda: dict(BASE_CFG)),
            patch("config.settings.get_ask_data_v3_enabled", lambda: v3_enabled),
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
    def chat_texts(app):
        """(role, markdown text) for every rendered st.chat_message, in order."""
        out = []
        for block in app.chat_message:
            md = block.markdown
            out.append((block.name, md[0].value if md else ""))
        return out

    def test_flag_off_shows_the_existing_one_shot_form_untouched(self):
        app = self.open_page(v3_enabled=False)
        self.assertEqual(len(app.chat_input), 0)
        self.assertTrue(any(b.label == "Ask" for b in app.button))

    def test_flag_on_shows_the_chat_ui_instead(self):
        app = self.open_page(v3_enabled=True)
        self.assertEqual(len(app.chat_input), 1)
        self.assertFalse(any(b.label == "Ask" for b in app.button))

    def test_a_question_and_answer_appear_as_chat_messages(self):
        payload = '{"tool": "get_debt", "arguments": {}}'
        app = self.open_page(v3_enabled=True, gemini_responses=[_gemini_ok(payload)])
        app.chat_input[0].set_value("How much debt do we have?").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        roles_and_texts = self.chat_texts(app)
        self.assertEqual([r for r, _ in roles_and_texts], ["user", "assistant"])
        self.assertEqual(roles_and_texts[0][1], "How much debt do we have?")
        self.assertIn("28,913.50", roles_and_texts[1][1])

    def test_a_follow_up_question_narrows_the_previous_answer(self):
        payload1 = '{"tool": "get_debt", "arguments": {}}'
        payload2 = '{"tool": "get_debt", "arguments": {"country": "Belgium"}}'
        app = self.open_page(v3_enabled=True, gemini_responses=[_gemini_ok(payload1), _gemini_ok(payload2)])
        app.chat_input[0].set_value("How much debt do we have?").run()
        app.chat_input[0].set_value("Only Belgium?").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        roles_and_texts = self.chat_texts(app)
        self.assertEqual([r for r, _ in roles_and_texts], ["user", "assistant", "user", "assistant"])
        self.assertIn("19,999.00", roles_and_texts[3][1])

    def test_show_invoices_button_drills_down_without_calling_gemini_again(self):
        payload = '{"tool": "get_debt", "arguments": {"country": "Belgium"}}'
        app = self.open_page(v3_enabled=True, gemini_responses=[_gemini_ok(payload)])
        app.chat_input[0].set_value("Debt in Belgium").run()
        self.assertEqual(len(self.posts), 1)
        [b for b in app.button if b.label == "📄 Show invoices"][0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual(len(self.posts), 1, "the rows drill-down must never call Gemini")
        roles_and_texts = self.chat_texts(app)
        self.assertEqual(roles_and_texts[-1][0], "assistant")
        self.assertIn("row(s) behind that figure", roles_and_texts[-1][1])
        # the initial "Debt in Belgium" answer already carries its own contributing-rows table (Phase 3
        # grounding); the drill-down answer renders a second, equivalent table -- both are expected.
        self.assertEqual(len(app.dataframe), 2)

    def test_new_conversation_button_clears_history_and_context(self):
        payload = '{"tool": "get_debt", "arguments": {"country": "Belgium"}}'
        app = self.open_page(v3_enabled=True, gemini_responses=[_gemini_ok(payload)])
        app.chat_input[0].set_value("Debt in Belgium").run()
        self.assertEqual(len(self.chat_texts(app)), 2)
        [b for b in app.button if b.label == "🧹 New conversation"][0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual(len(self.chat_texts(app)), 0)


if __name__ == "__main__":
    unittest.main()
