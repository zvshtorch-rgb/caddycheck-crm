"""End-to-end smoke test of the Ask Data v3 chat UI (Streamlit AppTest): the feature flag, the editor-only
gate (shadow-deployment plan), and the chat UI itself -- history, follow-ups, and the "Show invoices"/"Show
breakdown" drill-down buttons.

Flag-off behaviour (the existing v2 one-shot form) is already fully covered by test_ask_data_page.py.
Shadow-comparison logging itself (summaries, mismatch flags) is covered by test_ask_data_shadow.py; this
file only checks that the page WIRES one up for every editor-pilot question (it patches the logging call
itself to a no-op so no real file/Supabase write ever happens here).
"""
import importlib.util
import os
import unittest
from pathlib import Path
from unittest.mock import patch

from tests._ask_data_harness import BASE_CFG, _gemini_ok, _make_post, build_fixture

HAS_STREAMLIT = importlib.util.find_spec("streamlit") is not None
APP_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "streamlit_app.py")


class TestAppImportsAskDataV3SymbolsTolerantly(unittest.TestCase):
    """Streamlit Cloud can keep an OLD cached config.settings module in sys.modules while it reloads the new
    streamlit_app.py on redeploy. A hard top-level import of a brand-new name added to that already-imported
    module (get_ask_data_v3_enabled, append_ask_data_shadow_log, load_ask_data_shadow_log) would then crash
    the whole app with ImportError until a full reboot (this happened once already -- see commit 365ae09 for
    the same issue with the bank-payment PDF helpers). This test does not need streamlit installed."""

    def test_new_config_settings_symbols_are_imported_inside_try_except_importerror(self):
        source = Path(APP_PATH).read_text(encoding="utf-8")
        for name in ("get_ask_data_v3_enabled", "append_ask_data_shadow_log", "load_ask_data_shadow_log"):
            self.assertIn(name, source)
        self.assertIn("def get_ask_data_v3_enabled() -> bool:\n        return False", source)
        self.assertIn("def append_ask_data_shadow_log(entry: dict) -> None:\n        pass", source)
        self.assertIn("def load_ask_data_shadow_log() -> list:\n        return []", source)
        self.assertGreaterEqual(source.count("except ImportError:  # pragma: no cover - only hit mid-redeploy before a full reboot"), 2)


@unittest.skipUnless(HAS_STREAMLIT, "streamlit is not installed")
class TestAskDataV3Page(unittest.TestCase):
    def open_page(self, *, v3_enabled: bool, role: str = "admin", gemini_responses=(), log_shadow=False):
        import streamlit as st
        from streamlit.testing.v1 import AppTest

        st.cache_data.clear()
        self.addCleanup(st.cache_data.clear)
        projects, invoices = build_fixture()
        self.posts: list = []
        self.shadow_entries: list = []
        patches = [
            patch("services.supabase_service.load_projects", return_value=projects),
            patch("services.supabase_service.load_invoices", return_value=invoices),
            patch("services.order_approval_service.count_unread_notifications", return_value=0),
            patch("services.order_approval_service.list_notifications", return_value=[]),
            patch("config.settings.load_project_change_log", return_value=[]),
            patch("config.settings.get_gemini_config", lambda: dict(BASE_CFG)),
            patch("services.ask_data.planner.get_gemini_config", lambda: dict(BASE_CFG)),
            patch("config.settings.get_ask_data_v3_enabled", lambda: v3_enabled),
            patch("config.settings.load_ask_data_shadow_log", return_value=[]),
            patch("requests.post", _make_post(list(gemini_responses), self.posts)),
        ]
        if not log_shadow:
            # Most tests here are about the UI, not the shadow log itself (see test_ask_data_shadow.py and
            # the dedicated shadow-wiring tests below) -- never touch the real local/Supabase log for them.
            patches.append(patch("config.settings.append_ask_data_shadow_log", lambda entry: None))
        else:
            patches.append(patch("config.settings.append_ask_data_shadow_log", side_effect=self.shadow_entries.append))
        for p in patches:
            p.start()
            self.addCleanup(p.stop)
        self.projects, self.invoices = projects, invoices
        app = AppTest.from_file(APP_PATH, default_timeout=90)
        app.session_state["role"] = role
        app.run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        app.sidebar.radio[0].set_value("❓ Caddy Chat").run()
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

    # ── Feature flag + editor-only gate ──────────────────────────────────────────────────────────

    def test_flag_off_shows_the_existing_one_shot_form_untouched(self):
        app = self.open_page(v3_enabled=False, role="admin")
        self.assertEqual(len(app.chat_input), 0)
        self.assertTrue(any(b.label == "Ask" for b in app.button))
        self.assertEqual(len(app.main.radio), 0, "no mode switcher when the flag is off, even for an editor")

    def test_viewer_never_sees_v3_or_the_switcher_even_when_the_flag_is_on(self):
        """Goal 2: v3 is editor-only. A viewer session must stay on the v2 form regardless of the flag."""
        app = self.open_page(v3_enabled=True, role="viewer")
        self.assertEqual(len(app.chat_input), 0)
        self.assertTrue(any(b.label == "Ask" for b in app.button))
        self.assertEqual(len(app.main.radio), 0, "no mode switcher for a viewer")

    def test_editor_with_flag_on_sees_the_switcher_defaulting_to_v2(self):
        """Goal 1 + 4: off by default even for an approved editor; a switcher lets them opt into v3."""
        app = self.open_page(v3_enabled=True, role="admin")
        self.assertEqual(len(app.main.radio), 1)
        self.assertEqual(app.main.radio[0].value, "Caddy Chat v2")
        self.assertEqual(len(app.chat_input), 0, "still on v2 until the editor explicitly switches")
        self.assertTrue(any(b.label == "Ask" for b in app.button))

    def test_editor_can_switch_to_v3_and_back(self):
        app = self.open_page(v3_enabled=True, role="admin")
        app.main.radio[0].set_value("Caddy Chat v3 (preview)").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual(len(app.chat_input), 1)
        app.main.radio[0].set_value("Caddy Chat v2").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual(len(app.chat_input), 0)
        self.assertTrue(any(b.label == "Ask" for b in app.button))

    # ── Chat UI behaviour (editor, v3 selected) ──────────────────────────────────────────────────

    def _select_v3(self, app):
        app.main.radio[0].set_value("Caddy Chat v3 (preview)").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        return app

    def test_a_question_and_answer_appear_as_chat_messages(self):
        payload = '{"tool": "get_debt", "arguments": {}}'
        app = self._select_v3(self.open_page(v3_enabled=True, gemini_responses=[_gemini_ok(payload), _gemini_ok(payload)]))
        app.chat_input[0].set_value("How much debt do we have?").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        roles_and_texts = self.chat_texts(app)
        self.assertEqual([r for r, _ in roles_and_texts], ["user", "assistant"])
        self.assertEqual(roles_and_texts[0][1], "How much debt do we have?")
        self.assertIn("28,913.50", roles_and_texts[1][1])

    def test_a_follow_up_question_narrows_the_previous_answer(self):
        payload1 = '{"tool": "get_debt", "arguments": {}}'
        payload2 = '{"tool": "get_debt", "arguments": {"country": "Belgium"}}'
        # order of Gemini calls: v3(q1), shadow-v2(q1), v3(q2), shadow-v2(q2)
        responses = [_gemini_ok(payload1), _gemini_ok(payload1), _gemini_ok(payload2), _gemini_ok(payload2)]
        app = self._select_v3(self.open_page(v3_enabled=True, gemini_responses=responses))
        app.chat_input[0].set_value("How much debt do we have?").run()
        app.chat_input[0].set_value("Only Belgium?").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        roles_and_texts = self.chat_texts(app)
        self.assertEqual([r for r, _ in roles_and_texts], ["user", "assistant", "user", "assistant"])
        self.assertIn("19,999.00", roles_and_texts[3][1])

    def test_a_historical_question_gets_the_limitation_text_with_no_drilldown_buttons(self):
        """Regression test for the production bug: "What changed since September 1?" must never reach
        Gemini/an unrelated tool, and must never offer "Show invoices"/"Show breakdown" (goal 4).

        One Gemini response is still queued: the editor-pilot shadow comparison runs v2 (which has no
        historical-question guard) on the same text in the background for comparison -- that is a deliberate,
        separate feature (see test_ask_data_shadow.py), not v3 itself calling the model."""
        shadow_v2_payload = '{"tool": "get_license_payment_status", "arguments": {}}'
        app = self._select_v3(self.open_page(v3_enabled=True, gemini_responses=[_gemini_ok(shadow_v2_payload)]))
        app.chat_input[0].set_value("What changed since September 1?").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        roles_and_texts = self.chat_texts(app)
        self.assertEqual([r for r, _ in roles_and_texts], ["user", "assistant"])
        self.assertIn("I don't yet have a complete historical snapshot", roles_and_texts[1][1])
        self.assertFalse(any(b.label == "📄 Show invoices" for b in app.button))
        self.assertFalse(any(b.label == "📊 Show breakdown" for b in app.button))
        self.assertEqual(len(self.posts), 1, "only the background v2 shadow comparison may call Gemini, never v3 itself")

    def test_show_invoices_button_drills_down_without_calling_gemini_again(self):
        payload = '{"tool": "get_debt", "arguments": {"country": "Belgium"}}'
        app = self._select_v3(self.open_page(v3_enabled=True, gemini_responses=[_gemini_ok(payload), _gemini_ok(payload)]))
        app.chat_input[0].set_value("Debt in Belgium").run()
        posts_after_question = len(self.posts)
        [b for b in app.button if b.label == "📄 Show invoices"][0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual(len(self.posts), posts_after_question, "the rows drill-down must never call Gemini")
        roles_and_texts = self.chat_texts(app)
        self.assertEqual(roles_and_texts[-1][0], "assistant")
        self.assertIn("row(s) behind that figure", roles_and_texts[-1][1])
        # the initial "Debt in Belgium" answer already carries its own contributing-rows table (Phase 3
        # grounding); the drill-down answer renders a second, equivalent table -- both are expected.
        self.assertEqual(len(app.dataframe), 2)

    def test_new_conversation_button_clears_history_and_context(self):
        payload = '{"tool": "get_debt", "arguments": {"country": "Belgium"}}'
        app = self._select_v3(self.open_page(v3_enabled=True, gemini_responses=[_gemini_ok(payload), _gemini_ok(payload)]))
        app.chat_input[0].set_value("Debt in Belgium").run()
        self.assertEqual(len(self.chat_texts(app)), 2)
        [b for b in app.button if b.label == "🧹 New conversation"][0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual(len(self.chat_texts(app)), 0)

    # ── Shadow comparison wiring (goal 5/6/7) ────────────────────────────────────────────────────

    def test_v3_primary_question_logs_a_shadow_comparison_against_v2(self):
        payload = '{"tool": "get_debt", "arguments": {"country": "Belgium"}}'
        app = self._select_v3(self.open_page(
            v3_enabled=True, gemini_responses=[_gemini_ok(payload), _gemini_ok(payload)], log_shadow=True,
        ))
        app.chat_input[0].set_value("Debt in Belgium").run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual(len(self.shadow_entries), 1)
        entry = self.shadow_entries[0]
        self.assertEqual(entry["primary_mode"], "v3")
        self.assertNotIn("question", entry)
        self.assertEqual(entry["question_word_count"], 3)   # "Debt in Belgium"
        self.assertEqual(entry["v3_outcome_kind"], "tool")
        self.assertEqual(entry["v2_outcome_kind"], "tool")
        self.assertEqual(entry["mismatch_flags"], [])   # same question, same resolved scope -> no divergence

    def test_drilldown_button_does_not_log_a_shadow_comparison(self):
        payload = '{"tool": "get_debt", "arguments": {"country": "Belgium"}}'
        app = self._select_v3(self.open_page(
            v3_enabled=True, gemini_responses=[_gemini_ok(payload), _gemini_ok(payload)], log_shadow=True,
        ))
        app.chat_input[0].set_value("Debt in Belgium").run()
        self.assertEqual(len(self.shadow_entries), 1)
        [b for b in app.button if b.label == "📄 Show invoices"][0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual(len(self.shadow_entries), 1,
                          "the drill-down is answered from the result store, not a new comparable question")

    def test_v2_primary_question_logs_a_shadow_comparison_against_v3(self):
        payload = '{"tool": "get_debt", "arguments": {"country": "Belgium"}}'
        app = self.open_page(v3_enabled=True, gemini_responses=[_gemini_ok(payload), _gemini_ok(payload)], log_shadow=True)
        app.text_area[0].set_value("Debt in Belgium")
        [b for b in app.button if b.label == "Ask"][0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual(len(self.shadow_entries), 1)
        entry = self.shadow_entries[0]
        self.assertEqual(entry["primary_mode"], "v2")
        self.assertEqual(entry["v2_outcome_kind"], "tool")
        self.assertEqual(entry["v3_outcome_kind"], "tool")

    def test_viewer_or_flag_off_never_logs_a_shadow_comparison(self):
        payload = '{"tool": "get_debt", "arguments": {}}'
        app = self.open_page(v3_enabled=False, role="admin", gemini_responses=[_gemini_ok(payload)], log_shadow=True)
        app.text_area[0].set_value("How much debt do we have?")
        [b for b in app.button if b.label == "Ask"][0].click().run()
        self.assertFalse(app.exception, [e.value for e in app.exception])
        self.assertEqual(self.shadow_entries, [])


if __name__ == "__main__":
    unittest.main()
