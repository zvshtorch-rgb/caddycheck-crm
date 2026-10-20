"""Ask Data shadow log storage: local JSON fallback (round-trip, bounded retention) and graceful fallback
when Supabase is unavailable or the ask_data_shadow_log table doesn't exist yet (pre-migration)."""
import importlib
import unittest
from unittest import mock

import config.settings as settings


class TestAskDataShadowLogLocalFallback(unittest.TestCase):
    def setUp(self):
        self._tmp = self.enterContext(mock.patch.object(settings, "ASK_DATA_SHADOW_LOG_FILE",
                                                          settings.CONFIG_DIR / "_test_ask_data_shadow_log.json"))
        self.addCleanup(lambda: settings.ASK_DATA_SHADOW_LOG_FILE.unlink(missing_ok=True))
        settings.ASK_DATA_SHADOW_LOG_FILE.unlink(missing_ok=True)

    def test_append_then_load_round_trips_locally_when_supabase_is_unconfigured(self):
        with mock.patch("services.supabase_service.append_ask_data_shadow_log",
                        side_effect=RuntimeError("Supabase credentials not configured")), \
                mock.patch("services.supabase_service.load_ask_data_shadow_log",
                          side_effect=RuntimeError("Supabase credentials not configured")):
            settings.append_ask_data_shadow_log({"question": "How much debt?", "primary_mode": "v3"})
            loaded = settings.load_ask_data_shadow_log()
        self.assertEqual(len(loaded), 1)
        self.assertEqual(loaded[0]["question"], "How much debt?")

    def test_missing_table_error_also_falls_back_locally(self):
        with mock.patch("services.supabase_service.append_ask_data_shadow_log",
                        side_effect=Exception("Could not find the table 'public.ask_data_shadow_log'")), \
                mock.patch("services.supabase_service.load_ask_data_shadow_log",
                          side_effect=Exception("Could not find the table 'public.ask_data_shadow_log'")):
            settings.append_ask_data_shadow_log({"question": "q1"})
            loaded = settings.load_ask_data_shadow_log()
        self.assertEqual(len(loaded), 1)

    def test_local_fallback_is_bounded_to_the_configured_max(self):
        with mock.patch.object(settings, "ASK_DATA_SHADOW_LOG_LOCAL_MAX_ENTRIES", 3), \
                mock.patch("services.supabase_service.append_ask_data_shadow_log",
                          side_effect=RuntimeError("Supabase credentials not configured")), \
                mock.patch("services.supabase_service.load_ask_data_shadow_log",
                          side_effect=RuntimeError("Supabase credentials not configured")):
            for i in range(5):
                settings.append_ask_data_shadow_log({"question": f"q{i}"})
            loaded = settings.load_ask_data_shadow_log()
        self.assertEqual(len(loaded), 3)
        self.assertEqual([e["question"] for e in loaded], ["q2", "q3", "q4"])   # kept the most recent 3

    def test_supabase_success_is_preferred_over_local(self):
        remote = [{"question": "from supabase"}]
        with mock.patch("services.supabase_service.load_ask_data_shadow_log", return_value=remote):
            loaded = settings.load_ask_data_shadow_log()
        self.assertEqual(loaded, remote)


if __name__ == "__main__":
    unittest.main()
