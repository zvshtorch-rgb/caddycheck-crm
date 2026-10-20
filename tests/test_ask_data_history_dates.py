"""Tests for services/ask_data/history_dates.py: deterministic EN+Hebrew date/period phrase resolution.

Every case is anchored on a FIXED `today` (2026-10-20, a Tuesday) so results are reproducible regardless of
when the suite actually runs -- the model never supplies this date; only the caller's real wall-clock date
would be used in production (see planner.py's history rules)."""
import unittest
from datetime import date

from services.ask_data.history_dates import (
    ACCEPTED, NOT_FOUND, resolve_period_phrase, resolve_single_date_phrase,
)

TODAY = date(2026, 10, 20)


class TestResolveSingleDatePhrase(unittest.TestCase):
    def test_today(self):
        self.assertEqual(resolve_single_date_phrase("today", TODAY).value, TODAY)

    def test_yesterday(self):
        self.assertEqual(resolve_single_date_phrase("yesterday", TODAY).value, date(2026, 10, 19))

    def test_explicit_day_and_month_in_the_past_this_year(self):
        self.assertEqual(resolve_single_date_phrase("30 September", TODAY).value, date(2026, 9, 30))
        self.assertEqual(resolve_single_date_phrase("September 30", TODAY).value, date(2026, 9, 30))

    def test_explicit_date_in_the_future_rolls_back_a_year(self):
        # "1 December" has not happened yet in 2026 relative to TODAY -> most recent PAST occurrence (2025).
        result = resolve_single_date_phrase("1 December", TODAY)
        self.assertEqual(result.value, date(2025, 12, 1))

    def test_a_week_ago(self):
        self.assertEqual(resolve_single_date_phrase("a week ago", TODAY).value, date(2026, 10, 13))
        self.assertEqual(resolve_single_date_phrase("last week", TODAY).value, date(2026, 10, 13))

    def test_a_month_ago(self):
        self.assertEqual(resolve_single_date_phrase("a month ago", TODAY).value, date(2026, 9, 20))

    def test_n_days_ago(self):
        self.assertEqual(resolve_single_date_phrase("20 days ago", TODAY).value, date(2026, 9, 30))

    def test_hebrew_today_yesterday(self):
        self.assertEqual(resolve_single_date_phrase("היום", TODAY).value, TODAY)
        self.assertEqual(resolve_single_date_phrase("אתמול", TODAY).value, date(2026, 10, 19))

    def test_hebrew_a_week_ago(self):
        self.assertEqual(resolve_single_date_phrase("לפני שבוע", TODAY).value, date(2026, 10, 13))

    def test_range_only_phrases_are_not_found_as_a_single_point(self):
        for phrase in ("since September 1", "past 30 days", "this week"):
            with self.subTest(phrase=phrase):
                self.assertEqual(resolve_single_date_phrase(phrase, TODAY).status, NOT_FOUND)

    def test_empty_and_gibberish(self):
        self.assertEqual(resolve_single_date_phrase("", TODAY).status, NOT_FOUND)
        self.assertEqual(resolve_single_date_phrase("sometime last decade", TODAY).status, NOT_FOUND)

    def test_deterministic_and_idempotent(self):
        """Calling twice with the same inputs always gives the same answer (no hidden clock/state)."""
        a = resolve_single_date_phrase("30 September", TODAY)
        b = resolve_single_date_phrase("30 September", TODAY)
        self.assertEqual(a, b)


class TestResolvePeriodPhrase(unittest.TestCase):
    def test_today(self):
        result = resolve_period_phrase("today", TODAY)
        self.assertEqual((result.date_from, result.date_to), (TODAY, TODAY))

    def test_this_week_starts_monday_and_ends_today(self):
        result = resolve_period_phrase("this week", TODAY)
        self.assertEqual(result.date_from.weekday(), 0)
        self.assertEqual(result.date_to, TODAY)
        self.assertLessEqual(result.date_from, TODAY)

    def test_last_week_is_the_full_prior_monday_to_sunday(self):
        result = resolve_period_phrase("last week", TODAY)
        self.assertEqual(result.date_from, date(2026, 10, 12))
        self.assertEqual(result.date_to, date(2026, 10, 18))

    def test_this_month(self):
        result = resolve_period_phrase("this month", TODAY)
        self.assertEqual((result.date_from, result.date_to), (date(2026, 10, 1), TODAY))

    def test_last_month_is_the_full_prior_calendar_month(self):
        result = resolve_period_phrase("last month", TODAY)
        self.assertEqual((result.date_from, result.date_to), (date(2026, 9, 1), date(2026, 9, 30)))

    def test_since_a_specific_date(self):
        result = resolve_period_phrase("since September 1", TODAY)
        self.assertEqual((result.date_from, result.date_to), (date(2026, 9, 1), TODAY))

    def test_since_the_beginning_of_a_month(self):
        result = resolve_period_phrase("since the beginning of October", TODAY)
        self.assertEqual((result.date_from, result.date_to), (date(2026, 10, 1), TODAY))

    def test_past_n_days(self):
        result = resolve_period_phrase("past 30 days", TODAY)
        self.assertEqual((result.date_from, result.date_to), (date(2026, 9, 20), TODAY))

    def test_bare_past_month_name_is_the_whole_month(self):
        result = resolve_period_phrase("September", TODAY)
        self.assertEqual((result.date_from, result.date_to), (date(2026, 9, 1), date(2026, 9, 30)))

    def test_bare_current_month_name_is_clipped_to_today(self):
        result = resolve_period_phrase("October", TODAY)
        self.assertEqual((result.date_from, result.date_to), (date(2026, 10, 1), TODAY))

    def test_never_extrapolates_past_today(self):
        for phrase in ("today", "this week", "this month", "since September 1", "October"):
            with self.subTest(phrase=phrase):
                result = resolve_period_phrase(phrase, TODAY)
                self.assertLessEqual(result.date_to, TODAY)

    def test_hebrew_this_month_since_beginning(self):
        result = resolve_period_phrase("מתחילת החודש", TODAY)
        self.assertEqual((result.date_from, result.date_to), (date(2026, 10, 1), TODAY))

    def test_hebrew_in_october(self):
        result = resolve_period_phrase("באוקטובר", TODAY)
        self.assertEqual((result.date_from, result.date_to), (date(2026, 10, 1), TODAY))

    def test_hebrew_a_week_ago_falls_back_to_a_single_day_period(self):
        result = resolve_period_phrase("לפני שבוע", TODAY)
        self.assertEqual((result.date_from, result.date_to), (date(2026, 10, 13), date(2026, 10, 13)))

    def test_unrecognised_phrase_is_not_found(self):
        self.assertEqual(resolve_period_phrase("sometime last decade", TODAY).status, NOT_FOUND)
        self.assertEqual(resolve_period_phrase("", TODAY).status, NOT_FOUND)

    def test_str_label_is_human_readable_for_scope_text(self):
        result = resolve_period_phrase("this month", TODAY)
        self.assertIn("this month", str(result))
        self.assertIn("2026-10-01", str(result))
        self.assertIn("2026-10-20", str(result))

    def test_deterministic_and_idempotent(self):
        a = resolve_period_phrase("since September 1", TODAY)
        b = resolve_period_phrase("since September 1", TODAY)
        self.assertEqual(a, b)


if __name__ == "__main__":
    unittest.main()
