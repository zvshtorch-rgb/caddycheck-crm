import argparse
import imaplib
import unittest
from unittest.mock import patch

import order_intake_poll


class _ConfiguredProvider:
    name = "imap"

    def is_configured(self):
        return True

    def fetch_unread_with_pdf(self):
        raise NotImplementedError


class _ImapLoginFailureProvider(_ConfiguredProvider):
    def fetch_unread_with_pdf(self):
        raise imaplib.IMAP4.error(b"LOGIN failed.")


class _GenericFailureProvider(_ConfiguredProvider):
    def fetch_unread_with_pdf(self):
        raise RuntimeError("boom")


class OrderIntakePollTests(unittest.TestCase):
    def test_main_skips_imap_login_failures(self):
        with patch.object(order_intake_poll.argparse.ArgumentParser, "parse_args", return_value=argparse.Namespace(dry_run=False)):
            with patch.object(order_intake_poll, "get_email_provider", return_value=_ImapLoginFailureProvider()):
                self.assertEqual(order_intake_poll.main(), 0)

    def test_main_still_fails_other_fetch_errors(self):
        with patch.object(order_intake_poll.argparse.ArgumentParser, "parse_args", return_value=argparse.Namespace(dry_run=False)):
            with patch.object(order_intake_poll, "get_email_provider", return_value=_GenericFailureProvider()):
                self.assertEqual(order_intake_poll.main(), 1)


if __name__ == "__main__":
    unittest.main()
