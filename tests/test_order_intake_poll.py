"""Tests for order_intake_poll auth-error detection and graceful skip logic."""
import unittest
from unittest.mock import MagicMock, patch

from order_intake_poll import _is_auth_error, main


class TestIsAuthError(unittest.TestCase):
    def test_imap_login_failed(self):
        self.assertTrue(_is_auth_error(Exception("b'LOGIN failed.'")))

    def test_authentication_failed(self):
        self.assertTrue(_is_auth_error(Exception("Authentication failed")))

    def test_invalid_credentials(self):
        self.assertTrue(_is_auth_error(Exception("Invalid credentials")))

    def test_authenticationfailed_tag(self):
        self.assertTrue(_is_auth_error(Exception("[AUTHENTICATIONFAILED] Invalid credentials")))

    def test_http_401(self):
        self.assertTrue(_is_auth_error(Exception("HTTP 401 Unauthorized")))

    def test_http_403(self):
        self.assertTrue(_is_auth_error(Exception("403 Forbidden")))

    def test_non_auth_error(self):
        self.assertFalse(_is_auth_error(Exception("Connection refused")))

    def test_timeout_not_auth(self):
        self.assertFalse(_is_auth_error(Exception("timed out")))


class TestMainAuthSkip(unittest.TestCase):
    """main() should return 0 (not 1) when the mailbox rejects authentication."""

    @patch("order_intake_poll.get_email_provider")
    def test_auth_failure_exits_zero(self, mock_get_provider):
        provider = MagicMock()
        provider.is_configured.return_value = True
        provider.name = "imap"
        provider.fetch_unread_with_pdf.side_effect = Exception("b'LOGIN failed.'")
        mock_get_provider.return_value = provider

        with patch("sys.argv", ["order_intake_poll.py"]):
            result = main()

        self.assertEqual(result, 0)

    @patch("order_intake_poll.get_email_provider")
    def test_non_auth_failure_exits_one(self, mock_get_provider):
        provider = MagicMock()
        provider.is_configured.return_value = True
        provider.name = "imap"
        provider.fetch_unread_with_pdf.side_effect = Exception("Connection refused")
        mock_get_provider.return_value = provider

        with patch("sys.argv", ["order_intake_poll.py"]):
            result = main()

        self.assertEqual(result, 1)

    @patch("order_intake_poll.get_email_provider")
    def test_not_configured_exits_zero(self, mock_get_provider):
        provider = MagicMock()
        provider.is_configured.return_value = False
        provider.name = "imap"
        mock_get_provider.return_value = provider

        with patch("sys.argv", ["order_intake_poll.py"]):
            result = main()

        self.assertEqual(result, 0)


if __name__ == "__main__":
    unittest.main()
