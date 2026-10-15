"""Tests for the Microsoft Graph mailbox provider (auth, errors, paging, read-only behavior)."""
import datetime
import os
import unittest
from unittest.mock import patch

import requests

from services import email_intake_service as svc
from services.email_intake_service import GraphEmailProvider, IntakeError, get_email_provider

SECRET = "S3CRET-CLIENT-VALUE"
TOKEN = "ACCESS-TOKEN-VALUE"
ENV = {
    "ORDER_INTAKE_GRAPH_TENANT_ID": "tenant-123",
    "ORDER_INTAKE_GRAPH_CLIENT_ID": "client-456",
    "ORDER_INTAKE_GRAPH_CLIENT_SECRET": SECRET,
    "ORDER_INTAKE_GRAPH_MAILBOX": "orders@contoso.example",
    "ORDER_INTAKE_GRAPH_FOLDER": "Inbox",
}


class Resp:
    def __init__(self, status=200, json_data=None, content=b"", headers=None):
        self.status_code = status
        self._json = json_data
        self.content = content
        self.headers = headers or {}

    @property
    def ok(self):
        return self.status_code < 400

    def json(self):
        if self._json is None:
            raise ValueError("no json")
        return self._json


def token_response():
    return Resp(200, {"access_token": TOKEN, "expires_in": 3600})


class GraphTestCase(unittest.TestCase):
    def setUp(self):
        env_patch = patch.dict(os.environ, ENV)
        env_patch.start()
        self.addCleanup(env_patch.stop)
        self.calls = []
        self.handler = lambda method, url, kwargs: Resp(404, {"error": {"code": "ErrorItemNotFound"}})
        request_patch = patch("requests.request", side_effect=self._request)
        request_patch.start()
        self.addCleanup(request_patch.stop)
        sleep_patch = patch.object(svc, "_sleep")
        self.sleeps = sleep_patch.start()
        self.addCleanup(sleep_patch.stop)
        self.provider = GraphEmailProvider()

    def _request(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        if "login.microsoftonline.com" in url:
            return self.token_handler(method, url, kwargs)
        return self.handler(method, url, kwargs)

    token_handler = staticmethod(lambda method, url, kwargs: token_response())

    def graph_calls(self):
        return [c for c in self.calls if "graph.microsoft.com" in c[1]]


class TestAuthentication(GraphTestCase):
    def test_preflight_succeeds_with_valid_credentials(self):
        self.handler = lambda m, u, k: Resp(200, {"displayName": "Inbox", "totalItemCount": 7, "unreadItemCount": 2})
        info = self.provider.preflight()
        self.assertEqual((info["folder"], info["total_items"]), ("Inbox", 7))
        method, url, kwargs = self.graph_calls()[0]
        self.assertTrue(url.endswith("/users/orders@contoso.example/mailFolders/Inbox"))
        self.assertEqual(kwargs["headers"]["Authorization"], f"Bearer {TOKEN}")

    def test_invalid_client_secret_is_an_auth_error_without_leaking_secrets(self):
        self.token_handler = lambda m, u, k: Resp(401, {
            "error": "invalid_client",
            "error_description": "AADSTS7000215: Invalid client secret provided.\r\nTrace ID: abc",
        })
        with self.assertRaises(IntakeError) as ctx:
            self.provider.preflight()
        self.assertEqual((ctx.exception.kind, ctx.exception.status, ctx.exception.code), ("auth", 401, "invalid_client"))
        self.assertNotIn(SECRET, str(ctx.exception))
        self.assertNotIn("\n", str(ctx.exception))

    def test_error_text_keeps_the_aadsts_code_but_drops_tenant_and_client_ids(self):
        self.token_handler = lambda m, u, k: Resp(400, {
            "error": "invalid_request",
            "error_description": "AADSTS900021: Requested tenant identifier 'tenant-123' is not valid for app 'client-456'.",
        })
        with self.assertRaises(IntakeError) as ctx:
            self.provider.preflight()
        self.assertIn("AADSTS900021", str(ctx.exception))
        for leaked in ("tenant-123", "client-456", "orders@contoso"):
            self.assertNotIn(leaked, str(ctx.exception))

    def test_unknown_tenant_is_an_auth_error(self):
        self.token_handler = lambda m, u, k: Resp(400, {"error": "invalid_request", "error_description": "AADSTS90002"})
        with self.assertRaises(IntakeError) as ctx:
            self.provider.preflight()
        self.assertEqual(ctx.exception.kind, "auth")

    def test_mailbox_access_denied_is_an_auth_error(self):
        self.handler = lambda m, u, k: Resp(403, {"error": {"code": "ErrorAccessDenied", "message": "Access is denied."}})
        with self.assertRaises(IntakeError) as ctx:
            self.provider.preflight()
        self.assertEqual((ctx.exception.kind, ctx.exception.code), ("auth", "ErrorAccessDenied"))

    def test_missing_mailbox_is_a_mailbox_error(self):
        self.handler = lambda m, u, k: Resp(404, {"error": {"code": "ErrorInvalidUser", "message": "no such mailbox"}})
        with self.assertRaises(IntakeError) as ctx:
            self.provider.preflight()
        self.assertEqual(ctx.exception.kind, "mailbox")

    def test_missing_settings_are_reported_by_name_only(self):
        with patch.dict(os.environ, {"ORDER_INTAKE_GRAPH_CLIENT_SECRET": ""}):
            provider = GraphEmailProvider()
        self.assertFalse(provider.is_configured())
        self.assertEqual(provider.missing_settings(), ["ORDER_INTAKE_GRAPH_CLIENT_SECRET"])
        with self.assertRaises(IntakeError) as ctx:
            provider.preflight()
        self.assertEqual(ctx.exception.kind, "config")
        self.assertEqual(self.calls, [])

    def test_token_is_cached_between_calls(self):
        self.handler = lambda m, u, k: Resp(200, {"displayName": "Inbox"})
        self.provider.preflight()
        self.provider.preflight()
        self.assertEqual(len([c for c in self.calls if "login.microsoftonline.com" in c[1]]), 1)

    def test_graph_is_the_default_provider(self):
        with patch.dict(os.environ, {"ORDER_INTAKE_PROVIDER": ""}):
            self.assertIsInstance(get_email_provider(), GraphEmailProvider)


class TestResilience(GraphTestCase):
    def test_throttling_is_retried_with_retry_after(self):
        responses = [Resp(429, {"error": {"code": "TooManyRequests"}}, headers={"Retry-After": "3"}),
                     Resp(200, {"displayName": "Inbox"})]
        self.handler = lambda m, u, k: responses.pop(0)
        self.assertEqual(self.provider.preflight()["folder"], "Inbox")
        self.sleeps.assert_called_once_with(3.0)

    def test_persistent_server_errors_fail_as_transient(self):
        self.handler = lambda m, u, k: Resp(503, {"error": {"code": "ServiceUnavailable"}})
        with self.assertRaises(IntakeError) as ctx:
            self.provider.preflight()
        self.assertEqual((ctx.exception.kind, ctx.exception.status), ("transient", 503))
        self.assertEqual(len(self.graph_calls()), 3)

    def test_network_errors_fail_without_leaking_details(self):
        def boom(method, url, kwargs):
            raise requests.ConnectionError(f"failed for {SECRET}")

        self.handler = boom
        with self.assertRaises(IntakeError) as ctx:
            self.provider.preflight()
        self.assertEqual(ctx.exception.kind, "transient")
        self.assertNotIn(SECRET, str(ctx.exception))


class TestListing(GraphTestCase):
    @staticmethod
    def _page(ids, next_link=None):
        payload = {"value": [
            {"id": f"AAMk={i}", "internetMessageId": f"<{i}@x>", "subject": f"S{i}",
             "from": {"emailAddress": {"address": "tom@example.com"}},
             "receivedDateTime": "2026-10-15T12:15:00Z", "bodyPreview": "p"}
            for i in ids]}
        if next_link:
            payload["@odata.nextLink"] = next_link
        return Resp(200, payload)

    def test_all_pages_are_read_regardless_of_read_state_and_nothing_is_modified(self):
        pages = {
            "messages": self._page([1, 2], "https://graph.microsoft.com/v1.0/page2"),
            "page2": self._page([3, 4], "https://graph.microsoft.com/v1.0/page3"),
            "page3": self._page([5]),
        }
        self.handler = lambda m, u, k: pages["messages" if u.endswith("/messages") else u.rsplit("/", 1)[-1]]
        since = datetime.datetime(2026, 10, 8, 13, 0, tzinfo=datetime.timezone.utc)
        headers = self.provider.list_message_headers(since)

        self.assertEqual([h.subject for h in headers], ["S1", "S2", "S3", "S4", "S5"])
        calls = self.graph_calls()
        self.assertEqual(len(calls), 3)
        first_params = calls[0][2]["params"]
        self.assertIn("receivedDateTime ge 2026-10-08T13:00:00Z", first_params["$filter"])
        self.assertIn("hasAttachments eq true", first_params["$filter"])
        self.assertNotIn("isRead", first_params["$filter"])
        self.assertIsNone(calls[1][2]["params"])
        self.assertTrue(all(c[0] == "GET" for c in calls))  # no PATCH/POST: read state is never changed
        self.assertTrue(all(c[2]["headers"]["Prefer"] == 'IdType="ImmutableId"' for c in calls))

    def test_header_carries_rfc_and_immutable_ids(self):
        self.handler = lambda m, u, k: self._page([9])
        header = self.provider.list_message_headers(datetime.datetime.now(datetime.timezone.utc))[0]
        self.assertEqual((header.message_id, header.graph_id, header.from_address), ("<9@x>", "AAMk=9", "tom@example.com"))
        self.assertEqual(header.received_at, datetime.datetime(2026, 10, 15, 12, 15, tzinfo=datetime.timezone.utc))

    def test_runaway_paging_fails_loudly(self):
        self.handler = lambda m, u, k: self._page([1], "https://graph.microsoft.com/v1.0/again")
        with self.assertRaises(IntakeError):
            self.provider.list_message_headers(datetime.datetime.now(datetime.timezone.utc), max_pages=3)

    def test_listing_error_status_is_classified(self):
        self.handler = lambda m, u, k: Resp(403, {"error": {"code": "ErrorAccessDenied"}})
        with self.assertRaises(IntakeError) as ctx:
            self.provider.list_message_headers(datetime.datetime.now(datetime.timezone.utc))
        self.assertEqual(ctx.exception.kind, "auth")


class TestAttachments(GraphTestCase):
    HEADER = svc.MessageHeader("<1@x>", "AAMk=1/abc", "tom@example.com", "Order", None, "preview")

    def test_first_pdf_is_downloaded_and_inline_images_ignored(self):
        def handler(method, url, kwargs):
            if url.endswith("/attachments"):
                return Resp(200, {"value": [
                    {"id": "att1", "name": "logo.png", "contentType": "image/png", "isInline": True},
                    {"id": "att2", "name": "PO 123.PDF", "contentType": "application/octet-stream"},
                ]})
            return Resp(200, content=b"%PDF-1.7 bytes")

        self.handler = handler
        message = self.provider.fetch_pdf_message(self.HEADER)
        self.assertEqual((message.pdf_filename, message.pdf_bytes, message.raw_id), ("PO 123.PDF", b"%PDF-1.7 bytes", "AAMk=1/abc"))
        self.assertEqual(message.message_id, "<1@x>")
        urls = [c[1] for c in self.graph_calls()]
        self.assertIn("/messages/AAMk%3D1%2Fabc/attachments", urls[0])
        self.assertTrue(urls[1].endswith("/attachments/att2/$value"))

    def test_message_without_pdf_returns_none(self):
        self.handler = lambda m, u, k: Resp(200, {"value": [{"id": "a", "name": "logo.png", "contentType": "image/png"}]})
        self.assertIsNone(self.provider.fetch_pdf_message(self.HEADER))
        self.assertEqual(len(self.graph_calls()), 1)

    def test_attachment_download_error_is_classified(self):
        def handler(method, url, kwargs):
            if url.endswith("/attachments"):
                return Resp(200, {"value": [{"id": "a", "name": "x.pdf", "contentType": "application/pdf"}]})
            return Resp(404, {"error": {"code": "ErrorItemNotFound"}})

        self.handler = handler
        with self.assertRaises(IntakeError) as ctx:
            self.provider.fetch_pdf_message(self.HEADER)
        self.assertEqual(ctx.exception.kind, "mailbox")


if __name__ == "__main__":
    unittest.main()
