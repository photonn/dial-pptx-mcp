"""Tests for the DIAL Files API bridge (workstream 5.4), against a stub DIAL Core."""
import base64
import os
import re
import sys
import threading
import unittest
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


class _StubDial(BaseHTTPRequestHandler):
    last_upload = {}
    appdata = None  # set to a path to simulate an application context

    def log_message(self, *args):
        pass

    def _authed(self):
        return self.headers.get("Api-Key") == "test-key"

    def do_GET(self):
        files = re.match(r"^/v1/files/([^/]+)/(.+)$", self.path)
        if files and self._authed():
            # DIAL storage is per-user: anything outside this caller's own
            # bucket comes back 403, as Core does.
            if files.group(1) != "STUBBUCKET":
                self.send_response(403)
                self.end_headers()
                return
            body = b"FILEBYTES"
            self.send_response(200)
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if self.path == "/v1/bucket" and self._authed():
            body = {"bucket": "STUBBUCKET"}
            if _StubDial.appdata:
                body["appdata"] = _StubDial.appdata
            import json
            self._json(json.dumps(body).encode())
        else:
            self.send_response(401 if not self._authed() else 404)
            self.end_headers()

    def do_PUT(self):
        m = re.match(r"^/v1/files/([^/]+)/(.+)$", self.path)
        if not (m and self._authed()):
            self.send_response(401)
            self.end_headers()
            return
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)
        _StubDial.last_upload = {"path": self.path, "size": len(body)}
        url = f"files/{m.group(1)}/{m.group(2)}"
        self._json(('{"url": "%s"}' % url).encode())

    def _json(self, body):
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


class TestDialFileClient(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = HTTPServer(("127.0.0.1", 0), _StubDial)
        cls.port = cls.server.server_address[1]
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()
        os.environ["DIAL_CORE_URL"] = f"http://127.0.0.1:{cls.port}"
        os.environ["DIAL_API_KEY"] = "test-key"
        # No MCP request context in unit tests: allow the env-key fallback
        os.environ["DIAL_AUTH_MODE"] = "auto"

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        os.environ.pop("DIAL_CORE_URL", None)
        os.environ.pop("DIAL_API_KEY", None)
        os.environ.pop("DIAL_AUTH_MODE", None)

    def test_bucket_and_upload(self):
        from dial_client import DialFileClient
        client = DialFileClient()
        self.assertEqual(client.get_bucket(), "STUBBUCKET")
        url = client.upload(b"hello-pptx", "deck.pptx")
        self.assertTrue(url.startswith("files/STUBBUCKET/pptx-mcp/"))
        self.assertTrue(url.endswith("/deck.pptx"))
        self.assertGreater(_StubDial.last_upload["size"], 10)  # multipart envelope

    def test_upload_prefers_appdata_path(self):
        """Application context: files must land in the END USER's bucket via
        the appdata path so the user can download the attachment."""
        from dial_client import DialFileClient
        _StubDial.appdata = "USERBUCKET/appdata/dial-pptx-mcp"
        try:
            url = DialFileClient().upload(b"deck", "deck.pptx")
        finally:
            _StubDial.appdata = None
        self.assertTrue(url.startswith(
            "files/USERBUCKET/appdata/dial-pptx-mcp/pptx-mcp/"))
        self.assertTrue(url.endswith("/deck.pptx"))
        self.assertIn("/USERBUCKET/appdata/", _StubDial.last_upload["path"])

    def test_unique_upload_paths(self):
        from dial_client import DialFileClient
        client = DialFileClient()
        self.assertNotEqual(client.upload(b"a", "x.pptx"), client.upload(b"a", "x.pptx"))

    def test_missing_config(self):
        from dial_client import DialFileClient, DialConfigError
        saved = os.environ.pop("DIAL_CORE_URL")
        try:
            with self.assertRaises(DialConfigError):
                DialFileClient()
        finally:
            os.environ["DIAL_CORE_URL"] = saved

    def test_missing_credentials(self):
        from dial_client import DialFileClient, DialConfigError
        saved = os.environ.pop("DIAL_API_KEY")
        try:
            with self.assertRaises(DialConfigError):
                DialFileClient().get_bucket()
        finally:
            os.environ["DIAL_API_KEY"] = saved

    def test_download_refuses_foreign_host(self):
        from dial_client import DialFileClient, DialConfigError
        with self.assertRaises(DialConfigError) as ctx:
            DialFileClient().download("https://evil.example.com/files/x/y.pptx")
        self.assertIn("evil.example.com", str(ctx.exception))

    def test_download_refuses_a_non_file_reference(self):
        from dial_client import DialFileClient, DialConfigError
        with self.assertRaises(DialConfigError) as ctx:
            DialFileClient().download("just-a-name.png")
        self.assertIn("not a DIAL file reference", str(ctx.exception))

    def test_file_reference_forms_resolve_to_the_same_object(self):
        """The upload's relative URL, the Core API path and the chat
        frontend's api/files proxy path all name one object."""
        from dial_client import DialFileClient
        client = DialFileClient()
        expected = f"{client._base_url}/v1/files/BUCKET/img/x.png"
        for ref in ("files/BUCKET/img/x.png",
                    "/files/BUCKET/img/x.png",
                    "v1/files/BUCKET/img/x.png",
                    "api/files/BUCKET/img/x.png",
                    f"{client._base_url}/v1/files/BUCKET/img/x.png",
                    f"{client._base_url}/api/files/BUCKET/img/x.png"):
            self.assertEqual(client._file_request_url(ref), expected, ref)

    def test_public_alias_host_is_accepted_and_fetched_from_core(self):
        """DIAL reached in-cluster but linked publicly: the link's host is
        allowed by DIAL_PUBLIC_URL, the request still goes to Core."""
        from dial_client import DialFileClient, DialConfigError
        client = DialFileClient()
        link = "https://chat.example.com/api/files/BUCKET/img/x.png"
        with self.assertRaises(DialConfigError):
            client._file_request_url(link)
        os.environ["DIAL_PUBLIC_URL"] = "https://chat.example.com"
        try:
            self.assertEqual(client._file_request_url(link),
                             f"{client._base_url}/v1/files/BUCKET/img/x.png")
        finally:
            os.environ.pop("DIAL_PUBLIC_URL")

    def test_download_returns_bytes_from_the_callers_own_bucket(self):
        from dial_client import DialFileClient
        self.assertEqual(
            DialFileClient().download("files/STUBBUCKET/img/x.png"),
            b"FILEBYTES")

    def test_forbidden_download_explains_whose_storage_was_read(self):
        """A 403 is an identity problem; the message has to say which
        identity, whose bucket, and what to do — an agent cannot act on
        'Client error 403'."""
        from dial_client import DialFileClient, DialConfigError
        with self.assertRaises(DialConfigError) as ctx:
            DialFileClient().download(
                "files/OTHERBUCKET/appdata/gpt-image-1.5-square/img/x.png")
        message = str(ctx.exception)
        self.assertIn("403", message)
        self.assertIn("STUBBUCK", message)      # the identity's own bucket
        self.assertIn("OTHERBUC", message)      # the file's bucket
        self.assertIn("gpt-image-1.5-square", message)  # the appdata owner
        self.assertIn("DIAL_API_KEY", message)

    def test_caller_mode_refuses_env_fallback(self):
        from dial_client import DialFileClient, DialConfigError
        os.environ["DIAL_AUTH_MODE"] = "caller"
        try:
            with self.assertRaises(DialConfigError) as ctx:
                DialFileClient().get_bucket()
            self.assertIn("No caller credentials", str(ctx.exception))
        finally:
            os.environ["DIAL_AUTH_MODE"] = "auto"

    def test_caller_mode_uses_incoming_headers(self):
        from unittest.mock import patch
        import dial_client
        from dial_client import DialFileClient
        os.environ["DIAL_AUTH_MODE"] = "caller"
        try:
            with patch.object(dial_client, "_incoming_request_headers",
                              return_value={"api-key": "test-key"}):
                self.assertEqual(DialFileClient().get_bucket(), "STUBBUCKET")
        finally:
            os.environ["DIAL_AUTH_MODE"] = "auto"


class TestTemplateContentDecoding(unittest.TestCase):
    """The decoding contract of create_presentation_from_template_content:
    both plain base64 and data: URI forms must yield identical bytes."""

    def test_data_uri_and_plain_base64_equivalent(self):
        raw = b"PK\x03\x04fakepptx"
        plain = base64.b64encode(raw).decode()
        data_uri = "data:application/vnd.openxmlformats-officedocument" \
                   ".presentationml.presentation;base64," + plain
        for form in (plain, data_uri):
            payload = form.strip()
            if payload.startswith("data:"):
                _, _, payload = payload.partition(",")
            self.assertEqual(base64.b64decode(payload, validate=True), raw)


if __name__ == "__main__":
    unittest.main()
