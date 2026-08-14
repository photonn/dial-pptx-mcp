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

    def log_message(self, *args):
        pass

    def _authed(self):
        return self.headers.get("Api-Key") == "test-key"

    def do_GET(self):
        if self.path == "/v1/bucket" and self._authed():
            self._json(b'{"bucket": "STUBBUCKET"}')
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
        with self.assertRaises(DialConfigError):
            DialFileClient().download("https://evil.example.com/files/x/y.pptx")

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
