"""Tests for get_server_info: it must report what the server actually is,
not a hardcoded snapshot that drifts."""
import re
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import ppt_mcp_server as server

ROOT = Path(__file__).resolve().parent.parent


class TestServerInfo(unittest.TestCase):
    def setUp(self):
        self.info = server.get_server_info()

    def test_version_matches_pyproject(self):
        declared = re.search(r'^version\s*=\s*"([^"]+)"',
                             (ROOT / "pyproject.toml").read_text(),
                             re.MULTILINE).group(1)
        self.assertEqual(self.info["version"], declared)
        self.assertNotEqual(self.info["version"], "unknown")

    def test_tool_count_is_the_live_registry(self):
        self.assertEqual(self.info["tools"],
                         len(server.app._tool_manager._tools))
        self.assertGreater(self.info["tools"], 30)

    def test_reports_integration_state(self):
        self.assertIn(self.info["visual_qa"],
                      ("tools", "tools+export_gate", "disabled", "unavailable"))
        self.assertIn(self.info["dial_file_storage"], ("configured", "unset"))

    def test_is_annotated_read_only(self):
        tool = server.app._tool_manager._tools["get_server_info"]
        self.assertTrue(tool.annotations.readOnlyHint)


if __name__ == "__main__":
    unittest.main()
