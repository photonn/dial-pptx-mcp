"""Unit tests for single-line, LOG_LEVEL-controlled logging."""
import io
import logging
import os
import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import logging_utils


class TestResolveLevel(unittest.TestCase):
    def test_default_when_unset(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            self.assertEqual(logging_utils.resolve_level(), "INFO")

    def test_reads_env_case_insensitively(self):
        with mock.patch.dict(os.environ, {"LOG_LEVEL": "debug"}):
            self.assertEqual(logging_utils.resolve_level(), "DEBUG")

    def test_aliases_and_numeric(self):
        self.assertEqual(logging_utils.resolve_level("warn"), "WARNING")
        self.assertEqual(logging_utils.resolve_level("FATAL"), "CRITICAL")
        self.assertEqual(logging_utils.resolve_level("40"), "ERROR")

    def test_unknown_value_falls_back_rather_than_raising(self):
        # A typo in a log setting must never stop the server from starting.
        self.assertEqual(logging_utils.resolve_level("chatty"), "INFO")


class TestSingleLineOutput(unittest.TestCase):
    def _emit(self, level, action):
        """Configure logging onto a buffer, run action, return output lines."""
        buf = io.StringIO()
        logging_utils.configure_logging(level, stream=buf, force=True)
        try:
            action(logging_utils.get_logger("test"))
        finally:
            logging.getLogger().handlers = []
        return [ln for ln in buf.getvalue().splitlines() if ln]

    def test_multiline_message_becomes_one_line(self):
        lines = self._emit("INFO", lambda log: log.info("a\nb\nc"))
        self.assertEqual(len(lines), 1)
        self.assertIn("a | b | c", lines[0])

    def test_traceback_is_folded_into_the_same_line(self):
        def action(log):
            try:
                raise ValueError("boom")
            except ValueError:
                log.error("failed op=x", exc_info=True)

        lines = self._emit("ERROR", action)
        self.assertEqual(len(lines), 1)
        self.assertIn("Traceback", lines[0])
        self.assertIn("ValueError: boom", lines[0])

    def test_level_filters_records(self):
        lines = self._emit("WARNING", lambda log: (log.debug("d"), log.info("i"),
                                                   log.warning("w")))
        self.assertEqual(len(lines), 1)
        self.assertIn("WARNING", lines[0])

    def test_debug_level_lets_everything_through(self):
        lines = self._emit("DEBUG", lambda log: (log.debug("d"), log.info("i")))
        self.assertEqual(len(lines), 2)

    def test_line_carries_timestamp_level_and_logger_name(self):
        lines = self._emit("INFO", lambda log: log.info("event key=value"))
        self.assertRegex(
            lines[0],
            r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z INFO +dial_pptx\.test "
            r"event key=value$")


class TestFlatten(unittest.TestCase):
    def test_collapses_newlines_and_surrounding_whitespace(self):
        self.assertEqual(logging_utils.flatten("one\n  two\r\nthree"),
                         "one | two | three")

    def test_leaves_single_line_text_alone(self):
        self.assertEqual(logging_utils.flatten("already one line"),
                         "already one line")


class TestLibraryRouting(unittest.TestCase):
    def test_uvicorn_logging_config_is_mutated_in_place(self):
        # uvicorn.Config binds LOGGING_CONFIG as a default argument at import
        # time, so rebinding the module attribute would silently do nothing.
        uvicorn_config = mock.MagicMock()
        original = {"version": 1, "handlers": {"default": {}},
                    "loggers": {"uvicorn": {"propagate": False}}}
        config_module = mock.MagicMock(LOGGING_CONFIG=original)
        with mock.patch.dict(sys.modules, {"uvicorn": uvicorn_config,
                                           "uvicorn.config": config_module}):
            logging_utils._route_uvicorn_through_root()
        self.assertEqual(original, {"version": 1, "disable_existing_loggers": False})
        self.assertTrue(logging.getLogger("uvicorn").propagate)
        self.assertEqual(logging.getLogger("uvicorn.access").handlers, [])

    def test_http_client_logs_are_quiet_unless_debugging(self):
        logging_utils._quiet_third_party("INFO")
        self.assertEqual(logging.getLogger("httpx").level, logging.WARNING)
        logging_utils._quiet_third_party("DEBUG")
        self.assertEqual(logging.getLogger("httpx").level, logging.DEBUG)


class TestGetLogger(unittest.TestCase):
    def test_names_are_nested_under_the_server_namespace(self):
        self.assertEqual(logging_utils.get_logger("visual_qa").name,
                         "dial_pptx.visual_qa")
        self.assertEqual(logging_utils.get_logger("__main__").name, "dial_pptx")


if __name__ == "__main__":
    unittest.main()
