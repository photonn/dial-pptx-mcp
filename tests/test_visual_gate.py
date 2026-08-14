"""Tests for the automatic visual-QA export gate."""
import os
import sys
import unittest
from unittest.mock import patch
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from state import PresentationStore
from tools.presentation_tools import _visual_qa_gate
import visual_qa

VISION_ENV = {
    "VISION_LLM_ENDPOINT": "https://example.invalid/responses",
    "VISION_LLM_API_KEY": "k",
    "VISION_LLM_MODEL": "m",
}
GATE_VARS = list(VISION_ENV) + ["VISUAL_QA_ENFORCE", "VISUAL_QA_ON_ERROR"]


class GateTestCase(unittest.TestCase):
    def setUp(self):
        self._saved = {k: os.environ.get(k) for k in GATE_VARS}
        for k in GATE_VARS:
            os.environ.pop(k, None)
        os.environ.update(VISION_ENV)
        self.store = PresentationStore(ttl_seconds=60, max_items=10)
        self.pid = self.store.new_id()
        self.store[self.pid] = object()  # decks start dirty

    def tearDown(self):
        for k, v in self._saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v


class TestEnforcementSwitch(GateTestCase):
    def test_disabled_when_unconfigured(self):
        os.environ.pop("VISION_LLM_API_KEY")
        self.assertFalse(visual_qa.enforcement_enabled())
        self.assertIsNone(_visual_qa_gate(self.store, self.pid))

    def test_disabled_by_flag(self):
        os.environ["VISUAL_QA_ENFORCE"] = "false"
        self.assertIsNone(_visual_qa_gate(self.store, self.pid))

    def test_enabled_when_configured(self):
        self.assertTrue(visual_qa.enforcement_enabled())


class TestGateVerdicts(GateTestCase):
    def test_unresolved_deck_is_refused_and_stays_dirty(self):
        outcome = {"passed": False, "iterations": 3, "repair_rounds": [],
                   "issues": [{"slide": 1, "severity": "major",
                               "description": "overflow"}]}
        with patch.object(visual_qa, "inspect_and_repair", return_value=outcome):
            refusal = _visual_qa_gate(self.store, self.pid)
        self.assertIn("error", refusal)
        self.assertIn("not a request to retry", refusal["message"])
        self.assertEqual(refusal["unresolved_issues"][0]["slide"], 1)
        self.assertTrue(self.store.is_dirty(self.pid))

    def test_unresolved_policy_values(self):
        self.assertEqual(visual_qa.unresolved_policy(), "report")
        for value in ("export", "export_as_is", "EXPORT_AS_IS"):
            os.environ["VISUAL_QA_ON_UNRESOLVED"] = value
            self.assertEqual(visual_qa.unresolved_policy(), "export_as_is")
        for value in ("block", "report", "nonsense"):
            os.environ["VISUAL_QA_ON_UNRESOLVED"] = value
            self.assertEqual(visual_qa.unresolved_policy(), "report")
        os.environ.pop("VISUAL_QA_ON_UNRESOLVED")

    def test_unresolved_can_export_when_configured(self):
        os.environ["VISUAL_QA_ON_UNRESOLVED"] = "export_as_is"
        try:
            outcome = {"passed": False, "iterations": 3, "repair_rounds": [],
                       "issues": [{"slide": 1}]}
            with patch.object(visual_qa, "inspect_and_repair",
                              return_value=outcome):
                self.assertIsNone(_visual_qa_gate(self.store, self.pid))
        finally:
            os.environ.pop("VISUAL_QA_ON_UNRESOLVED")

    def test_passing_deck_clears_dirty_and_skips_reinspection(self):
        outcome = {"passed": True, "iterations": 1, "repair_rounds": []}
        with patch.object(visual_qa, "inspect_and_repair",
                          return_value=outcome) as mock:
            self.assertIsNone(_visual_qa_gate(self.store, self.pid))
            self.assertIsNone(_visual_qa_gate(self.store, self.pid))
        self.assertEqual(mock.call_count, 1)  # clean deck not re-inspected
        self.assertFalse(self.store.is_dirty(self.pid))

    def test_edit_after_pass_requires_reinspection(self):
        outcome = {"passed": True, "iterations": 1, "repair_rounds": []}
        with patch.object(visual_qa, "inspect_and_repair",
                          return_value=outcome) as mock:
            _visual_qa_gate(self.store, self.pid)
            self.store.mark_dirty(self.pid)  # what the tool wrapper does
            _visual_qa_gate(self.store, self.pid)
        self.assertEqual(mock.call_count, 2)


class TestGateInfraErrors(GateTestCase):
    def test_infra_error_blocks_by_default(self):
        with patch.object(visual_qa, "inspect_and_repair",
                          side_effect=visual_qa.VisualQAError("no soffice")):
            refusal = _visual_qa_gate(self.store, self.pid)
        self.assertIn("no soffice", refusal["error"])

    def test_infra_error_allows_when_configured(self):
        os.environ["VISUAL_QA_ON_ERROR"] = "allow"
        with patch.object(visual_qa, "inspect_and_repair",
                          side_effect=visual_qa.VisualQAError("down")):
            self.assertIsNone(_visual_qa_gate(self.store, self.pid))


class TestDirtyFlags(unittest.TestCase):
    def test_new_deck_starts_dirty(self):
        store = PresentationStore(ttl_seconds=60, max_items=5)
        pid = store.new_id()
        store[pid] = object()
        self.assertTrue(store.is_dirty(pid))
        store.clear_dirty(pid)
        self.assertFalse(store.is_dirty(pid))
        store.mark_dirty(pid)
        self.assertTrue(store.is_dirty(pid))

    def test_unknown_id_is_not_dirty(self):
        store = PresentationStore(ttl_seconds=60, max_items=5)
        self.assertFalse(store.is_dirty("nope"))


if __name__ == "__main__":
    unittest.main()
