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
    def test_failing_deck_is_refused_and_stays_dirty(self):
        verdict = {"passed": False, "issues": [
            {"slide": 1, "severity": "major", "description": "overflow"}]}
        with patch.object(visual_qa, "inspect_presentation", return_value=verdict):
            refusal = _visual_qa_gate(self.store, self.pid)
        self.assertIn("error", refusal)
        self.assertEqual(refusal["issues"][0]["slide"], 1)
        self.assertTrue(self.store.is_dirty(self.pid))

    def test_passing_deck_clears_dirty_and_skips_reinspection(self):
        verdict = {"passed": True, "issues": []}
        with patch.object(visual_qa, "inspect_presentation",
                          return_value=verdict) as mock:
            self.assertIsNone(_visual_qa_gate(self.store, self.pid))
            self.assertIsNone(_visual_qa_gate(self.store, self.pid))
        self.assertEqual(mock.call_count, 1)  # clean deck not re-inspected
        self.assertFalse(self.store.is_dirty(self.pid))

    def test_edit_after_pass_requires_reinspection(self):
        with patch.object(visual_qa, "inspect_presentation",
                          return_value={"passed": True, "issues": []}) as mock:
            _visual_qa_gate(self.store, self.pid)
            self.store.mark_dirty(self.pid)  # what the tool wrapper does
            _visual_qa_gate(self.store, self.pid)
        self.assertEqual(mock.call_count, 2)

    def test_unparseable_verdict_blocks(self):
        verdict = {"passed": None, "issues": [], "raw_review": "looks ok?"}
        with patch.object(visual_qa, "inspect_presentation", return_value=verdict):
            refusal = _visual_qa_gate(self.store, self.pid)
        self.assertIn("error", refusal)
        self.assertEqual(refusal["raw_review"], "looks ok?")


class TestGateInfraErrors(GateTestCase):
    def test_infra_error_blocks_by_default(self):
        with patch.object(visual_qa, "inspect_presentation",
                          side_effect=visual_qa.VisualQAError("no soffice")):
            refusal = _visual_qa_gate(self.store, self.pid)
        self.assertIn("no soffice", refusal["error"])

    def test_infra_error_allows_when_configured(self):
        os.environ["VISUAL_QA_ON_ERROR"] = "allow"
        with patch.object(visual_qa, "inspect_presentation",
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
