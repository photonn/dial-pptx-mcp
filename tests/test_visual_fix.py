"""Tests for the internal repair engine and the inspect-and-repair loop."""
import os
import sys
import unittest
from unittest.mock import patch
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx import Presentation
from pptx.util import Inches, Pt

import visual_fix
import visual_qa

ENV = {
    "VISION_LLM_ENDPOINT": "https://example.invalid/responses",
    "VISION_LLM_API_KEY": "k",
    "VISION_LLM_MODEL": "m",
}


def make_deck():
    pres = Presentation()
    slide = pres.slides.add_slide(pres.slide_layouts[6])  # blank
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    box.text_frame.text = "Hello"
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(40)
    return pres


class TestDescribeSlides(unittest.TestCase):
    def test_describe(self):
        desc = visual_fix.describe_slides(make_deck())
        self.assertEqual(desc[0]["slide"], 1)
        shape = desc[0]["shapes"][0]
        self.assertEqual(shape["shape_index"], 0)
        self.assertEqual(shape["left_in"], 1.0)
        self.assertEqual(shape["text"], "Hello")
        self.assertEqual(shape["font_sizes_pt"], [40.0])

    def test_slide_filter(self):
        self.assertEqual(visual_fix.describe_slides(make_deck(), [2]), [])


class TestApplyRepairs(unittest.TestCase):
    def test_move_resize_font_text(self):
        pres = make_deck()
        result = visual_fix.apply_repairs(pres, [
            {"op": "move_shape", "slide": 1, "shape_index": 0,
             "left_in": 2.0, "top_in": 3.0},
            {"op": "resize_shape", "slide": 1, "shape_index": 0,
             "width_in": 5.0},
            {"op": "set_font_size", "slide": 1, "shape_index": 0,
             "size_pt": 18},
            {"op": "set_text", "slide": 1, "shape_index": 0,
             "text": "Fixed"},
        ])
        self.assertEqual(len(result["applied"]), 4)
        self.assertEqual(result["skipped"], [])
        shape = pres.slides[0].shapes[0]
        self.assertEqual(shape.left, Inches(2.0))
        self.assertEqual(shape.width, Inches(5.0))
        self.assertEqual(shape.text_frame.text, "Fixed")

    def test_delete_shape(self):
        pres = make_deck()
        result = visual_fix.apply_repairs(pres, [
            {"op": "delete_shape", "slide": 1, "shape_index": 0}])
        self.assertEqual(len(result["applied"]), 1)
        self.assertEqual(len(pres.slides[0].shapes), 0)

    def test_invalid_operations_are_skipped(self):
        pres = make_deck()
        result = visual_fix.apply_repairs(pres, [
            {"op": "move_shape", "slide": 9, "shape_index": 0,
             "left_in": 1, "top_in": 1},               # bad slide
            {"op": "move_shape", "slide": 1, "shape_index": 5,
             "left_in": 1, "top_in": 1},               # bad shape
            {"op": "set_font_size", "slide": 1, "shape_index": 0,
             "size_pt": 500},                          # out of range
            {"op": "run_shell", "slide": 1, "shape_index": 0},  # unknown op
            "garbage",                                  # not an object
        ])
        self.assertEqual(result["applied"], [])
        self.assertEqual(len(result["skipped"]), 5)
        # deck untouched
        self.assertEqual(pres.slides[0].shapes[0].left, Inches(1.0))


class TestInspectAndRepairLoop(unittest.TestCase):
    def setUp(self):
        self._saved = {k: os.environ.get(k) for k in ENV}
        os.environ.update(ENV)

    def tearDown(self):
        for k, v in self._saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v

    def _run(self, verdicts, plans):
        """Drive the loop with scripted inspection verdicts and repair plans;
        rendering and HTTP are stubbed out."""
        calls = {"review": 0, "plan": 0}

        def fake_review(self_llm, images, prompt, timeout=None):
            v = verdicts[min(calls["review"], len(verdicts) - 1)]
            calls["review"] += 1
            return dict(v)

        def fake_plan(llm, issues, pres, images):
            p = plans[min(calls["plan"], len(plans) - 1)]
            calls["plan"] += 1
            return p

        with patch.object(visual_qa, "_render_deck",
                          return_value=[b"\x89PNG-fake"]), \
             patch.object(visual_qa.VisionLLM, "review", fake_review), \
             patch.object(visual_fix, "plan_repairs", fake_plan):
            outcome = visual_qa.inspect_and_repair(make_deck())
        return outcome, calls

    def test_pass_first_time(self):
        outcome, calls = self._run([{"passed": True, "issues": []}], [[]])
        self.assertTrue(outcome["passed"])
        self.assertEqual(outcome["iterations"], 1)
        self.assertEqual(calls["plan"], 0)

    def test_repair_then_pass(self):
        verdicts = [
            {"passed": False, "issues": [{"slide": 1, "description": "x"}]},
            {"passed": True, "issues": []},
        ]
        plan = [{"op": "set_font_size", "slide": 1, "shape_index": 0,
                 "size_pt": 20}]
        outcome, calls = self._run(verdicts, [plan])
        self.assertTrue(outcome["passed"])
        self.assertEqual(outcome["iterations"], 2)
        self.assertEqual(outcome["repair_rounds"][0]["operations_applied"], 1)

    def test_gives_up_after_max_iterations(self):
        failing = {"passed": False,
                   "issues": [{"slide": 1, "description": "x"}]}
        plan = [{"op": "set_font_size", "slide": 1, "shape_index": 0,
                 "size_pt": 20}]
        os.environ["VISUAL_QA_MAX_ITERATIONS"] = "3"
        try:
            outcome, calls = self._run([failing], [plan])
        finally:
            os.environ.pop("VISUAL_QA_MAX_ITERATIONS")
        self.assertFalse(outcome["passed"])
        self.assertEqual(calls["review"], 3)
        self.assertEqual(len(outcome["repair_rounds"]), 2)
        self.assertTrue(outcome["issues"])

    def test_stops_early_when_no_repairs_apply(self):
        failing = {"passed": False,
                   "issues": [{"slide": 1, "description": "x"}]}
        outcome, calls = self._run([failing], [[]])  # planner returns nothing
        self.assertFalse(outcome["passed"])
        self.assertEqual(calls["review"], 1)  # no progress -> stop immediately


if __name__ == "__main__":
    unittest.main()
