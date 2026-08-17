"""Tests for the single-call batch deck builder and template-slide removal."""
import sys
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import utils as ppt_utils
from tools.batch_tools import _body_placeholder, _fill_slide

DEMO = REPO / "mcp_all_tools_templates_effects_demo.pptx"


class TestRemoveAllSlides(unittest.TestCase):
    def test_removes_slides_but_keeps_layouts_and_masters(self):
        pres = ppt_utils.open_presentation_bytes(DEMO.read_bytes())
        layouts_before = len(pres.slide_layouts)
        masters_before = len(pres.slide_masters)
        self.assertGreater(len(pres.slides), 0)

        removed = ppt_utils.remove_all_slides(pres)

        self.assertGreater(removed, 0)
        self.assertEqual(len(pres.slides), 0)
        self.assertEqual(len(pres.slide_layouts), layouts_before)
        self.assertEqual(len(pres.slide_masters), masters_before)

    def test_idempotent_on_empty_deck(self):
        pres = ppt_utils.create_presentation()
        self.assertEqual(ppt_utils.remove_all_slides(pres), 0)

    def test_slides_can_be_added_after_removal(self):
        pres = ppt_utils.open_presentation_bytes(DEMO.read_bytes())
        ppt_utils.remove_all_slides(pres)
        slide, _ = ppt_utils.add_slide(pres, 1)
        ppt_utils.set_title(slide, "Fresh")
        self.assertEqual(len(pres.slides), 1)


class TestFillSlide(unittest.TestCase):
    def _blank_deck_slide(self, layout_index):
        pres = ppt_utils.create_presentation()
        slide, _ = ppt_utils.add_slide(pres, layout_index)
        return pres, slide

    def test_title_and_bullets(self):
        pres, slide = self._blank_deck_slide(1)
        warnings = _fill_slide(pres, slide, {
            "title": "T", "bullets": ["a", "b"], "notes": "n"})
        self.assertEqual(warnings, [])
        texts = [sh.text_frame.text for sh in slide.shapes if sh.has_text_frame]
        self.assertIn("T", texts)
        self.assertTrue(any("a" in t and "b" in t for t in texts))
        self.assertEqual(slide.notes_slide.notes_text_frame.text, "n")

    def test_body_text_alternative(self):
        pres, slide = self._blank_deck_slide(1)
        _fill_slide(pres, slide, {"title": "T", "body_text": "plain body"})
        texts = [sh.text_frame.text for sh in slide.shapes if sh.has_text_frame]
        self.assertIn("plain body", texts)

    def test_blank_layout_falls_back_to_textbox(self):
        pres, slide = self._blank_deck_slide(6)  # blank layout: no placeholders
        self.assertIsNone(_body_placeholder(slide))
        warnings = _fill_slide(pres, slide, {"bullets": ["x", "y"]})
        self.assertTrue(any("textbox" in w for w in warnings))
        texts = [sh.text_frame.text for sh in slide.shapes if sh.has_text_frame]
        self.assertTrue(any("x" in t and "y" in t for t in texts))

    def test_missing_title_placeholder_is_reported_not_fatal(self):
        pres, slide = self._blank_deck_slide(6)
        warnings = _fill_slide(pres, slide, {"title": "T"})
        self.assertTrue(any("title" in w for w in warnings))


if __name__ == "__main__":
    unittest.main()
