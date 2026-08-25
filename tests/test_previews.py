"""
Tests for the contact-sheet composer and the design-guidance document.

Rendering needs LibreOffice, so the tests that would drive it self-skip the way
the rest of the suite does; the composer and the guidance parser are pure and
are tested directly.
"""
import io
import unittest

from PIL import Image

import previews
from tools.guidance_tools import GUIDANCE_PATH, _load, _slug


def tile(width=320, height=180, colour=(120, 140, 160)):
    buffer = io.BytesIO()
    Image.new("RGB", (width, height), colour).save(buffer, "PNG")
    return buffer.getvalue()


def open_sheet(data):
    return Image.open(io.BytesIO(data))


class TestContactSheet(unittest.TestCase):
    def test_a_single_slide_makes_a_one_cell_sheet(self):
        sheet = open_sheet(previews.compose_contact_sheet([tile()], [1]))
        self.assertGreater(sheet.width, 320)
        self.assertGreater(sheet.height, 180)

    def test_the_grid_wraps_at_the_column_count(self):
        tiles = [tile() for _ in range(4)]
        two_across = open_sheet(
            previews.compose_contact_sheet(tiles, [1, 2, 3, 4], columns=2))
        four_across = open_sheet(
            previews.compose_contact_sheet(tiles, [1, 2, 3, 4], columns=4))
        self.assertGreater(four_across.width, two_across.width)
        self.assertGreater(two_across.height, four_across.height)

    def test_columns_never_exceed_the_number_of_slides(self):
        sheet = open_sheet(
            previews.compose_contact_sheet([tile()], [1], columns=6))
        # One tile plus padding, not six cells' worth of empty grid.
        self.assertLess(sheet.width, 320 * 2)

    def test_slides_of_different_sizes_are_not_stretched(self):
        tiles = [tile(320, 180), tile(160, 90)]
        sheet = previews.compose_contact_sheet(tiles, [1, 2], columns=2)
        self.assertTrue(open_sheet(sheet).width > 320)

    def test_an_empty_selection_is_rejected(self):
        with self.assertRaises(ValueError):
            previews.compose_contact_sheet([], [])

    def test_the_sheet_is_jpeg(self):
        sheet = previews.compose_contact_sheet([tile()], [1])
        self.assertEqual(open_sheet(sheet).format, "JPEG")

    def test_the_describe_prompt_maps_images_to_absolute_slide_numbers(self):
        prompt = previews.DESCRIBE_PROMPT.format(
            mapping="image 1 = slide 7, image 2 = slide 9")
        self.assertIn("image 1 = slide 7", prompt)
        self.assertIn("best_for", prompt)


class TestDesignGuidance(unittest.TestCase):
    def test_the_document_ships_with_the_server(self):
        self.assertTrue(GUIDANCE_PATH.exists(), GUIDANCE_PATH)

    def test_it_parses_into_sections(self):
        text, sections = _load()
        self.assertIsNotNone(text)
        self.assertGreaterEqual(len(sections), 8)
        for slug, meta in sections.items():
            with self.subTest(section=slug):
                self.assertTrue(meta["body"].startswith("## "))
                self.assertGreater(len(meta["body"]), 100)

    def test_sections_are_numbered_in_document_order(self):
        _, sections = _load()
        numbers = [meta["number"] for meta in
                   sorted(sections.values(), key=lambda m: m["number"])]
        self.assertEqual(numbers, list(range(1, len(numbers) + 1)))

    def test_it_covers_the_topics_the_tool_promises(self):
        text, _ = _load()
        for topic in ("duplicate_slide", "visual_inspect_slides",
                      "validate_presentation", "manage_speaker_notes",
                      "Liberation Sans", "Aptos", "Margins"):
            with self.subTest(topic=topic):
                self.assertIn(topic, text)

    def test_slugs_are_stable(self):
        self.assertEqual(_slug("Charts and tables"), "charts-and-tables")
        self.assertEqual(_slug("Type"), "type")


if __name__ == "__main__":
    unittest.main()
