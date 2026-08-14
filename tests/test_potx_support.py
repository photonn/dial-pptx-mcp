"""Tests for .potx template support via content-type coercion."""
import io
import sys
import unittest
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

from utils.presentation_utils import (
    _PRESENTATION_MAIN_CT,
    _TEMPLATE_MAIN_CT,
    coerce_template_bytes_to_presentation,
    open_presentation_bytes,
)

DEMO = REPO / "mcp_all_tools_templates_effects_demo.pptx"


def make_potx_bytes():
    """Synthesize a .potx from the bundled demo deck by re-declaring its
    main part with the template content type — the only difference between
    the two formats."""
    src = DEMO.read_bytes()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(src)) as zin, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    _PRESENTATION_MAIN_CT.encode(), _TEMPLATE_MAIN_CT.encode())
            zout.writestr(item, data)
    return out.getvalue()


class TestPotxSupport(unittest.TestCase):
    def test_pptx_passes_through_unchanged(self):
        data = DEMO.read_bytes()
        self.assertIs(coerce_template_bytes_to_presentation(data), data)

    def test_potx_is_coerced_and_opens(self):
        potx = make_potx_bytes()
        # Sanity: raw potx must be rejected by python-pptx
        from pptx import Presentation
        with self.assertRaises(ValueError):
            Presentation(io.BytesIO(potx))
        pres = open_presentation_bytes(potx)
        self.assertGreater(len(pres.slides), 0)
        self.assertGreater(len(pres.slide_layouts), 0)

    def test_coercion_touches_only_content_types(self):
        potx = make_potx_bytes()
        coerced = coerce_template_bytes_to_presentation(potx)
        zt = zipfile.ZipFile(io.BytesIO(potx))
        zc = zipfile.ZipFile(io.BytesIO(coerced))
        self.assertEqual(zt.namelist(), zc.namelist())
        for name in zt.namelist():
            if name == "[Content_Types].xml":
                self.assertNotIn(_TEMPLATE_MAIN_CT.encode(), zc.read(name))
            else:
                self.assertEqual(zt.read(name), zc.read(name), name)

    def test_non_zip_raises(self):
        with self.assertRaises(zipfile.BadZipFile):
            coerce_template_bytes_to_presentation(b"not a zip")


if __name__ == "__main__":
    unittest.main()
