"""Tests for the visual-inspection capability (rendering + vision LLM client)."""
import base64
import os
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from visual_qa import VisionLLM, VisionLLMConfigError, review_prompt

ENV = {
    "VISION_LLM_ENDPOINT": "https://example.invalid/openai/responses?api-version=x",
    "VISION_LLM_API_KEY": "k",
    "VISION_LLM_MODEL": "vision-model",
}


class WithEnv(unittest.TestCase):
    def setUp(self):
        self._saved = {k: os.environ.get(k) for k in ENV}
        os.environ.update(ENV)

    def tearDown(self):
        for k, v in self._saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v


class TestConfig(unittest.TestCase):
    def test_unconfigured_raises(self):
        saved = {k: os.environ.pop(k, None) for k in ENV}
        try:
            with self.assertRaises(VisionLLMConfigError):
                VisionLLM()
        finally:
            for k, v in saved.items():
                if v is not None:
                    os.environ[k] = v


class TestPayload(WithEnv):
    def test_payload_shape(self):
        llm = VisionLLM()
        png = b"\x89PNG-fake"
        payload = llm.build_payload([png, png], "check this")
        self.assertEqual(payload["model"], "vision-model")
        content = payload["input"][0]["content"]
        self.assertEqual(content[0], {"type": "input_text", "text": "check this"})
        self.assertEqual(len(content), 3)
        expected = "data:image/png;base64," + base64.b64encode(png).decode()
        self.assertEqual(content[1]["image_url"], expected)
        self.assertEqual(content[1]["type"], "input_image")


class TestResponseParsing(WithEnv):
    def test_extract_output_text_field(self):
        self.assertEqual(VisionLLM.extract_text({"output_text": "hi"}), "hi")

    def test_extract_from_output_list(self):
        resp = {"output": [
            {"type": "reasoning", "content": []},
            {"type": "message", "content": [
                {"type": "output_text", "text": '{"passed": true, "issues": []}'}]},
        ]}
        self.assertIn("passed", VisionLLM.extract_text(resp))

    def test_parse_verdict_plain(self):
        v = VisionLLM.parse_verdict('{"passed": false, "issues": [{"slide": 2}]}')
        self.assertFalse(v["passed"])
        self.assertEqual(v["issues"][0]["slide"], 2)

    def test_parse_verdict_fenced(self):
        v = VisionLLM.parse_verdict('```json\n{"passed": true, "issues": []}\n```')
        self.assertTrue(v["passed"])

    def test_parse_verdict_garbage_keeps_raw(self):
        v = VisionLLM.parse_verdict("The deck looks fine to me!")
        self.assertIsNone(v["passed"])
        self.assertIn("raw_review", v)


class TestPrompt(unittest.TestCase):
    def test_reference_and_focus_wording(self):
        p = review_prompt(True, focus="check logo size")
        self.assertIn("reference template", p)
        self.assertIn("check logo size", p)
        self.assertIn('"passed"', p)
        p2 = review_prompt(False)
        self.assertNotIn("FIRST images", p2)


class TestRendering(unittest.TestCase):
    def test_render_demo_deck(self):
        import shutil
        if not (os.environ.get("SOFFICE_PATH") or shutil.which("soffice")):
            self.skipTest("LibreOffice not installed")
        from visual_qa import render_pptx_bytes_to_pngs
        data = (Path(__file__).resolve().parent.parent /
                "mcp_all_tools_templates_effects_demo.pptx").read_bytes()
        images = render_pptx_bytes_to_pngs(data, max_slides=2)
        self.assertEqual(len(images), 2)
        for png in images:
            self.assertTrue(png.startswith(b"\x89PNG"))


if __name__ == "__main__":
    unittest.main()
