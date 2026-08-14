"""
Visual inspection of generated presentations (render + vision-LLM review).

Pipeline (inspired by the render/inspect loop of document-generation agents,
implemented independently):
1. Render the .pptx to one PNG per slide: LibreOffice headless converts the
   deck to PDF, PyMuPDF rasterizes the pages. LibreOffice must be installed
   (`soffice` on PATH, or SOFFICE_PATH env var); the Dockerfile includes it.
2. Send the slide images to a configurable external vision LLM with a
   fidelity/error checklist, and parse a structured verdict.

The LLM endpoint speaks the OpenAI Responses API with image input
(Azure OpenAI included). Configuration via environment (see .env.example):
- VISION_LLM_ENDPOINT   full URL, e.g.
  https://<resource>.openai.azure.com/openai/responses?api-version=2025-04-01-preview
- VISION_LLM_API_KEY    sent as both api-key (Azure) and Authorization: Bearer
- VISION_LLM_MODEL      model / Azure deployment name (must accept images)
- VISION_LLM_MAX_SLIDES cap on slides sent per inspection (default 15)
- SOFFICE_PATH          LibreOffice binary if not "soffice" on PATH
"""
import base64
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

import httpx


class VisualQAError(RuntimeError):
    pass


# ---- Rendering ----

def _soffice_binary():
    path = os.environ.get("SOFFICE_PATH") or shutil.which("soffice")
    if not path or not os.path.exists(path):
        raise VisualQAError(
            "LibreOffice is required for slide rendering but was not found. "
            "Install it (the server Docker image includes it) or set "
            "SOFFICE_PATH to the soffice binary."
        )
    return path


def render_pptx_bytes_to_pngs(pptx_data: bytes, dpi: int = 96,
                              max_slides: int = None) -> list:
    """Render presentation bytes to a list of PNG bytes, one per slide."""
    import pymupdf

    with tempfile.TemporaryDirectory(prefix="pptx-visual-qa-") as tmp:
        tmp = Path(tmp)
        src = tmp / "deck.pptx"
        src.write_bytes(pptx_data)
        # Isolated LibreOffice profile so concurrent renders don't fight
        # over the shared user profile lock.
        profile = tmp / "lo-profile"
        cmd = [
            _soffice_binary(), "--headless", "--norestore",
            f"-env:UserInstallation=file://{profile}",
            "--convert-to", "pdf", "--outdir", str(tmp), str(src),
        ]
        proc = subprocess.run(cmd, capture_output=True, timeout=180)
        pdf = tmp / "deck.pdf"
        if proc.returncode != 0 or not pdf.exists():
            raise VisualQAError(
                "LibreOffice failed to render the presentation: "
                + proc.stderr.decode(errors="replace")[-500:]
            )
        images = []
        with pymupdf.open(pdf) as doc:
            pages = doc.page_count if max_slides is None else min(doc.page_count, max_slides)
            for i in range(pages):
                pix = doc[i].get_pixmap(dpi=dpi)
                images.append(pix.tobytes("png"))
        return images


# ---- Vision LLM client (OpenAI Responses API shape, Azure-compatible) ----

REVIEW_PROMPT = """You are a meticulous presentation QA reviewer. You are shown \
rendered slide images of a PowerPoint deck generated from a corporate template{ref_note}.

Check every slide for:
1. Template/brand fidelity: consistent colors, fonts, logo placement, and layout \
usage matching the deck's own master style{ref_clause}.
2. Visible errors: text overflowing or clipped by its container, overlapping \
elements, elements off the slide edge, placeholder text left unfilled (e.g. \
"Click to add title"), broken or empty charts/tables/images, illegible text \
(too small or poor contrast), inconsistent alignment or spacing.

Respond with ONLY a JSON object, no markdown fence:
{{"passed": true|false, "issues": [{{"slide": <1-based number>, "severity": \
"critical"|"major"|"minor", "description": "...", "suggested_fix": "..."}}]}}
"passed" is true only when there are no critical or major issues."""


class VisionLLMConfigError(VisualQAError):
    pass


class VisionLLM:
    def __init__(self):
        self.endpoint = os.environ.get("VISION_LLM_ENDPOINT")
        self.api_key = os.environ.get("VISION_LLM_API_KEY")
        self.model = os.environ.get("VISION_LLM_MODEL")
        if not (self.endpoint and self.api_key and self.model):
            raise VisionLLMConfigError(
                "Visual inspection is not configured on this server: set "
                "VISION_LLM_ENDPOINT, VISION_LLM_API_KEY and VISION_LLM_MODEL "
                "(see .env.example)."
            )

    def build_payload(self, images: list, prompt: str) -> dict:
        content = [{"type": "input_text", "text": prompt}]
        for png in images:
            b64 = base64.b64encode(png).decode()
            content.append({
                "type": "input_image",
                "image_url": f"data:image/png;base64,{b64}",
            })
        return {"model": self.model,
                "input": [{"role": "user", "content": content}]}

    @staticmethod
    def extract_text(response_json: dict) -> str:
        """Pull the assistant text out of a Responses API result."""
        if isinstance(response_json.get("output_text"), str):
            return response_json["output_text"]
        parts = []
        for item in response_json.get("output", []):
            if item.get("type") == "message":
                for c in item.get("content", []):
                    if c.get("type") == "output_text":
                        parts.append(c.get("text", ""))
        if not parts:
            raise VisualQAError("Vision LLM returned no text output.")
        return "\n".join(parts)

    @staticmethod
    def parse_verdict(text: str) -> dict:
        """Parse the reviewer's JSON verdict, tolerating a stray code fence."""
        candidate = text.strip()
        m = re.search(r"\{.*\}", candidate, re.S)
        if m:
            candidate = m.group(0)
        try:
            verdict = json.loads(candidate)
            if isinstance(verdict, dict) and "passed" in verdict:
                verdict.setdefault("issues", [])
                return verdict
        except json.JSONDecodeError:
            pass
        return {"passed": None, "issues": [],
                "raw_review": text,
                "note": "Reviewer response was not valid JSON; see raw_review."}

    def ask(self, images: list, prompt: str, timeout: float = 300.0) -> str:
        """Send prompt + images, return the model's raw text answer."""
        headers = {
            "api-key": self.api_key,                       # Azure OpenAI
            "Authorization": f"Bearer {self.api_key}",     # OpenAI-compatible
            "Content-Type": "application/json",
        }
        r = httpx.post(self.endpoint, headers=headers,
                       json=self.build_payload(images, prompt), timeout=timeout)
        if r.status_code != 200:
            raise VisualQAError(
                f"Vision LLM request failed with HTTP {r.status_code}: "
                + r.text[:300]
            )
        return self.extract_text(r.json())

    def review(self, images: list, prompt: str, timeout: float = 300.0) -> dict:
        return self.parse_verdict(self.ask(images, prompt, timeout))

    def ask_json(self, images: list, prompt: str, timeout: float = 300.0) -> dict:
        """Like ask(), parsed as a JSON object ({} when unparseable)."""
        text = self.ask(images, prompt, timeout)
        candidate = text.strip()
        m = re.search(r"\{.*\}", candidate, re.S)
        if m:
            candidate = m.group(0)
        try:
            data = json.loads(candidate)
            return data if isinstance(data, dict) else {}
        except json.JSONDecodeError:
            return {}


def enforcement_enabled() -> bool:
    """Automatic visual QA is on when the vision LLM is configured, unless
    explicitly disabled with VISUAL_QA_ENFORCE=false."""
    configured = all(os.environ.get(k) for k in
                     ("VISION_LLM_ENDPOINT", "VISION_LLM_API_KEY",
                      "VISION_LLM_MODEL"))
    return configured and os.environ.get(
        "VISUAL_QA_ENFORCE", "true").lower() != "false"


def fail_open_on_error() -> bool:
    """VISUAL_QA_ON_ERROR=allow lets exports through when inspection itself
    fails (renderer missing, endpoint down). Default is to block."""
    return os.environ.get("VISUAL_QA_ON_ERROR", "block").lower() == "allow"


def inspect_presentation(pres, reference_pres=None, focus: str = None) -> dict:
    """Render a python-pptx Presentation (and optional reference) and return
    the vision reviewer's verdict. Raises VisualQAError on infrastructure
    failure (renderer/LLM)."""
    import io

    llm = VisionLLM()
    max_slides = int(os.environ.get("VISION_LLM_MAX_SLIDES", "15"))

    def render(p):
        buf = io.BytesIO()
        p.save(buf)
        return render_pptx_bytes_to_pngs(buf.getvalue(), max_slides=max_slides)

    deck_images = render(pres)
    ref_images = render(reference_pres) if reference_pres is not None else []

    prompt = review_prompt(bool(ref_images), focus)
    if ref_images:
        prompt += (
            f"\nImage order: images 1-{len(ref_images)} are the reference "
            f"template; images {len(ref_images) + 1}-"
            f"{len(ref_images) + len(deck_images)} are the deck under review. "
            "Report issue slide numbers relative to the deck under review."
        )
    verdict = llm.review(ref_images + deck_images, prompt)
    verdict["slides_reviewed"] = len(deck_images)
    return verdict


def _render_deck(pres, max_slides):
    import io
    buf = io.BytesIO()
    pres.save(buf)
    return render_pptx_bytes_to_pngs(buf.getvalue(), max_slides=max_slides)


def inspect_and_repair(pres) -> dict:
    """Internal QA loop: inspect the deck; on failure, repair it in place via
    the LLM-planned whitelisted operations (visual_fix.py) and inspect again,
    up to VISUAL_QA_MAX_ITERATIONS (default 3) inspections.

    Returns {"passed": bool, "iterations": n, "repair_rounds": [...],
    "issues": [...]} — "issues" holds what remains when passed is False.
    Raises VisualQAError on infrastructure failure (renderer/LLM).
    """
    import visual_fix

    llm = VisionLLM()
    max_slides = int(os.environ.get("VISION_LLM_MAX_SLIDES", "15"))
    max_iterations = max(1, int(os.environ.get("VISUAL_QA_MAX_ITERATIONS", "3")))

    repair_rounds = []
    verdict = {}
    for iteration in range(1, max_iterations + 1):
        deck_images = _render_deck(pres, max_slides)
        verdict = llm.review(deck_images, review_prompt(False))
        verdict["slides_reviewed"] = len(deck_images)
        if verdict.get("passed") is True:
            return {"passed": True, "iterations": iteration,
                    "repair_rounds": repair_rounds}
        issues = verdict.get("issues", [])
        if iteration == max_iterations or not issues:
            # Out of budget, or nothing actionable (e.g. unparseable review)
            break
        plan = visual_fix.plan_repairs(llm, issues, pres, deck_images)
        result = visual_fix.apply_repairs(pres, plan)
        repair_rounds.append({
            "iteration": iteration,
            "issues_found": len(issues),
            "operations_applied": len(result["applied"]),
            "operations_skipped": len(result["skipped"]),
        })
        if not result["applied"]:
            break  # no progress is possible; stop burning inspections

    out = {"passed": False,
           "iterations": len(repair_rounds) + 1,
           "repair_rounds": repair_rounds,
           "issues": verdict.get("issues", [])}
    for key in ("raw_review", "note"):
        if key in verdict:
            out[key] = verdict[key]
    return out


def review_prompt(has_reference: bool, focus: str = None) -> str:
    prompt = REVIEW_PROMPT.format(
        ref_note=(". The FIRST images are the reference template's slides; "
                  "the deck under review follows" if has_reference else ""),
        ref_clause=(" and matching the reference template images"
                    if has_reference else ""),
    )
    if focus:
        prompt += f"\nAdditional focus requested by the caller: {focus}"
    return prompt
