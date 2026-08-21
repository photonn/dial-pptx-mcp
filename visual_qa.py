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
import logging
import os
import re
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

import httpx

from logging_utils import get_logger, flatten

logger = get_logger("visual_qa")


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
                              max_slides: int = None,
                              slides: list = None) -> list:
    """Render presentation bytes to a list of PNG bytes.

    slides: 1-based slide numbers to rasterize (default: every slide).
    LibreOffice always converts the whole deck — there is no per-slide
    conversion — so a subset only skips rasterization and, more importantly,
    the vision call. max_slides still caps how many images come back.
    """
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
        started = time.monotonic()
        proc = subprocess.run(cmd, capture_output=True, timeout=180)
        pdf = tmp / "deck.pdf"
        if proc.returncode != 0 or not pdf.exists():
            logger.error("render_failed stage=libreoffice returncode=%d bytes=%d "
                         "stderr=%s", proc.returncode, len(pptx_data),
                         flatten(proc.stderr.decode(errors="replace")[-300:]))
            raise VisualQAError(
                "LibreOffice failed to render the presentation: "
                + proc.stderr.decode(errors="replace")[-500:]
            )
        images = []
        with pymupdf.open(pdf) as doc:
            page_count = doc.page_count
            if slides is None:
                wanted = list(range(1, page_count + 1))
            else:
                wanted = [n for n in slides if 1 <= n <= page_count]
            if max_slides is not None:
                wanted = wanted[:max_slides]
            for number in wanted:
                pix = doc[number - 1].get_pixmap(dpi=dpi)
                images.append(pix.tobytes("png"))
        if slides is None and page_count > len(images):
            logger.warning("render_truncated rendered=%d slides=%d cap=%s "
                           "hint=raise_VISION_LLM_MAX_SLIDES",
                           len(images), page_count, max_slides)
        logger.debug("render_ok slides=%d of=%d dpi=%d duration_ms=%d",
                     len(images), page_count, dpi,
                     int((time.monotonic() - started) * 1000))
        return images


def normalize_slides(pres, slides):
    """Validate a caller-supplied 1-based slide selection against the deck.

    Returns a sorted, de-duplicated list, or None for "the whole deck".
    Raises ValueError naming the out-of-range numbers, so the tool layer can
    hand the agent an actionable message.
    """
    if slides is None:
        return None
    if isinstance(slides, int):
        slides = [slides]
    total = len(pres.slides)
    numbers, bad = [], []
    for value in slides:
        if isinstance(value, bool) or not isinstance(value, int):
            bad.append(value)
        elif 1 <= value <= total:
            numbers.append(value)
        else:
            bad.append(value)
    if bad:
        raise ValueError(
            f"Invalid slide number(s) {bad}: this presentation has {total} "
            f"slide(s), numbered 1-{total}."
        )
    if not numbers:
        return None
    return sorted(set(numbers))


# ---- Vision LLM client (OpenAI Responses API shape, Azure-compatible) ----

REVIEW_PROMPT = """You are a meticulous presentation QA reviewer. You are shown \
rendered slide images of a PowerPoint deck generated from a corporate template{ref_note}.

Check every slide for:
1. Template/brand fidelity: consistent colors, fonts, logo placement, and layout \
usage matching the deck's own master style{ref_clause}.
2. Text placement and overlap, ANYWHERE text appears — not just in text boxes. \
Judge the rendered pixels, not what the text probably says:
   - Text overflowing, clipped, or spilling outside its container or the slide edge.
   - Text overlapping other text, or sitting on top of shapes, images or lines in \
a way that makes either hard to read.
   - Charts and graphs: axis tick labels colliding with each other or truncated \
("..." or cut-off words), data labels overlapping their bars/slices/points or each \
other, a legend covering the plot area or running off the chart, an axis title \
squeezed or rotated into illegibility, series labels detached from what they label.
   - Tables: cell text wrapping into an unreadable stack, clipped by the cell or \
row height, columns too narrow for their content, headers not aligned with their \
columns, a table extending past the slide.
   - Diagrams, SmartArt and grouped shapes: labels wider than the node or box that \
holds them, text escaping a connector or arrow, node labels overlapping neighbouring \
nodes or connectors.
   - Text that is too small to read at presentation size, or too low-contrast \
against what is behind it.
   - Text sized badly for the space it occupies: a heading or body block set so \
small that its box is mostly empty, or comparable elements on one slide set at \
visibly different sizes for no reason. Text should fill its container \
comfortably without crowding it or its neighbours — report both the starved and \
the overstuffed cases, but do not ask for larger text where growing it would \
eat the slide's white space.
3. Other visible errors: elements off the slide edge, placeholder text left \
unfilled (e.g. "Click to add title"), broken or empty charts/tables/images, \
inconsistent alignment or spacing between comparable elements.

Report each problem separately, naming the element it affects (e.g. "chart on the \
right: x-axis labels overlap"), and say in "suggested_fix" what change would resolve \
it (resize, reposition, shorten the text, smaller font, wider column, hide the \
legend).

Respond with ONLY a JSON object, no markdown fence:
{{"passed": true|false, "issues": [{{"slide": <1-based number>, "severity": \
"critical"|"major"|"minor", "description": "...", "suggested_fix": "..."}}]}}
Unreadable or overlapping text is at least a major issue. "passed" is true only \
when there are no critical or major issues."""


class VisionLLMConfigError(VisualQAError):
    pass


# Azure OpenAI (and DIAL Core's Azure upstream) require an api-version on
# every request. This preview version covers both the Responses API and
# chat completions with image input. Overridable with VISION_LLM_API_VERSION.
DEFAULT_API_VERSION = "2025-04-01-preview"


def _with_api_version(url: str) -> str:
    """Ensure the request URL carries an ?api-version= query parameter.

    Azure OpenAI rejects requests without it ("api-version is a required
    query parameter"), and DIAL Core passes the parameter through to its
    Azure upstream, so both providers need it. A version already present in
    the configured URL always wins; otherwise VISION_LLM_API_VERSION (or the
    default) is appended.
    """
    from urllib.parse import parse_qs, urlparse

    if "api-version" in parse_qs(urlparse(url).query):
        return url
    version = os.environ.get("VISION_LLM_API_VERSION", DEFAULT_API_VERSION).strip()
    if not version:
        return url
    return f"{url}{'&' if urlparse(url).query else '?'}api-version={version}"


def _resolve_provider() -> str:
    """Which backend serves the vision LLM:
    - "direct": VISION_LLM_ENDPOINT + VISION_LLM_API_KEY (OpenAI Responses
      API, Azure OpenAI included) — the default whenever an endpoint is set.
    - "dial": the model is a DIAL Core deployment, called at
      {DIAL_CORE_URL}/openai/deployments/{model}/chat/completions with DIAL
      credentials (caller headers first, DIAL_API_KEY fallback — the same
      resolution as file storage).
    VISION_LLM_PROVIDER=direct|dial overrides the inference."""
    value = os.environ.get("VISION_LLM_PROVIDER", "").lower()
    if value in ("direct", "azure", "openai"):
        return "direct"
    if value in ("dial", "dial-core", "dial_core"):
        return "dial"
    return "direct" if os.environ.get("VISION_LLM_ENDPOINT") else "dial"


class VisionLLM:
    def __init__(self):
        self.model = os.environ.get("VISION_LLM_MODEL")
        self.provider = _resolve_provider()
        self.endpoint = os.environ.get("VISION_LLM_ENDPOINT")
        self.api_key = os.environ.get("VISION_LLM_API_KEY")
        self.dial_url = os.environ.get("DIAL_CORE_URL")
        if not self.model:
            raise VisionLLMConfigError(
                "Visual inspection is not configured: set VISION_LLM_MODEL "
                "(see .env.example)."
            )
        if self.provider == "direct" and not (self.endpoint and self.api_key):
            raise VisionLLMConfigError(
                "Visual inspection (direct provider) needs VISION_LLM_ENDPOINT "
                "and VISION_LLM_API_KEY (see .env.example)."
            )
        if self.provider == "dial" and not self.dial_url:
            raise VisionLLMConfigError(
                "Visual inspection (dial provider) needs DIAL_CORE_URL so the "
                "model can be called as a DIAL deployment (see .env.example)."
            )

    def build_payload(self, images: list, prompt: str) -> dict:
        if self.provider == "dial":
            content = [{"type": "text", "text": prompt}]
            for png in images:
                b64 = base64.b64encode(png).decode()
                content.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:image/png;base64,{b64}"},
                })
            return {"messages": [{"role": "user", "content": content}]}
        content = [{"type": "input_text", "text": prompt}]
        for png in images:
            b64 = base64.b64encode(png).decode()
            content.append({
                "type": "input_image",
                "image_url": f"data:image/png;base64,{b64}",
            })
        return {"model": self.model,
                "input": [{"role": "user", "content": content}]}

    def _request_target(self):
        """(url, headers) for the configured provider."""
        if self.provider == "dial":
            from dial_client import DialConfigError, resolve_dial_auth_headers
            try:
                headers = resolve_dial_auth_headers()
            except DialConfigError as e:
                raise VisionLLMConfigError(str(e))
            headers["Content-Type"] = "application/json"
            url = (f"{self.dial_url.rstrip('/')}/openai/deployments/"
                   f"{self.model}/chat/completions")
            return _with_api_version(url), headers
        return _with_api_version(self.endpoint), {
            "api-key": self.api_key,                       # Azure OpenAI
            "Authorization": f"Bearer {self.api_key}",     # OpenAI-compatible
            "Content-Type": "application/json",
        }

    def extract_text(self, response_json: dict) -> str:
        """Pull the assistant text out of the provider's response."""
        if self.provider == "dial":
            try:
                content = response_json["choices"][0]["message"]["content"]
            except (KeyError, IndexError, TypeError):
                raise VisualQAError("Vision LLM returned no text output.")
            if isinstance(content, list):  # multimodal content parts
                content = "\n".join(
                    c.get("text", "") for c in content
                    if isinstance(c, dict) and c.get("type") == "text")
            if not isinstance(content, str) or not content:
                raise VisualQAError("Vision LLM returned no text output.")
            return content
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
        logger.warning("vision_verdict_unparseable chars=%d preview=%s",
                       len(text), flatten(text[:200]))
        return {"passed": None, "issues": [],
                "raw_review": text,
                "note": "Reviewer response was not valid JSON; see raw_review."}

    def ask(self, images: list, prompt: str, timeout: float = 300.0) -> str:
        """Send prompt + images, return the model's raw text answer."""
        url, headers = self._request_target()
        payload = self.build_payload(images, prompt)
        logger.debug("vision_request provider=%s model=%s images=%d "
                     "prompt_chars=%d timeout_s=%.0f",
                     self.provider, self.model, len(images), len(prompt), timeout)
        started = time.monotonic()
        r = httpx.post(url, headers=headers, json=payload, timeout=timeout)
        duration_ms = int((time.monotonic() - started) * 1000)
        if r.status_code != 200:
            detail = r.text[:300]
            logger.error("vision_request_failed provider=%s model=%s status=%d "
                         "duration_ms=%d detail=%s", self.provider, self.model,
                         r.status_code, duration_ms, flatten(detail))
            if "api-version" in detail:
                detail += (" — set VISION_LLM_API_VERSION to a version your "
                           "endpoint accepts (or put ?api-version=... in "
                           "VISION_LLM_ENDPOINT).")
            raise VisualQAError(
                f"Vision LLM request failed with HTTP {r.status_code}: {detail}"
            )
        text = self.extract_text(r.json())
        logger.debug("vision_response provider=%s model=%s duration_ms=%d "
                     "chars=%d", self.provider, self.model, duration_ms, len(text))
        return text

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
            logger.warning("vision_json_unparseable chars=%d preview=%s",
                           len(text), flatten(text[:200]))
            return {}


def enforcement_enabled() -> bool:
    """Visual QA is available when the vision LLM is configured — either
    directly (VISION_LLM_ENDPOINT + VISION_LLM_API_KEY) or as a DIAL Core
    deployment (DIAL_CORE_URL) — unless disabled with VISUAL_QA_ENFORCE=false.

    Gates registration of the inspect/repair tools; also required for the
    optional export gate (see export_gate_enabled)."""
    if not os.environ.get("VISION_LLM_MODEL"):
        return False
    if _resolve_provider() == "direct":
        configured = bool(os.environ.get("VISION_LLM_ENDPOINT")
                          and os.environ.get("VISION_LLM_API_KEY"))
    else:
        configured = bool(os.environ.get("DIAL_CORE_URL"))
    return configured and os.environ.get(
        "VISUAL_QA_ENFORCE", "true").lower() != "false"


def export_gate_enabled() -> bool:
    """Whether export/save should run the inspect-repair loop themselves.

    Off by default: QA is driven by the orchestrator through the
    visual_inspect_slides / visual_repair_slides tools, which it can call as
    often as it likes on whichever slides it just built. Operators who want
    the old always-on gate — a deck can never leave the server uninspected —
    set VISUAL_QA_EXPORT_GATE=true.
    """
    return (enforcement_enabled()
            and os.environ.get("VISUAL_QA_EXPORT_GATE", "false").lower() == "true")


def unresolved_policy() -> str:
    """What to do when the internal repair loop exhausts its iterations
    without a pass: "report" (default — the export fails with the unresolved
    issue list) or "export_as_is" (ship the best-effort deck).
    VISUAL_QA_ON_UNRESOLVED accepts report/block and export_as_is/export."""
    value = os.environ.get("VISUAL_QA_ON_UNRESOLVED", "report").lower()
    return "export_as_is" if value in ("export", "export_as_is") else "report"


def fail_open_on_error() -> bool:
    """VISUAL_QA_ON_ERROR=allow lets exports through when inspection itself
    fails (renderer missing, endpoint down). Default is to block."""
    return os.environ.get("VISUAL_QA_ON_ERROR", "block").lower() == "allow"


def _render_deck(pres, max_slides=None, slides=None):
    import io
    buf = io.BytesIO()
    pres.save(buf)
    return render_pptx_bytes_to_pngs(buf.getvalue(), max_slides=max_slides,
                                     slides=slides)


def _slide_cap():
    return int(os.environ.get("VISION_LLM_MAX_SLIDES", "15"))


def inspect_presentation(pres, reference_pres=None, focus: str = None,
                         slides: list = None) -> dict:
    """Render a python-pptx Presentation (and optional reference) and return
    the vision reviewer's verdict.

    slides: 1-based slide numbers to review; None reviews the whole deck
    (capped by VISION_LLM_MAX_SLIDES). Issue slide numbers in the verdict are
    always absolute deck positions, not positions within the selection.
    Raises VisualQAError on infrastructure failure (renderer/LLM).
    """
    llm = VisionLLM()
    max_slides = None if slides else _slide_cap()

    deck_images = _render_deck(pres, max_slides, slides)
    ref_images = _render_deck(reference_pres, _slide_cap()) \
        if reference_pres is not None else []

    prompt = review_prompt(bool(ref_images), focus, slides)
    if ref_images:
        prompt += (
            f"\nImage order: images 1-{len(ref_images)} are the reference "
            f"template; images {len(ref_images) + 1}-"
            f"{len(ref_images) + len(deck_images)} are the deck under review. "
            "Report issue slide numbers relative to the deck under review."
        )
    verdict = llm.review(ref_images + deck_images, prompt)
    verdict["slides_reviewed"] = slides or len(deck_images)
    logger.info("inspection_done slides=%d scope=%s reference=%s passed=%s "
                "issues=%d", len(deck_images),
                ",".join(map(str, slides)) if slides else "deck",
                bool(ref_images), verdict.get("passed"),
                len(verdict.get("issues", [])))
    return verdict


def inspect_and_repair(pres, slides: list = None, focus: str = None,
                       max_iterations: int = None) -> dict:
    """Inspect/repair loop: inspect the selected slides; on failure, repair
    them in place via LLM-planned whitelisted operations (visual_fix.py) and
    inspect again, up to VISUAL_QA_MAX_ITERATIONS (default 10) inspections.

    slides: 1-based slide numbers to work on; None means the whole deck.
    Repairs are confined to the reviewed slides — issues reported against
    other slides are ignored, so a caller iterating slide by slide never has
    the model rewrite a slide it did not ask about.

    Returns {"passed": bool, "iterations": n, "repair_rounds": [...],
    "issues": [...]} — "issues" holds what remains when passed is False.
    Raises VisualQAError on infrastructure failure (renderer/LLM).
    """
    import visual_fix

    llm = VisionLLM()
    max_slides = None if slides else _slide_cap()
    if max_iterations is None:
        max_iterations = int(os.environ.get("VISUAL_QA_MAX_ITERATIONS", "10"))
    max_iterations = max(1, max_iterations)

    repair_rounds = []
    verdict = {}
    loop_started = time.monotonic()
    logger.info("qa_loop_start scope=%s slides_cap=%s max_iterations=%d",
                ",".join(map(str, slides)) if slides else "deck",
                max_slides, max_iterations)
    for iteration in range(1, max_iterations + 1):
        round_started = time.monotonic()
        deck_images = _render_deck(pres, max_slides, slides)
        # Absolute slide number of each image, so issues and repairs address
        # deck positions even when only a subset was rendered.
        image_slides = slides or list(range(1, len(deck_images) + 1))
        verdict = llm.review(deck_images, review_prompt(False, focus, slides))
        verdict["slides_reviewed"] = slides or len(deck_images)
        issues = [i for i in verdict.get("issues", [])
                  if not slides or i.get("slide") in slides]
        logger.info("qa_round iteration=%d/%d slides=%d passed=%s issues=%d "
                    "duration_ms=%d", iteration, max_iterations,
                    len(deck_images), verdict.get("passed"), len(issues),
                    int((time.monotonic() - round_started) * 1000))
        if logger.isEnabledFor(logging.DEBUG):
            for issue in issues:
                logger.debug("qa_issue iteration=%d slide=%s severity=%s "
                             "description=%s", iteration, issue.get("slide"),
                             issue.get("severity"),
                             flatten(str(issue.get("description", ""))[:200]))
        if verdict.get("passed") is True or (slides and not issues
                                             and verdict.get("passed") is not None):
            # Passing verdict, or no issue left on the slides in scope.
            logger.info("qa_loop_passed iterations=%d repair_rounds=%d "
                        "duration_ms=%d", iteration, len(repair_rounds),
                        int((time.monotonic() - loop_started) * 1000))
            return {"passed": True, "iterations": iteration,
                    "repair_rounds": repair_rounds}
        if iteration == max_iterations or not issues:
            # Out of budget, or nothing actionable (e.g. unparseable review)
            logger.warning("qa_loop_stop reason=%s iteration=%d",
                           "budget_exhausted" if iteration == max_iterations
                           else "no_actionable_issues", iteration)
            break
        plan = visual_fix.plan_repairs(llm, issues, pres, deck_images,
                                       image_slides)
        result = visual_fix.apply_repairs(pres, plan, allowed_slides=slides)
        round_report = {
            "iteration": iteration,
            "issues_found": len(issues),
            "operations_applied": len(result["applied"]),
            "operations_skipped": len(result["skipped"]),
        }
        if result["skipped"]:
            # Why nothing changed matters more than that nothing changed:
            # "bad shape_index" means the fix targets something the repair
            # engine cannot reach (a layout/master placeholder, say), which
            # no amount of re-running will improve.
            round_report["skipped_reasons"] = visual_fix.skip_reason_summary(
                result["skipped"])
        repair_rounds.append(round_report)
        if not result["applied"]:
            logger.warning("qa_loop_stop reason=no_repair_progress iteration=%d "
                           "operations_planned=%d operations_skipped=%d",
                           iteration, len(plan), len(result["skipped"]))
            break  # no progress is possible; stop burning inspections

    remaining = [i for i in verdict.get("issues", [])
                 if not slides or i.get("slide") in slides]
    out = {"passed": False,
           "iterations": len(repair_rounds) + 1,
           "repair_rounds": repair_rounds,
           "issues": remaining}
    if repair_rounds and not repair_rounds[-1]["operations_applied"]:
        # Tell the agent what a zero-applied round means, so it stops the
        # deck rather than re-running an identical call.
        out["repair_note"] = (
            "The last round changed nothing: every planned operation was "
            "rejected (" + ", ".join(
                f"{reason} x{count}" for reason, count
                in repair_rounds[-1].get("skipped_reasons", {}).items())
            + "). Repairs are working; these issues are just outside what "
              "they can reach — 'bad shape_index' usually means the target "
              "belongs to the slide layout or master rather than the slide. "
              "Fix the content yourself with the editing tools, or report "
              "the issue to the user. Repeating this call will not help."
        )
    logger.warning("qa_loop_failed iterations=%d repair_rounds=%d "
                   "unresolved_issues=%d duration_ms=%d",
                   out["iterations"], len(repair_rounds), len(remaining),
                   int((time.monotonic() - loop_started) * 1000))
    for key in ("raw_review", "note"):
        if key in verdict:
            out[key] = verdict[key]
    return out


def review_prompt(has_reference: bool, focus: str = None,
                  slides: list = None) -> str:
    prompt = REVIEW_PROMPT.format(
        ref_note=(". The FIRST images are the reference template's slides; "
                  "the deck under review follows" if has_reference else ""),
        ref_clause=(" and matching the reference template images"
                    if has_reference else ""),
    )
    if slides:
        mapping = ", ".join(f"image {i} = slide {n}"
                            for i, n in enumerate(slides, start=1))
        prompt += (
            f"\nYou are shown only part of a larger deck: {mapping}. Report "
            "every issue with the slide number given here, not the image "
            "position, and judge each slide on its own merits."
        )
    if focus:
        prompt += f"\nAdditional focus requested by the caller: {focus}"
    return prompt
