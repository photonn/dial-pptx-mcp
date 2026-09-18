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
- VISION_LLM_MODEL_REVIEW / _PLAN  per-role override of the three variables
  above (also _ENDPOINT_/_API_KEY_): detection and repair planning are
  different jobs, so they can run on different models. Unset = one model for
  both, which is what every existing deployment has.
- VISION_LLM_MAX_SLIDES cap on slides sent per inspection (default 15)
- VISUAL_QA_RENDER_DPI  resolution of the review render (default 150)
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
import threading
import time
from pathlib import Path

import httpx

import fonts
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


_slots = None
_slots_guard = threading.Lock()


def _conversion_slots():
    """Semaphore bounding concurrent LibreOffice processes.

    Read lazily rather than at import so the .env file loaded in main() is
    already in effect.
    """
    global _slots
    with _slots_guard:
        if _slots is None:
            raw = os.environ.get("PPT_MCP_MAX_CONCURRENT_CONVERSIONS",
                                 "").strip()
            limit = 2
            if raw:
                try:
                    limit = max(1, int(raw))
                except ValueError:
                    logger.warning("invalid_max_concurrent_conversions "
                                   "value=%r falling_back_to_default", raw)
            logger.debug("conversion_limit slots=%d", limit)
            _slots = threading.Semaphore(limit)
        return _slots


def _seed_dir(env_name: str, destination: Path) -> bool:
    """Pre-fill a per-conversion directory from a template built at image
    build time (see the Dockerfile), when the deployment ships one.

    LibreOffice otherwise builds a user profile — and, under a fresh $HOME, a
    fontconfig cache — from nothing on every single conversion, which on a
    one-slide deck costs more than rendering the slide. The per-conversion
    isolation is not negotiable (it is what stops concurrent conversions
    fighting over the profile lock), so this seeds the private copy instead of
    sharing one. No template, or an unreadable one: soffice creates its own,
    exactly as before.
    """
    template = os.environ.get(env_name, "").strip()
    if not template or not os.path.isdir(template):
        return False
    try:
        shutil.copytree(template, destination, dirs_exist_ok=True,
                        symlinks=True)
        return True
    except OSError as e:
        logger.warning("soffice_seed_failed var=%s path=%s error=%s",
                       env_name, template, e)
        return False


def convert_with_soffice(data: bytes, source_suffix: str, target: str,
                         timeout: float = 180.0) -> bytes:
    """Run one headless LibreOffice conversion and return the output bytes.

    `source_suffix` is the input file's extension (".pptx", ".ppt"); `target`
    is LibreOffice's --convert-to argument ("pdf", "pptx"). Every caller here
    goes through this: the isolated user profile is what stops concurrent
    conversions fighting over the shared profile lock, and getting that wrong
    fails intermittently under load rather than in testing.
    """
    with tempfile.TemporaryDirectory(prefix="pptx-convert-") as tmp:
        tmp = Path(tmp)
        src = tmp / f"deck{source_suffix}"
        src.write_bytes(data)
        profile = tmp / "lo-profile"
        home = tmp / "home"
        (home / ".cache").mkdir(parents=True, exist_ok=True)
        (home / ".config").mkdir(parents=True, exist_ok=True)
        seeded = (_seed_dir("PPT_MCP_SOFFICE_PROFILE_TEMPLATE", profile),
                  _seed_dir("PPT_MCP_SOFFICE_CACHE_TEMPLATE", home / ".cache"))
        cmd = [
            _soffice_binary(), "--headless", "--norestore",
            f"-env:UserInstallation=file://{profile}",
            "--convert-to", target, "--outdir", str(tmp), str(src),
        ]
        # dconf and fontconfig key their caches off $HOME, not off the
        # UserInstallation profile, so a deployment whose $HOME is unset,
        # read-only or owned by another uid makes soffice fail before it
        # loads the document. Pointing $HOME at the per-conversion temp dir
        # keeps that working without the deployment having to supply one.
        env = {**os.environ, "HOME": str(home),
               "XDG_CACHE_HOME": str(home / ".cache"),
               "XDG_CONFIG_HOME": str(home / ".config")}
        started = time.monotonic()
        # One conversion peaks around half a gigabyte, so an unbounded burst
        # of them is what exhausts a pod's memory limit rather than its CPU
        # (which PPT_MCP_MAX_CONCURRENT_TOOL_CALLS already bounds). Queue
        # here instead: waiting is not counted against `timeout`, which
        # measures the conversion itself.
        with _conversion_slots():
            proc = subprocess.run(cmd, capture_output=True, timeout=timeout,
                                  env=env)
        # --convert-to may take a filter suffix ("pdf:impress_pdf_Export");
        # the file it writes is named after the bare extension.
        out = tmp / f"deck.{target.split(':', 1)[0]}"
        if proc.returncode != 0 or not out.exists():
            logger.error("convert_failed stage=libreoffice target=%s "
                         "returncode=%d bytes=%d stderr=%s", target,
                         proc.returncode, len(data),
                         flatten(proc.stderr.decode(errors="replace")[-300:]))
            raise VisualQAError(
                f"LibreOffice failed to convert the presentation to {target}: "
                + proc.stderr.decode(errors="replace")[-500:]
            )
        result = out.read_bytes()
        logger.debug("convert_ok target=%s in_bytes=%d out_bytes=%d "
                     "seeded_profile=%s seeded_cache=%s duration_ms=%d",
                     target, len(data), len(result), seeded[0], seeded[1],
                     int((time.monotonic() - started) * 1000))
        return result


def render_pptx_bytes_to_pdf(pptx_data: bytes) -> bytes:
    """Convert presentation bytes to PDF bytes."""
    return convert_with_soffice(pptx_data, ".pptx", "pdf")


def convert_legacy_ppt(data: bytes) -> bytes:
    """Convert a binary PowerPoint 97-2003 (.ppt) file to .pptx bytes.

    python-pptx reads only OOXML, so a legacy deck has to be converted before
    anything else in this server can touch it.
    """
    return convert_with_soffice(data, ".ppt", "pptx")


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

    started = time.monotonic()
    images = []
    with pymupdf.open(stream=render_pptx_bytes_to_pdf(pptx_data),
                      filetype="pdf") as doc:
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

Report what you find in two separate lists, because they are answered \
differently: a defect is repaired automatically, a judgement is handed to the \
author. Judge the rendered pixels, not what the text probably says.

"issues" — objective defects, visible in the pixels, that a reader would see as \
broken. These and only these decide "passed". Look for them ANYWHERE text \
appears, not just in text boxes:
   - Text overflowing, clipped, or spilling outside its container or the slide edge.
   - Text overlapping other text, or sitting on top of a shape, image or line in \
a way that makes either unreadable.
   - Elements partly or wholly off the slide.
   - Placeholder text left unfilled (e.g. "Click to add title").
   - Charts, tables or images that are broken or empty.
   - Charts: axis tick labels colliding with each other or truncated ("..." or \
cut-off words), data labels overlapping their bars/slices/points or each other, \
a legend covering the plot area or running off the chart, an axis title squeezed \
or rotated into illegibility, series labels detached from what they label.
   - Tables: cell text clipped by the cell or row height or wrapping into an \
unreadable stack, columns too narrow for their content, headers not aligned with \
their columns, a table extending past the slide.
   - Diagrams, SmartArt and grouped shapes: labels wider than the node or box \
that holds them, text escaping a connector or arrow, node labels overlapping \
neighbouring nodes or connectors.

"observations" — everything that is a matter of judgement rather than a defect: \
text sized badly for the space it occupies, comparable elements at visibly \
different sizes, uneven gaps, inconsistent alignment or spacing, and template \
fidelity (colors, fonts, logo placement, layout usage{ref_clause}). These are \
reported to the author and never make a deck fail — do not let them influence \
"passed", and do not repeat an entry from "issues" here.

Report each problem separately, naming the element it affects (e.g. "chart on \
the right: x-axis labels overlap"), and say in "suggested_fix" what change would \
resolve it (resize, reposition, shorten the text, smaller font, wider column, \
hide the legend).

Respond with ONLY a JSON object, no markdown fence:
{{"passed": true|false, "issues": [{{"slide": <1-based number>, "severity": \
"critical"|"major"|"minor", "description": "...", "suggested_fix": "..."}}], \
"observations": [{{"slide": <1-based number>, "description": "..."}}]}}
Unreadable or overlapping text is at least a major issue. "passed" is true when \
"issues" holds no critical or major entry, whatever "observations" contains."""


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


def _role_env(name: str, role: str) -> str:
    """VISION_LLM_<NAME>_<ROLE>, falling back to VISION_LLM_<NAME>.

    Reviewing is detection and planning a repair is reasoning; they are worth
    running on different models. Resolving per role here is what lets an
    operator split them by setting one variable, while every deployment that
    sets only the plain variable keeps one model for both.
    """
    if role:
        value = os.environ.get(f"{name}_{role.upper()}")
        if value:
            return value
    return os.environ.get(name)


def _resolve_provider(endpoint: str = None) -> str:
    """Which backend serves the vision LLM:
    - "direct": VISION_LLM_ENDPOINT + VISION_LLM_API_KEY (OpenAI Responses
      API, Azure OpenAI included) — the default whenever an endpoint is set.
    - "dial": the model is a DIAL Core deployment, called at
      {DIAL_CORE_URL}/openai/deployments/{model}/chat/completions with DIAL
      credentials (caller headers first, DIAL_API_KEY fallback — the same
      resolution as file storage).
    VISION_LLM_PROVIDER=direct|dial overrides the inference.

    `endpoint` is the endpoint of the role asking (see _role_env); the plain
    VISION_LLM_ENDPOINT is used when no role narrows it."""
    value = os.environ.get("VISION_LLM_PROVIDER", "").lower()
    if value in ("direct", "azure", "openai"):
        return "direct"
    if value in ("dial", "dial-core", "dial_core"):
        return "dial"
    if endpoint is None:
        endpoint = os.environ.get("VISION_LLM_ENDPOINT")
    return "direct" if endpoint else "dial"


class VisionLLM:
    def __init__(self, role: str = "review"):
        self.role = role
        self.model = _role_env("VISION_LLM_MODEL", role)
        self.endpoint = _role_env("VISION_LLM_ENDPOINT", role)
        self.api_key = _role_env("VISION_LLM_API_KEY", role)
        self.provider = _resolve_provider(self.endpoint)
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
            return self._with_reasoning(
                {"messages": [{"role": "user", "content": content}]})
        content = [{"type": "input_text", "text": prompt}]
        for png in images:
            b64 = base64.b64encode(png).decode()
            content.append({
                "type": "input_image",
                "image_url": f"data:image/png;base64,{b64}",
            })
        return self._with_reasoning(
            {"model": self.model,
             "input": [{"role": "user", "content": content}]})

    def _with_reasoning(self, payload: dict) -> dict:
        """Attach a reasoning-effort hint for this role, if one is configured.

        Nothing is sent when the variable is unset: Azure and DIAL both reject
        an unrecognised parameter on some deployments, so an unasked-for field
        is a failed call, not a no-op.
        """
        effort = os.environ.get(
            f"VISION_LLM_{self.role.upper()}_REASONING", "").strip()
        if not effort:
            return payload
        if self.provider == "dial":
            payload["reasoning_effort"] = effort
        else:
            payload["reasoning"] = {"effort": effort}
        return payload

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
                verdict.setdefault("observations", [])
                return verdict
        except json.JSONDecodeError:
            pass
        logger.warning("vision_verdict_unparseable chars=%d preview=%s",
                       len(text), flatten(text[:200]))
        return {"passed": None, "issues": [], "observations": [],
                "raw_review": text,
                "note": "Reviewer response was not valid JSON; see raw_review."}

    def ask(self, images: list, prompt: str, timeout: float = 300.0) -> str:
        """Send prompt + images, return the model's raw text answer."""
        url, headers = self._request_target()
        payload = self.build_payload(images, prompt)
        if logger.isEnabledFor(logging.DEBUG):
            # PNG -> base64 inflates every image by a third, on every round;
            # the payload is the one cost here that nothing else reports.
            logger.debug("vision_request provider=%s model=%s role=%s images=%d "
                         "prompt_chars=%d payload_bytes=%d timeout_s=%.0f",
                         self.provider, self.model, self.role, len(images),
                         len(prompt), sum(len(p) for p in images) * 4 // 3,
                         timeout)
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


def vision_configured() -> bool:
    """Whether a vision model can be reached at all — either directly
    (VISION_LLM_ENDPOINT + VISION_LLM_API_KEY) or as a DIAL Core deployment
    (DIAL_CORE_URL).

    Deliberately separate from enforcement_enabled(): VISUAL_QA_ENFORCE turns
    off *slide* inspection and repair, and a feature that merely uses the model
    for its own check — the icon review — should follow whether the model
    exists, not that switch."""
    if not os.environ.get("VISION_LLM_MODEL"):
        return False
    if _resolve_provider() == "direct":
        return bool(os.environ.get("VISION_LLM_ENDPOINT")
                    and os.environ.get("VISION_LLM_API_KEY"))
    return bool(os.environ.get("DIAL_CORE_URL"))


def enforcement_enabled() -> bool:
    """Visual QA is available when the vision LLM is configured, unless
    disabled with VISUAL_QA_ENFORCE=false.

    Gates registration of the inspect/repair tools; also required for the
    optional export gate (see export_gate_enabled)."""
    return vision_configured() and os.environ.get(
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


def _subset_deck_bytes(pres, slides):
    """Serialize a copy of `pres` containing only the given 1-based slides,
    kept in their original deck order (not the order of `slides`).

    LibreOffice always converts the whole file it is given, so trimming here
    — before the file ever reaches soffice — is what actually cuts render
    time for a slide-scoped inspect/repair call, instead of converting the
    full deck and discarding the unwanted pages afterward. Operates on a
    freshly reopened copy so the caller's live `pres` (and its bound
    `shapes`/spTree state, see CLAUDE.md) is never touched.
    """
    import io
    from pptx import Presentation
    from utils import delete_slide

    buf = io.BytesIO()
    pres.save(buf)
    if len(slides) >= len(pres.slides):
        # Nothing to trim: the selection already is the deck. Reopening it
        # only to delete no slides and serialize it again costs a second full
        # save of a deck that can run to tens of megabytes.
        return buf.getvalue()
    buf.seek(0)
    subset = Presentation(buf)
    keep = set(slides)
    for index in range(len(subset.slides) - 1, -1, -1):
        if (index + 1) not in keep:
            delete_slide(subset, index)
    out = io.BytesIO()
    subset.save(out)
    return out.getvalue()


def _render_deck(pres, max_slides=None, slides=None, dpi=None):
    """Render `pres` (or a slide subset of it) to PNGs.

    dpi=None keeps the renderer's own default, which is what the preview and
    summary-card composers want; the review path passes _render_dpi().
    """
    import io
    extra = {} if dpi is None else {"dpi": dpi}
    if slides:
        # slides is always pre-sorted (normalize_slides / the repair loop's
        # own image_slides), so the subset deck's page order already matches
        # what callers expect back — no slides= filtering needed downstream.
        return render_pptx_bytes_to_pngs(_subset_deck_bytes(pres, slides),
                                         max_slides=max_slides, **extra)
    buf = io.BytesIO()
    pres.save(buf)
    return render_pptx_bytes_to_pngs(buf.getvalue(), max_slides=max_slides,
                                     slides=slides, **extra)


DEFAULT_RENDER_DPI = 150


def _render_dpi():
    """Resolution of the images the reviewer is shown.

    The renderer's own default (96 dpi, 1280x720) leaves small text marginal,
    and a reviewer that cannot quite read a label reports it as unreadable —
    a false finding costs a whole repair round, which is far more than the
    larger image costs. Lower it with VISUAL_QA_RENDER_DPI where the payload
    size matters more than the false-finding rate.
    """
    raw = os.environ.get("VISUAL_QA_RENDER_DPI", "").strip()
    if not raw:
        return DEFAULT_RENDER_DPI
    try:
        return max(36, int(raw))
    except ValueError:
        logger.warning("invalid_render_dpi value=%r falling_back_to_default",
                       raw)
        return DEFAULT_RENDER_DPI


def _slide_cap():
    return int(os.environ.get("VISION_LLM_MAX_SLIDES", "15"))


# ---- What the loop acts on ----

# Severity of an issue, ranked. An issue with no severity, or one the
# reviewer invented, counts as "major": acting on it is what this loop did
# before severities were filtered at all, so it is the conservative reading.
_SEVERITY_RANK = {"critical": 3, "major": 2, "minor": 1}
_DEFAULT_SEVERITY = "major"


def _severity_rank(issue) -> int:
    value = str(issue.get("severity", "") or "").strip().lower()
    return _SEVERITY_RANK.get(value, _SEVERITY_RANK[_DEFAULT_SEVERITY])


def repair_severity_floor() -> str:
    """Lowest severity that may drive another repair round
    (VISUAL_QA_REPAIR_SEVERITY, default "major").

    Below the floor an issue is still reported to the agent; it just does not
    buy a render, a plan and a re-review of its own. A cosmetic finding the
    reviewer will report again next round is otherwise indistinguishable from
    a defect, and burns the budget at the same rate."""
    value = os.environ.get("VISUAL_QA_REPAIR_SEVERITY", "").strip().lower()
    return value if value in _SEVERITY_RANK else _DEFAULT_SEVERITY


def _actionable(issues) -> list:
    floor = _SEVERITY_RANK[repair_severity_floor()]
    return [i for i in issues if _severity_rank(i) >= floor]


def _iteration_budget(slides, max_iterations=None) -> int:
    """How many inspect/repair rounds this call may run.

    An explicit argument always wins. Otherwise the budget follows the scope:
    VISUAL_QA_MAX_ITERATIONS is a whole-deck number, and spending it on a
    single slide is how a 20-second call becomes a 200-second one — a slide
    that two rounds cannot fix needs its content rebuilt, not a third round.
    """
    if max_iterations is None:
        name = ("VISUAL_QA_MAX_ITERATIONS_SLIDE" if slides
                else "VISUAL_QA_MAX_ITERATIONS")
        default = "2" if slides else "10"
        raw = os.environ.get(name, "").strip() or default
        try:
            max_iterations = int(raw)
        except ValueError:
            logger.warning("invalid_max_iterations var=%s value=%r "
                           "falling_back_to_default", name, raw)
            max_iterations = int(default)
    return max(1, max_iterations)


# deck_validation codes naming a geometry defect that a whitelisted repair
# operation can actually fix, and the severity each is worth to the planner.
# Nothing else it reports belongs here: a broken relationship is not
# geometry, and a distorted picture is geometry nothing in the whitelist can
# undo (see deck_validation._check_picture) — that one is an observation.
_REPAIRABLE_VALIDATION_CODES = {
    "shape_off_slide": "critical",
    "zero_sized_shape": "critical",
    "partial_transform": "major",
}
_OBSERVED_VALIDATION_CODES = ("distorted_picture",)


def _to_slide_number(problem):
    """deck_validation numbers slides from 0 (`slide_index`); visual_qa and
    visual_fix number them from 1 (`slide`). This is the only place the two
    conventions meet — convert here, nowhere else."""
    index = problem.get("slide_index")
    if not isinstance(index, int) or isinstance(index, bool):
        return None
    return index + 1


def _renumber(message, number):
    """Rewrite deck_validation's own 0-based "Slide N," prefix, so everything
    the planner and the agent read counts slides the same way."""
    return re.sub(r"^Slide \d+,", f"Slide {number},", message or "", count=1)


def structural_findings(pres, slides=None):
    """Geometry defects found without rendering anything.

    Returns (issues, observations) in the reviewer's own vocabulary, so the
    loop can feed them to the planner beside what the model saw. A shape
    parked off the slide is a fact about the XML, not a judgement about
    pixels: paying a vision round to discover it — and another to confirm the
    fix — is the most expensive way to learn something that takes
    milliseconds. deck_validation's "fix" strings already name the repair
    operation that resolves each one, which is exactly what the planner needs.
    """
    import deck_validation

    try:
        report = deck_validation.validate_presentation(pres)
    except Exception as e:  # never fail a QA call over the cheap check
        logger.warning("qa_validation_skipped error=%s", e)
        return [], []
    issues, observations = [], []
    for problem in report.get("problems", []):
        code = problem.get("code")
        number = _to_slide_number(problem)
        if number is None or (slides and number not in slides):
            continue
        message = _renumber(problem.get("message", ""), number)
        if code in _REPAIRABLE_VALIDATION_CODES:
            issues.append({
                "slide": number,
                "severity": _REPAIRABLE_VALIDATION_CODES[code],
                "source": "structure",
                "description": message,
                "suggested_fix": problem.get("fix", ""),
            })
        elif code in _OBSERVED_VALIDATION_CODES:
            observations.append({"slide": number, "source": "structure",
                                 "description": message,
                                 "suggested_fix": problem.get("fix", "")})
    return issues, observations


def _in_scope(entries, slides):
    """Findings on the slides this call is about. An entry that names no
    slide is kept: it is about the selection as a whole."""
    if not slides:
        return list(entries)
    return [e for e in entries
            if not isinstance(e, dict) or e.get("slide") in slides
            or e.get("slide") is None]


def _accept_seed(initial_verdict, slides):
    """Validate a verdict the caller already paid for, or None.

    It arrives from the orchestrator, so it is checked rather than trusted,
    and its issues go through the same scope filter as a fresh review's —
    a seeded verdict must not be able to move a slide nobody put in scope.
    """
    if initial_verdict is None:
        return None
    reason = None
    if not isinstance(initial_verdict, dict):
        reason = "not_an_object"
    elif "passed" not in initial_verdict:
        reason = "no_passed_key"
    elif not isinstance(initial_verdict.get("issues", []), list):
        reason = "issues_not_a_list"
    if reason:
        logger.warning("qa_seed_ignored reason=%s", reason)
        return None
    seed = dict(initial_verdict)
    seed["issues"] = [i for i in seed.get("issues", [])
                      if isinstance(i, dict)
                      and (not slides or i.get("slide") in slides)]
    seed.setdefault("observations", [])
    return seed


def inspect_presentation(pres, reference_pres=None, focus: str = None,
                         slides: list = None) -> dict:
    """Render a python-pptx Presentation (and optional reference) and return
    the vision reviewer's verdict.

    slides: 1-based slide numbers to review; None reviews the whole deck
    (capped by VISION_LLM_MAX_SLIDES). Issue slide numbers in the verdict are
    always absolute deck positions, not positions within the selection.

    "issues" holds objective defects and decides "passed"; "observations"
    holds the reviewer's judgements about composition, which are reported but
    never fail a deck.
    Raises VisualQAError on infrastructure failure (renderer/LLM).
    """
    llm = VisionLLM()
    max_slides = None if slides else _slide_cap()
    dpi = _render_dpi()

    deck_images = _render_deck(pres, max_slides, slides, dpi)
    ref_images = _render_deck(reference_pres, _slide_cap(), dpi=dpi) \
        if reference_pres is not None else []

    structural, observed = structural_findings(pres, slides)
    prompt = review_prompt(bool(ref_images), focus, slides,
                           fonts.unreliable_fonts_in(pres), structural)
    if ref_images:
        prompt += (
            f"\nImage order: images 1-{len(ref_images)} are the reference "
            f"template; images {len(ref_images) + 1}-"
            f"{len(ref_images) + len(deck_images)} are the deck under review. "
            "Report issue slide numbers relative to the deck under review."
        )
    verdict = llm.review(ref_images + deck_images, prompt)
    verdict["slides_reviewed"] = slides or len(deck_images)
    # The structural pass is cheap and certain, so its findings join the
    # reviewer's rather than waiting for the model to notice them.
    if structural:
        verdict["issues"] = structural + list(verdict.get("issues", []))
        verdict["passed"] = False
    if observed:
        verdict["observations"] = observed + list(
            verdict.get("observations", []))
    logger.info("inspection_done slides=%d scope=%s reference=%s passed=%s "
                "issues=%d structural=%d", len(deck_images),
                ",".join(map(str, slides)) if slides else "deck",
                bool(ref_images), verdict.get("passed"),
                len(verdict.get("issues", [])), len(structural))
    return verdict


def inspect_and_repair(pres, slides: list = None, focus: str = None,
                       max_iterations: int = None, *,
                       initial_verdict: dict = None) -> dict:
    """Inspect/repair loop: inspect the selected slides; on failure, repair
    them in place via LLM-planned whitelisted operations (visual_fix.py) and
    inspect again, up to the iteration budget (see _iteration_budget).

    slides: 1-based slide numbers to work on; None means the whole deck.
    Repairs are confined to the reviewed slides — issues reported against
    other slides are ignored, so a caller iterating slide by slide never has
    the model rewrite a slide it did not ask about.

    initial_verdict: a verdict the caller already holds for these same slides
    (typically the one visual_inspect_slides just returned). It stands in for
    the first review, which is the single largest saving available here: the
    first round of a repair otherwise re-asks the model a question that was
    answered seconds ago. It is validated, not trusted; anything unusable is
    logged and a normal first review runs instead.

    Returns {"passed": bool, "iterations": n, "repair_rounds": [...],
    "issues": [...], "observations": [...]} — "issues" holds what remains when
    passed is False.
    Raises VisualQAError on infrastructure failure (renderer/LLM).
    """
    import visual_fix

    reviewer = VisionLLM("review")
    planner = VisionLLM("plan")
    max_slides = None if slides else _slide_cap()
    dpi = _render_dpi()
    max_iterations = _iteration_budget(slides, max_iterations)
    seed = _accept_seed(initial_verdict, slides)

    # Constant for the whole loop: repairs never change which fonts the deck
    # names, and re-scanning per round would only cost time.
    risky_fonts = fonts.unreliable_fonts_in(pres)
    repair_rounds = []
    verdict = {}
    issues, observations = [], []
    previous_count = None
    stop_reason = None
    loop_started = time.monotonic()
    logger.info("qa_loop_start scope=%s slides_cap=%s max_iterations=%d "
                "seeded=%s severity_floor=%s",
                ",".join(map(str, slides)) if slides else "deck",
                max_slides, max_iterations, str(seed is not None).lower(),
                repair_severity_floor())
    for iteration in range(1, max_iterations + 1):
        round_started = time.monotonic()
        # The structural pass runs every round, not only the first: it is
        # milliseconds, and re-running it is how a geometry repair gets
        # confirmed without a vision call.
        structural, observed = structural_findings(pres, slides)
        if iteration == 1 and seed is not None:
            # The images are only needed to plan a repair, and a seeded round
            # that already passes never plans one — so do not render yet.
            deck_images, verdict = None, seed
        else:
            deck_images = _render_deck(pres, max_slides, slides, dpi)
            verdict = reviewer.review(
                deck_images,
                review_prompt(False, focus, slides, risky_fonts, structural))
        reviewed_count = (len(deck_images) if deck_images is not None
                          else len(slides or pres.slides))
        verdict["slides_reviewed"] = slides or reviewed_count
        reviewed = [i for i in verdict.get("issues", [])
                    if not slides or i.get("slide") in slides]
        issues = structural + reviewed
        observations = observed + _in_scope(
            verdict.get("observations", []), slides)
        actionable = _actionable(issues)
        logger.info("qa_round iteration=%d/%d slides=%d passed=%s issues=%d "
                    "structural=%d actionable=%d duration_ms=%d",
                    iteration, max_iterations, reviewed_count,
                    verdict.get("passed"), len(issues), len(structural),
                    len(actionable),
                    int((time.monotonic() - round_started) * 1000))
        if logger.isEnabledFor(logging.DEBUG):
            for issue in issues:
                logger.debug("qa_issue iteration=%d slide=%s severity=%s "
                             "description=%s", iteration, issue.get("slide"),
                             issue.get("severity"),
                             flatten(str(issue.get("description", ""))[:200]))
        if repair_rounds:
            # Filled in a round late, because "did that round help?" is a
            # question only the next round's count can answer. Counted the
            # same way as issues_found, so the pair is comparable.
            repair_rounds[-1]["issues_remaining"] = len(actionable)
        if not structural and (
                verdict.get("passed") is True
                or (slides and not reviewed
                    and verdict.get("passed") is not None)):
            # Passing verdict, or no issue left on the slides in scope.
            logger.info("qa_loop_passed iterations=%d repair_rounds=%d "
                        "duration_ms=%d", iteration, len(repair_rounds),
                        int((time.monotonic() - loop_started) * 1000))
            out = {"passed": True, "iterations": iteration,
                   "repair_rounds": repair_rounds}
            if observations:
                out["observations"] = observations
            return out
        if iteration == max_iterations:
            stop_reason = "budget_exhausted"
        elif not issues:
            # Nothing actionable (e.g. an unparseable review)
            stop_reason = "no_actionable_issues"
        elif not actionable:
            stop_reason = "no_severity_match"
        elif previous_count is not None and len(actionable) >= previous_count:
            # Repairs are landing and the findings are not going down. Another
            # round costs the same and reaches the same place.
            stop_reason = "issues_not_reducing"
        if stop_reason:
            logger.warning("qa_loop_stop reason=%s iteration=%d",
                           stop_reason, iteration)
            break
        previous_count = len(actionable)
        if deck_images is None:
            deck_images = _render_deck(pres, max_slides, slides, dpi)
        # Absolute slide number of each image, so issues and repairs address
        # deck positions even when only a subset was rendered.
        image_slides = slides or list(range(1, len(deck_images) + 1))
        plan = visual_fix.plan_repairs(planner, actionable, pres, deck_images,
                                       image_slides)
        result = visual_fix.apply_repairs(pres, plan, allowed_slides=slides)
        round_report = {
            "iteration": iteration,
            "issues_found": len(actionable),
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
            stop_reason = "no_repair_progress"
            logger.warning("qa_loop_stop reason=no_repair_progress iteration=%d "
                           "operations_planned=%d operations_skipped=%d",
                           iteration, len(plan), len(result["skipped"]))
            break  # no progress is possible; stop burning inspections

    out = {"passed": False,
           "iterations": len(repair_rounds) + 1,
           "repair_rounds": repair_rounds,
           "issues": issues,
           "stop_reason": stop_reason}
    if observations:
        out["observations"] = observations
    note = _repair_note(stop_reason, repair_rounds)
    if note:
        out["repair_note"] = note
    logger.warning("qa_loop_failed iterations=%d repair_rounds=%d "
                   "unresolved_issues=%d reason=%s duration_ms=%d",
                   out["iterations"], len(repair_rounds), len(issues),
                   stop_reason, int((time.monotonic() - loop_started) * 1000))
    for key in ("raw_review", "note"):
        if key in verdict:
            out[key] = verdict[key]
    return out


def _repair_note(stop_reason, repair_rounds):
    """What a failed loop should tell the agent to do next.

    Every one of these ends the same way — do not call this tool again — but
    for different reasons, and the reason is what decides the agent's next
    move: reach for a different tool, or accept the finding.
    """
    if stop_reason == "no_repair_progress" and repair_rounds:
        return (
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
    if stop_reason == "issues_not_reducing":
        return (
            "Repairs are being applied but the findings are not going down, "
            "so the loop stopped rather than spend the rest of its budget. "
            "The remaining issues need different content, not different "
            "geometry: rebuild this slide with the editing tools — shorten "
            "the text, split it across two slides, or use a layout with more "
            "room — and inspect again. Repeating this call will not help."
        )
    if stop_reason == "no_severity_match":
        return (
            "The issues that remain are all below "
            f"VISUAL_QA_REPAIR_SEVERITY ({repair_severity_floor()}), so no "
            "repair round was spent on them. They are listed for you to "
            "judge: fix them with the editing tools if they matter, or leave "
            "them."
        )
    if stop_reason == "budget_exhausted":
        return (
            "The iteration budget ran out with issues still open. Another "
            "identical call would start from the same place — edit the slide "
            "content yourself and inspect again, or tell the user what "
            "remains."
        )
    return None


def review_prompt(has_reference: bool, focus: str = None,
                  slides: list = None, risky_fonts=None,
                  structural=None) -> str:
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
    prompt += fonts.qa_font_caveat(risky_fonts)
    if structural:
        # Already known, already being repaired. Telling the reviewer keeps it
        # from spending its attention — and a line of the verdict — on a
        # finding that is on its way out anyway.
        listed = "\n".join(f"- slide {i.get('slide')}: {i.get('description')}"
                            for i in structural)
        prompt += (
            "\nA structural check of the file has already found these, and "
            "they are being fixed; do not report them again:\n" + listed
        )
    if focus:
        prompt += f"\nAdditional focus requested by the caller: {focus}"
    return prompt
