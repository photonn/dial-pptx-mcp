"""
Custom icons: the orchestrator writes the SVG, the server renders and checks it.

Same split as images (`tools/image_tools.py`): the creative half stays with the
calling model and the mechanical half stays here. Deciding what path data reads
as a "gear" or a "supply chain" is reasoning over a style guide — that is the
orchestrator's job, and `docs/ICON_GUIDANCE.md` is the guide it works from.
Turning that SVG into something PowerPoint can embed is deterministic, and it
has to happen server-side because OOXML has no route to SVG: python-pptx
embeds raster pictures, so an icon reaches a slide as a PNG or not at all.

Why the render is checked before the icon is placed. A hand-written path is
plausible XML long before it is a recognisable pictogram: an unclosed subpath
fills into a blob, a stray coordinate leaves a hairline across the canvas, a
mistyped viewBox pushes the drawing off the edge. None of that is visible to
the agent, which cannot look at its own icon, and by the time visual QA sees it
on a slide the only repairs available are move, resize and delete — the deck
loses the icon rather than getting a correct one. So the same vision model that
reviews slides reviews the icon on its own, immediately, while regenerating it
costs one tool call.

Rendering is PyMuPDF, already a dependency for the QA rasteriser, so custom
icons work wherever the server runs — LibreOffice is not involved and neither
is a new native library.
"""
import io
import os
import re
import threading
import time
import uuid

from logging_utils import get_logger

logger = get_logger("svg_icons")

PNG_MIME = "image/png"

DEFAULT_PX = 800          # what the STADA-style guides ship; ~0.7-1.2in on a slide
MIN_PX = 64
MAX_PX = 2048
DEFAULT_MAX_KB = 64.0

# Below this fraction of inked pixels the "icon" is a few stray hairlines, and
# above the upper bound it is a solid block — both are the classic outcomes of
# a malformed path, and both are worth naming before a vision call is spent.
MIN_INK_COVERAGE = 0.004
MAX_INK_COVERAGE = 0.97


class SvgIconError(RuntimeError):
    pass


class IconStore:
    """Rendered icons, held server-side under unguessable handles.

    An icon does not go through DIAL file storage, and the reason is an
    identity boundary rather than a size one. A file this server writes lands
    in `{user}/appdata/{this-deployment}/`, which only the end user and this
    deployment may read. Placing it again means the *orchestrator* asking DIAL
    Core to grant that file to the toolset key before the call — and the
    orchestrator is neither of those two identities, so Core refuses with 403
    before the tool is even entered. (Exports do not hit this: their URL goes
    to the end user, who owns the bucket. An image-model PNG does not either:
    it reaches the conversation as an attachment the orchestrator can see and
    therefore share.)

    So the bytes stay where they were rendered. The handle is a UUID4
    capability like a presentation id, and the store is bounded the same way —
    a remote server must bound its memory.
    """

    def __init__(self, ttl_seconds=None, max_items=None):
        self._ttl = ttl_seconds if ttl_seconds is not None else int(
            os.environ.get("PPT_MCP_STATE_TTL_SECONDS", "3600"))
        self._max = max_items if max_items is not None else int(
            os.environ.get("PPT_MCP_MAX_ICONS", "100"))
        self._lock = threading.RLock()
        self._items = {}

    def _purge(self):
        """Caller holds _lock. Expiry first, then LRU down to max size."""
        now = time.monotonic()
        for key in [k for k, v in self._items.items()
                    if now - v["last_used"] > self._ttl]:
            del self._items[key]
        while len(self._items) > self._max:
            oldest = min(self._items, key=lambda k: self._items[k]["last_used"])
            del self._items[oldest]
            logger.warning("icon_evicted icon_id=%s reason=lru max=%d",
                           oldest[:8] + "…", self._max)

    def put(self, png, meta=None):
        icon_id = uuid.uuid4().hex
        with self._lock:
            self._items[icon_id] = {"png": png, "meta": meta or {},
                                    "last_used": time.monotonic()}
            self._purge()
        return icon_id

    def get(self, icon_id):
        """(png, meta) for a live handle, or None."""
        with self._lock:
            self._purge()
            entry = self._items.get(icon_id)
            if entry is None:
                return None
            entry["last_used"] = time.monotonic()
            return entry["png"], entry["meta"]

    def __contains__(self, icon_id):
        with self._lock:
            self._purge()
            return icon_id in self._items

    def __len__(self):
        with self._lock:
            return len(self._items)


# The SVG is written by the calling model and rendered by a parser in this
# process, so it is untrusted input. Rejecting these outright is cheaper than
# reasoning about what MuPDF does with them: external references are the SSRF /
# file-read surface, and the rest cannot appear in line-art anyway. Text is
# refused for a second reason as well — glyphs come from whatever font the
# renderer happens to substitute, which is exactly the artifact class this
# module exists to catch, and the style guide forbids text in an icon.
_FORBIDDEN = (
    (re.compile(r"<!DOCTYPE", re.I), "a DOCTYPE declaration"),
    (re.compile(r"<!ENTITY", re.I), "an ENTITY declaration"),
    (re.compile(r"<\s*script", re.I), "a <script> element"),
    (re.compile(r"<\s*foreignObject", re.I), "a <foreignObject> element"),
    (re.compile(r"<\s*(image|iframe)\b", re.I), "an embedded <image>/<iframe>"),
    (re.compile(r"<\s*(text|tspan|textPath)\b", re.I),
     "a text element — icons are pure line art, and glyphs render in whatever "
     "font the rasteriser substitutes"),
    (re.compile(r"\bon[a-z]+\s*=", re.I), "an event-handler attribute"),
)

# Any reference that leaves the document: url(http://…), href="file:…",
# href="//host/…". Internal fragment references (href="#clip") are fine.
_EXTERNAL_REF = re.compile(
    r"""(?:href|src)\s*=\s*["']\s*(?:[a-z][a-z0-9+.-]*:|//)"""
    r"""|url\(\s*["']?\s*(?:[a-z][a-z0-9+.-]*:|//)""", re.I)


def max_svg_bytes():
    raw = os.environ.get("SVG_ICON_MAX_KB", DEFAULT_MAX_KB)
    try:
        return int(float(raw) * 1024)
    except (TypeError, ValueError):
        logger.warning("svg_max_kb_invalid value=%s using=%s", raw,
                       DEFAULT_MAX_KB)
        return int(DEFAULT_MAX_KB * 1024)


def parse_color(value, what="background"):
    """Accept "transparent"/"none", "#RGB", "#RRGGBB" or a bare hex triplet.

    Returns an (r, g, b) tuple, or None for transparent.
    """
    if value is None:
        return None
    text = str(value).strip().lower()
    if text in ("", "transparent", "none"):
        return None
    hex_digits = text[1:] if text.startswith("#") else text
    if len(hex_digits) == 3 and re.fullmatch(r"[0-9a-f]{3}", hex_digits):
        hex_digits = "".join(c * 2 for c in hex_digits)
    if not re.fullmatch(r"[0-9a-f]{6}", hex_digits):
        raise SvgIconError(
            f"Invalid {what} '{value}': give a hex colour such as \"#FFFFFF\", "
            "or \"transparent\"."
        )
    return tuple(int(hex_digits[i:i + 2], 16) for i in (0, 2, 4))


def validate_svg(svg):
    """Reject SVG this server will not render, with a message the agent can
    act on. Returns the stripped source."""
    source = (svg or "").strip()
    if not source:
        raise SvgIconError("svg is empty: pass the SVG source as a string.")
    limit = max_svg_bytes()
    encoded = len(source.encode("utf-8"))
    if encoded > limit:
        raise SvgIconError(
            f"The SVG is {encoded / 1024:.1f} KB, over the "
            f"{limit / 1024:.0f} KB limit. An icon is a handful of paths — "
            "simplify it rather than embedding traced artwork."
        )
    if "<svg" not in source.lower():
        raise SvgIconError(
            "That is not SVG source: it must contain an <svg> element. Pass "
            "the markup itself, not a file path or a URL."
        )
    for pattern, what in _FORBIDDEN:
        if pattern.search(source):
            raise SvgIconError(
                f"The SVG contains {what}, which this server does not render. "
                "Draw the icon with <path>, <circle>, <rect>, <line>, "
                "<polyline> and <polygon> only."
            )
    if _EXTERNAL_REF.search(source):
        raise SvgIconError(
            "The SVG references an external resource. Icons must be "
            "self-contained: no remote or local file references, only shapes "
            "drawn inline."
        )
    return source


def _ink_coverage(image, background):
    """Fraction of the canvas that carries drawing.

    On a transparent render that is the alpha channel; on a solid background it
    is the pixels that differ from it. Either way it is what separates a real
    pictogram from the two failure modes a malformed path produces.
    """
    from PIL import Image, ImageChops

    if background is None:
        alpha = image.getchannel("A")
        inked = sum(count for value, count in
                    zip(range(256), alpha.histogram()) if value > 16)
    else:
        flat = Image.new("RGB", image.size, background)
        diff = ImageChops.difference(image.convert("RGB"), flat).convert("L")
        inked = sum(count for value, count in
                    zip(range(256), diff.histogram()) if value > 16)
    return inked / float(image.width * image.height)


def render_svg(svg, size=DEFAULT_PX, background=None):
    """Rasterize validated SVG source to PNG bytes.

    size is the longest side in pixels; the drawing's own aspect ratio is
    kept. background is an (r, g, b) tuple or None for a transparent PNG —
    transparency is the point for an icon that has to sit on a coloured panel,
    which is the case a pre-rendered icon library cannot serve.

    Returns (png_bytes, {"width", "height", "ink_coverage"}).
    """
    import pymupdf
    from PIL import Image

    source = validate_svg(svg)
    try:
        with pymupdf.open(stream=source.encode("utf-8"), filetype="svg") as doc:
            if not doc.page_count:
                raise SvgIconError("The SVG produced no drawing.")
            page = doc[0]
            longest = max(page.rect.width, page.rect.height)
            if longest <= 0:
                raise SvgIconError(
                    "The SVG has no drawable area: give the <svg> element a "
                    "viewBox such as viewBox=\"0 0 200 200\"."
                )
            zoom = size / longest
            pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=True)
            raw = pix.tobytes("png")
    except SvgIconError:
        raise
    except Exception as e:
        logger.warning("svg_render_failed error=%s detail=%s", type(e).__name__, e)
        raise SvgIconError(
            f"The SVG could not be rendered ({e}). Check that every tag is "
            "closed and that path data is well formed."
        )

    image = Image.open(io.BytesIO(raw)).convert("RGBA")
    if background is not None:
        flat = Image.new("RGBA", image.size, background + (255,))
        flat.alpha_composite(image)
        image = flat
    coverage = _ink_coverage(image, background)
    if coverage <= 0:
        raise SvgIconError(
            "The SVG rendered to an empty image — nothing was drawn. The "
            "usual causes are coordinates outside the viewBox, or shapes with "
            "fill=\"none\" and no stroke."
        )

    out = io.BytesIO()
    image.save(out, format="PNG", optimize=True)
    logger.info("svg_rendered size=%dx%d bytes=%d coverage=%.4f",
                image.width, image.height, out.tell(), coverage)
    return out.getvalue(), {"width": image.width, "height": image.height,
                            "ink_coverage": round(coverage, 4)}


def coverage_note(coverage):
    """A warning about a render that is almost certainly wrong, or None.

    Cheap enough to run always, and it names the defect more precisely than a
    vision model will.
    """
    if coverage < MIN_INK_COVERAGE:
        return ("Almost nothing was drawn (%.2f%% of the canvas). The shapes "
                "are probably outside the viewBox, or far too small in it."
                % (coverage * 100))
    if coverage > MAX_INK_COVERAGE:
        return ("The canvas is almost entirely filled (%.1f%% of it). A shape "
                "meant as line art is probably filled instead of stroked — "
                "set fill=\"none\" and give it a stroke."
                % (coverage * 100))
    return None


# ---- Vision review of the rendered icon ----

REVIEW_PROMPT = """You are reviewing a single icon that was generated from \
hand-written SVG for use in a PowerPoint deck. You are shown it twice: image 1 \
at full resolution, image 2 at the size it will actually appear on the slide \
(about 1 inch). {concept_note}

Judge only the icon in front of you, as pixels. Report:
1. Rendering artifacts: stray marks, hairlines or dots that belong to no shape; \
paths that did not close, or closed into a solid blob; shapes filled that should \
be line art; overlapping strokes that read as a smudge; parts of the drawing \
clipped by the edge of the canvas.
2. Composition: the drawing badly off-centre, or so small in the canvas that it \
will be a speck on the slide; strokes of visibly uneven weight; a shape that \
does not sit inside its containing circle or frame.
3. Legibility at slide size (image 2): whether it still reads as a distinct \
symbol, or collapses into a dark smear.
4. Whether it is recognisable as what it is meant to depict.

Be strict about artifacts and forgiving about style: a plain, simple pictogram \
that reads clearly is a pass.

Respond with ONLY a JSON object, no markdown fence:
{{"passed": true|false, "reads_as": "<what the icon looks like to you, in a \
few words>", "issues": [{{"severity": "critical"|"major"|"minor", \
"description": "...", "suggested_fix": "<what to change in the SVG>"}}]}}
"passed" is true only when there are no critical or major issues."""


def _preview_size(png, pixels, background):
    """The same icon as it will appear on the slide, on the slide's colour.

    A transparent icon is composited onto that colour rather than sent with an
    alpha channel, so the reviewer judges the contrast the audience gets.
    """
    from PIL import Image

    image = Image.open(io.BytesIO(png)).convert("RGBA")
    scale = min(1.0, pixels / max(image.width, image.height))
    if scale < 1.0:
        image = image.resize((max(1, round(image.width * scale)),
                              max(1, round(image.height * scale))),
                             Image.LANCZOS)
    flat = Image.new("RGBA", image.size, (background or (255, 255, 255)) + (255,))
    flat.alpha_composite(image)
    out = io.BytesIO()
    flat.convert("RGB").save(out, format="PNG")
    return out.getvalue()


def review_icon(png, concept=None, background=None, timeout=120.0):
    """Ask the vision model whether the rendered icon is clean and readable.

    Returns the parsed verdict ({"passed", "reads_as", "issues"}); raises
    visual_qa.VisualQAError (or its config subclass) if the model cannot be
    reached, which the tool layer reports as a note rather than a failure.
    """
    import visual_qa

    concept_note = (
        f"It is meant to depict: {concept}." if concept else
        "You were not told what it is meant to depict; say what it reads as."
    )
    llm = visual_qa.VisionLLM()
    images = [png, _preview_size(png, 96, background)]
    verdict = llm.review(images, REVIEW_PROMPT.format(concept_note=concept_note),
                         timeout=timeout)
    verdict.pop("slides_reviewed", None)
    logger.info("icon_reviewed concept=%s passed=%s issues=%d",
                (concept or "-")[:40], verdict.get("passed"),
                len(verdict.get("issues", [])))
    return verdict
