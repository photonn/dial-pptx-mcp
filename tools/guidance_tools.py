"""
Design guidance for the agent building the deck.

The ~30 content tools answer "how do I place this"; none of them answers "what
should this slide look like", and a deck assembled from correct tool calls can
still be a wall of 11pt bullets. docs/DESIGN_GUIDANCE.md holds that knowledge
and this tool serves it, so the orchestrator can pull it at planning time
without the server having to prepend a few thousand tokens to every session.

The document is the single source of truth — it is read from disk, not
duplicated here — and it is split on its own "## N. Title" headings so an agent
can ask for just the part it needs mid-build.

`get_template_instructions` is the per-deck half of the same idea: the generic
document cannot know which slide of *this* template is the section divider, so
that knowledge travels with the template as a markdown sidecar, is loaded by
the tool that loads the template, and is served from the deck's own entry in
the store (see template_instructions).
"""
import re
from pathlib import Path
from typing import Dict, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from logging_utils import get_logger

logger = get_logger("tools.guidance")

UNKNOWN_ID = (
    "Unknown or expired presentation_id. Pass the presentation_id returned "
    "when the deck was created: by create_presentation, "
    "create_presentation_from_template, "
    "create_presentation_from_template_content or open_presentation"
)

GUIDANCE_PATH = (Path(__file__).resolve().parent.parent / "docs"
                 / "DESIGN_GUIDANCE.md")

_SECTION_RE = re.compile(r"^## (\d+)\.\s+(.*)$", re.MULTILINE)

_cache = {}


def _load():
    """Parse the guidance document into (full_text, {slug: (title, body)}).

    Cached on the file's mtime so an edit during development is picked up
    without a restart, and a deployed server parses it once.
    """
    try:
        stamp = GUIDANCE_PATH.stat().st_mtime
    except OSError as e:
        logger.warning("design_guidance_unreadable path=%s error=%s",
                       GUIDANCE_PATH, e)
        return None, {}

    if _cache.get("stamp") == stamp:
        return _cache["text"], _cache["sections"]

    text = GUIDANCE_PATH.read_text(encoding="utf-8")
    matches = list(_SECTION_RE.finditer(text))
    sections = {}
    for i, match in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        number, title = match.group(1), match.group(2).strip()
        body = text[match.start():end].strip()
        sections[_slug(title)] = {"number": int(number), "title": title,
                                  "body": body}
    _cache.update(stamp=stamp, text=text, sections=sections)
    return text, sections


def _slug(title):
    return re.sub(r"[^a-z0-9]+", "-", title.lower()).strip("-")


def register_guidance_tools(app: FastMCP, presentations):
    """Register the design-guidance tool."""

    @app.tool(
        annotations=ToolAnnotations(
            title="Get Deck Design Guidance",
            readOnlyHint=True,
        ),
    )
    def get_design_guidance(section: Optional[str] = None) -> Dict:
        """How to make the deck look like someone designed it.

        Read this BEFORE planning a deck, and re-read the relevant section when
        you hit a decision it covers. It carries what the individual tools
        cannot: deck structure, layout and spacing, type scale, colour, chart
        and table choices, the visual habits that make a deck look
        machine-generated, and the build loop that catches problems while they
        are still cheap.

        It also covers two things specific to this server: how differently to
        behave when the user supplied a corporate template (the default — you
        inherit their design rather than inventing one), and which fonts this
        server's renderer reproduces at true width, which is what determines
        whether you can trust a visual_inspect_slides verdict on text fit.

        section: omit for the whole document (a few pages), or name one —
        call once with no argument to see the available section names in
        "sections".
        """
        text, sections = _load()
        if text is None:
            return {"error": "The design guidance document is not available on "
                             "this server. Proceed without it: prefer the "
                             "template's own slides and fonts, keep 0.5in "
                             "margins, and inspect each slide as you build it."}

        available = [{"section": slug, "title": meta["title"]}
                     for slug, meta in sorted(sections.items(),
                                              key=lambda kv: kv[1]["number"])]
        if section is None:
            return {"guidance": text, "sections": available}

        wanted = _slug(section)
        if wanted in sections:
            return {"section": wanted, "title": sections[wanted]["title"],
                    "guidance": sections[wanted]["body"],
                    "sections": available}

        matches = [slug for slug in sections if wanted in slug]
        if len(matches) == 1:
            found = sections[matches[0]]
            return {"section": matches[0], "title": found["title"],
                    "guidance": found["body"], "sections": available}

        return {"error": f"No guidance section named '{section}'. Available "
                         f"sections: "
                         f"{', '.join(item['section'] for item in available)}.",
                "sections": available}

    @app.tool(
        annotations=ToolAnnotations(
            title="Get This Template's Instructions",
            readOnlyHint=True,
        ),
    )
    def get_template_instructions(presentation_id: str,
                                  section: Optional[str] = None) -> Dict:
        """The build instructions that came with this deck's template.

        Where get_design_guidance is the same for every deck, this is the
        template author's own rules for the template in front of you: which
        slide to duplicate for what, the accent colour, the footer text, the
        slides that must not be touched. It exists only when the template was
        loaded with an instructions document beside it — the loading tool
        says so in its "template_instructions" field.

        Read it before planning the deck, and re-read the section you need
        before building a slide: a long build will push the first read out of
        your context, and these rules override the generic guidance wherever
        the two disagree.

        section: omit for the whole document, or name one from the "sections"
        list that every response carries.
        """
        import template_instructions

        if presentation_id not in presentations:
            return {"error": UNKNOWN_ID}

        doc = presentations.get_instructions(presentation_id)
        if doc is None:
            return {"error": "This deck's template came with no instructions "
                             "document. Use get_design_guidance and follow "
                             "the template's own example slides.",
                    "loaded": False}
        return template_instructions.select(doc, section)
