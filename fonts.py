"""
Which fonts this server's visual QA can actually be trusted about.

Fonts named in a .pptx are rendered by whatever machine opens it. This server's
visual QA renders through LibreOffice, which does not have Microsoft's fonts and
substitutes its own. Some of those substitutes are *metric-compatible* — every
glyph has the same advance width as the font it replaces — so a line that wraps
in the render wraps identically in PowerPoint. The rest are merely similar, and
their widths differ enough that a QA screenshot can show text overflowing a box
that fits fine in PowerPoint, or fitting one that will not.

That distinction matters twice: the reviewer should discount text-fit findings
on an unreliable font, and `fit_text`'s geometric estimate is calibrated for
widths it cannot verify. So the deck's font choices are surfaced to the agent
rather than silently trusted.

The metric-compatible pairs are the ones LibreOffice ships for exactly this
purpose (Liberation Sans/Serif/Mono for the classic Microsoft core fonts,
Carlito for Calibri, Caladea for Cambria). Everything else is a substitution by
similarity.
"""
from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

from logging_utils import get_logger

logger = get_logger("fonts")

# Font -> the metric-compatible face LibreOffice renders in its place.
# A deck restricted to these keys renders at true width in visual QA *and*
# ships with Microsoft Office, so it is safe on both ends.
METRIC_SAFE_SUBSTITUTES = {
    "arial": "Liberation Sans",
    "helvetica": "Liberation Sans",
    "times new roman": "Liberation Serif",
    "courier new": "Liberation Mono",
    "calibri": "Carlito",
    "cambria": "Caladea",
}

# Common choices with no metric-compatible substitute here. Fine for titles and
# short accents with room to spare; a poor choice for body text whose fit the
# QA pass is expected to verify.
QA_UNRELIABLE_FONTS = frozenset({
    "georgia", "verdana", "tahoma", "trebuchet ms", "garamond", "impact",
    "arial black", "consolas", "palatino linotype", "book antiqua",
    "century gothic", "century schoolbook", "bookman old style",
    "franklin gothic book", "candara", "corbel", "constantia", "segoe ui",
    "calibri light", "gill sans mt", "rockwell", "futura", "helvetica neue",
})

# Office's post-2023 default. It has no substitute in this renderer and is
# missing from Office installs older than 2024, so it is unreliable at both
# ends — never a good default to leave in a deck.
DISCOURAGED_FONTS = frozenset({"aptos", "aptos display", "aptos narrow"})

_TYPEFACE_TAGS = (qn("a:latin"), qn("a:ea"), qn("a:cs"))


def is_metric_safe(name):
    """True when visual QA renders `name` at its true character widths."""
    return bool(name) and name.strip().lower() in METRIC_SAFE_SUBSTITUTES


def substitute_for(name):
    """The face LibreOffice renders in place of `name`, when it is a known
    metric-compatible pair; None otherwise."""
    return METRIC_SAFE_SUBSTITUTES.get((name or "").strip().lower())


def _theme_fonts(pres):
    """The major/minor typefaces the master's theme defines.

    Text that inherits its font (+mj-lt / +mn-lt, and anything that names no
    typeface at all) is rendered in these, so they count as fonts in use even
    though no slide mentions them.
    """
    names = set()
    for master in pres.slide_masters:
        try:
            theme = master.part.part_related_by(RT.THEME)
        except KeyError:
            continue
        # The theme is an opaque Part here, not an XmlPart: python-pptx keeps
        # it as bytes, so it has to be parsed rather than walked.
        try:
            root = etree.fromstring(theme.blob)
        except etree.XMLSyntaxError as e:
            logger.debug("theme_unparsed part=%s error=%s", theme.partname, e)
            continue
        for element in root.iter(qn("a:latin")):
            typeface = element.get("typeface")
            if typeface and not typeface.startswith("+"):
                names.add(typeface)
    return names


def fonts_in(pres, include_theme=True):
    """Every typeface the deck's slides name, plus the theme's own.

    Layouts and masters are skipped on purpose: their text is template
    decoration this server does not author, and reporting it would flag every
    corporate template as risky.
    """
    names = set()
    for slide in pres.slides:
        for element in slide._element.iter():
            if element.tag not in _TYPEFACE_TAGS:
                continue
            typeface = element.get("typeface")
            # "+mj-lt" / "+mn-lt" are theme references, resolved separately.
            if typeface and not typeface.startswith("+"):
                names.add(typeface)
    if include_theme:
        names |= _theme_fonts(pres)
    return names


def unreliable_fonts_in(pres):
    """Fonts in the deck whose rendered widths visual QA cannot be trusted on.

    Anything that is not a known metric-compatible pair qualifies, not just the
    names in QA_UNRELIABLE_FONTS: an unrecognised font is substituted by
    similarity like any other.
    """
    return {name for name in fonts_in(pres) if not is_metric_safe(name)}


def discouraged_fonts_in(pres):
    """Fonts that are a problem on the user's machine too, not just in QA."""
    return {name for name in fonts_in(pres)
            if name.strip().lower() in DISCOURAGED_FONTS}


def safe_font_list():
    """The metric-safe names, title-cased, for guidance text."""
    return ["Arial", "Calibri", "Cambria", "Times New Roman", "Courier New"]


# Appended to the visual review prompt when the deck uses fonts the renderer
# substitutes non-metrically, so the reviewer weighs fit findings accordingly.
def qa_font_caveat(risky_fonts):
    if not risky_fonts:
        return ""
    return (
        "\nRendering caveat: these images were rendered with substitute fonts "
        "for " + ", ".join(sorted(risky_fonts)) + ", and the substitutes have "
        "different character widths from the real fonts. For text in those "
        "fonts, report only clear, substantial overflow or clipping — text "
        "that ends marginally close to its container's edge is within the "
        "margin of error here and must not be reported. Text that is fully "
        "inside its box but merely near the edge is fine. Judge every other "
        "kind of issue (overlap, alignment, contrast, placement) normally, and "
        "judge text in other fonts normally too."
    )
