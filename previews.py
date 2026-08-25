"""
Contact sheets of a deck's slides, for choosing what to build on.

`get_presentation_info` reports layout names and indices, which is not enough to
choose between them: "Title and Content" tells you nothing about where the
template puts the content, whether the slide carries a logo panel, or which of
the eight near-identical layouts is the one with the three-card row. The agent
ends up picking a layout blind and finding out at visual QA.

This module renders the deck and tiles the slides into a labelled grid, which
answers the question two ways at once: the grid is uploaded to DIAL so a person
can look at it, and — when a vision model is configured — the same images are
described back to the agent, which is the half the agent can act on.

Rendering is the existing visual-QA pipeline (LibreOffice → PDF → raster), so a
deck that previews is a deck that renders.
"""
import io

from PIL import Image, ImageDraw

from logging_utils import get_logger

logger = get_logger("previews")

# Deliberately low: these are thumbnails for picking a slide, not the
# full-resolution renders visual QA judges. A 12-slide sheet at this scale is a
# few hundred KB.
PREVIEW_DPI = 60
MAX_PER_SHEET = 12
DEFAULT_COLUMNS = 4

_LABEL_HEIGHT = 22
_PADDING = 10
_BACKGROUND = (245, 245, 247)
_LABEL_BACKGROUND = (30, 32, 38)
_LABEL_TEXT = (255, 255, 255)
_BORDER = (200, 200, 206)

JPEG_MIME = "image/jpeg"


def _label(draw, box, text):
    left, top, right, _ = box
    draw.rectangle([left, top, right, top + _LABEL_HEIGHT],
                   fill=_LABEL_BACKGROUND)
    draw.text((left + 6, top + 5), text, fill=_LABEL_TEXT)


def compose_contact_sheet(images, numbers, columns=DEFAULT_COLUMNS):
    """Tile rendered slides into one labelled grid image (JPEG bytes).

    Every cell is captioned with its absolute slide number, because the whole
    point is to name a slide afterwards — an unlabelled grid means counting.
    """
    tiles = [Image.open(io.BytesIO(data)).convert("RGB") for data in images]
    if not tiles:
        raise ValueError("no slides to preview")

    columns = max(1, min(columns, len(tiles)))
    rows = (len(tiles) + columns - 1) // columns
    cell_w = max(tile.width for tile in tiles)
    cell_h = max(tile.height for tile in tiles)

    sheet_w = columns * cell_w + (columns + 1) * _PADDING
    sheet_h = rows * (cell_h + _LABEL_HEIGHT) + (rows + 1) * _PADDING
    sheet = Image.new("RGB", (sheet_w, sheet_h), _BACKGROUND)
    draw = ImageDraw.Draw(sheet)

    for index, (tile, number) in enumerate(zip(tiles, numbers)):
        row, column = divmod(index, columns)
        left = _PADDING + column * (cell_w + _PADDING)
        top = _PADDING + row * (cell_h + _LABEL_HEIGHT + _PADDING)
        box = (left, top, left + cell_w, top + cell_h + _LABEL_HEIGHT)
        _label(draw, box, f"Slide {number}")
        # Centre a slide narrower than the widest one rather than stretching it.
        offset = left + (cell_w - tile.width) // 2
        sheet.paste(tile, (offset, top + _LABEL_HEIGHT))
        draw.rectangle([left, top, left + cell_w - 1,
                        top + cell_h + _LABEL_HEIGHT - 1], outline=_BORDER)

    buffer = io.BytesIO()
    sheet.save(buffer, format="JPEG", quality=80, optimize=True)
    return buffer.getvalue()


def render_contact_sheets(pres, slides=None, columns=DEFAULT_COLUMNS,
                          per_sheet=MAX_PER_SHEET):
    """Render a deck to one or more contact sheets.

    Returns (sheets, images, numbers): the JPEG bytes per sheet, the individual
    slide PNGs (reused for the description pass, so the deck is rendered once),
    and the absolute slide numbers behind them.
    """
    import visual_qa

    images = visual_qa._render_deck(pres, None, slides)
    numbers = slides or list(range(1, len(images) + 1))
    numbers = numbers[:len(images)]

    sheets = []
    for start in range(0, len(images), per_sheet):
        chunk = images[start:start + per_sheet]
        sheets.append(compose_contact_sheet(
            chunk, numbers[start:start + per_sheet], columns))
    logger.debug("contact_sheets_built slides=%d sheets=%d", len(images),
                 len(sheets))
    return sheets, images, numbers


DESCRIBE_PROMPT = """You are helping another agent choose which slide of a \
PowerPoint template to build its next slide on. You are shown rendered slides \
of the template, in order: {mapping}.

For each slide, describe what it is STRUCTURALLY useful for — not its current \
placeholder wording. Cover: the arrangement of its regions (title position, how \
many content areas and where), what kind of content it holds (bullets, a large \
statement, cards, a chart area, an image area, a table), roughly how much text \
it can take, and any fixed template decoration (logo, panel, rule, background \
treatment).

Then say, in "best_for", the kind of content this slide should be reused for.

Respond with ONLY a JSON object, no markdown fence:
{{"slides": [{{"slide": <number as given above>, "description": "...", \
"best_for": "..."}}]}}"""


def describe_slides(images, numbers, timeout=300.0):
    """Ask the vision model what each rendered slide is structurally good for.

    This is what makes previews usable by the agent rather than only by a
    person: the grid image is not something the orchestrator can look at.
    """
    import visual_qa

    mapping = ", ".join(f"image {i} = slide {n}"
                        for i, n in enumerate(numbers, start=1))
    llm = visual_qa.VisionLLM()
    result = llm.ask_json(images, DESCRIBE_PROMPT.format(mapping=mapping),
                          timeout=timeout)
    described = result.get("slides")
    if not isinstance(described, list):
        return None
    return [item for item in described if isinstance(item, dict)]
