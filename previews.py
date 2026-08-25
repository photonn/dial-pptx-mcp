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

from PIL import Image, ImageDraw, ImageFont

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

# The summary card is one image of a whole deck, so its cells shrink as the
# deck grows. These bound the result: a card wider than CARD_MAX_WIDTH is
# scaled down in a chat window anyway, and CARD_MIN_CELL_WIDTH is the point
# below which a thumbnail stops being recognisable. When the two conflict the
# minimum wins and the card gets taller — a tall card scrolls, an illegible
# one is useless.
CARD_MAX_WIDTH = 2000
CARD_MAX_HEIGHT = 2600
CARD_MIN_CELL_WIDTH = 150
CARD_TARGET_ASPECT = 1.4
# Cell captions scale with the cells: an 11px bitmap caption under a 480px
# thumbnail is as hard to read as a full-size one under a 150px thumbnail.
_CARD_MIN_FONT = 11
_CARD_MAX_FONT = 18
_HEADER_HEIGHT = 44
_HEADER_FONT = 20


def _font(size):
    """A scalable font, falling back to the fixed 11px bitmap default.

    Pillow only grew `load_default(size=)` in 10.1; on anything older the
    labels are small rather than absent.
    """
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def _text_height(font):
    return font.getbbox("Ag")[3] if font is not None else 11


def _label(draw, box, text, height=_LABEL_HEIGHT, font=None):
    left, top, right, _ = box
    draw.rectangle([left, top, right, top + height], fill=_LABEL_BACKGROUND)
    draw.text((left + 6, top + (height - _text_height(font)) // 2), text,
              fill=_LABEL_TEXT, font=font)


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


def auto_columns(count, tile_aspect=16 / 9, target=CARD_TARGET_ASPECT):
    """Pick a column count that keeps the whole grid roughly `target`-shaped.

    Laying 30 wide slides out in 4 columns gives a card three times taller
    than it is wide, which a chat window shows as a stripe. Solving
    (columns * w) / (rows * h) = target for a square-ish grid gives this.
    """
    if count <= 1:
        return 1
    return max(1, min(count, round((count * target / tile_aspect) ** 0.5)))


def _card_label_height(cell_w):
    size = max(_CARD_MIN_FONT, min(_CARD_MAX_FONT, cell_w // 22))
    return size + 8


def _fit(tile, cell_w, cell_h):
    scale = min(cell_w / tile.width, cell_h / tile.height)
    if scale >= 1:
        return tile
    return tile.resize((max(1, round(tile.width * scale)),
                        max(1, round(tile.height * scale))),
                       Image.LANCZOS)


def _header(draw, width, title, count):
    font = _font(_HEADER_FONT)
    baseline = (_HEADER_HEIGHT - _text_height(font)) // 2
    draw.rectangle([0, 0, width, _HEADER_HEIGHT], fill=_LABEL_BACKGROUND)
    draw.text((_PADDING, baseline), title, fill=_LABEL_TEXT, font=font)
    tally = f"{count} slide{'s' if count != 1 else ''}"
    draw.text((width - _PADDING - draw.textlength(tally, font=font), baseline),
              tally, fill=_LABEL_TEXT, font=font)


def compose_summary_card(images, numbers, title=None, columns=None,
                         max_width=CARD_MAX_WIDTH,
                         max_height=CARD_MAX_HEIGHT):
    """Tile every slide into a single labelled card image (JPEG bytes).

    Unlike compose_contact_sheet, which pages a long deck into several sheets
    of same-size thumbnails, this always returns exactly one image and scales
    the cells to fit it — the card is a finished-work summary a person looks
    at once, and two of them is not a summary.
    """
    tiles = [Image.open(io.BytesIO(data)).convert("RGB") for data in images]
    if not tiles:
        raise ValueError("no slides to summarize")

    native_w = max(tile.width for tile in tiles)
    native_h = max(tile.height for tile in tiles)
    if columns is None:
        columns = auto_columns(len(tiles), native_w / native_h)
    columns = max(1, min(int(columns), len(tiles)))
    rows = (len(tiles) + columns - 1) // columns

    header = _HEADER_HEIGHT if title else 0
    cell_w = min(native_w,
                 (max_width - (columns + 1) * _PADDING) // columns)
    label_h = _card_label_height(cell_w)
    room = max_height - header - (rows + 1) * _PADDING - rows * label_h
    if room > 0:
        cell_w = min(cell_w, int(room / rows * native_w / native_h))
    cell_w = max(CARD_MIN_CELL_WIDTH, cell_w)
    label_h = _card_label_height(cell_w)
    label_font = _font(label_h - 8)
    cell_h = max(1, round(cell_w * native_h / native_w))

    sheet_w = columns * cell_w + (columns + 1) * _PADDING
    sheet_h = header + rows * (cell_h + label_h) + (rows + 1) * _PADDING
    card = Image.new("RGB", (sheet_w, sheet_h), _BACKGROUND)
    draw = ImageDraw.Draw(card)
    if title:
        _header(draw, sheet_w, title, len(tiles))

    for index, (tile, number) in enumerate(zip(tiles, numbers)):
        row, column = divmod(index, columns)
        left = _PADDING + column * (cell_w + _PADDING)
        top = header + _PADDING + row * (cell_h + label_h + _PADDING)
        _label(draw, (left, top, left + cell_w, top + cell_h + label_h),
               f"Slide {number}", label_h, label_font)
        scaled = _fit(tile, cell_w, cell_h)
        card.paste(scaled, (left + (cell_w - scaled.width) // 2,
                            top + label_h + (cell_h - scaled.height) // 2))
        draw.rectangle([left, top, left + cell_w - 1,
                        top + cell_h + label_h - 1], outline=_BORDER)

    buffer = io.BytesIO()
    card.save(buffer, format="JPEG", quality=82, optimize=True)
    logger.debug("summary_card_built slides=%d columns=%d size=%dx%d bytes=%d",
                 len(tiles), columns, sheet_w, sheet_h, buffer.tell())
    return buffer.getvalue(), (sheet_w, sheet_h), columns


def render_summary_card(pres, title=None, columns=None):
    """Render every slide of a deck into one summary card.

    Returns (jpeg_bytes, (width, height), columns, slide_count).
    """
    import visual_qa

    images = visual_qa._render_deck(pres, None, None)
    card, size, columns = compose_summary_card(
        images, list(range(1, len(images) + 1)), title, columns)
    return card, size, columns, len(images)
