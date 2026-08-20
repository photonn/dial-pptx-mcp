"""
Internal repair engine for the visual-QA loop (see visual_qa.py).

When inspection finds issues, the server — not the calling agent — repairs
the deck: the configured vision LLM is shown the issues, a structural
description of the affected slides, and their rendered images, and asked for
a plan made only of whitelisted, validated operations, which are applied
with python-pptx. The gate then re-inspects; see _visual_qa_gate in
tools/presentation_tools.py for the loop.

Whitelisted operations (anything else in a plan is skipped and reported):
- move_shape    {slide, shape_index, left_in, top_in}
- resize_shape  {slide, shape_index, width_in?, height_in?}
- set_font_size {slide, shape_index, size_pt}
- set_text      {slide, shape_index, text}
- set_word_wrap {slide, shape_index, wrap: true|false}
- delete_shape  {slide, shape_index}

Slide numbers are 1-based (matching the inspector's issue reports); shape
indexes are the 0-based positions reported by describe_slides.
"""
import json
import logging

from pptx.util import Emu, Inches, Pt

from logging_utils import get_logger, flatten

logger = get_logger("visual_fix")

MAX_TEXT_CHARS = 4000
FONT_PT_RANGE = (6, 96)
POSITION_IN_RANGE = (-5.0, 60.0)
SIZE_IN_RANGE = (0.05, 60.0)


def _emu_to_in(v):
    return round(Emu(v).inches, 3) if v is not None else None


def describe_slides(pres, slide_numbers=None):
    """Compact structural description of slides for the repair planner.
    slide_numbers: 1-based; None describes every slide."""
    described = []
    for idx, slide in enumerate(pres.slides, start=1):
        if slide_numbers and idx not in slide_numbers:
            continue
        shapes = []
        for s_idx, shape in enumerate(slide.shapes):
            info = {
                "shape_index": s_idx,
                "name": shape.name,
                "type": str(shape.shape_type),
                "left_in": _emu_to_in(shape.left),
                "top_in": _emu_to_in(shape.top),
                "width_in": _emu_to_in(shape.width),
                "height_in": _emu_to_in(shape.height),
                "is_placeholder": shape.is_placeholder,
            }
            if shape.has_text_frame:
                text = shape.text_frame.text
                info["text"] = text[:300] + ("…" if len(text) > 300 else "")
                sizes = sorted({
                    run.font.size.pt
                    for para in shape.text_frame.paragraphs
                    for run in para.runs
                    if run.font.size is not None
                })
                if sizes:
                    info["font_sizes_pt"] = sizes
            shapes.append(info)
        described.append({"slide": idx, "shapes": shapes})
    return described


REPAIR_PROMPT = """You are a presentation repair engine. A visual QA reviewer \
found the following issues in a PowerPoint deck:

{issues}

Below is the structural description of the affected slides (positions/sizes \
in inches, 1-based slide numbers, 0-based shape_index), followed by their \
rendered images.

{structure}

Slide dimensions: {width_in} x {height_in} inches.

Produce repair operations that fix the issues. Respond with ONLY a JSON \
object, no markdown fence:
{{"operations": [
  {{"op": "move_shape", "slide": N, "shape_index": N, "left_in": X, "top_in": Y}},
  {{"op": "resize_shape", "slide": N, "shape_index": N, "width_in": X, "height_in": Y}},
  {{"op": "set_font_size", "slide": N, "shape_index": N, "size_pt": X}},
  {{"op": "set_text", "slide": N, "shape_index": N, "text": "..."}},
  {{"op": "set_word_wrap", "slide": N, "shape_index": N, "wrap": true}},
  {{"op": "delete_shape", "slide": N, "shape_index": N}}
]}}
Only these operation types are allowed. Keep every shape fully inside the \
slide bounds. Prefer minimal changes: resize/move/shrink text before deleting \
anything. Do not touch shapes that have no reported issue."""


def plan_repairs(llm, issues, pres, deck_images):
    """Ask the vision LLM for a repair plan. Returns a list of operation
    dicts (possibly empty)."""
    slide_numbers = sorted({
        i.get("slide") for i in issues if isinstance(i.get("slide"), int)
    }) or None
    structure = describe_slides(pres, slide_numbers)
    prompt = REPAIR_PROMPT.format(
        issues=json.dumps(issues, indent=1),
        structure=json.dumps(structure, indent=1),
        width_in=_emu_to_in(pres.slide_width),
        height_in=_emu_to_in(pres.slide_height),
    )
    images = deck_images
    if slide_numbers:
        images = [deck_images[n - 1] for n in slide_numbers
                  if 0 < n <= len(deck_images)]
    logger.debug("repair_plan_request issues=%d slides=%s images=%d",
                 len(issues), slide_numbers or "all", len(images))
    data = llm.ask_json(images, prompt)
    ops = data.get("operations", [])
    if not isinstance(ops, list):
        logger.warning("repair_plan_malformed keys=%s",
                       ",".join(sorted(data)) or "-")
        return []
    logger.info("repair_plan_received operations=%d issues=%d",
                len(ops), len(issues))
    return ops


def _in_range(v, lo_hi):
    return isinstance(v, (int, float)) and lo_hi[0] <= v <= lo_hi[1]


def apply_repairs(pres, operations):
    """Validate and apply whitelisted operations. Returns
    {"applied": [...], "skipped": [{"op":..., "reason":...}]}."""
    applied, skipped = [], []
    slides = list(pres.slides)

    for op in operations if isinstance(operations, list) else []:
        if not isinstance(op, dict):
            skipped.append({"op": op, "reason": "not an object"})
            continue
        kind = op.get("op")
        slide_no, shape_idx = op.get("slide"), op.get("shape_index")
        if not (isinstance(slide_no, int) and 1 <= slide_no <= len(slides)):
            skipped.append({"op": op, "reason": "bad slide number"})
            continue
        shapes = list(slides[slide_no - 1].shapes)
        if not (isinstance(shape_idx, int) and 0 <= shape_idx < len(shapes)):
            skipped.append({"op": op, "reason": "bad shape_index"})
            continue
        shape = shapes[shape_idx]

        try:
            if kind == "move_shape":
                if not (_in_range(op.get("left_in"), POSITION_IN_RANGE)
                        and _in_range(op.get("top_in"), POSITION_IN_RANGE)):
                    skipped.append({"op": op, "reason": "position out of range"})
                    continue
                shape.left = Inches(op["left_in"])
                shape.top = Inches(op["top_in"])
            elif kind == "resize_shape":
                w, h = op.get("width_in"), op.get("height_in")
                if w is None and h is None:
                    skipped.append({"op": op, "reason": "no dimensions"})
                    continue
                if w is not None:
                    if not _in_range(w, SIZE_IN_RANGE):
                        skipped.append({"op": op, "reason": "width out of range"})
                        continue
                    shape.width = Inches(w)
                if h is not None:
                    if not _in_range(h, SIZE_IN_RANGE):
                        skipped.append({"op": op, "reason": "height out of range"})
                        continue
                    shape.height = Inches(h)
            elif kind == "set_font_size":
                if not _in_range(op.get("size_pt"), FONT_PT_RANGE):
                    skipped.append({"op": op, "reason": "font size out of range"})
                    continue
                if not shape.has_text_frame:
                    skipped.append({"op": op, "reason": "shape has no text frame"})
                    continue
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(op["size_pt"])
            elif kind == "set_text":
                text = op.get("text")
                if not isinstance(text, str) or len(text) > MAX_TEXT_CHARS:
                    skipped.append({"op": op, "reason": "bad text"})
                    continue
                if not shape.has_text_frame:
                    skipped.append({"op": op, "reason": "shape has no text frame"})
                    continue
                shape.text_frame.text = text
            elif kind == "set_word_wrap":
                if not shape.has_text_frame:
                    skipped.append({"op": op, "reason": "shape has no text frame"})
                    continue
                shape.text_frame.word_wrap = bool(op.get("wrap", True))
            elif kind == "delete_shape":
                shape._element.getparent().remove(shape._element)
            else:
                skipped.append({"op": op, "reason": f"unknown op '{kind}'"})
                continue
        except Exception as e:
            logger.debug("repair_op_failed op=%s slide=%s shape_index=%s error=%s",
                         kind, slide_no, shape_idx, flatten(str(e)))
            skipped.append({"op": op, "reason": f"apply failed: {e}"})
            continue
        logger.debug("repair_op_applied op=%s slide=%s shape_index=%s",
                     kind, slide_no, shape_idx)
        applied.append(op)

    level = logging.WARNING if skipped else logging.INFO
    logger.log(level, "repairs_applied applied=%d skipped=%d%s",
               len(applied), len(skipped),
               " reasons=" + ",".join(sorted({s["reason"].split(":")[0]
                                              for s in skipped})) if skipped else "")
    return {"applied": applied, "skipped": skipped}
