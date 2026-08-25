"""
Combo charts and per-series chart formatting.

python-pptx builds one chart group per chart: every series is the same type and
sits on the same value axis. The two things it therefore cannot express are the
two a business deck most often needs — bars with a line across them (actual vs.
target, volume vs. rate), and a second series whose units are nothing like the
first's (revenue in millions against a percentage), which needs a secondary
axis or renders as a flat line along the bottom.

Both are ordinary OOXML: a chart's plot area holds a *list* of chart-group
elements (`c:barChart`, `c:lineChart`, …), each naming the pair of axis ids it
plots against. So the deck is built the way python-pptx can — one chart with
every series and a correct embedded workbook — and the series are then
redistributed into the groups they belong in. The workbook, the caches and the
category references are untouched, which is what keeps the chart editable in
PowerPoint afterwards.

Rendering it as an image instead would lose all of that, and the visual repair
loop can only move or delete a picture.
"""
import random

from pptx.oxml.ns import nsdecls, qn
from pptx.oxml import parse_xml

from logging_utils import get_logger

logger = get_logger("utils.combo_chart")

__all__ = [
    "COMBO_SERIES_TYPES",
    "apply_combo_layout",
    "value_axis_with_id",
    "set_series_color",
    "set_series_data_labels",
    "add_series_trendline",
    "TRENDLINE_TYPES",
]

# The series types that can share one plot area. Pie and scatter cannot: pie
# has no axes to share, and scatter uses value axes on both sides.
COMBO_SERIES_TYPES = {
    "column": ("barChart", {"barDir": "col", "grouping": "clustered"}),
    "stacked_column": ("barChart", {"barDir": "col", "grouping": "stacked"}),
    "bar": ("barChart", {"barDir": "bar", "grouping": "clustered"}),
    "stacked_bar": ("barChart", {"barDir": "bar", "grouping": "stacked"}),
    "line": ("lineChart", {"grouping": "standard", "marker": "0"}),
    "line_markers": ("lineChart", {"grouping": "standard", "marker": "1"}),
    "area": ("areaChart", {"grouping": "standard"}),
    "stacked_area": ("areaChart", {"grouping": "stacked"}),
}

TRENDLINE_TYPES = ("linear", "movingAvg", "exp", "log", "poly", "power")

_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"

# Where each chart group's <c:axId> pair sits relative to its other children:
# the schema fixes the order, and PowerPoint rejects a group that breaks it.
_GROUP_TAILS = {
    "barChart": ('<c:gapWidth val="150"/>',),
    "lineChart": ('<c:marker val="{marker}"/>',),
    "areaChart": (),
}


def _c(tag):
    return f"{{{_C}}}{tag}"


def _new_ax_id(taken):
    # PowerPoint's own ids are large signed 32-bit values; any number works as
    # long as the group and its axes agree — and as long as it is not already
    # in use, since two axes sharing an id wires a series to the wrong scale.
    while True:
        candidate = random.randint(100000000, 2000000000)
        if candidate not in taken:
            taken.add(candidate)
            return candidate


def _element(xml):
    return parse_xml(xml.format(nsdecls=nsdecls("c", "a")))


def _group_element(kind, options, series, ax_ids):
    """Build one chart-group element holding the given <c:ser> elements."""
    parts = [f'<c:{kind} {nsdecls("c", "a")}>']
    if "barDir" in options:
        parts.append(f'<c:barDir val="{options["barDir"]}"/>')
    parts.append(f'<c:grouping val="{options["grouping"]}"/>')
    parts.append('<c:varyColors val="0"/>')
    parts.append("<!--series-->")
    for tail in _GROUP_TAILS[kind]:
        parts.append(tail.format(marker=options.get("marker", "0")))
    if kind == "barChart" and options["grouping"] == "stacked":
        # Stacked bars must overlap completely or they render as a clustered
        # chart with the segments side by side.
        parts.append('<c:overlap val="100"/>')
    parts.append(f'<c:axId val="{ax_ids[0]}"/>')
    parts.append(f'<c:axId val="{ax_ids[1]}"/>')
    parts.append(f"</c:{kind}>")

    group = parse_xml("".join(parts).replace("<!--series-->", ""))
    # Series go after grouping/varyColors and before the tail elements.
    anchor = group.find(_c("varyColors"))
    for element in series:
        anchor.addnext(element)
        anchor = element
    return group


def _secondary_axes(cat_id, val_id):
    """The hidden category axis and right-hand value axis a secondary group
    plots against. The category axis is deleted, not drawn: two sets of
    category labels on one chart is never what anyone wants."""
    cat = _element(
        '<c:catAx {nsdecls}>'
        f'<c:axId val="{cat_id}"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="1"/>'
        '<c:axPos val="b"/>'
        '<c:tickLblPos val="none"/>'
        f'<c:crossAx val="{val_id}"/>'
        '<c:crosses val="autoZero"/>'
        '<c:auto val="1"/>'
        '<c:lblAlgn val="ctr"/>'
        '<c:lblOffset val="100"/>'
        '<c:noMultiLvlLbl val="0"/>'
        '</c:catAx>')
    val = _element(
        '<c:valAx {nsdecls}>'
        f'<c:axId val="{val_id}"/>'
        '<c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/>'
        '<c:axPos val="r"/>'
        '<c:majorTickMark val="out"/>'
        '<c:minorTickMark val="none"/>'
        '<c:tickLblPos val="nextTo"/>'
        f'<c:crossAx val="{cat_id}"/>'
        '<c:crosses val="max"/>'
        '</c:valAx>')
    return cat, val


def apply_combo_layout(chart, series_specs):
    """Redistribute a chart's series into per-type, per-axis chart groups.

    `series_specs` is one dict per series, in the chart's own series order:
    ``{"type": "line", "secondary_axis": True}``. Series sharing a (type, axis)
    pair end up in one group, which is how PowerPoint models them.

    Returns ``{"groups": n, "primary": (cat_id, val_id), "secondary": ... }``.
    The axis ids matter to the caller: with two value axes present,
    python-pptx's ``chart.value_axis`` returns the *second* one (it assumes a
    scatter chart, where the first is the X axis), so anything wanting the
    left-hand axis has to find it by id.
    """
    plot_area = chart._chartSpace.find(_c("chart")).find(_c("plotArea"))
    existing = [child for child in plot_area
                if child.tag.endswith("Chart")]
    if not existing:
        raise ValueError("this chart has no plot group to rearrange")

    series = []
    for group in existing:
        for element in group.findall(_c("ser")):
            series.append(element)
            group.remove(element)
    if len(series) != len(series_specs):
        raise ValueError(
            f"the chart has {len(series)} series but {len(series_specs)} "
            f"series specification(s) were given")

    primary_cat, primary_val = _primary_axis_ids(existing[0])
    taken = {int(element.get("val"))
             for element in plot_area.iter(_c("axId"))}
    secondary_ids = None

    # Preserve the caller's series order, and keep one group per distinct
    # (type, axis) pair — merging non-adjacent series of the same kind, which
    # is what PowerPoint does too.
    groups = {}
    order = []
    for element, spec in zip(series, series_specs):
        key = (spec["type"], bool(spec.get("secondary_axis")))
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(element)

    new_groups = []
    for key in order:
        kind, options = COMBO_SERIES_TYPES[key[0]]
        if key[1]:
            if secondary_ids is None:
                secondary_ids = (_new_ax_id(taken), _new_ax_id(taken))
            ax_ids = secondary_ids
        else:
            ax_ids = (primary_cat, primary_val)
        new_groups.append(_group_element(kind, options, groups[key], ax_ids))

    first_axis = _first_axis(plot_area)
    for group in existing:
        plot_area.remove(group)
    for group in new_groups:
        if first_axis is None:
            plot_area.append(group)
        else:
            first_axis.addprevious(group)

    if secondary_ids is not None:
        cat, val = _secondary_axes(*secondary_ids)
        last_axis = _last_axis(plot_area)
        if last_axis is None:
            plot_area.append(cat)
            plot_area.append(val)
        else:
            last_axis.addnext(val)
            last_axis.addnext(cat)

    logger.debug("combo_layout_applied groups=%d secondary=%s",
                 len(new_groups), secondary_ids is not None)
    return {"groups": len(new_groups),
            "primary": (primary_cat, primary_val),
            "secondary": secondary_ids}


def _primary_axis_ids(group):
    ids = [int(element.get("val")) for element in group.findall(_c("axId"))]
    if len(ids) != 2:
        raise ValueError("the chart's plot group does not name an axis pair")
    return ids[0], ids[1]


_AXIS_TAGS = (_c("catAx"), _c("valAx"), _c("dateAx"), _c("serAx"))


def _first_axis(plot_area):
    for child in plot_area:
        if child.tag in _AXIS_TAGS:
            return child
    return None


def _last_axis(plot_area):
    found = None
    for child in plot_area:
        if child.tag in _AXIS_TAGS:
            found = child
    return found


def value_axis_with_id(chart, ax_id):
    """The ValueAxis whose <c:axId> is `ax_id`, or None.

    python-pptx's chart.value_axis cannot be trusted once a chart has two
    value axes — see apply_combo_layout — so axes are addressed by id.
    """
    from pptx.chart.axis import ValueAxis

    plot_area = chart._chartSpace.find(_c("chart")).find(_c("plotArea"))
    for axis in plot_area.findall(_c("valAx")):
        element = axis.find(_c("axId"))
        if element is not None and int(element.get("val")) == int(ax_id):
            return ValueAxis(axis)
    return None


# ---- Per-series formatting ----

def set_series_color(series, rgb):
    """Give one series an explicit solid fill.

    `rgb` is an (r, g, b) tuple. Chart palettes are the fastest way for a deck
    to stop matching its own template, so this exists to pin a series to a
    brand colour rather than whatever the theme's nth accent happens to be.
    """
    from pptx.dml.color import RGBColor

    fill = series.format.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(*rgb)


def set_series_data_labels(series, show, position=None, number_format=None):
    """Turn one series' data labels on or off, and place them.

    A stacked chart rejects the outside positions — the segment has no outside
    — so those are caught here rather than producing a chart PowerPoint
    refuses to open.
    """
    labels = series.data_labels
    labels.show_value = bool(show)
    if number_format:
        labels.number_format = number_format
        labels.number_format_is_linked = False
    if show and position:
        from pptx.enum.chart import XL_LABEL_POSITION

        labels.position = getattr(XL_LABEL_POSITION, position.upper())


def add_series_trendline(series, kind="linear", period=None):
    """Add a trendline to one series.

    python-pptx has no trendline API; the element is plain OOXML and belongs
    directly after the series' shape properties. Charting a computed extra
    series instead would put the fit in the legend and in the data, which is
    not the same thing.
    """
    if kind not in TRENDLINE_TYPES:
        raise ValueError(f"unknown trendline type: {kind}")

    xml = ['<c:trendline %s>' % nsdecls("c", "a"),
           f'<c:trendlineType val="{kind}"/>']
    if kind == "movingAvg":
        xml.append(f'<c:period val="{max(2, int(period or 2))}"/>')
    elif kind == "poly":
        xml.append(f'<c:order val="{max(2, min(6, int(period or 2)))}"/>')
    xml.append("</c:trendline>")

    element = parse_xml("".join(xml))
    ser = series._element
    # Schema order: idx, order, tx, spPr, ..., trendline, ..., cat, val.
    anchor = ser.find(_c("cat"))
    if anchor is None:
        anchor = ser.find(_c("val"))
    if anchor is not None:
        anchor.addprevious(element)
    else:
        ser.append(element)
    return element
