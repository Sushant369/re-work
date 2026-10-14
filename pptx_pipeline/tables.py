"""Native tables, charts and SmartArt to HTML.

Tables and charts are ports of Docling's ``_handle_tables`` and
``_chart_to_table_data`` / ``_build_chart_table``. SmartArt is an addition:
Docling skips SmartArt graphic frames, so their text is lost there.
"""

from __future__ import annotations

import html
import logging
import math
import re
from datetime import datetime, time, timedelta
from typing import Any, Optional

from lxml import etree
from pptx.chart.plot import PlotFactory, PlotTypeInspector
from pptx.oxml.ns import qn

from .ooxml import NS, SAFE_XML_PARSER

_log = logging.getLogger(__name__)

MAX_CHART_TABLE_CELLS = 100_000
_EXCEL_EPOCH = datetime(1899, 12, 30)
_EXCEL_1904_EPOCH = datetime(1904, 1, 1)
_NUMBER_FORMAT_LITERALS = re.compile(r'"[^"]*"|\\.|\[[^\]]*\]|[_*].')
SMARTART_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


def _cell_html(text: str) -> str:
    return html.escape(text.strip()).replace("\n", "<br>").replace("\v", "<br>")


def grid_to_html(rows: list[list[tuple[str, int, int]]], header_rows: int = 1) -> str:
    """``rows`` holds ``(text, rowspan, colspan)`` per origin cell (spanned cells omitted)."""
    out = ["<table>"]
    for r, row in enumerate(rows):
        tag = "th" if r < header_rows else "td"
        cells = []
        for text, rowspan, colspan in row:
            attrs = (f' rowspan="{rowspan}"' if rowspan > 1 else "") + (
                f' colspan="{colspan}"' if colspan > 1 else ""
            )
            cells.append(f"<{tag}{attrs}>{_cell_html(text)}</{tag}>")
        out.append("<tr>" + "".join(cells) + "</tr>")
    out.append("</table>")
    return "".join(out)


# ------------------------------------------------------------------------ tables


def table_to_html(shape) -> Optional[str]:
    """Docling treats row 0 as the column header and drops fully empty tables.

    Docling reads ``gridSpan``/``rowSpan`` from XML; python-pptx exposes the
    same via ``span_width``/``span_height``, and ``is_spanned`` lets merged-over
    cells be skipped even when they still hold stale hidden text.
    """
    table = shape.table
    rows: list[list[tuple[str, int, int]]] = []
    any_text = False
    for row in table.rows:
        out_row = []
        for cell in row.cells:
            if cell.is_spanned:
                continue
            text = cell.text
            any_text = any_text or bool(text.strip())
            rowspan = cell.span_height if cell.is_merge_origin else 1
            colspan = cell.span_width if cell.is_merge_origin else 1
            out_row.append((text, rowspan, colspan))
        rows.append(out_row)
    return grid_to_html(rows) if any_text else None


# ------------------------------------------------------------------------ charts


def _is_date_format(format_code: str) -> bool:
    codes = _NUMBER_FORMAT_LITERALS.sub("", format_code).lower()
    if "d" in codes or "y" in codes:
        return True
    return "m" in codes and "h" not in codes and "s" not in codes


def _serial_to_iso(text: str, date1904: bool) -> str:
    try:
        serial = float(text)
    except ValueError:
        return text
    if not math.isfinite(serial) or serial < 0:
        return text
    if date1904:
        epoch = _EXCEL_1904_EPOCH
    else:
        epoch = _EXCEL_EPOCH + timedelta(days=1 if serial < 60 else 0)
    try:
        moment = epoch + timedelta(seconds=round(serial * 86400))
    except OverflowError:
        return text
    if moment.time() == time(0):
        return moment.date().isoformat()
    return moment.isoformat(timespec="seconds")


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _category_axis(chart, x_chart):
    axis_ids = set(x_chart.xpath("./c:axId/@val"))
    for axis in chart.element.chart.plotArea.xpath("./c:catAx | ./c:dateAx"):
        if axis_ids.intersection(axis.xpath("./c:axId/@val")):
            return axis
    return None


def _categories_are_dates(axis, x_chart) -> bool:
    if axis.tag == qn("c:dateAx"):
        return True
    number_format = axis.find(qn("c:numFmt"))
    if number_format is not None and number_format.get("sourceLinked") not in ("1", "true"):
        format_code = number_format.get("formatCode", "")
    else:
        format_code = "".join(x_chart.xpath("./c:ser[1]/c:cat//c:numCache/c:formatCode/text()"))
    return _is_date_format(format_code)


def _axis_title_text(axis) -> str:
    paragraphs = [
        "".join(p.xpath(".//a:t/text()")).strip() for p in axis.xpath("./c:title/c:tx/c:rich/a:p")
    ]
    text = " ".join(p for p in paragraphs if p)
    if not text:
        text = "".join(axis.xpath("./c:title/c:tx/c:strRef//c:v/text()"))
    return text.strip()


def chart_to_html(shape) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """Return ``(title, chart_type, html_table)`` from the chart's cached data.

    Categories run down the first column and each series is a column, across
    every plot of a combination chart. Trailing empty rows are trimmed.
    """
    chart = shape.chart
    title = None
    if chart.has_title:
        title = chart.chart_title.text_frame.text.strip() or None

    plots = []
    for x_chart in chart.element.chart.plotArea.xCharts:
        try:
            plots.append((x_chart, PlotFactory(x_chart, chart)))
        except ValueError:
            plots.append((x_chart, None))

    chart_types = []
    for _, plot in plots:
        try:
            chart_types.append(PlotTypeInspector.chart_type(plot).name if plot else "OTHER")
        except (ValueError, NotImplementedError):
            chart_types.append("OTHER")
    chart_type = "+".join(dict.fromkeys(chart_types)) or None

    categories: list[str] = []
    corner = ""
    columns: list[tuple[str, list[str]]] = []
    for x_chart, plot in plots:
        if plot is None:
            _log.warning("Skipping unreadable chart plot %s", etree.QName(x_chart).localname)
            continue
        if not categories:
            axis = _category_axis(chart, x_chart)
            categories = [_cell_text(c) for c in plot.categories]
            if axis is not None:
                if _categories_are_dates(axis, x_chart):
                    date1904 = bool(chart.element.date_1904)
                    categories = [_serial_to_iso(c, date1904) for c in categories]
                corner = _axis_title_text(axis)
        for series in plot.series:
            name = _cell_text(series.name) if series.name is not None else ""
            columns.append((name, [_cell_text(v) for v in series.values]))

    if not columns:
        return title, chart_type, None
    num_rows = max([len(categories)] + [len(v) for _, v in columns])
    while num_rows > 0:  # trim trailing empty rows
        r = num_rows - 1
        if (r < len(categories) and categories[r]) or any(r < len(v) and v[r] for _, v in columns):
            break
        num_rows -= 1
    if num_rows == 0:
        return title, chart_type, None
    num_rows = min(num_rows, max(1, MAX_CHART_TABLE_CELLS // (1 + len(columns))))

    grid = [[(corner, 1, 1)] + [(name, 1, 1) for name, _ in columns]]
    for r in range(num_rows):
        row = [(categories[r] if r < len(categories) else "", 1, 1)]
        row += [((v[r] if r < len(v) else ""), 1, 1) for _, v in columns]
        grid.append(row)
    return title, chart_type, grid_to_html(grid)


# ---------------------------------------------------------------------- smartart


def is_smartart(shape) -> bool:
    return bool(shape._element.xpath(f'.//a:graphicData[@uri="{SMARTART_URI}"]'))


def smartart_texts(shape) -> list[str]:
    """Node texts from the SmartArt data model, in document order."""
    rel_ids = shape._element.xpath(".//dgm:relIds/@r:dm", namespaces=NS)
    if not rel_ids:
        return []
    try:
        part = shape.part.related_part(rel_ids[0])
        root = etree.fromstring(part.blob, SAFE_XML_PARSER)
    except Exception as exc:
        _log.warning("Could not read SmartArt data: %s", exc)
        return []
    texts = []
    for pt in root.iterfind(".//dgm:pt", NS):
        if pt.get("type", "node") != "node":
            continue
        for p in pt.iterfind("dgm:t/a:p", NS):
            text = "".join(p.itertext()).strip()
            if text:
                texts.append(text)
    return texts
