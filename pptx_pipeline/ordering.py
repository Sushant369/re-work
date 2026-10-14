"""Reading order and geometry.

``iter_shapes_by_position`` is a port of Docling's ``_iter_shapes_by_position``.
``Transform`` is an addition: python-pptx reports a grouped shape's position in
the group's *child* coordinate space, which Docling uses as-is. The heading
stage compares positions across shapes, so group offsets are mapped back to
slide coordinates here.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Iterator, Optional

from .models import BBox
from .ooxml import NS

# 0.05 inch: roughly half a title line-height. Shapes whose tops are within this
# of the previous shape's top are one visual row, read left to right.
SHAPE_ROW_TOLERANCE_EMU = 45720


def _position(shape, attr: str) -> Optional[int]:
    try:
        value = shape.top if attr == "top" else shape.left
    except (AttributeError, ValueError, TypeError):
        return None
    return None if value is None else int(value)


def iter_shapes_by_position(shapes: Iterable) -> Iterator:
    """Yield shapes top-to-bottom, then left-to-right within a row.

    Rows chain: a shape joins the current row when its top is within the
    tolerance of the *previous* shape's top. Shapes without a position go last
    in their original z-order.
    """
    fallback = 2**63 - 1
    entries = []
    for index, shape in enumerate(shapes):
        top = _position(shape, "top")
        left = _position(shape, "left")
        entries.append(
            (index, shape, top if top is not None else fallback, left if left is not None else fallback)
        )
    entries.sort(key=lambda e: (e[2], e[0]))

    rows: list[list] = []
    current: list = []
    prev_top = None
    for entry in entries:
        if prev_top is None or entry[2] - prev_top <= SHAPE_ROW_TOLERANCE_EMU:
            current.append(entry)
        else:
            rows.append(current)
            current = [entry]
        prev_top = entry[2]
    if current:
        rows.append(current)

    for row in rows:
        for entry in sorted(row, key=lambda e: (e[3], e[0])):
            yield entry[1]


@dataclass(frozen=True)
class Transform:
    """Affine map from a (possibly nested) group's child space to slide EMUs."""

    sx: float = 1.0
    sy: float = 1.0
    dx: float = 0.0
    dy: float = 0.0

    def apply(self, x: float, y: float) -> tuple[int, int]:
        return round(self.dx + self.sx * x), round(self.dy + self.sy * y)

    def into_group(self, group_shape) -> "Transform":
        xfrm = group_shape._element.find("p:grpSpPr/a:xfrm", NS)
        if xfrm is None:
            return self
        off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
        ch_off, ch_ext = xfrm.find("a:chOff", NS), xfrm.find("a:chExt", NS)
        if None in (off, ext, ch_off, ch_ext):
            return self
        cx, cy = int(ch_ext.get("cx", 0)), int(ch_ext.get("cy", 0))
        gsx = int(ext.get("cx", 0)) / cx if cx else 1.0
        gsy = int(ext.get("cy", 0)) / cy if cy else 1.0
        ox, oy = int(off.get("x", 0)), int(off.get("y", 0))
        cox, coy = int(ch_off.get("x", 0)), int(ch_off.get("y", 0))
        return Transform(
            sx=self.sx * gsx,
            sy=self.sy * gsy,
            dx=self.dx + self.sx * (ox - cox * gsx),
            dy=self.dy + self.sy * (oy - coy * gsy),
        )

    def bbox(self, shape, slide_w: int, slide_h: int) -> BBox:
        try:
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
        except (AttributeError, ValueError, TypeError):
            left = top = width = height = None
        if None in (left, top, width, height):
            return (0, 0, slide_w, slide_h)  # Docling's fallback: the whole slide
        l, t = self.apply(left, top)
        r, b = self.apply(left + width, top + height)
        return (l, t, r, b)


def x_overlap_ratio(a: BBox, b: BBox) -> float:
    overlap = min(a[2], b[2]) - max(a[0], b[0])
    narrower = min(a[2] - a[0], b[2] - b[0])
    return max(0.0, overlap) / narrower if narrower > 0 else 0.0


def vertical_gap(a: BBox, b: BBox) -> int:
    return max(a[1] - b[3], b[1] - a[3], 0)
