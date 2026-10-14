"""Stage 6: heading-aware reading order (runs after classification).

Docling's order (rows of shapes, left to right) interleaves multi-column
slides: "Heading A, Heading B, body A, body B". Once headings are known, a
recursive XY-cut over shapes fixes this:

* split the slide at the widest horizontal whitespace band (rows first),
  but never at a band directly below a heading, so a heading stays with the
  content under it;
* when no such band exists, split at the widest vertical band (columns);
* when nothing can be split (overlapping shapes), keep Docling's order.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass

from .models import EMU_PER_INCH, BBox, Deck, Element

_HEADING_TOL = int(EMU_PER_INCH * 0.15)


@dataclass
class _Unit:
    key: str
    bbox: BBox
    first_order: int
    heading_last: bool  # the shape ends with a heading, so what is below belongs to it


def _best_cut(units: list[_Unit], axis: int) -> tuple[int, float] | None:
    """Widest valid gap along ``axis`` (1 = y, 0 = x): ``(split position, gap)``."""
    lo, hi = axis, axis + 2
    ordered = sorted(units, key=lambda u: (u.bbox[lo], u.first_order))
    best = None
    reach = ordered[0].bbox[hi]
    upper_heading_bottoms = [ordered[0].bbox[3]] if ordered[0].heading_last else []
    for i in range(1, len(ordered)):
        start = ordered[i].bbox[lo]
        gap = start - reach
        blocked = axis == 1 and any(b >= reach - _HEADING_TOL for b in upper_heading_bottoms)
        if gap > 0 and not blocked and (best is None or gap > best[1]):
            best = (start, gap)
        reach = max(reach, ordered[i].bbox[hi])
        if ordered[i].heading_last:
            upper_heading_bottoms.append(ordered[i].bbox[3])
    return best


def _xy_cut(units: list[_Unit]) -> list[_Unit]:
    if len(units) <= 1:
        return units
    for axis in (1, 0):
        cut = _best_cut(units, axis)
        if cut is not None:
            pos = cut[0]
            before = [u for u in units if u.bbox[axis] < pos]
            after = [u for u in units if u.bbox[axis] >= pos]
            return _xy_cut(before) + _xy_cut(after)
    return sorted(units, key=lambda u: u.first_order)


def refine_reading_order(deck: Deck) -> Deck:
    out: list[Element] = []
    for slide in range(1, deck.slide_count + 1):
        els = deck.slide_elements(slide)
        positioned = [e for e in els if e.kind != "notes" and e.bbox[2] > e.bbox[0]]
        rest = [e for e in els if e not in positioned]

        by_key: dict[str, list[Element]] = defaultdict(list)
        for e in positioned:
            by_key[e.shape_key or e.id].append(e)
        units = []
        for key, group in by_key.items():
            group.sort(key=lambda e: e.order)
            units.append(
                _Unit(
                    key=key,
                    bbox=group[0].bbox,
                    first_order=group[0].order,
                    heading_last=group[-1].kind == "heading",
                )
            )
        for unit in _xy_cut(units):
            out.extend(by_key[unit.key])
        out.extend(sorted(rest, key=lambda e: e.order))

    for i, e in enumerate(out):
        e.order = i
    deck.elements = out
    return deck
