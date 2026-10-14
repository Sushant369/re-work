"""Stage 6: serialize classified elements to Markdown (headings, lists, HTML tables)."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

from .models import Deck, Element
from .sections import Section


def _block(e: Element, image_root: Optional[Path]) -> str:
    if e.kind in ("title", "heading"):
        return "#" * max(1, e.level) + " " + e.text
    if e.kind == "list_item":
        marker = f"{e.enum_index}." if e.enumerated else "-"
        return "  " * e.level + f"{marker} {e.text}"
    if e.kind == "table":
        return e.html or ""
    if e.kind == "chart":
        label = f"**Chart ({e.chart_type}){': ' + e.text if e.text else ''}**"
        return label + ("\n\n" + e.html if e.html else "")
    if e.kind == "picture":
        path = e.image_path or ""
        if path and image_root is not None:
            path = os.path.relpath(path, image_root)
        parts = [f"![{e.alt_text}]({path})"]
        if e.description:
            parts.append(e.description)
        return "\n\n".join(parts)
    if e.kind == "notes":
        return "> **Speaker notes:** " + e.text.replace("\n", "\n> ")
    return e.text  # paragraph, kpi


def _layers(include_notes: bool, include_furniture: bool, include_hidden: bool) -> set[str]:
    layers = {"body"}
    if include_notes:
        layers.add("notes")
    if include_furniture:
        layers.add("furniture")
    if include_hidden:
        layers.add("invisible")
    return layers


def _render_elements(els: list[Element], layers: set[str], image_root: Optional[Path]) -> list[str]:
    out: list[str] = []
    prev_list = False
    for e in els:
        if e.layer not in layers or (e.kind == "title" and e.continued):
            continue  # a repeated "(cont.)" title does not open a second section
        text = _block(e, image_root)
        if not text.strip():
            continue
        is_list = e.kind == "list_item"
        # consecutive list items stay together; everything else is its own block
        out.append(("\n" if is_list and prev_list else "\n\n") + text)
        prev_list = is_list
    return out


def to_markdown(
    deck: Deck,
    include_notes: bool = False,
    include_furniture: bool = False,
    include_hidden: bool = False,
    image_root: Optional[Path] = None,
) -> str:
    layers = _layers(include_notes, include_furniture, include_hidden)
    out: list[str] = []
    for slide in range(1, deck.slide_count + 1):
        els = deck.slide_elements(slide)
        marker = f"<!-- slide {slide}"
        starts_section = any(e.kind == "title" and not e.continued and e.layer == "body" for e in els)
        if els and not starts_section and els[0].section_path:
            marker += " | continues: " + " > ".join(els[0].section_path)
        out.append(marker + " -->")
        out.extend(_render_elements(els, layers, image_root))
        out.append("\n\n")
    return "".join(out).strip() + "\n"


def to_section_chunks(
    sections: list[Section],
    include_notes: bool = False,
    image_root: Optional[Path] = None,
) -> list[dict]:
    """One chunk per section, spanning all its slides; slide markers keep citations."""
    layers = _layers(include_notes, False, False)
    chunks = []
    for sec in sections:
        parts: list[str] = []
        for slide in sec.slides:
            rendered = _render_elements([e for e in sec.elements if e.slide == slide], layers, image_root)
            if rendered:
                parts.append(f"\n\n<!-- slide {slide} -->")
                parts.extend(rendered)
        body = "".join(parts).strip()
        if not body:
            continue
        chunks.append({"title": sec.title, "slides": sec.slides, "markdown": body})
    return chunks
