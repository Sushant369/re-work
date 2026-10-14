"""End-to-end pipeline.

  1. load + reading order      (Docling port: row-banded top/left sort, groups recursed)
  2. element extraction        (Docling port: title placeholder, paragraphs, inherited
                                bullets, tables, chart data, pictures, notes, furniture,
                                hidden content; + SmartArt text and effective font style)
  3. KPI guard                 (numeric call-outs and their labels never become headings)
  4. sub-heading candidates    (emphasis vs deck body text, must introduce content)
  5. deck-level consistency    (heading level = recurring visual style, ranked by size;
                                one-offs flagged, optional constrained LLM arbitration)
  6. reading order refinement  (heading-aware XY-cut: multi-column slides read column by
                                column, a heading stays with the content under it)
  7. sections across slides    (untitled or "(cont.)" slides continue the previous
                                section and keep its heading path)
  8. render                    (Markdown, per-section chunks, elements JSON for audit)
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

from .extract import ImageDescriber, extract
from .headings import Arbiter, HeadingConfig, classify
from .models import Deck
from .reading_order import refine_reading_order
from .render import to_markdown, to_section_chunks
from .sections import assign_sections


def run(
    pptx_path: str | Path,
    out_dir: str | Path,
    config: Optional[HeadingConfig] = None,
    describe_image: Optional[ImageDescriber] = None,
    arbitrate: Optional[Arbiter] = None,
    include_notes: bool = False,
    column_aware_order: bool = True,
) -> Deck:
    pptx_path, out_dir = Path(pptx_path), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    deck = extract(pptx_path, image_dir=out_dir / "images", describe_image=describe_image)
    classify(deck, config, arbitrate)
    if column_aware_order:
        refine_reading_order(deck)
    sections = assign_sections(deck)

    stem = pptx_path.stem
    (out_dir / f"{stem}.md").write_text(
        to_markdown(deck, include_notes=include_notes, image_root=out_dir), encoding="utf-8"
    )
    (out_dir / f"{stem}.elements.json").write_text(
        json.dumps([e.to_dict() for e in deck.elements], indent=1, ensure_ascii=False),
        encoding="utf-8",
    )
    (out_dir / f"{stem}.sections.json").write_text(
        json.dumps(to_section_chunks(sections, include_notes, out_dir), indent=1, ensure_ascii=False),
        encoding="utf-8",
    )
    return deck
