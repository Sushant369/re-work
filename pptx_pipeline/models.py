from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Optional

EMU_PER_INCH = 914400

BBox = tuple[int, int, int, int]  # left, top, right, bottom in slide EMUs (top-left origin)


@dataclass
class Element:
    """One extracted unit of content. Text is one element per paragraph, like Docling."""

    slide: int  # 1-based slide number
    kind: str  # title | heading | paragraph | list_item | kpi | table | chart | picture | notes
    text: str = ""
    order: int = 0  # reading order within the deck
    uid: int = 0  # extraction order; stable id across re-ordering
    level: int = 0  # heading level (1-3) or list nesting depth (0-based)
    bbox: BBox = (0, 0, 0, 0)
    layer: str = "body"  # body | furniture | notes | invisible  (Docling content layers)

    # Shape provenance
    shape_id: Optional[int] = None
    shape_name: str = ""
    shape_key: str = ""  # "<slide>:<shape_id>", groups paragraphs of one text frame
    placeholder: Optional[str] = None  # PP_PLACEHOLDER name, e.g. "TITLE", "BODY"
    para_index: int = 0  # index among the non-empty paragraphs of the shape
    para_count: int = 1

    # Lists
    enumerated: bool = False
    enum_index: Optional[int] = None

    # Effective (inherited) text style, used by the heading stage
    font_size: Optional[float] = None  # points, after autofit scaling
    bold: bool = False
    caps: bool = False

    # Non-text payloads
    html: Optional[str] = None  # tables and chart data
    chart_type: Optional[str] = None
    image_path: Optional[str] = None
    alt_text: str = ""
    description: str = ""  # filled by an image-description hook (e.g. a vision LLM)

    # Sections (sections.py): heading path in effect, carried across slides
    section_path: list[str] = field(default_factory=list)
    continued: bool = False  # a title repeating the previous section's ("Risks (cont.)")

    # Classification audit trail
    reason: str = ""
    ambiguous: bool = False

    @property
    def id(self) -> str:
        return f"s{self.slide}-e{self.uid}"

    @property
    def words(self) -> int:
        return len(self.text.split())

    def to_dict(self) -> dict:
        d = asdict(self)
        d["id"] = self.id
        return d


@dataclass
class Deck:
    source: str
    slide_width: int
    slide_height: int
    slide_count: int
    elements: list[Element] = field(default_factory=list)

    def slide_elements(self, slide: int) -> list[Element]:
        return [e for e in self.elements if e.slide == slide]
