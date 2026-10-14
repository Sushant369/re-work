"""Stage 7: sections that span slides.

A deck's structure does not stop at slide boundaries. A slide continues the
previous section when it

* has no title at all (no title placeholder, and no top text that passes the
  fallback-title rules in headings.py), or
* repeats the previous section's title, with or without a continuation marker:
  "Risks (cont.)", "Risks - continued", "Risks (2/3)", "Risks, part 2".

A continuation slide keeps the heading path in effect at the end of the previous
slide (title > sub-heading), so its content stays under the right headings, and
its repeated title is marked ``continued`` instead of opening a second section.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

from .models import Deck, Element

_SEP = r"[\s\-–—:,|]*"
_CONT_WORD = re.compile(
    _SEP + r"""(?:
        [(\[]\s*(?:cont(?:inued|inue|'d|d|\.)?|contd\.?)\s*[)\]]   # (cont.) (continued) [contd]
      | \b(?:cont(?:inued|'d|d)?|contd)\.?                        # - continued / cont'd
    )\s*$""",
    re.IGNORECASE | re.VERBOSE,
)
# "(1/2)", "2 of 3", "part 2 of 3": page numbering, not part of a larger number
# ("Plan 2025/26" is a fiscal year, not page 25 of 26).
_PAGE = re.compile(
    _SEP + r"[(\[]?\s*(?:part\s*)?(?<![\w/.])(\d{1,2})\s*(?:/|of)\s*(\d{1,2})(?![\w/.])\s*[)\]]?\s*$",
    re.IGNORECASE,
)
_PART = re.compile(_SEP + r"[(\[]?\s*\bpart\s*\d{1,2}\s*[)\]]?\s*$", re.IGNORECASE)
# Deliberately not markers: a bare "(2)" or "2" ("Wave (1)" / "Wave (2)", "Phase 2")
# usually names a different topic.


def _strip_once(title: str) -> str:
    m = _PAGE.search(title)
    if m and 1 <= int(m.group(1)) <= int(m.group(2)) <= 20:
        return title[: m.start()]
    return _PART.sub("", _CONT_WORD.sub("", title))


def strip_continuation(title: str) -> str:
    """'Risks (cont.)' -> 'Risks'. Repeated so 'Risks (cont.) (2/3)' also strips."""
    prev = None
    title = title.strip()
    while prev != title:
        prev, title = title, _strip_once(title).strip()
    return title or prev


def title_key(title: str) -> str:
    return re.sub(r"\W+", " ", strip_continuation(title)).strip().lower()


@dataclass
class Section:
    title: Optional[str]  # None for content before the first title
    slides: list[int] = field(default_factory=list)
    elements: list[Element] = field(default_factory=list)


def assign_sections(deck: Deck, carry_subheadings: bool = True) -> list[Section]:
    """Group slides into sections and set ``section_path`` on every element.

    ``carry_subheadings``: on a continuation slide, keep the sub-heading that
    was open at the end of the previous slide (False keeps only the title).
    """
    sections: list[Section] = []
    current: Optional[Section] = None
    current_key: Optional[str] = None
    stack: list[str] = []

    for slide in range(1, deck.slide_count + 1):
        els = deck.slide_elements(slide)
        title = next((e for e in els if e.kind == "title" and e.layer == "body"), None)

        if title is not None and current is not None and title_key(title.text) == current_key:
            title.continued = True
        if title is None or title.continued:
            if current is None:
                current = Section(title=None)
                sections.append(current)
            if not carry_subheadings:
                stack = stack[:1]
        else:
            current = Section(title=strip_continuation(title.text))
            sections.append(current)
            current_key = title_key(title.text)
            stack = [current.title]

        current.slides.append(slide)
        for e in els:
            if e.kind == "heading":
                stack = stack[: max(e.level - 1, 0)] + [e.text]
            e.section_path = list(stack)
            current.elements.append(e)
    return sections
