"""Stages 3-5: structure classification. Deterministic: the same deck always gives the same tree.

Docling only knows one heading: a TITLE / CENTER_TITLE placeholder. It never
looks at font size, which is why it never mistakes a KPI for a heading -- and
also why it never produces sub-headings. This module keeps Docling's rule as
the top level and adds sub-headings without a "largest font wins" rule:

3. KPI guard: numeric call-outs ("$4.2M", "35%", "+12% YoY") and the short
   labels next to them can never become headings, whatever their size.
4. Sub-heading candidates: short, emphasised text that introduces content
   below it, with emphasis measured against the deck's body text (or the rest
   of its own text box), not against the largest font on the slide.
5. Deck-level consistency: a heading level is a property of a *visual style*
   (size, bold), not of one paragraph. A candidate style must recur on several
   slides to become a heading; recurring styles are ranked by size into H2/H3,
   so the same style always gets the same level. One-off styles are flagged
   ``ambiguous`` and can be settled by an optional constrained LLM call
   (``arbitrate``).
"""

from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from statistics import median
from typing import Callable, Optional

from .models import EMU_PER_INCH, Deck, Element
from .ordering import vertical_gap, x_overlap_ratio


@dataclass
class HeadingConfig:
    max_level: int = 3  # deepest heading level emitted (title is 1)
    subtitle_as_heading: bool = True  # Docling treats subtitles as plain text
    size_ratio: float = 1.12  # "emphasised" = this much larger than the reference text
    strong_size_ratio: float = 1.3  # a one-off style this emphatic (and bold) is still a heading
    max_heading_words: int = 12
    max_heading_chars: int = 90
    max_title_words: int = 30  # fallback titles may be full sentences
    max_title_chars: int = 250
    min_support: Optional[int] = None  # slides a style must appear on; default 2 (1 for tiny decks)
    kpi_label_gap_in: float = 0.6  # a short text this close to a KPI is its label
    fallback_title_zone: float = 0.25  # top fraction of the slide searched when there is no title placeholder
    title_size_ratio: float = 1.5  # one-off top text this much larger than body text can be a fallback title


Arbiter = Callable[[list[dict]], dict[str, str]]

_KPI_TOKEN = re.compile(r"^[+\-−~≈<>]?[$€£¥₹]?\d[\d,.]*\s*(%|x|×|k|m|bn|mn|b|cr|pp|bps|pts)?\+?$", re.I)
# Years and reporting periods read as words, not figures: "2025 Strategy", "FY26 plan", "Q3".
_PERIOD_TOKEN = re.compile(
    r"\b(?:(?:19|20)\d{2}(?:[/\-–](?:\d{4}|\d{2}))?|[FC]Y\s?\d{2,4}|[QH][1-4])\b", re.I
)
_TRAILING_PUNCT = (".", ",", ";", "!")


def is_kpi(text: str) -> bool:
    """A numeric call-out: mostly digits, or a short text led by a *decorated* number.

    "$4.2M", "35%", "+12% YoY", "500+ clients" and "3x faster" are KPIs.
    "3 Key Risks", "100 Day Plan" and "2025 Strategy" are not: a bare integer
    followed by words, and years or periods, are ordinary heading text.
    """
    core = _PERIOD_TOKEN.sub(" ", text.strip())
    alnum = [c for c in core if c.isalnum()]
    digits = sum(c.isdigit() for c in alnum)
    if not digits:
        return False
    if digits / len(alnum) >= 0.5:
        return True
    tokens = core.split()
    return len(tokens) <= 3 and bool(_KPI_TOKEN.match(tokens[0])) and not tokens[0].isdigit()


def _body_baseline(elements: list[Element], exclude: set[int]) -> tuple[float, bool]:
    """Character-weighted median size and majority boldness of the deck's body text.

    ``exclude`` holds ids of elements that may be titles (top text on slides
    without a title placeholder); long sentence titles would otherwise pull the
    baseline up to their own size.
    """
    weighted: list[float] = []
    bold_chars = total = 0
    for e in elements:
        if id(e) in exclude:
            continue
        if e.kind in ("paragraph", "list_item") and e.font_size and not is_kpi(e.text):
            n = len(e.text)
            weighted.extend([e.font_size] * min(n, 500))
            bold_chars += n if e.bold else 0
            total += n
    if not weighted:
        return 18.0, False
    return median(weighted), bold_chars > total / 2


def _emphasised(e: Element, ref_size: float, ref_bold: bool, cfg: HeadingConfig) -> bool:
    size = e.font_size or 0
    return size >= ref_size * cfg.size_ratio or (e.bold and not ref_bold and size >= ref_size * 0.95)


def _signature(e: Element) -> tuple[int, bool]:
    """The style as designed: shrink-to-fit varies per slide, so it is undone here."""
    return (round((e.font_size or 0) / (e.font_scale or 1.0)), e.bold)


def classify(deck: Deck, cfg: Optional[HeadingConfig] = None, arbitrate: Optional[Arbiter] = None) -> Deck:
    cfg = cfg or HeadingConfig()
    body = [e for e in deck.elements if e.layer == "body"]
    zone = deck.slide_height * cfg.fallback_title_zone
    titled = {e.slide for e in body if e.kind == "title"}
    maybe_titles = {id(e) for e in body if e.slide not in titled and e.bbox[1] <= zone}
    baseline_size, baseline_bold = _body_baseline(body, maybe_titles)

    by_slide: dict[int, list[Element]] = defaultdict(list)
    by_shape: dict[str, list[Element]] = defaultdict(list)
    for e in body:
        by_slide[e.slide].append(e)
        if e.kind in ("paragraph", "list_item"):
            by_shape[e.shape_key].append(e)

    # -- stage 3: KPI guard
    kpis = [e for e in body if e.kind == "paragraph" and is_kpi(e.text)]
    for e in kpis:
        e.kind, e.reason = "kpi", "kpi:numeric"
    # A KPI has at most one label: short text in its own box, or else the
    # nearest short text just above/below it in its column. A larger-than-body
    # text that is not tight against the KPI is a heading under the KPI row.
    gap = cfg.kpi_label_gap_in * EMU_PER_INCH
    tight = EMU_PER_INCH * 0.25
    for k in kpis:
        own = [s for s in by_shape[k.shape_key] if s is not k and s.kind == "paragraph" and s.words <= 6]
        for s in own:
            s.reason = "kpi_label"
        if own:
            continue
        near = [
            e for e in by_slide[k.slide]
            if e.kind == "paragraph" and e.shape_key != k.shape_key and e.words <= 6
            and e.reason != "kpi_label"
            and vertical_gap(k.bbox, e.bbox) <= gap and x_overlap_ratio(k.bbox, e.bbox) >= 0.5
        ]
        if near:
            e = min(near, key=lambda e: vertical_gap(k.bbox, e.bbox))
            if vertical_gap(k.bbox, e.bbox) <= tight or (e.font_size or 0) < baseline_size * cfg.size_ratio:
                e.reason = "kpi_label"

    def heading_shaped(e: Element) -> bool:
        return (
            e.kind == "paragraph"
            and e.reason != "kpi_label"
            and 0 < e.words <= cfg.max_heading_words
            and len(e.text) <= cfg.max_heading_chars
            and not e.text.endswith(_TRAILING_PUNCT)
        )

    # -- stage 4: sub-heading candidates
    def introduces_content(e: Element, siblings: list[Element]) -> bool:
        if e.para_index < len(siblings) - 1:
            return True
        return any(
            o.shape_key != e.shape_key
            and o.kind not in ("title",)
            and o.bbox[1] >= e.bbox[3] - EMU_PER_INCH * 0.1
            and x_overlap_ratio(o.bbox, e.bbox) >= 0.3
            for o in by_slide[e.slide]
        )

    # -- structural: subtitle placeholder (Docling: plain text). A heading on a
    # title-style slide, or when it introduces content; otherwise (e.g. used as a
    # footnote under a table) it falls through to the candidate rules below.
    if cfg.subtitle_as_heading:
        for e in body:
            if e.kind != "paragraph" or e.reason != "placeholder:subtitle":
                continue
            others = [
                o for o in by_slide[e.slide]
                if o.kind != "title" and o.reason != "placeholder:subtitle" and o.layer == "body"
            ]
            if not others or introduces_content(e, by_shape[e.shape_key]):
                e.kind, e.level = "heading", 2

    candidates: list[Element] = []
    for e in body:
        if not heading_shaped(e) or (e.font_size or 0) < baseline_size * 0.9:
            continue
        siblings = by_shape[e.shape_key]
        if len(siblings) == 1:
            ref_size, ref_bold = baseline_size, baseline_bold
        else:
            # Compare against the rest of the text box; the next paragraph must be plainer.
            rest = [s for s in siblings if s is not e]
            ref_size = median(s.font_size or baseline_size for s in rest)
            ref_bold = all(s.bold for s in rest)
            nxt = next((s for s in siblings if s.para_index == e.para_index + 1), None)
            if nxt is None or _emphasised(nxt, ref_size, ref_bold, cfg):
                continue
        if _emphasised(e, ref_size, ref_bold, cfg) and introduces_content(e, siblings):
            candidates.append(e)

    # -- fallback title for slides without a title placeholder. Decided after the
    # candidates, so that a sub-heading at the top of a continuation slide is not
    # promoted: a style used as a sub-heading elsewhere in the deck is never a title.
    # A slide left without a title continues the previous slide's section (sections.py).
    # Titles are often full sentences ("Revenue grew 12% as renewals closed early."),
    # so they get looser length limits than sub-headings and may end with a period.
    def title_shaped(e: Element) -> bool:
        return (
            e.kind == "paragraph"
            and e.reason != "kpi_label"
            and 0 < e.words <= cfg.max_title_words
            and len(e.text) <= cfg.max_title_chars
        )

    picks: dict[int, Element] = {}
    for slide, els in by_slide.items():
        if slide in titled:
            continue
        pool = [
            e for e in els
            if title_shaped(e) and e.bbox[1] <= zone and _emphasised(e, baseline_size * 1.07, baseline_bold, cfg)
        ]
        if pool:
            picks[slide] = max(pool, key=lambda e: (e.font_size or 0, -e.bbox[1]))

    pick_ids = {id(p) for p in picks.values()}
    subheading_sigs = {_signature(c) for c in candidates if id(c) not in pick_ids}
    pick_support = Counter(_signature(p) for p in picks.values())
    title_sizes = [e.font_size for e in body if e.kind == "title" and e.font_size]
    for p in picks.values():
        sig, size = _signature(p), p.font_size or 0
        if sig in subheading_sigs:
            verdict = None  # a sub-heading at the top of a continuation slide
        elif pick_support[sig] >= 2:
            verdict = "fallback:recurring_title_style"
        elif title_sizes and size >= 0.85 * median(title_sizes):
            verdict = "fallback:title_sized"
        elif size >= cfg.title_size_ratio * baseline_size:
            verdict = "fallback:large_top_text"
        else:
            verdict = None
        if verdict:
            p.kind, p.level, p.reason = "title", 1, verdict
            if p in candidates:
                candidates.remove(p)

    # -- stage 5: deck-level style clustering
    min_support = cfg.min_support or (2 if deck.slide_count >= 3 else 1)
    slides_by_sig: dict[tuple[int, bool], set[int]] = defaultdict(set)
    for e in candidates:
        slides_by_sig[_signature(e)].add(e.slide)
    accepted = sorted(
        (sig for sig, slides in slides_by_sig.items() if len(slides) >= min_support),
        key=lambda sig: (-sig[0], not sig[1]),
    )
    level_of = {sig: min(2 + i, cfg.max_level) for i, sig in enumerate(accepted)}

    def level_for_unlisted(sig: tuple[int, bool]) -> int:
        bigger = sum(1 for a in accepted if (a[0], a[1]) > (sig[0], sig[1]))
        return min(2 + bigger, cfg.max_level)

    for e in candidates:
        sig = _signature(e)
        if sig in level_of:
            e.kind, e.level = "heading", level_of[sig]
            e.reason = f"style:{sig[0]}pt{'/b' if sig[1] else ''}@{len(slides_by_sig[sig])}slides"
        elif e.bold and (e.font_size or 0) >= baseline_size * cfg.strong_size_ratio:
            e.kind, e.level, e.reason = "heading", level_for_unlisted(sig), "style:strong_one_off"
        else:
            e.reason, e.ambiguous = "candidate:one_off_style", True

    if arbitrate:
        apply_arbitration(deck, arbitrate, baseline_size)
    return deck


# ------------------------------------------------------------- optional LLM step

ARBITRATION_PROMPT = """You label lines extracted from a PowerPoint deck.
For each item, answer with exactly one label: "heading2", "heading3" or "paragraph".
A heading introduces the content below it on the same slide. KPI values, KPI labels,
captions and diagram labels are "paragraph". Reply with JSON only: {"<id>": "<label>", ...}

Items:
"""


def arbitration_payload(deck: Deck, baseline_size: float) -> list[dict]:
    items = []
    for e in deck.elements:
        if not e.ambiguous:
            continue
        slide_els = deck.slide_elements(e.slide)
        title = next((s.text for s in slide_els if s.kind == "title"), "")
        after = [s.text[:80] for s in slide_els if s.order > e.order and s.text][:3]
        items.append({
            "id": e.id, "text": e.text, "slide_title": title, "font_pt": e.font_size,
            "body_font_pt": round(baseline_size, 1), "bold": e.bold, "followed_by": after,
        })
    return items


def apply_arbitration(deck: Deck, arbitrate: Arbiter, baseline_size: float) -> None:
    """``arbitrate`` receives the payload and returns ``{id: label}``; anything else is ignored."""
    payload = arbitration_payload(deck, baseline_size)
    if not payload:
        return
    labels = arbitrate(payload)
    by_id = {e.id: e for e in deck.elements}
    for el_id, label in labels.items():
        e = by_id.get(el_id)
        if e is None or not e.ambiguous:
            continue
        if label in ("heading2", "heading3"):
            e.kind, e.level, e.reason = "heading", int(label[-1]), "llm:" + label
        elif label == "paragraph":
            e.reason = "llm:paragraph"
        e.ambiguous = False


def arbitration_prompt(payload: list[dict]) -> str:
    return ARBITRATION_PROMPT + json.dumps(payload, ensure_ascii=False, indent=1)
