"""Build a deck that breaks "largest font = heading" rules."""

import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt


def textbox(slide, x, y, w, h, lines, group=None):
    """lines: [(text, size_pt or None, bold)]"""
    shapes = group.shapes if group is not None else slide.shapes
    tf = shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    for i, (text, size, bold) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = text
        if size:
            r.font.size = Pt(size)
        r.font.bold = bold
    return tf


def main(path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    title_only, title_content, title_slide = prs.slide_layouts[5], prs.slide_layouts[1], prs.slide_layouts[0]

    s = prs.slides.add_slide(title_slide)
    s.shapes.title.text = "Q3 Business Review"
    s.placeholders[1].text = "Finance & Operations"

    # KPI slide: 40pt numbers dwarf the 32pt-ish title and the 20pt subheading
    s = prs.slides.add_slide(title_only)
    s.shapes.title.text = "Quarter at a glance"
    for i, (val, label) in enumerate([("$4.2M", "Revenue"), ("35%", "Gross margin"), ("+12% YoY", "Growth")]):
        textbox(s, 0.8 + i * 4, 1.8, 3.5, 1.4, [(val, 40, True), (label, 14, False)])
    textbox(s, 0.8, 4.0, 11, 0.5, [("Key drivers", 20, True)])
    textbox(s, 0.8, 4.6, 11, 1.5, [("Enterprise renewals closed early.", 14, False), ("Pricing change in APAC lifted margin.", 14, False)])

    # Body placeholder with inherited sizes + bullets; subheadings in separate boxes
    s = prs.slides.add_slide(title_only)
    s.shapes.title.text = "Risks and mitigations"
    textbox(s, 0.8, 1.6, 5.5, 0.5, [("Key drivers", 20, True)])
    textbox(s, 0.8, 2.2, 5.5, 2, [("Supplier concentration in one region.", 14, False)])
    textbox(s, 7, 1.6, 5.5, 0.5, [("Mitigations", 20, True)])
    textbox(s, 7, 2.2, 5.5, 2, [("Dual-source critical parts by Q1.", 14, False)])
    textbox(s, 0.8, 6.8, 11, 0.4, [("Source: internal finance data, unaudited.", 10, False)])

    # Bullets inherited from the master (no explicit buChar on the paragraphs)
    s = prs.slides.add_slide(title_content)
    s.shapes.title.text = "Next steps"
    body = s.placeholders[1].text_frame
    body.text = "Finalise FY plan"
    for text, lvl in [("Budget sign-off", 1), ("Headcount", 1), ("Launch partner pilot", 0)]:
        p = body.add_paragraph()
        p.text, p.level = text, lvl

    # No title placeholder: title is a plain text box; grouped card with a heading
    s = prs.slides.add_slide(prs.slide_layouts[6])
    textbox(s, 0.8, 0.4, 11, 0.9, [("Regional performance", 32, False)])
    grp = s.shapes.add_group_shape()
    textbox(s, 0.8, 1.6, 5, 0.5, [("Mitigations", 20, True)], group=grp)
    textbox(s, 0.8, 2.2, 5, 1, [("Hedging FX exposure in LATAM.", 14, False)], group=grp)
    rows, cols = 3, 3
    tbl = s.shapes.add_table(rows, cols, Inches(0.8), Inches(3.6), Inches(6), Inches(1.5)).table
    for c, h in enumerate(["Region", "Q2", "Q3"]):
        tbl.cell(0, c).text = h
    tbl.cell(1, 0).text, tbl.cell(1, 1).text, tbl.cell(1, 2).text = "EMEA", "1.1", "1.3"
    tbl.cell(2, 0).text = "APAC (merged)"
    tbl.cell(2, 1).merge(tbl.cell(2, 2))
    tbl.cell(2, 1).text = "n/a"
    cd = CategoryChartData()
    cd.categories = ["Q1", "Q2", "Q3"]
    cd.add_series("Revenue", (3.1, 3.8, 4.2))
    chart = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(7.2), Inches(1.6), Inches(5.5), Inches(4), cd).chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Revenue by quarter ($M)"

    s.notes_slide.notes_text_frame.text = "Mention LATAM hedging."

    # Continuation without any title: starts with a sub-heading near the top
    s = prs.slides.add_slide(prs.slide_layouts[6])
    textbox(s, 0.8, 0.5, 11, 0.5, [("Mitigations", 20, True)])
    textbox(s, 0.8, 1.1, 11, 1, [("Renegotiate supplier terms in Brazil.", 14, False)])

    # Continuation with a repeated, marked title
    s = prs.slides.add_slide(title_only)
    s.shapes.title.text = "Regional performance (cont.)"
    textbox(s, 0.8, 1.6, 11, 1, [("APAC recovery expected in Q4.", 14, False)])

    # Continuation with body text only
    s = prs.slides.add_slide(prs.slide_layouts[6])
    textbox(s, 0.8, 0.8, 11, 1, [("EMEA flat; no change to guidance.", 14, False)])

    s = prs.slides.add_slide(title_only)
    s.shapes.title.text = "Appendix"
    textbox(s, 0.8, 1.6, 11, 1, [("Definitions and data sources.", 14, False)])
    prs.save(path)


def no_placeholder_deck(path):
    """Every slide uses the blank layout: titles are plain text boxes."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    s = prs.slides.add_slide(blank)
    textbox(s, 0.8, 0.3, 11, 0.8, [("Market overview", 28, False)])
    textbox(s, 0.8, 1.3, 11, 0.5, [("Highlights", 20, True)])
    textbox(s, 0.8, 1.9, 11, 1.5, [("Demand grew in all segments.", 14, False)])

    s = prs.slides.add_slide(blank)  # continues "Market overview > Highlights"
    textbox(s, 0.8, 0.4, 11, 0.5, [("Highlights", 20, True)])
    textbox(s, 0.8, 1.0, 11, 1.5, [("Two new competitors entered.", 14, False)])

    s = prs.slides.add_slide(blank)
    textbox(s, 0.8, 0.3, 11, 0.8, [("Outlook", 28, False)])
    textbox(s, 0.8, 1.3, 11, 1.5, [("Moderate growth through FY27.", 14, False)])
    prs.save(path)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "sample.pptx")
    no_placeholder_deck(sys.argv[2] if len(sys.argv) > 2 else "no_titles.pptx")
