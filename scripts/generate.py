#!/usr/bin/env python3
"""
PPT Master — Core generation engine.
Generates .pptx from outline.json.

Usage:
    python3 generate.py outline.json [output.pptx]
"""

import json
import sys
import os

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu, Cm
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
except ImportError:
    print("❌ python-pptx not installed. Run: pip install python-pptx")
    sys.exit(1)

# ─── Theme Definitions ─────────────────────────────────────
THEMES = {
    "blue": {
        "primary": RGBColor(0x1A, 0x3A, 0x5C),
        "accent": RGBColor(0x47, 0x89, 0xDB),
        "bg": RGBColor(0xF0, 0xF4, 0xF8),
        "text": RGBColor(0x2C, 0x3E, 0x50),
        "text_light": RGBColor(0x7F, 0x8C, 0x8D),
        "white": RGBColor(0xFF, 0xFF, 0xFF),
        "dark_bg": RGBColor(0x1A, 0x3A, 0x5C),
    },
    "red": {
        "primary": RGBColor(0x8B, 0x00, 0x00),
        "accent": RGBColor(0xE7, 0x4C, 0x3C),
        "bg": RGBColor(0xFD, 0xF2, 0xF0),
        "text": RGBColor(0x4A, 0x1A, 0x1A),
        "text_light": RGBColor(0x95, 0x5F, 0x5F),
        "white": RGBColor(0xFF, 0xFF, 0xFF),
        "dark_bg": RGBColor(0x6B, 0x00, 0x00),
    },
    "green": {
        "primary": RGBColor(0x00, 0x64, 0x3C),
        "accent": RGBColor(0x27, 0xAE, 0x60),
        "bg": RGBColor(0xE8, 0xF5, 0xE9),
        "text": RGBColor(0x1B, 0x3A, 0x2B),
        "text_light": RGBColor(0x61, 0x8C, 0x72),
        "white": RGBColor(0xFF, 0xFF, 0xFF),
        "dark_bg": RGBColor(0x00, 0x4D, 0x2E),
    },
    "purple": {
        "primary": RGBColor(0x4B, 0x00, 0x6E),
        "accent": RGBColor(0x8E, 0x44, 0xAD),
        "bg": RGBColor(0xF3, 0xE5, 0xF5),
        "text": RGBColor(0x2E, 0x0E, 0x3E),
        "text_light": RGBColor(0x7B, 0x5E, 0x8C),
        "white": RGBColor(0xFF, 0xFF, 0xFF),
        "dark_bg": RGBColor(0x38, 0x00, 0x55),
    },
    "orange": {
        "primary": RGBColor(0xE6, 0x5C, 0x00),
        "accent": RGBColor(0xF3, 0x9C, 0x12),
        "bg": RGBColor(0xFE, 0xF5, 0xE7),
        "text": RGBColor(0x4E, 0x34, 0x20),
        "text_light": RGBColor(0x99, 0x78, 0x54),
        "white": RGBColor(0xFF, 0xFF, 0xFF),
        "dark_bg": RGBColor(0xB6, 0x47, 0x00),
    },
    "dark": {
        "primary": RGBColor(0x1A, 0x1A, 0x2E),
        "accent": RGBColor(0x63, 0x66, 0xF1),
        "bg": RGBColor(0xF8, 0xF9, 0xFA),
        "text": RGBColor(0x1A, 0x1A, 0x2E),
        "text_light": RGBColor(0x6B, 0x72, 0x80),
        "white": RGBColor(0xFF, 0xFF, 0xFF),
        "dark_bg": RGBColor(0x11, 0x11, 0x22),
    },
    "teal": {
        "primary": RGBColor(0x00, 0x50, 0x64),
        "accent": RGBColor(0x00, 0x96, 0x88),
        "bg": RGBColor(0xE0, 0xF2, 0xF1),
        "text": RGBColor(0x15, 0x33, 0x3E),
        "text_light": RGBColor(0x54, 0x7A, 0x7D),
        "white": RGBColor(0xFF, 0xFF, 0xFF),
        "dark_bg": RGBColor(0x00, 0x3B, 0x4B),
    },
}

SLIDE_WIDTH = Inches(13.333)   # 16:9
SLIDE_HEIGHT = Inches(7.5)


# ─── Helpers ───────────────────────────────────────────────

def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_box(slide, left, top, width, height, text, size=18,
            bold=False, color=None, align=PP_ALIGN.LEFT, font="Microsoft YaHei"):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tb.text_frame.word_wrap = True
    p = tb.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color or RGBColor(0x2C, 0x3E, 0x50)
    p.font.name = font
    p.alignment = align
    return tb


def add_shape(slide, stype, left, top, width, height, fill=None, line_color=None, lw=None):
    s = slide.shapes.add_shape(stype, left, top, width, height)
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line_color:
        s.line.color.rgb = line_color
        if lw:
            s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    return s


def top_bar(slide, theme):
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0), Inches(0), SLIDE_WIDTH, Pt(4),
              fill=theme["accent"])


def slide_title(slide, theme, title):
    add_box(slide, Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.7),
            title, size=28, bold=True, color=theme["primary"])
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0.8), Inches(1.1), Inches(2), Pt(2),
              fill=theme["accent"])


# ─── Layouts ───────────────────────────────────────────────

def make_cover(prs, sd, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, fill=t["dark_bg"])
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(2), Inches(3.0), Inches(9.333), Pt(3), fill=t["accent"])
    add_box(s, Inches(2), Inches(1.5), Inches(9.333), Inches(1.5),
            sd.get("title", "PPT Title"), size=40, bold=True, color=t["white"], align=PP_ALIGN.CENTER)
    sub = sd.get("content", {}).get("subtitle", "")
    if sub:
        add_box(s, Inches(2), Inches(3.3), Inches(9.333), Inches(0.8),
                sub, size=20, color=t["accent"], align=PP_ALIGN.CENTER)
    author = sd.get("content", {}).get("author", "")
    company = sd.get("content", {}).get("company", "")
    info = " · ".join(filter(None, [author, company]))
    if info:
        add_box(s, Inches(2), Inches(4.5), Inches(9.333), Inches(0.5),
                info, size=14, color=t["text_light"], align=PP_ALIGN.CENTER)


def make_section_divider(prs, sd, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, fill=t["dark_bg"])
    num = sd.get("content", {}).get("number", "01")
    add_box(s, Inches(2), Inches(1.5), Inches(9.333), Inches(1.2),
            num, size=60, bold=True, color=t["accent"])
    add_box(s, Inches(2), Inches(3.0), Inches(9.333), Inches(1.0),
            sd.get("title", ""), size=32, bold=True, color=t["white"])
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(2), Inches(4.2), Inches(3), Pt(2), fill=t["accent"])


def make_text_only(prs, sd, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, t["white"])
    top_bar(s, t)
    slide_title(s, t, sd.get("title", ""))
    paras = sd.get("content", {}).get("paragraphs", [])
    y = Inches(1.5)
    for para in paras:
        add_box(s, Inches(0.8), y, Inches(11.733), Inches(0.6),
                para, size=16, color=t["text"])
        y += Inches(0.5)


def make_text_image(prs, sd, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, t["white"])
    top_bar(s, t)
    slide_title(s, t, sd.get("title", ""))
    text = sd.get("content", {}).get("text", "")
    add_box(s, Inches(0.8), Inches(1.5), Inches(6), Inches(5),
            text, size=16, color=t["text"])
    img = sd.get("content", {}).get("image_path", "")
    if img and os.path.exists(img):
        try:
            s.shapes.add_picture(img, Inches(7.5), Inches(1.5), Inches(5), Inches(5))
            return
        except Exception:
            pass
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(7.5), Inches(1.5), Inches(5), Inches(5),
              fill=t["bg"], line_color=t["text_light"], lw=1)
    add_box(s, Inches(7.5), Inches(3.8), Inches(5), Inches(0.5),
            "[ 插入图片 ]", size=14, color=t["text_light"], align=PP_ALIGN.CENTER)


def make_two_column(prs, sd, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, t["white"])
    top_bar(s, t)
    slide_title(s, t, sd.get("title", ""))
    cols = sd.get("content", {}).get("columns", [])
    cw = Inches(5.867)
    gap = Inches(0.5)
    for i, col in enumerate(cols):
        x = Inches(0.8) + i * (cw + gap)
        add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.5), cw, Inches(5.3), fill=t["bg"])
        add_box(s, x + Inches(0.3), Inches(1.7), cw - Inches(0.6), Inches(0.5),
                col.get("title", ""), size=18, bold=True, color=t["primary"], align=PP_ALIGN.CENTER)
        add_box(s, x + Inches(0.3), Inches(2.4), cw - Inches(0.6), Inches(4),
                col.get("text", ""), size=14, color=t["text"])


def make_data_chart(prs, sd, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, t["white"])
    top_bar(s, t)
    slide_title(s, t, sd.get("title", ""))
    cd = sd.get("content", {})
    ct = cd.get("chart_type", "bar")
    add_shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(1.5), Inches(7.5), Inches(5.3), fill=t["bg"])
    add_box(s, Inches(1), Inches(3.8), Inches(7.5), Inches(0.5),
            f"[ {ct.upper()} CHART ] — 在 PowerPoint 中编辑数据即可",
            size=16, color=t["text_light"], align=PP_ALIGN.CENTER)
    labels = cd.get("labels", [])
    y = Inches(1.7)
    for lbl in labels:
        add_box(s, Inches(1.3), y, Inches(3), Inches(0.3),
                f"\u25a0 {lbl}", size=12, color=t["text"])
        y += Inches(0.3)


def make_data_table(prs, sd, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, t["white"])
    top_bar(s, t)
    slide_title(s, t, sd.get("title", ""))
    td = sd.get("content", {}).get("table", [])
    if td:
        rows = len(td)
        cols = max(len(r) for r in td)
        tbl = s.shapes.add_table(rows, cols, Inches(1), Inches(1.5), Inches(11.333), Inches(5)).table
        cw = int(Inches(11.333) / cols)
        for i in range(cols):
            tbl.columns[i].width = cw
        for r, rd in enumerate(td):
            for c, ct in enumerate(rd):
                if c < cols:
                    cell = tbl.cell(r, c)
                    cell.text = str(ct)
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    for p in cell.text_frame.paragraphs:
                        p.font.size = Pt(12)
                        p.font.color.rgb = t["white"] if r == 0 else t["text"]
                        p.alignment = PP_ALIGN.CENTER
                    if r == 0:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = t["primary"]
                    else:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = t["white"] if r % 2 == 0 else t["bg"]


def make_timeline(prs, sd, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, t["white"])
    top_bar(s, t)
    slide_title(s, t, sd.get("title", ""))
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(1), Inches(3.7), Inches(11.333), Pt(3), fill=t["accent"])
    ms = sd.get("content", {}).get("milestones", [])
    n = len(ms)
    if n > 0:
        sp = Inches(11.333) / n
        for i, m in enumerate(ms):
            x = Inches(1) + sp * i + sp / 2 - Inches(0.3)
            add_shape(s, MSO_SHAPE.OVAL, x, Inches(3.5), Inches(0.4), Inches(0.4), fill=t["accent"])
            add_box(s, x - Inches(0.8), Inches(2.5), Inches(2), Inches(0.4),
                    m.get("date", ""), size=12, bold=True, color=t["primary"], align=PP_ALIGN.CENTER)
            add_box(s, x - Inches(0.8), Inches(4.2), Inches(2), Inches(1.5),
                    m.get("title", ""), size=11, color=t["text"], align=PP_ALIGN.CENTER)


def make_list(prs, sd, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, t["white"])
    top_bar(s, t)
    slide_title(s, t, sd.get("title", ""))
    items = sd.get("content", {}).get("items", [])
    y = Inches(1.5)
    for i, item in enumerate(items):
        c = add_shape(s, MSO_SHAPE.OVAL, Inches(0.8), y + Inches(0.05), Inches(0.4), Inches(0.4), fill=t["accent"])
        tf = c.text_frame
        tf.paragraphs[0].text = str(i + 1)
        tf.paragraphs[0].font.size = Pt(11)
        tf.paragraphs[0].font.color.rgb = t["white"]
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        title = item.get("title", "")
        desc = item.get("description", "")
        txt = title + ("\n" + desc if desc else "")
        add_box(s, Inches(1.4), y, Inches(11.2), Inches(0.6),
                txt, size=16, color=t["text"])
        y += Inches(0.7)


def make_contact(prs, sd, t):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(s, MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT, fill=t["dark_bg"])
    add_box(s, Inches(2), Inches(1.5), Inches(9.333), Inches(1),
            "Thank You", size=44, bold=True, color=t["white"], align=PP_ALIGN.CENTER)
    contact = sd.get("content", {}).get("contact", "")
    if contact:
        add_box(s, Inches(2), Inches(3.0), Inches(9.333), Inches(2),
                contact, size=18, color=t["accent"], align=PP_ALIGN.CENTER)


# ─── Router ────────────────────────────────────────────────

LAYOUTS = {
    "cover": make_cover,
    "section-divider": make_section_divider,
    "text-only": make_text_only,
    "text-image": make_text_image,
    "two-column": make_two_column,
    "data-chart": make_data_chart,
    "data-table": make_data_table,
    "timeline": make_timeline,
    "list": make_list,
    "contact": make_contact,
}


def generate(outline_path, output_path="presentation.pptx"):
    """Generate PPTX from outline JSON."""
    with open(outline_path, "r", encoding="utf-8") as f:
        outline = json.load(f)

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    theme_key = outline.get("presentation", {}).get("theme", "blue")
    theme = THEMES.get(theme_key, THEMES["blue"])

    for sd in outline.get("slides", []):
        layout = sd.get("layout", "text-only")
        func = LAYOUTS.get(layout)
        if func:
            func(prs, sd, theme)
        else:
            make_text_only(prs, sd, theme)

    prs.save(output_path)
    return output_path


if __name__ == "__main__":
    infile = sys.argv[1] if len(sys.argv) > 1 else "outline.json"
    outfile = sys.argv[2] if len(sys.argv) > 2 else "presentation.pptx"
    result = generate(infile, outfile)
    print(f"\u2705 PPTX generated: {result}")
