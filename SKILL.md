---
name: ppt-master
description: >
  AI-powered PowerPoint slide generator that creates professional .pptx presentations from
  user input. Supports business pitches, product launches, work reports, annual summaries,
  project reviews and more. Built-in layout library with 12 slide types, 8 color schemes,
  and interactive editing workflow.
  Use when: (1) User asks to create a PowerPoint/PPT/slides, (2) User wants to generate a
  business presentation, (3) User mentions 路演PPT, 产品介绍, 年终总结, 述职报告, 项目汇报,
  or commercial slides, (4) User has content (text/charts/images) and wants presentation output.
---

# PPT Master

Generate professional, real `.pptx` PowerPoint files from user input — no PowerPoint needed.
Built with python-pptx. Works on any platform with Python 3.

## Pipeline Overview

```
需求收集 → 大纲确认 → 逐页生成 → PPTX导出 → 交互修改循环
```

## Phase 0: Environment Check

### 0.1 Check python-pptx

```bash
python3 -c "from pptx import Presentation; print('OK')"
```

If missing, install:
```bash
pip install python-pptx
```

If pip is slow, use mirror:
```bash
pip install python-pptx -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 0.2 Check Dependencies

This skill uses only python-pptx (pure Python, no native deps). Verify it works:

```bash
python3 -c "
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
print('All imports OK')
"
```

## Phase 1: Understand User's Needs

### 1.1 Gather Requirements (In-Conversation)

Ask user these questions **one at a time** in natural conversation:

1. **场景**: 这是什么类型的 PPT？
   - 💼 商业路演 / 融资计划
   - 🚀 产品发布 / 新品介绍
   - 📊 工作汇报 / 年度总结
   - 📈 项目复盘 / 数据分析
   - 🏢 企业介绍 / 公司宣传
   - 🎓 学术报告 / 开题答辩 (用 beamer-academic？)
   - ✏️ 其他（请描述）

2. **核心信息**: 想传达什么核心信息？（一句话）

3. **目标受众**: 听众是谁？
   - 投资人 / 高管
   - 客户 / 合作伙伴
   - 内部团队
   - 公众 / 行业大会

4. **内容材料**: 有现成的文档/数据吗？
   - 文字稿 / 大纲
   - 图片 / 截图
   - 数据表格 / CSV
   - 还没有，需要我从头写

5. **页数**: 目标页数？（建议 10-20 页）

6. **色调偏好**: 想要什么风格的配色？
   - 🔵 专业蓝（科技/金融）
   - 🔴 热情红（电商/零售）
   - 🟢 自然绿（环保/农业）
   - 🟣 优雅紫（美妆/文创）
   - 🟠 活力橙（互联网/教育）
   - ⚫ 简约黑（高端/设计）
   - 🔷 深海蓝（医疗/科研）
   - 🎨 自定义

7. **公司/品牌元素**: 需要加 logo 或品牌色吗？

### 1.2 Collect Content

If user has existing content (Word doc, notes, outline):

1. Ask them to paste the content or point to a file
2. Read the content and extract key sections
3. Identify which parts go on which slides

If user has no content:
1. Propose an outline based on the presentation type
2. Fill in details together in conversation

## Phase 2: Brainstorm Outline (Interactive, In-Conversation)

### 2.1 Propose High-Level Structure

Based on user's answers, propose a slide structure directly in conversation:

**Example for 产品发布**:
> ## 📐 PPT 结构建议
>
> | # | 版式 | 内容 |
> |---|------|------|
> | 1 | 封面 | 产品名 + 标语 + 公司 |
> | 2 | 痛点 | 用户面临什么问题 |
> | 3 | 方案 | 我们的产品如何解决 |
> | 4 | 核心功能 | 3大亮点功能介绍 |
> | 5 | 数据对比 | 与传统方案对比 |
> | 6 | 案例 | 客户成功案例 |
> | 7 | 团队 | 核心团队介绍 |
> | 8 | 融资 | 融资计划与用途 |
> | 9 | 联系方式 | CTA + 二维码 |
>
> **你觉得这个结构可以吗？** 可以说：
> - "加一页市场分析"
> - "功能页拆成两页"
> - "去掉案例页"

**Example for 年终总结**:
> ## 📐 PPT 结构建议
>
> | # | 版式 | 内容 |
> |---|------|------|
> | 1 | 封面 | 年度总结 + 姓名 + 部门 |
> | 2 | 目录 | 工作概览目录 |
> | 3 | KPI总览 | 关键指标仪表盘 |
> | 4 | 重点项目 | 重点成果展示 |
> | 5 | 数据趋势 | 月度数据趋势 |
> | 6 | 团队建设 | 团队成长 |
> | 7 | 经验反思 | 不足与改进 |
> | 8 | Q1计划 | 下季度规划 |
> | 9 | 致谢 | 感谢 |
>
> **你觉得这个结构可以吗？**

Loop until user approves.

### 2.2 Per-Slide Detail

Once structure is approved, go through each slide **one at a time** in conversation:

> ## 📋 第 3 页 - 核心功能
>
> **版式**: 三栏功能展示
>
> **内容要点**:
> - 左栏: 智能分析 — AI自动识别数据模式
> - 中栏: 实时监控 — 7×24小时系统监控
> - 右栏: 一键报表 — 自动生成周报月报
>
> **这页可以吗？** 可以说：
> - "第三个改成'多维度对比'"
> - "加一张截图"
> - "ok，继续"

**HARD GATE**: Do NOT proceed to Phase 3 until user explicitly confirms the complete outline.

### 2.3 Save Outline

After final confirmation, save the approved structure to `outline.json`:

```json
{
  "presentation": {
    "type": "product_launch",
    "title": "AI数据分析平台 - 产品发布会",
    "author": "张三",
    "company": "智数科技",
    "theme": "blue",
    "total_slides": 9,
    "aspect_ratio": "16x9"
  },
  "slides": [
    {
      "id": 1,
      "layout": "cover",
      "title": "AI数据分析平台",
      "content": { "subtitle": "让数据说话", "company": "智数科技" }
    },
    {
      "id": 2,
      "layout": "pain-points",
      "title": "企业数据之痛",
      "content": { "items": ["数据分散", "分析耗时", "决策滞后"] }
    }
  ]
}
```

### Layout Selection Rules

1. **cover** — 封面页，第一页
2. **toc** — 目录页，展示结构
3. **section-divider** — 章节过渡，视觉节奏
4. **text-only** — 纯文字叙述，背景/概念
5. **text-image** — 左文右图，图文配合
6. **image-text** — 左图右文，图片为主
7. **two-column** — 双栏对比/并列
8. **three-column** — 三栏特性展示
9. **data-chart** — 数据图表（柱状/折线/饼图）
10. **data-table** — 数据表格
11. **timeline** — 时间线/里程碑
12. **quote** — 大号引用/金句
13. **list** — 列表/要点
14. **team** — 团队介绍
15. **contact** — 联系方式/CTA

Avoid 3 consecutive slides with the same layout. Vary between text-heavy and visual slides.

## Phase 3: Generate PPTX

### 3.1 Create Generator Script

Create a Python script in the user's working directory (e.g. `generate_ppt.py`) that uses
python-pptx to build the presentation. The script's core logic:

1. Create a `Presentation` object with proper slide dimensions
2. Apply theme (colors, fonts) from theme definitions
3. For each slide in `outline.json`:
   - Select the correct layout function
   - Populate content
4. Save as `.pptx`

**IMPORTANT**: Generate the actual Python script on disk, then run it. Do NOT try to generate
the PPTX directly through AI action (can't handle binary file output).

### 3.2 Slide Generation Template

Each slide type has a corresponding Python function. Here are the core implementations:

```python
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import json, os

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

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

def set_slide_bg(slide, color):
    """Set solid background color for a slide."""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_textbox(slide, left, top, width, height, text, font_size=18,
                bold=False, color=None, alignment=PP_ALIGN.LEFT, font_name="Microsoft YaHei"):
    """Add a text box to a slide."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color or RGBColor(0x2C, 0x3E, 0x50)
    p.alignment = alignment
    return txBox

def add_shape(slide, shape_type, left, top, width, height, fill_color=None,
              line_color=None, line_width=None):
    """Add an auto shape."""
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
        if line_width:
            shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape

# ─── Layout Functions ──────────────────────────────────────

def make_cover(prs, slide_data, theme):
    """Cover slide — dark background with centered title."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    
    # Full background
    bg_shape = add_shape(slide, MSO_SHAPE.RECTANGLE,
                         Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT,
                         fill_color=theme["dark_bg"])
    bg_shape.line.fill.background()
    
    # Accent line
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(2), Inches(3.0), Inches(9.333), Pt(3),
              fill_color=theme["accent"])
    
    # Title
    add_textbox(slide, Inches(2), Inches(1.5), Inches(9.333), Inches(1.5),
                slide_data.get("title", "PPT Title"),
                font_size=40, bold=True, color=theme["white"],
                alignment=PP_ALIGN.CENTER)
    
    # Subtitle
    add_textbox(slide, Inches(2), Inches(3.3), Inches(9.333), Inches(0.8),
                slide_data.get("content", {}).get("subtitle", ""),
                font_size=20, color=theme["accent"],
                alignment=PP_ALIGN.CENTER)
    
    # Author/Company
    author = slide_data.get("content", {}).get("author", "")
    company = slide_data.get("content", {}).get("company", "")
    info = " · ".join(filter(None, [author, company]))
    if info:
        add_textbox(slide, Inches(2), Inches(4.5), Inches(9.333), Inches(0.5),
                    info, font_size=14, color=theme["text_light"],
                    alignment=PP_ALIGN.CENTER)

def make_section_divider(prs, slide_data, theme):
    """Section divider — full color with number and title."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT,
              fill_color=theme["dark_bg"])
    
    # Section number
    num = slide_data.get("content", {}).get("number", "01")
    add_textbox(slide, Inches(2), Inches(1.5), Inches(9.333), Inches(1.2),
                num, font_size=60, bold=True, color=theme["accent"],
                alignment=PP_ALIGN.LEFT)
    
    # Section title
    add_textbox(slide, Inches(2), Inches(3.0), Inches(9.333), Inches(1.0),
                slide_data.get("title", ""),
                font_size=32, bold=True, color=theme["white"],
                alignment=PP_ALIGN.LEFT)
    
    # Divider line
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(2), Inches(4.2), Inches(3), Pt(2),
              fill_color=theme["accent"])

def make_text_only(prs, slide_data, theme):
    """Text content slide — clean, professional."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["white"])
    
    # Top accent bar
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0), Inches(0), SLIDE_WIDTH, Pt(4),
              fill_color=theme["accent"])
    
    # Title
    add_textbox(slide, Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.7),
                slide_data.get("title", ""),
                font_size=28, bold=True, color=theme["primary"])
    
    # Divider under title
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0.8), Inches(1.1), Inches(2), Pt(2),
              fill_color=theme["accent"])
    
    # Content paragraphs
    paras = slide_data.get("content", {}).get("paragraphs", [])
    y = Inches(1.5)
    for para in paras:
        tb = add_textbox(slide, Inches(0.8), y, Inches(11.733), Inches(0.6),
                         para, font_size=16, color=theme["text"])
        y += Inches(0.5)

def make_text_image(prs, slide_data, theme):
    """Text + Image (left text, right image area)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["white"])
    
    # Top accent
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0), Inches(0), SLIDE_WIDTH, Pt(4),
              fill_color=theme["accent"])
    
    # Title
    add_textbox(slide, Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.7),
                slide_data.get("title", ""),
                font_size=28, bold=True, color=theme["primary"])
    
    # Left text
    text = slide_data.get("content", {}).get("text", "")
    add_textbox(slide, Inches(0.8), Inches(1.5), Inches(6), Inches(5),
                text, font_size=16, color=theme["text"])
    
    # Right image placeholder (gray rect)
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(7.5), Inches(1.5), Inches(5), Inches(5),
              fill_color=theme["bg"],
              line_color=theme["text_light"], line_width=1)
    
    # Image label
    add_textbox(slide, Inches(7.5), Inches(3.8), Inches(5), Inches(0.5),
                "[ 插入图片 ]", font_size=14, color=theme["text_light"],
                alignment=PP_ALIGN.CENTER)

def make_two_column(prs, slide_data, theme):
    """Two-column content slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["white"])
    
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0), Inches(0), SLIDE_WIDTH, Pt(4),
              fill_color=theme["accent"])
    
    add_textbox(slide, Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.7),
                slide_data.get("title", ""),
                font_size=28, bold=True, color=theme["primary"])
    
    cols = slide_data.get("content", {}).get("columns", [])
    col_w = Inches(5.867)
    gap = Inches(0.5)
    start_x = Inches(0.8)
    
    for i, col in enumerate(cols):
        x = start_x + i * (col_w + gap)
        
        # Column card bg
        card = add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                         x, Inches(1.5), col_w, Inches(5.3),
                         fill_color=theme["bg"])
        card.line.fill.background()
        
        # Column title
        add_textbox(slide, x + Inches(0.3), Inches(1.7), col_w - Inches(0.6), Inches(0.5),
                    col.get("title", ""),
                    font_size=18, bold=True, color=theme["primary"],
                    alignment=PP_ALIGN.CENTER)
        
        # Column text
        add_textbox(slide, x + Inches(0.3), Inches(2.4), col_w - Inches(0.6), Inches(4),
                    col.get("text", ""),
                    font_size=14, color=theme["text"])

def make_data_chart(prs, slide_data, theme):
    """Data chart slide with placeholder for chart."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["white"])
    
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0), Inches(0), SLIDE_WIDTH, Pt(4),
              fill_color=theme["accent"])
    
    add_textbox(slide, Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.7),
                slide_data.get("title", ""),
                font_size=28, bold=True, color=theme["primary"])
    
    # Chart area placeholder
    chart_data = slide_data.get("content", {})
    chart_type = chart_data.get("chart_type", "bar")
    
    # Chart background
    add_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
              Inches(1), Inches(1.5), Inches(7.5), Inches(5.3),
              fill_color=theme["bg"])
    
    # Chart label
    add_textbox(slide, Inches(1), Inches(3.8), Inches(7.5), Inches(0.5),
                f"[ {chart_type.upper()} CHART ] — 在 PowerPoint 中编辑数据即可",
                font_size=16, color=theme["text_light"],
                alignment=PP_ALIGN.CENTER)
    
    # Legend placeholder
    labels = chart_data.get("labels", [])
    y = Inches(1.7)
    for lbl in labels:
        add_textbox(slide, Inches(1.3), y, Inches(3), Inches(0.3),
                    f"■ {lbl}", font_size=12, color=theme["text"])
        y += Inches(0.3)

def make_data_table(prs, slide_data, theme):
    """Table slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["white"])
    
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0), Inches(0), SLIDE_WIDTH, Pt(4),
              fill_color=theme["accent"])
    
    add_textbox(slide, Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.7),
                slide_data.get("title", ""),
                font_size=28, bold=True, color=theme["primary"])
    
    table_data = slide_data.get("content", {}).get("table", [])
    if table_data:
        rows = len(table_data)
        cols = max(len(r) for r in table_data) if table_data else 3
        left = Inches(1)
        top = Inches(1.5)
        width = Inches(11.333)
        height = Inches(5)
        
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table
        
        # Set column widths
        col_width = int(width / cols)
        for i in range(cols):
            table.columns[i].width = col_width
        
        for r, row_data in enumerate(table_data):
            for c, cell_text in enumerate(row_data):
                if c < cols:
                    cell = table.cell(r, c)
                    cell.text = str(cell_text)
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.size = Pt(12)
                        paragraph.font.color.rgb = theme["white"] if r == 0 else theme["text"]
                        paragraph.alignment = PP_ALIGN.CENTER
                    # Header row
                    if r == 0:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = theme["primary"]
                    else:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = theme["white"] if r % 2 == 0 else theme["bg"]

def make_list(prs, slide_data, theme):
    """Bullet list / numbered items."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["white"])
    
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0), Inches(0), SLIDE_WIDTH, Pt(4),
              fill_color=theme["accent"])
    
    add_textbox(slide, Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.7),
                slide_data.get("title", ""),
                font_size=28, bold=True, color=theme["primary"])
    
    items = slide_data.get("content", {}).get("items", [])
    y = Inches(1.5)
    for i, item in enumerate(items):
        # Number circle
        circle = add_shape(slide, MSO_SHAPE.OVAL,
                           Inches(0.8), y + Inches(0.05), Inches(0.4), Inches(0.4),
                           fill_color=theme["accent"])
        tf = circle.text_frame
        p = tf.paragraphs[0]
        p.text = str(i + 1)
        p.font.size = Pt(11)
        p.font.color.rgb = theme["white"]
        p.alignment = PP_ALIGN.CENTER
        tf.word_wrap = False
        
        # Item text
        title = item.get("title", "")
        desc = item.get("description", "")
        item_text = f"{title}" + (f"\n{desc}" if desc else "")
        add_textbox(slide, Inches(1.4), y, Inches(11.2), Inches(0.6),
                    item_text, font_size=16, color=theme["text"])
        y += Inches(0.7)

def make_contact(prs, slide_data, theme):
    """Contact / CTA page."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0), Inches(0), SLIDE_WIDTH, SLIDE_HEIGHT,
              fill_color=theme["dark_bg"])
    
    add_textbox(slide, Inches(2), Inches(1.5), Inches(9.333), Inches(1),
                "Thank You", font_size=44, bold=True, color=theme["white"],
                alignment=PP_ALIGN.CENTER)
    
    content = slide_data.get("content", {})
    contact_info = content.get("contact", "")
    if contact_info:
        add_textbox(slide, Inches(2), Inches(3.0), Inches(9.333), Inches(2),
                    contact_info, font_size=18, color=theme["accent"],
                    alignment=PP_ALIGN.CENTER)

def make_timeline(prs, slide_data, theme):
    """Timeline / Milestone slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["white"])
    
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0), Inches(0), SLIDE_WIDTH, Pt(4),
              fill_color=theme["accent"])
    
    add_textbox(slide, Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.7),
                slide_data.get("title", ""),
                font_size=28, bold=True, color=theme["primary"])
    
    # Timeline line
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(1), Inches(3.7), Inches(11.333), Pt(3),
              fill_color=theme["accent"])
    
    milestones = slide_data.get("content", {}).get("milestones", [])
    n = len(milestones)
    if n > 0:
        spacing = Inches(11.333) / n
        for i, ms in enumerate(milestones):
            x = Inches(1) + spacing * i + spacing / 2 - Inches(0.3)
            
            # Circle node
            add_shape(slide, MSO_SHAPE.OVAL,
                      x, Inches(3.5), Inches(0.4), Inches(0.4),
                      fill_color=theme["accent"])
            
            # Date label
            add_textbox(slide, x - Inches(0.8), Inches(2.5), Inches(2), Inches(0.4),
                        ms.get("date", ""),
                        font_size=12, bold=True, color=theme["primary"],
                        alignment=PP_ALIGN.CENTER)
            
            # Description
            add_textbox(slide, x - Inches(0.8), Inches(4.2), Inches(2), Inches(1.5),
                        ms.get("title", ""),
                        font_size=11, color=theme["text"],
                        alignment=PP_ALIGN.CENTER)

def make_quote(prs, slide_data, theme):
    """Quote / Key message slide — large text."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, theme["bg"])
    
    add_shape(slide, MSO_SHAPE.RECTANGLE,
              Inches(0), Inches(0), SLIDE_WIDTH, Pt(4),
              fill_color=theme["accent"])
    
    # Big quote mark
    add_textbox(slide, Inches(0.8), Inches(1.0), Inches(11.733), Inches(1.5),
                "❝", font_size=72, color=theme["accent"],
                alignment=PP_ALIGN.LEFT)
    
    # Quote text
    add_textbox(slide, Inches(1.5), Inches(2.2), Inches(10.333), Inches(2),
                slide_data.get("content", {}).get("quote", ""),
                font_size=28, bold=False, color=theme["primary"],
                alignment=PP_ALIGN.LEFT)
    
    # Attribution
    attr = slide_data.get("content", {}).get("attribution", "")
    if attr:
        add_textbox(slide, Inches(1.5), Inches(4.5), Inches(10.333), Inches(0.5),
                    f"— {attr}", font_size=16, color=theme["text_light"],
                    alignment=PP_ALIGN.LEFT)

# ─── Layout Router ─────────────────────────────────────────

LAYOUT_FUNCTIONS = {
    "cover": make_cover,
    "section-divider": make_section_divider,
    "text-only": make_text_only,
    "text-image": make_text_image,
    "two-column": make_two_column,
    "data-chart": make_data_chart,
    "data-table": make_data_table,
    "list": make_list,
    "contact": make_contact,
    "timeline": make_timeline,
    "quote": make_quote,
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
    
    for slide_data in outline.get("slides", []):
        layout = slide_data.get("layout", "text-only")
        func = LAYOUT_FUNCTIONS.get(layout)
        if func:
            func(prs, slide_data, theme)
        else:
            # Fallback to text-only
            make_text_only(prs, slide_data, theme)
    
    prs.save(output_path)
    return output_path

if __name__ == "__main__":
    import sys
    outline_file = sys.argv[1] if len(sys.argv) > 1 else "outline.json"
    output_file = sys.argv[2] if len(sys.argv) > 2 else "presentation.pptx"
    result = generate(outline_file, output_file)
    print(f"✅ PPTX generated: {result}")
```

### 3.3 Run Generation

```bash
python3 generate_ppt.py outline.json my_presentation.pptx
```

On success, inform the user:
> ✅ PPT 已生成：`my_presentation.pptx`
> 说修改意见，或说"满意"结束。

### 3.4 Error Handling

If generation fails:
1. Read the Python error traceback
2. Check: missing outline field? Wrong data type? Missing image?
3. Fix the generate script or outline JSON
4. Re-run

## Phase 4: Interactive Editing (Guided Choices)

When user gives feedback, offer concrete choices:

| User says | Respond with |
|-----------|-------------|
| "第3页不好看" | "你想：A. 换个两栏版式？B. 改纯文字版式？C. 内容拆成两页？" |
| "整体风格太素" | "调整：A. 换深色背景版封面？B. 加更多强调色？C. 换个配色方案？" |
| "再加两页" | "加在哪里？A. 第4页后加案例页 B. 最后加总结页 C. 发给我内容我来安排" |
| "能加图表吗" | "支持：A. 柱状图对比 B. 饼图占比 C. 表格数据 D. 时间线" |

After each edit, regenerate and deliver. Loop until user says done.

## Phase 5: Delivery

### 5.1 Final Checks

Before delivering, verify:
- [ ] File extension is `.pptx`
- [ ] All slide titles are populated
- [ ] No empty text boxes
- [ ] Placeholder images noted for user to insert

### 5.2 Deliver

Output the file path and key info:
> ✅ 生成完成！文件：`my_presentation.pptx`
>
> 📊 共 12 页 · 蓝色科技主题 · 16:9 宽屏
>
> **注意**：
> - 占位图区域已标注 `[ 插入图片 ]`，替换为实际图片即可
> - 图表数据在 PowerPoint 中双击即可编辑

## Reference Files

- `references/layouts.md` — Detailed layout definitions with visual descriptions
- `references/layout-registry.yaml` — Layout selection rules in structured format
- `SKILL.md` — This file (main instructions)

## Assets

- `assets/themes/` — Individual theme color definitions
- `assets/config.yaml` — Configuration template
