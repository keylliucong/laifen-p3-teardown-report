#!/usr/bin/env python3
# Build two 16:9 PPTX decks for the shaver research reports.
from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.parts.image import Image
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import os
import json

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
W, H = 13.333, 7.5
INK = RGBColor(0x1C, 0x19, 0x17)
MUTED = RGBColor(0x6B, 0x65, 0x60)
LINE = RGBColor(0xE7, 0xE2, 0xDC)
BG = RGBColor(0xF7, 0xF5, 0xF2)
CARD = RGBColor(0xFF, 0xFF, 0xFF)
COPPER = RGBColor(0xB4, 0x53, 0x09)
COPPER_SOFT = RGBColor(0xFE, 0xF3, 0xE2)
GREEN = RGBColor(0x15, 0x80, 0x3D)
GREEN_SOFT = RGBColor(0xE8, 0xF7, 0xEE)
BLUE = RGBColor(0x1D, 0x4E, 0xD8)
BLUE_SOFT = RGBColor(0xE8, 0xEF, 0xFF)
AMBER = RGBColor(0xA1, 0x62, 0x07)
AMBER_SOFT = RGBColor(0xFD, 0xF3, 0xD7)
RED = RGBColor(0xB9, 0x1C, 0x1C)
RED_SOFT = RGBColor(0xFD, 0xE8, 0xE8)
FONT = "PingFang SC"
with open(os.path.join(BASE, "assets", "sources.json"), encoding="utf-8") as source_file:
    IMAGE_SOURCES = json.load(source_file)


def set_font(run, size=14, bold=False, color=INK, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = FONT
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", FONT)


def add_text(slide, text, x, y, w, h, size=16, bold=False, color=INK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.0,
             margins=(0, 0, 0, 0)):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(margins[0])
    tf.margin_right = Inches(margins[1])
    tf.margin_top = Inches(margins[2])
    tf.margin_bottom = Inches(margins[3])
    tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(0)
        p.space_before = Pt(0)
        p.line_spacing = spacing
        r = p.add_run()
        r.text = line
        set_font(r, size, bold, color)
    return tb


def add_rich(slide, parts, x, y, w, h, size=14, spacing=1.05,
             anchor=MSO_ANCHOR.TOP):
    """parts: list of paragraphs; each paragraph is list of (text,bold,color,size)"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, para in enumerate(parts):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(0)
        p.space_before = Pt(0)
        p.line_spacing = spacing
        for text, bold, color, sz in para:
            r = p.add_run()
            r.text = text
            set_font(r, sz or size, bold, color)
    return tb


def add_rect(slide, x, y, w, h, fill=CARD, line=LINE, radius=False,
             line_w=0.75):
    shp_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shp_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    if radius:
        try:
            shp.adjustments[0] = 0.06
        except Exception:
            pass
    shp.shadow.inherit = False
    return shp


def add_badge(slide, text, x, y, color=GREEN, bg=GREEN_SOFT, w=0.9, h=0.28, size=9):
    add_rect(slide, x, y, w, h, fill=bg, line=None, radius=True)
    add_text(slide, text, x, y + 0.01, w, h - 0.02, size=size, bold=True,
             color=color, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def add_pic_fit(slide, path, x, y, w, h, border=True):
    iw, ih = Image.from_file(path).size
    scale = min(w / iw, h / ih)
    nw, nh = iw * scale, ih * scale
    if border:
        add_rect(slide, x, y, w, h, fill=CARD, line=LINE, radius=True)
    slide.shapes.add_picture(path, Inches(x + (w - nw) / 2), Inches(y + (h - nh) / 2),
                             width=Inches(nw), height=Inches(nh))


def add_header(slide, num, title, subtitle):
    add_text(slide, num, 0.55, 0.42, 1.0, 0.28, size=11, bold=True, color=COPPER)
    add_text(slide, title, 0.55, 0.68, 12.2, 0.55, size=28, bold=True, color=INK)
    if subtitle:
        add_text(slide, subtitle, 0.55, 1.28, 12.2, 0.42, size=13, color=MUTED)
    add_rect(slide, 0.55, 1.78 if subtitle else 1.62, 12.2, 0.012,
             fill=LINE, line=None)


def add_footer(slide, idx, label):
    prs = slide.part.package.presentation_part.presentation
    num = len(prs.slides)
    add_text(slide, label, 0.55, 7.05, 9.5, 0.25, size=9, color=MUTED)
    add_text(slide, f"{num:02d}", 12.1, 7.05, 0.65, 0.25, size=9, color=MUTED,
             align=PP_ALIGN.RIGHT)


def image_credit(slide, key, x, y, w):
    source = IMAGE_SOURCES[key]
    box = add_text(slide, source["credit"], x, y, w, 0.25, size=9, color=MUTED,
                   align=PP_ALIGN.CENTER)
    run = box.text_frame.paragraphs[0].runs[0]
    run.hyperlink.address = source["page"]
    run.font.color.rgb = MUTED


def image_path(key):
    return os.path.join(BASE, "assets", key)


def add_table(slide, headers, rows, x, y, w, h, col_widths=None,
              font_size=11, header_size=10, row_h=None):
    shape = slide.shapes.add_table(len(rows) + 1, len(headers),
                                   Inches(x), Inches(y), Inches(w), Inches(h))
    table = shape.table
    if col_widths:
        total = sum(col_widths)
        for i, cw in enumerate(col_widths):
            table.columns[i].width = Inches(w * cw / total)
    # header
    for j, head in enumerate(headers):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0xF4, 0xF1, 0xEC)
        cell.margin_left = Inches(0.08)
        cell.margin_right = Inches(0.06)
        cell.margin_top = Inches(0.04)
        cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = head
        set_font(r, header_size, True, MUTED)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.06)
            cell.margin_top = Inches(0.035)
            cell.margin_bottom = Inches(0.035)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = str(val)
            set_font(r, font_size, j == 0, INK if j == 0 else MUTED)
    if row_h:
        for i in range(len(rows) + 1):
            table.rows[i].height = Inches(row_h if i else row_h + 0.05)
    return table


def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    return prs


def new_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    return s


def cover(prs, kicker, title, subtitle, meta, stats, foot, image=None):
    s = new_slide(prs)
    # left accent
    add_rect(s, 0, 0, 0.12, H, fill=COPPER, line=None)
    add_text(s, kicker, 0.75, 0.62, 7.0, 0.3, size=11, bold=True, color=COPPER)
    add_text(s, title, 0.75, 1.0, 7.4, 0.95, size=44, bold=True, color=INK, spacing=0.95)
    add_text(s, subtitle, 0.78, 2.05, 6.6, 0.6, size=22, bold=True, color=COPPER)
    add_text(s, meta, 0.78, 2.78, 6.8, 0.4, size=12, color=MUTED)
    # stats row
    x = 0.78
    for n, l in stats:
        add_text(s, n, x, 3.58, 1.65, 0.55, size=24, bold=True, color=INK)
        add_text(s, l, x, 4.15, 1.65, 0.4, size=10, color=MUTED)
        x += 1.65
    # right illustration panel
    add_rect(s, 8.55, 0.8, 4.1, 5.0, fill=CARD, line=LINE, radius=True)
    if image:
        add_pic_fit(s, image_path(image), 8.73, 1.02, 3.74, 4.08, border=False)
        image_credit(s, image, 8.73, 5.29, 3.74)
    add_text(s, foot, 0.78, 6.55, 11.8, 0.5, size=10, color=MUTED)
    return s


def section_cards(s, items, x=0.55, y=1.95, w=12.2, h=4.7, cols=2, gap=0.18):
    cw = (w - gap * (cols - 1)) / cols
    rows = (len(items) + cols - 1) // cols
    ch = (h - gap * (rows - 1)) / rows if rows else h
    for i, (head, body) in enumerate(items):
        col = i % cols
        row = i // cols
        cx = x + col * (cw + gap)
        cy = y + row * (ch + gap)
        add_rect(s, cx, cy, cw, ch, fill=CARD, line=LINE, radius=True)
        add_text(s, head, cx + 0.18, cy + 0.14, cw - 0.36, 0.4, size=14, bold=True)
        add_text(s, body, cx + 0.18, cy + 0.62, cw - 0.36, ch - 0.72, size=10.5,
                 color=MUTED, spacing=1.05)


def product_gallery(prs, num, title, subtitle, images, note, label=""):
    """Each image pairs an asset key with a model-specific caption."""
    s = new_slide(prs)
    add_header(s, num, title, subtitle)
    n = len(images)
    gap = 0.16
    w = (12.2 - gap * (n - 1)) / n
    for i, (key, cap) in enumerate(images):
        x = 0.55 + i * (w + gap)
        add_rect(s, x, 1.95, w, 4.25, fill=CARD, line=LINE, radius=True)
        add_pic_fit(s, image_path(key), x + 0.12, 2.07, w - 0.24, 3.2, border=False)
        add_text(s, cap, x + 0.14, 5.36, w - 0.28, 0.33, size=11, bold=True,
                 color=INK, align=PP_ALIGN.CENTER)
        image_credit(s, key, x + 0.10, 5.80, w - 0.20)
    add_text(s, note, 0.55, 6.38, 12.2, 0.4, size=10, color=MUTED)
    add_footer(s, num, label)
    return s


def figure_label(slide, text, x, y, w, h=0.32, size=11, color=INK):
    add_text(slide, text, x, y, w, h, size=size, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def arrow(slide, x, y, w, h, kind=MSO_SHAPE.RIGHT_ARROW, color=COPPER):
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def flow(slide, nodes, x, y, w, h=0.68):
    gap = 0.45
    width = (w - gap * (len(nodes) - 1)) / len(nodes)
    for i, label in enumerate(nodes):
        left = x + i * (width + gap)
        add_rect(slide, left, y, width, h, fill=COPPER_SOFT, line=None, radius=True)
        figure_label(slide, label, left + 0.04, y, width - 0.08, h, color=COPPER)
        if i < len(nodes) - 1:
            arrow(slide, left + width + 0.10, y + (h - 0.17) / 2, 0.25, 0.17)


def motor_diagram(slide, x, y, w):
    add_rect(slide, x, y, w, 3.2, radius=True)
    add_text(slide, "音圈式直线电机 · 通用原理（推断）", x + 0.25, y + 0.18,
             w - 0.5, 0.33, size=14, bold=True)
    center = x + w / 2
    arrow(slide, center - 0.72, y + 0.69, 1.44, 0.20, MSO_SHAPE.LEFT_RIGHT_ARROW)
    figure_label(slide, "往复方向", center + 0.78, y + 0.61, 0.96, 0.32, size=9, color=COPPER)
    add_rect(slide, center - 1.13, y + 1.08, 2.26, 0.45, fill=BLUE_SOFT, line=BLUE)
    figure_label(slide, "动子 · 线圈 / 刀头座", center - 1.13, y + 1.08, 2.26, 0.45, color=BLUE)
    for direction in (-1, 1):
        # Each flexure joins the moving coil to the stationary frame.
        points = [(center + direction * dx, y + dy) for dx, dy in
                  ((1.13, 1.30), (1.33, 1.30), (1.57, 1.55), (1.33, 1.80), (1.57, 2.13))]
        for start, end in zip(points, points[1:]):
            segment = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                                *(Inches(v) for v in (*start, *end)))
            segment.line.color.rgb = AMBER
            segment.line.width = Pt(2.2)
    figure_label(slide, "弹片", center + 1.66, y + 1.45, 0.61, size=10, color=AMBER)
    figure_label(slide, "气隙", center - 0.40, y + 1.53, 0.80, 0.22, size=8, color=MUTED)
    for offset, pole, soft, color in ((-0.91, "N", RED_SOFT, RED), (0.01, "S", BLUE_SOFT, BLUE)):
        add_rect(slide, center + offset, y + 1.76, 0.90, 0.37, fill=soft, line=None)
        figure_label(slide, pole, center + offset, y + 1.76, 0.90, 0.37, color=color)
    figure_label(slide, "永磁体", center + 1.68, y + 1.84, 0.74, size=10)
    add_rect(slide, center - 1.94, y + 2.13, 3.88, 0.42, fill=LINE, line=None)
    figure_label(slide, "定子 · 磁铁 / 轭铁", center - 1.94, y + 2.13, 3.88, 0.42)
    figure_label(slide, "弹片提供回复刚度 k · 本机固定/运动件归属待确认",
                 x + 0.22, y + 2.70, w - 0.44, size=9, color=MUTED)


def score_bar(slide, score, total):
    for i in range(total):
        add_rect(slide, 0.65 + 0.164 * i, 3.46, 0.127, 0.2,
                 fill=COPPER if i < score else LINE, line=None)
    figure_label(slide, "报告估分", 0.64, 3.77, 1.60, size=10, color=MUTED)


def fmea_table(slide, headers, rows, widths):
    x, y, table_w, h = 0.55, 1.95, 8.18, 3.9
    add_table(slide, headers, rows, x, y, table_w, h, col_widths=widths, font_size=10.5)
    panel_x = 8.93
    add_rect(slide, panel_x, y, 3.82, h, radius=True)
    add_text(slide, "RPN · 共同尺度 0–200", panel_x + 0.16, y + 0.12, 3.48, 0.27,
             size=11, bold=True)
    row_h = h / (len(rows) + 1)
    for i, row in enumerate(rows, start=1):
        value = int(row[-1])
        mid = y + (i + 0.5) * row_h
        kind = row[0].split()[-1]
        color = COPPER if kind.startswith("D") else BLUE
        add_text(slide, kind, panel_x + 0.15, mid - 0.10, 0.44, 0.24, size=9, color=MUTED)
        add_rect(slide, panel_x + 0.63, mid - 0.095, 2.56, 0.19, fill=BG, line=None)
        add_rect(slide, panel_x + 0.63, mid - 0.095, 2.56 * value / 200, 0.19, fill=color, line=None)
        add_text(slide, str(value), panel_x + 3.29, mid - 0.10, 0.42, 0.24, size=9, color=color)


def comparison_panel(slide, rows, x, y, w, title, unit, axis_max):
    add_rect(slide, x, y, w, 2.08, radius=True)
    add_text(slide, title, x + 0.17, y + 0.12, w - 0.34, 0.26, size=11, bold=True)
    start, track = x + 1.67, w - 2.55
    for i, (name, low, high, label) in enumerate(rows):
        top = y + 0.47 + i * 0.28
        add_text(slide, name, x + 0.17, top, 1.47, 0.23, size=8.5, color=MUTED)
        add_rect(slide, start, top + 0.048, track, 0.12, fill=BG, line=None)
        add_rect(slide, start, top + 0.048, track * low / axis_max, 0.12,
                 fill=COPPER if unit == "元" else BLUE, line=None)
        if high > low:
            add_rect(slide, start + track * low / axis_max, top + 0.048,
                     track * (high - low) / axis_max, 0.12, fill=AMBER_SOFT, line=AMBER)
        add_text(slide, label, x + w - 0.78, top, 0.72, 0.23, size=8.5, color=INK)


def competitor_comparison(slide, rows, widths):
    add_table(slide, ["机型", "路线", "刀头", "重量", "续航", "价格"], rows,
              0.55, 1.95, 12.2, 2.45, col_widths=widths, font_size=10)
    prices, weights = [], []
    for model, _, _, mass, _, price in rows:
        name = model.removeprefix("徕芬 ").removeprefix("yoose ")
        if "$" not in price:
            bounds = price.replace("~", "").replace(",", "").split("–")
            prices.append((name, int(bounds[0]), int(bounds[-1]), price))
        # MINI 5 PRO's weight in the existing table is a MINI 5 series figure.
        if mass.endswith("g") and name != "MINI 5 PRO":
            value = float(mass[:-1])
            weights.append((name, value, value, mass[:-1]))
    comparison_panel(slide, prices, 0.55, 4.55, 6.0, "公开价格（元） · 0–1,200", "元", 1200)
    comparison_panel(slide, weights, 6.75, 4.55, 6.0, "可对照重量（g） · 0–200", "g", 200)
    note = "价格沿用报告的发布/渠道口径，价区间用浅色尾段表示，美元价格未画。"
    if any("MINI 5 PRO" in row[0] for row in rows):
        note += " 117 g 为 MINI 5 系列口径，PRO 实重待确认。"
    add_text(slide, note, 0.55, 6.73, 12.2, 0.23, size=9, color=MUTED)


def photo_panel(slide, key, x, y, w, h):
    add_rect(slide, x, y, w, h, radius=True)
    add_pic_fit(slide, image_path(key), x + 0.14, y + 0.14, w - 0.28, h - 0.61, border=False)
    image_credit(slide, key, x + 0.08, y + h - 0.37, w - 0.16)


def layer_stack(slide, layers):
    for i, (name, desc, color) in enumerate(layers):
        y = 1.95 + i * 0.72
        # Offset layers to make their separation and order visible.
        left = 0.75 + i * 0.14
        add_rect(slide, left + 0.08, y + 0.09, 1.22, 0.39, fill=LINE, line=None, radius=True)
        add_rect(slide, left, y, 1.22, 0.39, fill=color, line=None, radius=True)
        figure_label(slide, name[0], left, y, 1.22, 0.39, color=CARD)
        add_rect(slide, 2.9, y, 9.85, 0.58, radius=True)
        add_rect(slide, 2.9, y, 0.07, 0.58, fill=color, line=None)
        add_text(slide, name[2:], 3.1, y + 0.13, 3.4, 0.30, size=12, bold=True)
        add_text(slide, desc, 6.45, y + 0.15, 6.1, 0.29, size=10.5, color=MUTED)


def evidence_panel(slide, items, y=1.95, h=3.9):
    x, w = 9.0, 3.75
    add_rect(slide, x, y, w, h, radius=True)
    step = (h - 0.36) / len(items)
    for i, (heading, desc, color, bg) in enumerate(items):
        top = y + 0.18 + i * step
        add_rect(slide, x + 0.17, top, w - 0.34, step - 0.11, fill=bg, line=None, radius=True)
        add_text(slide, heading, x + 0.32, top + 0.09, w - 0.64, 0.26, size=11, bold=True, color=color)
        add_text(slide, desc, x + 0.32, top + 0.39, w - 0.64, step - 0.43, size=9.5, color=MUTED)


def source_note(slide):
    add_text(slide, "图片授权见 assets/sources.json。第三方资料图未查到开放转载许可。",
             0.55, 6.49, 12.2, 0.26, size=9, color=MUTED)

def build_laifen(out):
    prs = new_deck()
    L = "徕芬 P3 标准版 拆机研究报告 · v0.1 · 2026-11-27"

    s = cover(prs, "TEARDOWN RESEARCH · 非官方独立研究",
              "徕芬 P3 标准版", "拆机研究报告",
              "LFESP3-P · 2026.05 产 · 曜石黑 · 599 元",
              [("187.1 g", "主机 + 刀头"), ("200 Hz", "≈12,000 CPM"),
               ("3 刀头", "2 穹顶 + 1 长须"), ("3/10", "可维修性（推断）")],
              "证据分级：官方/包装 · 拆解/媒体 · 推断 · 待确认。拆机由机主执行，本报告不代替实测。",
              image="laifen/product.jpg")

    product_gallery(prs, "产品图", "产品外观与官方结构图",
                    "中国版 P3 标准版发布图与徕芬 P3 官网电机细节。",
                    [("laifen/product.jpg", "P3 标准版整机 · 四色"),
                     ("laifen/blade.jpg", "开放式长须刀与短须刀网"),
                     ("laifen/exploded.jpg", "机身分解渲染"),
                     ("laifen/motor.jpg", "L2 双直线电机")],
                    "机身分解渲染来自发布物料。内部零件、固定方式和装配关系仍需本机验证。",
                    label=L)

    # 2 summary
    s = new_slide(prs)
    add_header(s, "00", "一页结论", "先给判断，再给依据；未实测的一律标推断。")
    section_cards(s, [
        ("1. 它是 P3 Pro 的同平台降本机",
         "同 L2 电机、同三刀头、同 1000 mAh；差异在外壳（玻纤 PC）、刀网（冲压）、续航（90 vs 100 min）。"),
        ("2. 内部大概率沿用 T1 Pro 机头方案",
         "压板 + 螺丝 + 双道密封圈；双电机反向、连杆带三刀头。这是最该被本机证伪的一条。"),
        ("3. 刀头易维护，机身内部不是",
         "刀头磁吸可换；电池、PCBA、电机、密封件不供件，开壳即失去 IPX7。iFixit 口径 3/10。"),
        ("4. 四个硬缺口",
         "视窗固定方式、电池批次与并联、双电机到三刀头的连杆拓扑、电机-主板焊/插。"),
    ], h=4.35)
    add_text(s, "一句话：动力与切剃平台共用、外壳与刀网重新做；判断共用件时，电机与刀头可共用结论，刀网/外壳/密封必须分开写。",
             0.55, 6.42, 12.2, 0.4, size=11, bold=True, color=COPPER)
    add_footer(s, 2, L)

    # 3 baseline
    s = new_slide(prs)
    add_header(s, "01", "产品基线与批次", "包装 + 官网交叉核对；续航保留 80/90/100 三种口径。")
    add_table(s, ["项目", "值", "证据"], [
        ["型号 / 版本", "LFESP3-P，标准版；曜石黑（另有极白、海湾蓝、薄雾灰）", "官方 / 包装"],
        ["上市 / 价格", "2026-04-23 开售 509 元起；2026-05-11 发布 599 元", "媒体"],
        ["尺寸 / 重量", "59×72.5×35.4 mm；187.1 g（主机+刀头）", "官方"],
        ["外壳", "玻纤增强 PC + Monocoat；正面透明探索视窗", "官方"],
        ["电机 / 刀头", "双直线 L2，200±10 Hz；3 刀头（2 穹顶+1 长须）", "官方"],
        ["电池 / 续航", "1000 mAh / 3.8 Wh；约 90 min；USB-C 1.5 h", "官方 / 包装"],
        ["防护 / 标准", "IPX7；GB/T 4706.1-2024、GB/T 4706.9-2024", "包装"],
        ["制造", "深圳术叶创新（制造商）/ 珠海徕芬创新（生产商）", "包装"],
    ], 0.55, 1.95, 8.46, 4.35, col_widths=[1.6, 6.6, 1.7], font_size=10.5)
    photo_panel(s, "laifen/product.jpg", 9.18, 1.95, 3.57, 4.35)
    add_footer(s, 3, L)

    # 4 conflicts
    s = new_slide(prs)
    add_header(s, "02", "三处公开口径冲突", "不取平均，拆机前并列保留。")
    section_cards(s, [
        ("续航：80 / 90 / 100 min",
         "4 月商品页 80；5 月发布会/官网 90；台湾渠道 100。三者都没有负载、停机判据或测试批次。说明书另规定约 10 min 自动停机。"),
        ("刀网：55 µm 是 Pro 口径",
         "中国官网 55 µm 的脚注明确用于 P3 Pro；标准版用冲压刀网，厚度未官宣。台湾渠道写 0.055 mm，不能直接当标准版实测。"),
        ("电池：双电芯已见，颗数待核",
         "拆解视频可见两颗电芯，作者称每块 500 mAh；整机标 1000 mAh/3.8 Wh，最可能并联，但本批次要拆开确认。"),
    ], h=3.6)
    add_text(s, "报告写法：官方声明 ≠ 实测；同平台拆解 ≠ 本机拆解。冲突处保留来源与日期，不替厂商圆。",
             0.55, 6.0, 12.2, 0.4, size=11, bold=True, color=COPPER)
    add_footer(s, 4, L)

    # 5 structure layers
    s = new_slide(prs)
    add_header(s, "03", "结构分层与 BOM", "9 个一级件、约 34 个二级件；12 项待本机确认。")
    layers = [
        ("① 刀头总成", "2 穹顶短须 + 1 长须 · 冲压不锈钢网 · 共用托架", GREEN),
        ("② 上壳 + 透明视窗", "玻纤 PC · Monocoat · 视窗固定方式待确认", AMBER),
        ("③ 机芯：双直线电机 L2", "两台 L1 反向 · 弹性减震悬挂 · 金属回旋弹片", GREEN),
        ("④ PCBA", "MCU · 双电机驱动 · 压感 · 充电 · 触觉", AMBER),
        ("⑤ 电池组", "1000 mAh · 双电芯（视频）· 并联待确认", RED),
        ("⑥ 密封与减振", "机头双圈 · 尾部双圈 · USB-C 密封 · 吸震器", COPPER),
    ]
    layer_stack(s, layers)
    add_text(s, "四个决定性未知：视窗固定 / 电池颗数与并联 / 双电机→三刀头连杆 / 电机-主板焊插。",
             0.55, 6.45, 12.2, 0.35, size=10.5, bold=True, color=RED)
    add_footer(s, 5, L)

    # 6 BOM grid
    s = new_slide(prs)
    add_header(s, "04", "BOM 一级件", "二级约 34 项；绿色=可见/官方，琥珀=推断，红色=待拆。")
    items = ["刀头总成", "上壳 + 视窗", "底板 / 压盖", "机芯（双 L1）", "PCBA",
             "电池组", "USB-C 小板", "密封 / 减振", "按键 / 导光"]
    for i, name in enumerate(items):
        col = i % 3
        row = i // 3
        x = 0.55 + col * 4.05
        y = 1.95 + row * 1.28
        add_rect(s, x, y, 3.85, 1.05, fill=CARD, line=LINE, radius=True)
        add_text(s, f"{i+1:02d}", x + 0.16, y + 0.12, 0.5, 0.25, size=9, bold=True, color=COPPER)
        add_text(s, name, x + 0.55, y + 0.1, 3.1, 0.32, size=13, bold=True)
        add_text(s, ["2 短须+1 长须+托架", "玻纤 PC + Monocoat", "铭牌 / 尾部密封",
                     "双 L1 + 悬挂 + 连杆", "主控 / 驱动 / 充电 / 压感",
                     "1000mAh；颗数待确认", "插座 + 密封 + OTG", "O 圈 / 轴封 / 吸震器",
                     "压感区 + 双指示灯"][i],
                 x + 0.55, y + 0.52, 3.1, 0.4, size=9.5, color=MUTED)
    add_text(s, "完整 BOM 树（含 12 项待确认）见 HTML 报告与 data/bom.csv。",
             0.55, 6.25, 12.2, 0.35, size=10, color=MUTED)
    add_footer(s, 6, L)

    # 7 motor
    s = new_slide(prs)
    add_header(s, "05", "双直线电机 L2", "200 Hz 与 12,000 CPM 是同一个量；但“次”可能按单程计。")
    motor_diagram(s, 0.55, 1.95, 6.0)
    add_rect(s, 6.95, 1.95, 5.8, 3.2, fill=CARD, line=LINE, radius=True)
    add_text(s, "三个判断", 7.2, 2.1, 5.2, 0.35, size=15, bold=True)
    add_text(s, "• 200 Hz 放在机械谐振附近，弹片储能、电磁力补损耗\n"
                "• 双电机水平反向，惯量对消；装回必须成对对位\n"
                "• 金属回旋弹片是疲劳件；官方称电机 >500 h",
             7.2, 2.6, 5.2, 2.2, size=11.5, color=MUTED, spacing=1.2)
    add_text(s, "12,000 ÷ 60 = 200 Hz", 0.55, 5.5, 6.0, 0.5, size=22, bold=True, color=COPPER)
    add_text(s, "判别：加速度计看主峰 200 Hz 还是 100 Hz；100 Hz 表示“次”= 单程。",
             0.55, 6.05, 12.2, 0.4, size=11, color=INK)
    add_footer(s, 7, L)

    # 8 numbers
    s = new_slide(prs)
    add_header(s, "06", "关键数字（均为推断/设计区间）", "不是实测；拆机当天按 A 组对照。")
    metrics = [("0.4–0.8 mm", "峰峰行程"), ("2–4 N", "峰值推力"),
               ("1.3–1.6 W", "空载功率"), ("55–65 dB(A)", "空载 1 m 噪声")]
    for i, (n, l) in enumerate(metrics):
        x = 0.55 + i * 3.08
        add_rect(s, x, 1.95, 2.9, 1.55, fill=CARD, line=LINE, radius=True)
        add_text(s, n, x + 0.18, 2.12, 2.55, 0.5, size=21, bold=True, color=COPPER)
        add_text(s, l, x + 0.18, 2.72, 2.55, 0.35, size=10.5, color=MUTED)
    add_text(s, "推算链\n"
                "• 频率 200 Hz → 峰值加速度 316–632 m/s²（32–64 g）\n"
                "• 动子质量 3–8 g、气隙磁密 0.3–0.5 T → 推力 2–4 N\n"
                "• 3.8 Wh / 1.5 h → 空载平均 1.3–1.6 W\n"
                "• 刀网 55 µm 是 Pro 口径；标准版冲压网厚度待测",
             0.55, 3.95, 12.2, 2.4, size=12.5, color=INK, spacing=1.3)
    add_footer(s, 8, L)

    # 9 A1-A4
    s = new_slide(prs)
    add_header(s, "07", "拆前通电可测 A1–A4", "整机还能跑的时候做完；拆开就没法测了。")
    add_table(s, ["#", "测什么", "怎么测", "预期"], [
        ["A1", "刀头座加速度主峰", "加速度计 + FFT，采样 ≥10 kHz", "200±5 Hz；壳体 ≤ 刀头座 1/5"],
        ["A2", "峰峰行程", "高速摄影 / 激光位移计，3 次", "0.4–0.8 mm；离散 <5%"],
        ["A3", "空载功率", "USB 功率计 + 支路电流", "整机 1.3–1.6 W；每台 0.17–0.21 A"],
        ["A4", "噪声频谱", "声级计 1 m，1/3 倍频程", "200/400/600 Hz 窄带；55–65 dB(A)"],
    ], 0.55, 1.95, 12.2, 3.6, col_widths=[0.8, 2.8, 5.0, 4.2], font_size=11.5)
    add_text(s, "B1–B6：弹片形式、磁钢排布、绕组电阻、刀网厚度/槽宽、刃角、刀片-刀网间隙。",
             0.55, 5.8, 12.2, 0.4, size=11, color=MUTED)
    flow(s, ("A1 主峰", "A2 行程", "A3 空载功率", "A4 噪声"), 0.55, 6.31, 12.2, h=0.42)
    add_footer(s, 9, L)

    # 10 blades
    s = new_slide(prs)
    add_header(s, "08", "三刀头与刀网", "2 穹顶短须 + 1 开放式长须；标准版与 Pro 的刀网工艺不同。")
    add_rect(s, 0.55, 1.95, 4.17, 3.8, fill=CARD, line=LINE, radius=True)
    add_text(s, "已知", 0.78, 2.1, 3.76, 0.35, size=15, bold=True)
    add_text(s, "• 布局：2 穹顶短须 + 1 长须，条形往复\n"
                "• 外刀网：冲压精钢/低镍不锈钢，Ra0.23\n"
                "• 内刃：301 不锈钢，18 次冲压\n"
                "• 拆解视频：错列圆角长孔；短须动刀下有弹簧\n"
                "• 刀头磁吸，唯一官方可换件",
             0.78, 2.6, 3.76, 2.9, size=11.5, color=MUTED, spacing=1.2)
    add_rect(s, 4.88, 1.95, 4.17, 3.8, fill=CARD, line=LINE, radius=True)
    add_text(s, "未知", 5.10, 2.1, 3.76, 0.35, size=15, bold=True)
    add_text(s, "• 标准版刀网厚度：55 µm 是 Pro 口径\n"
                "• 孔型、孔距、开孔率、倒角\n"
                "• 单条网法向顺应行程、整头摆角\n"
                "• 动刀牌号/硬度/镀层\n"
                "• P3 与 Pro 刀头是否真能互换（SKU 待核）",
             5.10, 2.6, 3.76, 2.9, size=11.5, color=MUTED, spacing=1.2)
    add_text(s, "不要直接引用 55 µm：标准版是冲压网，Pro 是电铸网，不是同一工艺。",
             0.55, 6.0, 12.2, 0.4, size=11, bold=True, color=RED)
    photo_panel(s, "laifen/blade.jpg", 9.21, 1.95, 3.54, 3.8)
    add_footer(s, 10, L)

    # 11 electrical
    s = new_slide(prs)
    add_header(s, "09", "电气架构：以 T1 Pro 为参照", "P3 没有公开器件级拆解；同平台 T1 Pro 是最可靠的参照。")
    add_table(s, ["功能", "T1 Pro 已知器件", "P3 预期"], [
        ["主控 MCU", "中微 CMS32F033（Cortex-M0）", "同系列或升级，待丝印"],
        ["直线电机驱动", "拓尔微 TMI8261（5A）", "双电机需两路或双通道"],
        ["充电管理", "钰泰 ETA6198（4.35V）", "可能沿用；OTG 需额外通路"],
        ["压感", "芯海 CSU18M68", "全域感应按钮"],
        ["触觉驱动", "艾为 AW86927", "线性马达反馈"],
        ["电池", "亿纬 L1318 530mAh", "1000mAh；双电芯待确认"],
    ], 0.55, 1.95, 12.2, 3.05, col_widths=[2.2, 5.0, 4.8], font_size=11)
    flow(s, ("USB-C · 5 W 上限", "充电管理（推断）", "1000 mAh · 3.8 Wh", "约 1.5 小时"), 0.55, 5.17, 12.2)
    add_text(s, "5V⎓1A 是额定输入上限（约 5W），不等于电芯充电电流；1.5h 充满 3.8Wh，平均约 2.5W。",
             0.55, 6.1, 12.2, 0.4, size=10.5, color=MUTED)
    add_footer(s, 11, L)

    # 12 seal
    s = new_slide(prs)
    add_header(s, "10", "密封与 IPX7", "IPX7 是短时浸水，不等于可长期泡水；开壳即破坏密封。")
    add_table(s, ["密封点", "预计形式", "风险"], [
        ["视窗与壳体", "卡扣 / 胶 / 双色注塑（待确认）", "1（最高）"],
        ["USB-C 口", "胶塞 + 插座密封圈", "2"],
        ["尾部底板", "两道胶圈", "3"],
        ["机头压板 / 输出轴", "压板 O 圈 + 轴封", "4"],
    ], 0.55, 1.95, 12.2, 2.9, col_widths=[3.0, 6.6, 1.6], font_size=12)
    flow(s, ("视窗 / 壳体", "USB-C 口", "尾部底板", "压板 / 输出轴"), 0.55, 6.46, 12.2, h=0.39)
    add_text(s, "维修含义：任何一次开壳都会破坏至少一道密封；官方不单卖密封件。\n"
                "“能装回去”不等于“还能防水”。复装后没有新圈，报告写“不再保证 IPX7”。",
             0.55, 5.15, 12.2, 1.2, size=13, bold=True, color=RED, spacing=1.25)
    add_footer(s, 12, L)

    # 13 repairability
    s = new_slide(prs)
    add_header(s, "11", "可维修性：3/10（推断）", "按 iFixit 10 分制；拆机后按实际密封件/供件情况复核。")
    add_text(s, "3", 0.6, 2.0, 1.8, 1.3, size=64, bold=True, color=COPPER)
    add_text(s, "/10", 1.55, 2.55, 1.0, 0.5, size=18, bold=True, color=MUTED)
    score_bar(s, 3, 10)
    add_table(s, ["维度", "得分", "依据"], [
        ["进入机身", "2 / 3", "刀头磁吸；压板/底板需螺丝；隐藏螺丝未知"],
        ["可换件", "1 / 3", "刀头可买；电池/PCBA/电机/密封件不供件"],
        ["工具与步骤", "0 / 2", "精密批头 + 撬片 + 溶胶；电池胶粘 + 点焊"],
        ["复装 / 防水", "0 / 1", "开壳破坏 IPX7；官方不供密封件"],
        ["文档 / 诊断", "0 / 1", "无维修手册、无故障码、无调试口"],
    ], 2.5, 1.95, 10.2, 3.8, col_widths=[2.2, 1.3, 6.7], font_size=11)
    add_text(s, "对刀头是易维护设计，对机身内部不是。", 0.55, 6.15, 12.2, 0.4,
             size=12, bold=True, color=COPPER)
    add_footer(s, 13, L)

    # 14 FMEA
    s = new_slide(prs)
    add_header(s, "12", "FMEA Top 风险", "使用期 F 与拆解引入 D 分列；拆解项沿用 v1 单次拆装口径。")
    fmea_table(s, ["类型", "风险", "早期信号", "RPN"], [
        ["拆解引入 D1", "密封圈错位 / 漏装", "无；拆过就回不到 IPX7", "200"],
        ["拆解引入 D7", "静电损伤 PCBA", "当场测不出，几周后失效", "168"],
        ["使用期 F5", "电机弹片疲劳 / 断裂", "噪音变大、震感增强", "147"],
        ["使用期 F1", "进水 → 内部短路", "水雾 / 充电异常", "144"],
        ["使用期 F2", "刀网磨损 / 变形", "拉扯、剃不净", "144"],
        ["拆解引入 D2", "排线撕裂", "压感 / 指示灯 / 电机断联", "140"],
    ], [1.8, 3.0, 3.0, 0.8])
    add_text(s, "其余 7 条：电池 105 · 压感 120 · USB-C 120 · 磁吸 75 · 传动轴 120 · 触觉 27 · PID 72。",
             0.55, 6.1, 12.2, 0.4, size=10.5, color=MUTED)
    add_footer(s, 14, L)

    # 15 compliance
    s = new_slide(prs)
    add_header(s, "13", "合规与标准", "包装标 2024 版合规；要核 CCC 证书是否已换版。")
    add_table(s, ["文件", "要点"], [
        ["GB/T 4706.1-2024", "通用安全；2024-07-24 发布、2026-08-01 实施；10/11/13/15-16/19/21-24 章"],
        ["GB/T 4706.9-2024", "剃须刀专用；6.2 至少 IPX7；19.101 非正常；21.101 900mm 五方位跌落"],
        ["GB 44246-2024", "家电安全强制规范，2026-08-01 实施；4706 系列转为推荐性配套"],
        ["CQC 换版通知", "2026-07-31 前可自愿按新版；旧证 2026-11-01 前未换撤销"],
        ["IPX7", "1m/30min 浸水、无有害进水；不覆盖“USB-C 口带水充电”"],
    ], 0.55, 1.95, 8.24, 3.9, col_widths=[2.55, 5.69], font_size=10.5)
    evidence_panel(s, [
        ("通用安全", "GB/T 4706.1", BLUE, BLUE_SOFT),
        ("剃须刀专用", "GB/T 4706.9", BLUE, BLUE_SOFT),
        ("认证资料待取", "CCC 证书号与 IPX7 原始报告", AMBER, AMBER_SOFT),
    ])
    add_text(s, "未找到：CCC 证书号、IPX7 原始报告、电池 UN 38.3 摘要。报告只写“未找到”，不假设已认证。",
             0.55, 6.1, 12.2, 0.4, size=10.5, color=MUTED)
    add_footer(s, 15, L)

    # 16 patents
    s = new_slide(prs)
    add_header(s, "14", "专利与供应链", "专利是设计意图，不是本机一定用了；权项与法律状态未取。")
    add_table(s, ["公开号", "内容", "与本机的关系"], [
        ["CN223124712U", "双往复电机反向 + 弹性悬挂", "与“双电机反向”强线索"],
        ["WO2026066028A1", "双往复电机 / 毛发修剪器", "同族对照"],
        ["CN116015013A", "往复式磁悬浮线性电机", "L2 首选核对；权项未取"],
        ["CN221111915U", "动刀组件 / 刀头模组", "穹顶刀头对照；权项未取"],
        ["CN309370742S", "深圳术叶剃须刀外观", "外观图样，不证明内部"],
    ], 0.55, 1.95, 8.24, 3.6, col_widths=[2.1, 3.25, 2.89], font_size=10.5)
    evidence_panel(s, [
        ("已取", "专利著录资料与摘要", BLUE, BLUE_SOFT),
        ("待核", "权利要求、法律状态与本机应用", AMBER, AMBER_SOFT),
    ], h=3.6)
    add_text(s, "Google Patents 对权利要求页 503；本轮只到著录/摘要级，不能写“本机使用了该专利”。",
             0.55, 5.9, 12.2, 0.4, size=10.5, color=MUTED)
    add_footer(s, 16, L)

    # 17 competitors
    s = new_slide(prs)
    add_header(s, "15", "竞品对照", "同品牌看平台差异，跨品牌看技术路线。")
    competitor_comparison(s, [
        ["徕芬 P3", "双直线 L2，200Hz", "3", "187.1g", "90min", "599"],
        ["徕芬 P3 Pro", "双直线 L2", "3", "179g", "100min", "699–799"],
        ["徕芬 T1 Pro", "单直线 L1", "1", "102g", "110min", "499"],
        ["松下 ES-LV67", "直线往复 14,000/min", "5", "—", "50min", "~$150"],
        ["飞利浦 S7888", "旋转式", "3", "—", "60min", "~1,149"],
        ["博朗 3040s", "三组浮动往复", "3", "—", "45min", "284–369"],
    ], [2.4, 4.2, 1.2, 1.5, 1.5, 1.4])
    add_footer(s, 17, L)

    # 18 checklist
    s = new_slide(prs)
    add_header(s, "16", "拆机观察清单", "先拍照、后动手；每步螺丝单独装袋。")
    add_text(s, "必拍 P01–P20：四视图 → 刀头 → 压板与密封圈 → 机芯 → 电机端子 → PCBA 正反 → 电池 → USB-C → 视窗",
             0.55, 1.95, 12.2, 0.5, size=11, color=INK)
    add_text(s, "必量 D01–D10：上壳、视窗、压板、电机中心距、刀头、电池、密封圈截面、行程、重量、壁厚",
             0.55, 2.55, 12.2, 0.5, size=11, color=INK)
    add_text(s, "必记 S01–S10：MCU、电机驱动、充电 IC、压感、触觉驱动、保护 IC、电芯、电机铭牌、铭牌、认证标识",
             0.55, 3.15, 12.2, 0.45, size=11, color=INK)
    flow(s, ("刀头", "压板", "机芯", "PCBA / 电池", "Type-C"), 0.55, 3.7, 12.2, h=0.51)
    add_text(s, "拆机观察顺序（推断，开壳路径待确认）", 0.55, 4.28, 12.2, 0.25, size=9, color=MUTED)
    add_rect(s, 0.55, 4.63, 12.2, 1.35, fill=COPPER_SOFT, line=None, radius=True)
    add_text(s, "四个决定性问题", 0.8, 4.75, 4.0, 0.35, size=14, bold=True, color=COPPER)
    add_text(s, "① 视窗怎么固定？  ② 电池 1 颗还是 2 颗？\n③ 双电机怎么带三个刀头？  ④ 电机与主板是焊还是插？",
             0.8, 5.12, 11.6, 0.71, size=13, bold=True, color=INK, spacing=1.3)
    add_text(s, "安全：电池不刺穿、先放电、防火袋；密封圈拆一次即换；拆前录噪音基线。",
             0.55, 6.2, 12.2, 0.4, size=10.5, color=RED)
    add_footer(s, 18, L)

    # 19 gaps
    s = new_slide(prs)
    add_header(s, "17", "资料缺口与边界", "公开资料能回答“大概怎么做的”，不能回答“这台具体怎么做的”。")
    add_rect(s, 0.55, 5.20, 12.2, 0.77, fill=COPPER_SOFT, line=None, radius=True)
    add_text(s, "拆机可回答", 0.55, 1.95, 5.8, 0.35, size=15, bold=True)
    add_text(s, "• 视窗固定、电池颗数/并联、连杆拓扑\n• 电机-主板焊/插、驱动 IC 颗数\n• 密封圈数量/截面、螺丝规格\n• PCBA 器件型号、板层、测试点",
             0.55, 2.45, 5.8, 2.5, size=11.5, color=MUTED, spacing=1.25)
    add_text(s, "拆机也不够", 6.95, 1.95, 5.8, 0.35, size=15, bold=True)
    add_text(s, "• 真实行程、负载频率、推力-电流\n• 弹片寿命与失效模式\n• 刀网厚度/刃角/切割效率\n• 复装后 IPX7、续航真值、噪声频谱",
             6.95, 2.45, 5.8, 2.5, size=11.5, color=MUTED, spacing=1.25)
    add_text(s, "本报告不回答：值不值得买、比 Pro 好多少、能用几年。这些需要统一样本实测与长期数据。",
             0.75, 5.35, 11.8, 0.6, size=13, bold=True, color=COPPER)
    add_footer(s, 19, L)

    # 20 sources
    s = new_slide(prs)
    add_header(s, "18", "来源与下载", "查询日 2026-11-27；官方/包装优先，媒体只作线索。")
    add_text(s, "关键来源", 0.55, 1.95, 3.0, 0.35, size=14, bold=True)
    add_text(s, "• laifen.net P3 规格 / 对比 / P3 Pro 技术页\n"
                "• 充电头网 T1 Pro 拆解（2025-06）\n"
                "• 避坑拆解 P3 标准版（B 站）\n"
                "• GB/T 4706.1/.9-2024、GB/T 4208-2017\n"
                "• Google Patents: CN223124712U 等",
             0.55, 2.4, 5.6, 2.5, size=11, color=MUTED, spacing=1.25)
    add_rect(s, 6.6, 1.95, 6.15, 3.3, fill=CARD, line=LINE, radius=True)
    add_text(s, "在线报告", 6.85, 2.1, 5.6, 0.35, size=14, bold=True)
    add_text(s, "https://keylliucong.github.io/laifen-p3-teardown-report/laifen-p3/",
             6.85, 2.55, 5.6, 0.8, size=10.5, color=COPPER)
    add_text(s, "仓库：github.com/keylliucong/laifen-p3-teardown-report\n"
                "本 PPT 与 HTML 同源；HTML 含完整数据表与可点击来源。",
             6.85, 3.4, 5.6, 1.2, size=10.5, color=MUTED, spacing=1.2)
    add_text(s, "v0.1 · 研究版 · 未做本机拆解与台架实测 · 非官方独立研究",
             0.55, 5.65, 12.2, 0.4, size=11, color=MUTED)
    add_footer(s, 20, L)

    source_note(s)
    prs.save(out)
    return out


# ------------------------- yoose deck -------------------------
def build_yoose(out):
    prs = new_deck()
    L = "有色 MINI 5 PRO 拆机研究报告 · v0.1 · 2026-11-27"

    s = cover(prs, "TEARDOWN RESEARCH · 非官方独立研究",
              "有色 MINI 5 PRO", "拆机研究报告",
              "YSFSW50-Z · 电镀银 · 599 元",
              [("117 g", "MINI 5 系列口径"), ("233 Hz", "≈14,000 RPM"),
               ("5 刀头", "4 Lockcut + 1 Precut"), ("15 W", "5V⎓3A · 35min")],
              "证据分级：官方/包装 · 拆解/媒体 · 推断 · 待确认。拆机由机主执行，本报告不代替实测。",
              image="yoose/front.jpg")

    product_gallery(prs, "产品图", "产品外观与官方结构图",
                    "MINI 5 PRO 的电镀银外观、握持比例与礼盒配置。",
                    [("yoose/front.jpg", "整机 · 电镀银正面及侧面"),
                     ("yoose/hand.jpg", "手持比例"),
                     ("yoose/box.jpg", "品牌礼盒与手提袋"),
                     ("yoose/accessories.jpg", "渠道礼盒配件陈列")],
                    "官方内部结构图未取到，原理见第 7 页。礼盒随附配件以实际 SKU 为准。",
                    label=L)

    # 2 summary
    s = new_slide(prs)
    add_header(s, "00", "一页结论", "先给判断，再给依据；未实测的一律标推断。")
    section_cards(s, [
        ("1. 五刀头不是三刀头简单加宽",
         "4 Lockcut + 1 Precut，至少需要两组传动（4+1 或 3+2）。这是最该先确认的结构点，也决定刀头总成能不能整组换。"),
        ("2. 它是 MINI 5 的另一代电气方案",
         "包装 15W / 5V3A / 35min / 420mAh×2；公开的 MINI 5 是 800mAh / 60min / 5W 级。不要把 W5 的拆解当 PRO。"),
        ("3. “14,000 RPM”是借来的单位",
         "直线电机没有转轴，14,000 RPM ≈ 233 Hz；3,240,000 次/分钟与 12,000 CPM 不是同一量纲，不能直接说“270 倍”。"),
        ("4. 可维修性 3/10（推断）",
         "刀头是耗材；双电芯胶粘、五刀头托架、合金壳，开壳即失去 IPX7。与徕芬 P3 同级或略低。"),
    ], h=4.35)
    add_text(s, "一句话：单电机高频 + 五刀头 + 双电芯快充 + 合金壳，和 P3 的“双电机 + 大电池 + 塑壳”不是同一套结构。",
             0.55, 6.42, 12.2, 0.4, size=11, bold=True, color=COPPER)
    add_footer(s, 2, L)

    # 3 baseline
    s = new_slide(prs)
    add_header(s, "01", "产品基线与口径冲突", "以包装与实物铭牌为准；公开页只作对照。")
    add_table(s, ["项目", "值", "证据"], [
        ["型号 / 颜色", "YSFSW50-Z，电镀银（另有午夜蓝/暗夜黑/亚麻白等）", "包装 / 渠道"],
        ["定位 / 价格", "旗舰全合金往复式；官方 599 元，活动价约 509–539", "媒体"],
        ["尺寸 / 重量", "62×39×52 mm；约 117 g（±6 g）", "官方系列资料\n本机待确认"],
        ["外壳 / 刀头", "全合金；5 刀头（4+1），0.05mm 网，45° 逆角", "官方"],
        ["电机", "磁悬浮直线电机，14,000 RPM；3,240,000 次/分钟", "官方"],
        ["电池 / 充电", "420mAh×2；5V⎓3A / 15W；约 35 min", "包装"],
        ["续航 / 防护", "约 90 min；IPX7", "官方 / 包装"],
        ["标准 / 制造", "GB/T 4706.1/.9-2024；深圳加减生活科技", "包装"],
    ], 0.55, 1.95, 8.46, 4.35, col_widths=[1.7, 6.7, 1.6], font_size=10.5)
    photo_panel(s, "yoose/front.jpg", 9.18, 1.95, 3.57, 4.35)
    add_footer(s, 3, L)

    # 4 conflicts
    s = new_slide(prs)
    add_header(s, "02", "电池 / 充电的三套口径", "不是笔误，是不同 SKU 或不同批次；拆机前不取平均。")
    add_table(s, ["来源", "电池", "充电", "备注"], [
        ["包装（本机）", "420 mAh × 2", "15W / 5V3A / 35min", "本报告以此为准"],
        ["海外官方（MINI 5）", "800 mAh", "约 5W / 60min", "MINI 5，不一定同 PRO"],
        ["香港渠道（MINI 5）", "900 mAh", "7 h", "重量 180g，明显另一版"],
    ], 0.55, 1.95, 8.24, 2.6, col_widths=[2.35, 1.31, 2.45, 2.13], font_size=11)
    evidence_panel(s, [
        ("本机包装", "420 mAh × 2 · 15 W · 35 min", GREEN, GREEN_SOFT),
        ("系列 / 渠道口径", "电芯与充电标称需分 SKU 核对", AMBER, AMBER_SOFT),
    ], h=2.6)
    add_text(s, "判断：包装 420mAh×2 / 35min 与 15W 自洽（见后页能量账）；800/900mAh 更像 MINI 5 或不同批次。\n"
                "拆机以电芯丝印、极耳连接、保护板为准，不拿渠道页改包装。",
             0.55, 4.85, 12.2, 1.2, size=12, color=INK, spacing=1.25)
    add_text(s, "“无线充电”：仅媒体口径；包装只列 Type-C 线，没有无线充底座。拆机看底盖内是否有线圈。",
             0.55, 6.05, 12.2, 0.5, size=10.5, color=RED)
    add_footer(s, 4, L)

    # 5 structure
    s = new_slide(prs)
    add_header(s, "03", "结构分层与 BOM", "9 个一级件、约 30 个二级件；11 项待本机确认。")
    layers = [
        ("① 五刀头总成", "4 Lockcut + 1 Precut · 0.05mm 网 · 45° 逆角", GREEN),
        ("② 合金前盖 / 中框", "全合金 · 电镀/烤漆 · 前盖-中框密封（推断）", GREEN),
        ("③ 压板 / 机头固定", "压板螺丝 · 输出轴密封（推断）", AMBER),
        ("④ 机芯：磁悬浮电机", "14,000 RPM ≈ 233 Hz · 五刀头至少两组传动", GREEN),
        ("⑤ PCBA + 双电芯", "15W 充电 · 420mAh×2 · 并/串待确认", RED),
        ("⑥ 密封与附件", "前盖/压板/底盖 · Type-C 密封 · 磁吸保护盖", COPPER),
    ]
    layer_stack(s, layers)
    add_text(s, "四个决定性未知：几组传动 / 双电芯并串 / 电机-主板焊插 / 是否有无线充电线圈。",
             0.55, 6.45, 12.2, 0.35, size=10.5, bold=True, color=RED)
    add_footer(s, 5, L)

    # 6 motor
    s = new_slide(prs)
    add_header(s, "04", "14,000 RPM 与 3,240,000 次", "一个能换算，一个只能复原；厂商没给公式。")
    motor_diagram(s, 0.55, 1.95, 5.95)
    add_rect(s, 6.68, 1.95, 6.07, 1.72, radius=True)
    add_text(s, "14,000 RPM ≈ 233 Hz", 6.88, 2.08, 5.67, 0.4, size=16, bold=True, color=COPPER)
    add_text(s, "14,000 ÷ 60 = 233.3 Hz\n"
                "对照 P3 的 200 Hz，高约 17%，同档。\n"
                "主峰应落在 233±10 Hz；若在 117 Hz，\n说明“14,000”按单程计数。",
             6.88, 2.48, 5.67, 1.08, size=10.5, color=MUTED, spacing=1.25)
    add_rect(s, 6.68, 3.83, 6.07, 1.86, fill=CARD, line=LINE, radius=True)
    add_text(s, "3,240,000 次/分钟：可复原，不是定义", 6.88, 3.95, 5.67, 0.4, size=15, bold=True)
    add_text(s, "233.3 × 2 × 23 × 5 × 60 ≈ 3,220,000\n"
                "与 3,240,000 差 0.6%，取 23 齿、\n14,000 RPM 的取整误差内可以对上。\n"
                "但厂商只给了两个数，没给公式；\n报告只写“可复原”，不写“厂商定义”。",
             6.88, 4.36, 5.67, 1.20, size=10.5, color=MUTED, spacing=1.2)
    add_text(s, "不要把 3,240,000 和 12,000 CPM 并列说“270 倍”：量纲不同。正确比法是 233 Hz vs 200 Hz。",
             0.55, 6.07, 12.2, 0.7, size=12, bold=True, color=RED, spacing=1.2)
    add_footer(s, 6, L)

    # 7 blades
    s = new_slide(prs)
    add_header(s, "05", "五刀头系统", "4 Lockcut 抓短须，1 Precut 先梳后切；五片网 = 五倍破裂概率。")
    add_table(s, ["拓扑", "现象", "含义"], [
        ["A：1 电机 + 5 路", "1 颗驱动 IC；连杆从一处分出", "结构最简，震感大"],
        ["B：1 电机 + 2 组（4+1）", "1 颗 IC，两条连杆", "折中：短须同组、长须独立"],
        ["C：2 电机 + 2 组", "2 颗驱动 IC；两组反向", "成本最高，震感最好"],
    ], 0.55, 1.95, 8.18, 2.7, col_widths=[3.4, 5.0, 3.8], font_size=11)
    photo_panel(s, "yoose/foil.jpg", 8.93, 1.95, 3.82, 2.7)
    add_text(s, "拆机一眼分：看连杆数量、驱动 IC 颗数、手动推刀头的相位（4 短须同相？长须反相？）。\n"
                "倾向 B 或 C：5 刀头 + 117g 仍标低震感，单电机推五组负载偏大。",
             0.55, 4.9, 12.2, 1.0, size=12, color=INK, spacing=1.25)
    add_text(s, "0.05mm 对照 P3 Pro 的 55µm；P3 标准版厚度待确认。刀网凹陷/破裂会刮脸。清洁时勿按压网面。",
             0.55, 6.1, 12.2, 0.5, size=10.5, color=RED)
    add_footer(s, 7, L)

    # 8 power
    s = new_slide(prs)
    add_header(s, "06", "15W 快充与双电芯", "15W 对 840mAh 是 ~1.7C；35 分钟充满，先算能量账。")
    add_rect(s, 0.55, 1.95, 5.8, 3.5, fill=CARD, line=LINE, radius=True)
    add_text(s, "能量账", 0.8, 2.1, 5.2, 0.4, size=15, bold=True)
    add_text(s, "840mAh × 3.7V ≈ 3.11Wh\n"
                "35min = 0.583h\n"
                "3.11 ÷ 0.583 ≈ 5.3W 入电池\n"
                "按 85% 效率 ≈ 6.3W 输入\n\n"
                "15W 只能是峰值；全程平均应在 4.5–7W。",
             0.8, 2.6, 5.2, 2.7, size=11.5, color=MUTED, spacing=1.25)
    add_rect(s, 6.75, 1.95, 6.0, 3.5, fill=CARD, line=LINE, radius=True)
    add_text(s, "双电芯：并还是串", 7.0, 2.1, 5.5, 0.4, size=15, bold=True)
    add_text(s, "并联（最可能）：3.8V / 840mAh；每颗 ~0.85C\n"
                "串联：7.6V / 420mAh；需升压充电\n\n"
                "判据：极耳并到同一焊盘？电池座 2 pin？\n保护板 1 块还是 2 块？",
             7.0, 2.6, 5.5, 2.7, size=11.5, color=MUTED, spacing=1.25)
    flow(s, ("USB-C · 15 W 上限", "充电管理（推断）", "420 mAh × 2（合计 840 mAh）", "约 35 分钟"), 0.55, 5.60, 12.2)
    add_text(s, "90min 续航对应平均放电约 2.1W（840mAh ÷ 1.5h）。实测远高于此，续航标称就偏乐观。",
             0.55, 6.46, 12.2, 0.6, size=11.5, bold=True, color=COPPER, spacing=1.2)
    add_footer(s, 8, L)

    # 9 seal
    s = new_slide(prs)
    add_header(s, "07", "密封与湿态充电", "IPX7 不覆盖“口内有水还能充”；15W/3A 把缺口放大。")
    add_rect(s, 0.55, 1.86, 6.12, 3.02, radius=True)
    add_rect(s, 6.78, 1.86, 5.97, 3.02, fill=RED_SOFT, line=None, radius=True)
    add_text(s, "密封点（推断）", 0.55, 1.95, 6.0, 0.35, size=14, bold=True)
    add_text(s, "• 前盖/刀头面与中框：最大开口，五刀头开窗\n"
                "• 压板与输出轴：轴封随传动组数增加\n"
                "• 底盖与 Type-C 口：高频插拔面\n"
                "• 无线充电线圈窗（若有）：第五个未知",
             0.55, 2.45, 6.0, 2.4, size=11.5, color=MUTED, spacing=1.25)
    add_text(s, "湿态充电为什么更危险", 6.95, 1.95, 5.8, 0.35, size=14, bold=True)
    add_text(s, "• 电解腐蚀：水 + 5V/3A，速率远高于 5V/1A\n"
                "• 局部发热：接触电阻上升 → 发热 → 加速腐蚀\n"
                "• 判据不覆盖：IPX7 不考核带水通电后的性能\n"
                "• 金属壳：水 + 金属 + 充电，接触电流要额外看",
             6.95, 2.45, 5.8, 2.4, size=11.5, color=MUTED, spacing=1.25)
    add_text(s, "“IPX7 合格”只能推出“浸水后无有害进水”，推不出“湿手/湿口可充电”。",
             0.55, 5.35, 12.2, 0.6, size=13, bold=True, color=RED)
    add_footer(s, 9, L)

    # 10 repairability
    s = new_slide(prs)
    add_header(s, "08", "可维修性：3/10（推断）", "合金壳 + 五刀头托架 + 双电芯，复装比 P3 更难。")
    add_text(s, "3", 0.6, 2.0, 1.8, 1.3, size=64, bold=True, color=COPPER)
    add_text(s, "/10", 1.55, 2.55, 1.0, 0.5, size=18, bold=True, color=MUTED)
    score_bar(s, 3, 10)
    add_table(s, ["维度", "得分", "依据"], [
        ["进入机身", "2 / 3", "刀头磁吸；合金壳从底盖/刀头面进，螺丝未知"],
        ["可换件", "1 / 3", "刀头可买；电芯/PCBA/电机/密封件不供件"],
        ["工具与步骤", "0 / 2", "精密批头 + 撬片；五刀头托架 + 双电芯"],
        ["复装 / 防水", "0 / 1", "合金卡扣屈服不回弹；开壳失去 IPX7"],
        ["文档 / 诊断", "0 / 1", "无维修手册、无调试口"],
    ], 2.5, 1.95, 10.2, 3.8, col_widths=[2.2, 1.3, 6.7], font_size=11)
    add_text(s, "对刀头是易维护设计，对机身内部不是。", 0.55, 6.15, 12.2, 0.4,
             size=12, bold=True, color=COPPER)
    add_footer(s, 10, L)

    # 11 FMEA
    s = new_slide(prs)
    add_header(s, "09", "FMEA Top 风险", "两条寿命线：磁悬浮结构洁净度、五片刀网完好度。")
    fmea_table(s, ["ID", "风险", "早期信号", "RPN"], [
        ["F7", "磁悬浮结构 / 弹片疲劳", "噪音变大、震感增强、卡死", "147"],
        ["F1", "5 刀头磨损 / 变形", "剃不净、拉扯、局部漏剃", "144"],
        ["F8", "进液短路", "水雾 / 充电异常", "144"],
        ["F5", "Type-C 口腐蚀", "充不进电、需掰角度", "120"],
        ["F2", "0.05mm 刀网破裂", "刮伤、刀头报废", "112"],
        ["F4", "15W 快充发热", "充电烫手、降速、老化", "108"],
    ], [0.65, 3.5, 3.35, 0.75])
    add_text(s, "完整 F1–F10 与 D1–D8 见 data/fmea.csv；D1 撬合金壳、D2 拆刀头托架、D3 拆双电芯。",
             0.55, 6.1, 12.2, 0.4, size=10.5, color=MUTED)
    add_footer(s, 11, L)

    # 12 compliance
    s = new_slide(prs)
    add_header(s, "10", "合规与专利", "同一套 4706 标准；15W 让温升/输入电流/充电回路权重更高。")
    add_table(s, ["文件", "要点"], [
        ["GB/T 4706.1-2024", "通用安全；10 输入功率、11 温升、13 泄漏/电气强度、15-16 耐潮湿、19 非正常、22 结构"],
        ["GB/T 4706.9-2024", "剃须刀专用；6.2 至少 IPX7；21.101 900mm 五方位跌落；22 结构"],
        ["GB 44246-2024", "家电安全强制规范；4706 系列转为推荐性配套；核 CCC 是否换版"],
        ["IPX7 + 15W", "1m/30min 浸水；不覆盖湿态充电；3A 下腐蚀/发热风险更高"],
    ], 0.55, 1.95, 8.24, 3.5, col_widths=[2.55, 5.69], font_size=10.5)
    evidence_panel(s, [
        ("通用与专用标准", "GB/T 4706.1 与 GB/T 4706.9", BLUE, BLUE_SOFT),
        ("证书 / 专利资料待取", "CCC、IPX7 报告与相关权项", AMBER, AMBER_SOFT),
    ], h=3.5)
    add_text(s, "专利：仅 CN119399865A（解锁控制）一条线索；Google Patents 503，磁悬浮/五刀头权项未取。\n"
                "证书：无 CCC 号、无 IPX7 报告、无说明书 PDF、无保修原文——只写“未找到”。",
             0.55, 5.55, 12.2, 1.0, size=10.5, color=MUTED, spacing=1.2)
    add_footer(s, 12, L)

    # 13 competitors
    s = new_slide(prs)
    add_header(s, "11", "竞品对照", "同价位看路线，跨品牌看技术。")
    competitor_comparison(s, [
        ["yoose MINI 5 PRO", "磁悬浮直线 14,000 RPM", "5", "117g", "90min", "599"],
        ["徕芬 P3", "双直线 L2，200Hz", "3", "187.1g", "90min", "599"],
        ["徕芬 P3 Pro", "双直线 L2", "3", "179g", "100min", "699–799"],
        ["松下 ES-LV67", "直线往复 14,000/min", "5", "—", "50min", "~$150"],
        ["飞利浦 S7888", "旋转式", "3", "—", "60min", "~1,149"],
        ["博朗 3040s", "三组浮动往复", "3", "—", "45min", "284–369"],
    ], [3.0, 4.0, 1.2, 1.4, 1.4, 1.4])
    add_footer(s, 13, L)

    # 14 checklist
    s = new_slide(prs)
    add_header(s, "12", "拆机观察清单", "先拍照、后动手；合金壳先探隐藏螺丝。")
    add_text(s, "必拍 Y01–Y40：四视图 → 五刀头布局 → 刀头背面 → 拆壳 → 密封 → 双电芯 → 电机 → PCBA → 装配",
             0.55, 1.95, 12.2, 0.5, size=11, color=INK)
    add_text(s, "必量：上壳、五刀头托架、电机中心距、电芯尺寸、密封圈截面、行程、重量、刀网厚度",
             0.55, 2.55, 12.2, 0.5, size=11, color=INK)
    add_text(s, "必记：MCU、电机驱动（几颗）、15W 充电 IC、霍尔、保护板、电芯丝印、电机铭牌、认证标识",
             0.55, 3.15, 12.2, 0.45, size=11, color=INK)
    flow(s, ("刀头", "压板", "机芯", "PCBA / 电池", "Type-C"), 0.55, 3.7, 12.2, h=0.51)
    add_text(s, "拆机观察顺序（推断，开壳路径待确认）", 0.55, 4.28, 12.2, 0.25, size=9, color=MUTED)
    add_rect(s, 0.55, 4.63, 12.2, 1.35, fill=COPPER_SOFT, line=None, radius=True)
    add_text(s, "四个决定性问题", 0.8, 4.75, 4.0, 0.35, size=14, bold=True, color=COPPER)
    add_text(s, "① 五刀头几组传动？  ② 双电芯并还是串？\n③ 电机与主板焊还是插？  ④ 有没有无线充电线圈？",
             0.8, 5.12, 11.6, 0.71, size=13, bold=True, color=INK, spacing=1.3)
    add_text(s, "安全：双电芯壁薄、不刺穿；合金卡扣别单点发力；0.05mm 网一按就凹；磁铁旁只用非铁工具。",
             0.55, 6.2, 12.2, 0.4, size=10.5, color=RED)
    add_footer(s, 14, L)

    # 15 gaps
    s = new_slide(prs)
    add_header(s, "13", "资料缺口与边界", "公开资料能回答“大概怎么做的”，不能回答“这台具体怎么做的”。")
    add_rect(s, 0.55, 5.50, 12.2, 0.78, fill=COPPER_SOFT, line=None, radius=True)
    add_text(s, "拆机可回答", 0.55, 1.95, 5.8, 0.35, size=15, bold=True)
    add_text(s, "• 五刀头几组传动、连杆/相位\n• 双电芯并/串、电芯型号、保护板\n• 电机-主板焊/插、驱动 IC 颗数\n• 磁悬浮是真无接触还是弹片混用\n• 外壳固定、密封件、无线充电线圈",
             0.55, 2.45, 5.8, 2.8, size=11.5, color=MUTED, spacing=1.25)
    add_text(s, "拆机也不够", 6.95, 1.95, 5.8, 0.35, size=15, bold=True)
    add_text(s, "• 真实频率、行程、推力-电流\n• 3,240,000 次/分钟的厂商定义\n• 0.05mm 网的刃角/孔型/切割效率\n• 15W 完整充电曲线与温升\n• 复装后 IPX7、长期可靠性",
             6.95, 2.45, 5.8, 2.8, size=11.5, color=MUTED, spacing=1.25)
    add_text(s, "本报告不回答：值不值得买、比 P3 好多少、能用几年。需要统一样本实测与长期数据。",
             0.75, 5.65, 11.8, 0.6, size=13, bold=True, color=COPPER)
    add_footer(s, 15, L)

    # 16 sources
    s = new_slide(prs)
    add_header(s, "14", "来源与下载", "查询日 2026-11-27；官方/包装优先，媒体只作线索。")
    add_text(s, "关键来源", 0.55, 1.95, 3.0, 0.35, size=14, bold=True)
    add_text(s, "• yoose 官网 MINI 5 系列 / EU 页\n"
                "• 京东 MINI 5 PRO 页、CityLink HK\n"
                "• 东方财富 MINI 5 PRO 发布\n"
                "• 火星同学 MINI 5 开箱\n"
                "• GB/T 4706.1/.9-2024",
             0.55, 2.4, 5.6, 2.5, size=11, color=MUTED, spacing=1.25)
    add_rect(s, 6.6, 1.95, 6.15, 3.3, fill=CARD, line=LINE, radius=True)
    add_text(s, "在线报告", 6.85, 2.1, 5.6, 0.35, size=14, bold=True)
    add_text(s, "https://keylliucong.github.io/laifen-p3-teardown-report/yoose-mini5pro/",
             6.85, 2.55, 5.6, 0.8, size=10.5, color=COPPER)
    add_text(s, "仓库：github.com/keylliucong/laifen-p3-teardown-report\n"
                "本 PPT 与 HTML 同源；HTML 含完整数据表与可点击来源。",
             6.85, 3.4, 5.6, 1.2, size=10.5, color=MUTED, spacing=1.2)
    add_text(s, "v0.1 · 研究版 · 未做本机拆解与台架实测 · 非官方独立研究",
             0.55, 5.65, 12.2, 0.4, size=11, color=MUTED)
    add_footer(s, 16, L)

    source_note(s)
    prs.save(out)
    return out


if __name__ == "__main__":
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    p1 = os.path.join(base, "laifen-p3", "laifen-p3-report.pptx")
    p2 = os.path.join(base, "yoose-mini5pro", "yoose-mini5pro-report.pptx")
    build_laifen(p1)
    build_yoose(p2)
    print(p1)
    print(p2)
