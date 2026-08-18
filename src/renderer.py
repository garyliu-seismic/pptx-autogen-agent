"""PPTX 渲染器：用 python-pptx 把 DeckPlan 渲染为标准的、PowerPoint 可打开且可编辑的 .pptx。

所有文本 / 表格 / 图表都是原生 PowerPoint 对象：
- 文本   -> 文本框（可改字）
- 表格   -> 原生表格（可改数）
- 图表   -> 原生图表（双击可编辑数据）
"""
from __future__ import annotations

import datetime
import io
import math
import os
import re

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

import config
from .models import DeckPlan, SlideData
from . import text_metrics

# 16:9 画布
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.6)
CONTENT_W = Emu(SLIDE_W - 2 * MARGIN)   # 保持为 Length，保证 .inches 可用
FONT = "Microsoft YaHei"

# 布局可调参数（供“质检-自动重排”使用）
_FONT_SCALE = 1.0   # 正文字号缩放
_HPAD = 0.0         # 每个文本框额外高度（英寸），防溢出
_GAP_MULT = 1.0     # 行距/卡片间距倍率
# 逐图旋转表：{文件名: 度数(90/180/270)}，来自素材目录的 rotate.json
_ROTATE_MAP: dict = {}


def _fs(size: int) -> int:
    """按 _FONT_SCALE 缩放字号（不缩标题）。"""
    return max(8, int(round(size * _FONT_SCALE)))


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------
def _rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _tint(hex_str: str, factor: float = 0.88) -> str:
    """把颜色向白色混合 factor 比例，生成浅色（用于卡片底色）。"""
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    r2 = int(r + (255 - r) * factor)
    g2 = int(g + (255 - g) * factor)
    b2 = int(b + (255 - b) * factor)
    return f"{r2:02X}{g2:02X}{b2:02X}"


def _style_run(run, size: int, color: str, bold: bool = False,
               font: str = FONT) -> None:
    """设置 run 字体，含东亚字体，保证中文正常显示。"""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", font)


def _set_bg(slide, color: str) -> None:
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(color)


def _textbox(slide, x, y, w, h, text: str, size: int = 18, color: str = "222222",
             bold: bool = False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             font: str = FONT, wrap: bool = True):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = 1.15
    r = p.add_run()
    r.text = text
    _style_run(r, size, color, bold, font)
    return tb


def _bar(slide, x, y, w, h, color: str) -> None:
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(color)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _est_height(text: str, size_pt: int, width_in: float | None = None) -> float:
    """估算文本高度（英寸），中英文混排精确模型，避免重叠。"""
    w = width_in if width_in is not None else CONTENT_W.inches
    return text_metrics.est_height(text, size_pt, w, pad=_HPAD)


# 页型对应的右上角装饰图标（对应“背景图/图标”需求）
_KIND_ICONS = {
    "agenda": "📋", "bullets": "📝", "cards": "🃏", "timeline": "⏱️",
    "table": "🗂️", "chart": "📈", "gallery": "🖼️", "image": "🖼️",
    "figure": "🧭", "summary": "✅",
}


# 卡片文本关键词 -> 图标
_CARD_ICON_RULES = [
    ("工具", "🔧"), ("接口", "🔌"), ("API", "🔌"), ("函数", "⚙️"), ("搜索", "🔍"),
    ("数据", "📊"), ("分析", "📈"), ("图表", "📉"), ("指标", "📉"),
    ("目标", "🎯"), ("规划", "🗺️"), ("步骤", "🗺️"), ("拆解", "🗺️"), ("计划", "🗺️"),
    ("记忆", "🧠"), ("学习", "📚"), ("训练", "📚"), ("预训练", "📚"), ("知识", "📚"),
    ("感知", "👁️"), ("理解", "👁️"), ("观察", "👀"),
    ("安全", "🛡️"), ("风险", "🛡️"), ("隐私", "🔒"),
    ("发布", "🚀"), ("部署", "🚀"), ("上线", "🚀"),
    ("对话", "💬"), ("沟通", "💬"), ("聊天", "💬"),
    ("复用", "📦"), ("封装", "📦"), ("打包", "📦"),
    ("加载", "⚡"), ("速度", "⚡"), ("效率", "⚡"), ("性能", "⚡"),
    ("完成", "✅"), ("稳定", "✅"), ("交付", "✅"),
    ("代码", "💻"), ("编程", "💻"), ("调试", "🐞"),
    ("反馈", "🔁"), ("循环", "🔁"), ("迭代", "🔁"),
    ("示例", "🧪"), ("测试", "🧪"), ("触发", "🎛️"), ("指令", "📣"),
    ("用户", "👤"), ("团队", "👥"), ("时间", "🕐"), ("成本", "💰"),
]


def _card_icon(text: str) -> str:
    for kw, icon in _CARD_ICON_RULES:
        if kw in text:
            return icon
    return "⭐"


def _content_header(slide, theme: config.Theme, title: str,
                    page_no: int, total: int, kind: str = "") -> None:
    _set_bg(slide, theme.bg)
    _bar(slide, 0, 0, SLIDE_W, Inches(0.12), theme.primary)          # 顶部色条
    _textbox(slide, MARGIN, Inches(0.32), CONTENT_W, Inches(0.9),
             title, size=30, color=theme.primary, bold=True)
    _bar(slide, MARGIN, Inches(1.32), CONTENT_W, Pt(2.5), theme.accent)  # 分隔线
    # 右上角“类型图标”（对应背景图/图标）
    icon = _KIND_ICONS.get(kind)
    if icon:
        _textbox(slide, SLIDE_W - Inches(1.5), Inches(0.30), Inches(0.9),
                 Inches(0.9), icon, size=38, color=theme.primary,
                 align=PP_ALIGN.RIGHT)
    _textbox(slide, SLIDE_W - Inches(1.6), SLIDE_H - Inches(0.55),
             Inches(1.0), Inches(0.4), f"{page_no} / {total}",
             size=12, color=theme.secondary, align=PP_ALIGN.RIGHT)


def _footer(slide, theme: config.Theme, text: str) -> None:
    _textbox(slide, MARGIN, SLIDE_H - Inches(0.55), CONTENT_W, Inches(0.4),
             text, size=12, color=theme.secondary)


# ---------------------------------------------------------------------------
# 各页面渲染
# ---------------------------------------------------------------------------
def _render_title(slide, sd: SlideData, theme: config.Theme, plan: DeckPlan) -> None:
    _set_bg(slide, theme.primary)
    _bar(slide, 0, 0, SLIDE_W, Inches(0.18), theme.accent)
    # 封面照片：文字左置、右侧放代表性照片
    photo = sd.images[0] if sd.images else ""
    txt_w = CONTENT_W - Inches(6.4) if photo else CONTENT_W
    _textbox(slide, MARGIN, Inches(2.4), txt_w, Inches(1.9),
             sd.title or plan.title, size=44, color="FFFFFF", bold=True)
    sub = plan.subtitle or (plan.title if sd.title else "")
    _textbox(slide, MARGIN, Inches(4.25), txt_w, Inches(0.8),
             sub, size=20, color="E8E8E8")
    author = plan.author or "Auto-Generated by Agent"
    _textbox(slide, MARGIN, SLIDE_H - Inches(0.9), CONTENT_W, Inches(0.6),
             f"{author}   |   {datetime.date.today().isoformat()}",
             size=14, color="D0D0D0")
    if photo and os.path.isfile(photo):
        _place_image(slide, photo,
                     Emu(int(MARGIN + Inches(7.6))), Emu(int(Inches(1.15))),
                     Inches(4.9), Inches(5.2), theme, valign="center")


def _render_section(slide, sd: SlideData, theme: config.Theme,
                    page_no: int, total: int) -> None:
    _set_bg(slide, theme.primary)
    _bar(slide, 0, Inches(3.0), Inches(0.18), Inches(1.2), theme.accent)
    _textbox(slide, Inches(0.9), Inches(2.9), CONTENT_W, Inches(1.4),
             sd.title, size=40, color="FFFFFF", bold=True)
    _footer(slide, theme, f"{page_no} / {total}")


def _render_agenda(slide, sd: SlideData, theme: config.Theme,
                   page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="agenda")
    for i, item in enumerate(sd.bullets[:8]):
        y = Inches(1.85 + i * 0.62)
        _textbox(slide, MARGIN, y, Inches(0.6), Inches(0.5),
                 f"{i + 1:02d}", size=20, color=theme.accent, bold=True)
        _textbox(slide, MARGIN + Inches(0.65), y, CONTENT_W - Inches(0.65),
                 Inches(0.5), item, size=18, color=theme.fg)


def _render_bullet_col(slide, items: list[str], x, y, w, size: int,
                       gap_in: float, theme: config.Theme) -> None:
    """渲染一列要点；y/gap_in 单位为英寸，w 为该列可用宽度（英寸）。"""
    cy = y
    for b in items:
        h = _est_height(b, size, w - 0.4)
        _textbox(slide, Emu(int(x * 914400)), Emu(int(cy * 914400)),
                 Inches(0.35), Inches(0.42), "▸", size=size + 1,
                 color=theme.accent, bold=True)
        _textbox(slide, Emu(int((x + 0.4) * 914400)), Emu(int(cy * 914400)),
                 Emu(int((w - 0.4) * 914400)), Inches(h), b, size=size,
                 color=theme.fg)
        cy += h + gap_in * _GAP_MULT


def _render_bullets(slide, sd: SlideData, theme: config.Theme,
                    page_no: int, total: int) -> None:
    """要点页：条数少时居中放大留白均匀，条数多时双栏，追求平衡少空白。"""
    _content_header(slide, theme, sd.title, page_no, total, kind="bullets")
    items = [b for b in sd.bullets if b.strip()]
    n = len(items)
    if n == 0:
        return
    if n <= 4:
        size, start_y, gap, two_col = 20, 2.4, 0.34, False
    elif n <= 6:
        size, start_y, gap, two_col = 18, 2.0, 0.26, False
    elif n <= 9:
        size, start_y, gap, two_col = 16, 1.85, 0.16, True
    else:
        size, start_y, gap, two_col = 14, 1.8, 0.12, True
    # 含超长文本时不双栏，避免窄列文字溢出
    if two_col and any(len(b) > 60 for b in items):
        two_col = False
        size = min(size, 15)
        start_y = 1.75
        gap = 0.22
    size = _fs(size)

    if two_col:
        half = math.ceil(n / 2)
        col_w = (CONTENT_W.inches - 0.6) / 2
        _render_bullet_col(slide, items[:half], MARGIN.inches, start_y,
                           col_w, size, gap, theme)
        _render_bullet_col(slide, items[half:], MARGIN.inches + col_w + 0.6,
                           start_y, col_w, size, gap, theme)
    else:
        _render_bullet_col(slide, items, MARGIN.inches, start_y,
                           CONTENT_W.inches, size, gap, theme)


def _render_cards(slide, sd: SlideData, theme: config.Theme,
                  page_no: int, total: int) -> None:
    """卡片页：2~4 条短要点 -> 彩色卡片网格，专业且平衡。"""
    _content_header(slide, theme, sd.title, page_no, total, kind="cards")
    items = [b for b in sd.bullets if b.strip()]
    n = len(items)
    if n == 0:
        return
    cols = 2 if n <= 4 else 3
    rows = math.ceil(n / cols)
    gap = 0.3 * _GAP_MULT
    top = 1.8
    avail_w = CONTENT_W.inches
    avail_h = 4.9
    cell_w = (avail_w - gap * (cols - 1)) / cols
    cell_h = (avail_h - gap * (rows - 1)) / rows
    palette = [theme.primary, theme.secondary, theme.accent]
    for i, item in enumerate(items):
        c, r = i % cols, i // cols
        x = MARGIN.inches + c * (cell_w + gap)
        y = top + r * (cell_h + gap)
        color = palette[i % len(palette)]
        _draw_card(slide, x, y, cell_w - 0.04, cell_h - 0.04, i + 1, item, color, theme)


def _draw_card(slide, x: float, y: float, w: float, h: float,
               idx: int, text: str, color: str, theme: config.Theme) -> None:
    """绘制一张卡片：浅色底 + 彩色描边 + 编号圆角徽章 + 文本。"""
    x0, y0 = int(x * 914400), int(y * 914400)
    w0, h0 = int(w * 914400), int(h * 914400)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x0, y0, w0, h0)
    try:
        card.adjustments[0] = 0.06
    except Exception:
        pass
    card.fill.solid()
    card.fill.fore_color.rgb = _rgb(_tint(color, 0.9))
    card.line.color.rgb = _rgb(color)
    card.line.width = Pt(1.25)
    card.shadow.inherit = False

    # 编号徽章（左上角彩色圆角方块）
    badge = 0.42
    bx, by = int((x + 0.16) * 914400), int((y + 0.16) * 914400)
    badge_shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, int(badge * 914400), int(badge * 914400))
    badge_shp.fill.solid()
    badge_shp.fill.fore_color.rgb = _rgb(color)
    badge_shp.line.fill.background()
    badge_shp.shadow.inherit = False
    tf = badge_shp.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(idx)
    _style_run(r, 16, "FFFFFF", True)

    # 内容图标（右上角，与卡片主题对应）
    icon = _card_icon(text)
    _textbox(slide, Emu(int((x + w - 0.52) * 914400)), Emu(int((y + 0.12) * 914400)),
             Inches(0.4), Inches(0.4), icon, size=20, color=theme.fg,
             align=PP_ALIGN.RIGHT)

    # 文本（垂直居中，避免空白）
    size = _fs(16 if len(text) <= 28 else 14)
    tb = slide.shapes.add_textbox(
        int((x + 0.18) * 914400), int((y + 0.72) * 914400),
        int((w - 0.36) * 914400), int((h - 0.9) * 914400))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = 1.2
    r = p.add_run()
    r.text = text
    _style_run(r, size, theme.fg, False)


def _render_timeline(slide, sd: SlideData, theme: config.Theme,
                     page_no: int, total: int) -> None:
    """时间轴页：适合里程碑 / 分阶段步骤。加高节点、标题下补引导语。"""
    _content_header(slide, theme, sd.title, page_no, total, kind="timeline")
    items = [b for b in sd.bullets if b.strip()]
    n = min(len(items), 6)
    if n == 0:
        return

    # 标题下加一句引导（填补上部空白）
    _textbox(slide, MARGIN, Inches(1.48), CONTENT_W, Inches(0.4),
             f"▎按阶段推进 · 共 {len(items)} 个关键节点", size=13,
             color=theme.secondary)

    line_y = 4.05
    start_x, end_x = 1.1, SLIDE_W.inches - 1.1
    span = end_x - start_x

    # 主轴
    _bar(slide, Emu(int(start_x * 914400)), Emu(int((line_y - 0.02) * 914400)),
         Emu(int(span * 914400)), Pt(3), theme.secondary)

    for i in range(n):
        t = i / (n - 1) if n > 1 else 0.5
        nx = start_x + t * span
        # 节点圆（加大）
        d = 0.60
        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, int((nx - d / 2) * 914400),
            int((line_y - d / 2) * 914400), int(d * 914400), int(d * 914400))
        circ.fill.solid()
        circ.fill.fore_color.rgb = _rgb(theme.primary)
        circ.line.color.rgb = _rgb(theme.bg)
        circ.line.width = Pt(2.5)
        circ.shadow.inherit = False
        ctf = circ.text_frame
        ctf.word_wrap = False
        ctf.margin_left = ctf.margin_right = ctf.margin_top = ctf.margin_bottom = 0
        cp = ctf.paragraphs[0]
        cp.alignment = PP_ALIGN.CENTER
        cr = cp.add_run()
        cr.text = str(i + 1)
        _style_run(cr, 16, "FFFFFF", True)

        # 节点下方文本：等宽平铺，绝不互相重叠/越界
        box_w = span / n - 0.15
        bx = start_x + i * (span / n)
        tb = slide.shapes.add_textbox(
            int(bx * 914400), int((line_y + 0.6) * 914400),
            int(box_w * 914400), int(2.3 * 914400))
        ttf = tb.text_frame
        ttf.word_wrap = True
        ttf.auto_size = MSO_AUTO_SIZE.NONE
        ttf.vertical_anchor = MSO_ANCHOR.TOP
        ttf.margin_left = ttf.margin_right = ttf.margin_top = ttf.margin_bottom = 0
        tp = ttf.paragraphs[0]
        tp.alignment = PP_ALIGN.CENTER
        tp.line_spacing = 1.2
        tr = tp.add_run()
        tr.text = items[i]
        _style_run(tr, _fs(14), theme.fg, False)

    # 底部补充提示
    if n < len(items):
        _footer(slide, theme, f"… 共 {len(items)} 项")


def _render_table(slide, sd: SlideData, theme: config.Theme,
                  page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="table")
    table = sd.table
    nrows = len(table.rows) + 1
    ncols = max(len(table.headers),
                max((len(r) for r in table.rows), default=1))
    nrows = min(nrows, 11)
    top = Inches(1.7)
    height = Inches(0.42 * nrows)
    tbl_shape = slide.shapes.add_table(nrows, ncols, MARGIN, top,
                                       CONTENT_W, height)
    tbl = tbl_shape.table
    col_w = int(CONTENT_W / ncols)
    for j in range(ncols):
        tbl.columns[j].width = col_w

    for j in range(ncols):
        hdr = table.headers[j] if j < len(table.headers) else ""
        _set_cell(tbl.cell(0, j), str(hdr), theme.primary, True)
    for i, row in enumerate(table.rows[: nrows - 1]):
        for j in range(ncols):
            val = row[j] if j < len(row) else ""
            _set_cell(tbl.cell(i + 1, j), str(val),
                      theme.bg if i % 2 == 0 else "FFFFFF", False)


def _set_cell(cell, text: str, fill: str, bold: bool) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = _rgb(fill)
    cell.margin_left = cell.margin_right = Inches(0.08)
    cell.margin_top = cell.margin_bottom = Inches(0.03)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    _style_run(r, 13, "FFFFFF" if bold else "222222", bold)


_CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
}


def _render_chart(slide, sd: SlideData, theme: config.Theme,
                  page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="chart")
    chart = sd.chart
    cdata = CategoryChartData()
    cdata.categories = list(chart.categories)
    for s in chart.series:
        cdata.add_series(s["name"], tuple(s["values"]))
    ctype = _CHART_TYPES.get(chart.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)
    gf = slide.shapes.add_chart(ctype, MARGIN, Inches(1.75),
                                CONTENT_W, Inches(4.35), cdata)
    ch = gf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    try:
        ch.plots[0].gap_width = 80
    except Exception:
        pass

    # 系列配色跟随主题（代替 Office 默认蓝橙）
    palette = [theme.primary, theme.secondary, theme.accent,
               theme.primary, theme.secondary, theme.accent]
    for i, ser in enumerate(ch.series):
        color = palette[i % len(palette)]
        try:
            if chart.chart_type == "line":
                ser.format.line.color.rgb = _rgb(color)
                ser.format.line.width = Pt(2.5)
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = _rgb(color)
        except Exception:
            pass

    # 底部附加精简数据表（PowerPoint 原生表格，可编辑）
    if sd.table and len(sd.table.rows) <= 6:
        t = sd.table
        nr = len(t.rows) + 1
        nc = max(len(t.headers), max((len(r) for r in t.rows), default=1))
        tiny = slide.shapes.add_table(nr, nc, MARGIN, Inches(6.35),
                                      CONTENT_W, Inches(0.85))
        tbl = tiny.table
        cw = int(CONTENT_W / nc)
        for j in range(nc):
            tbl.columns[j].width = cw
        for j in range(nc):
            _set_cell(tbl.cell(0, j), t.headers[j] if j < len(t.headers) else "",
                      theme.secondary, True)
        for i, row in enumerate(t.rows):
            for j in range(nc):
                val = row[j] if j < len(row) else ""
                _set_cell(tbl.cell(i + 1, j), str(val),
                          theme.bg if i % 2 == 0 else "FFFFFF", False)


def _gallery_caption(path: str) -> str:
    """图片说明：优先 EXIF 拍摄时间（MM-DD HH:MM），否则从文件名提炼。"""
    t = _exif_datetime(path)
    if t:
        return t
    name = os.path.splitext(os.path.basename(path))[0]
    name = re.sub(r"^(图_|图-|IMG_|img_|图)", "", name)
    # 中文与英文/数字之间补空格
    name = re.sub(r"(?<=[\u4e00-\u9fff])(?=[A-Za-z0-9])", " ", name)
    name = re.sub(r"(?<=[A-Za-z0-9])(?=[\u4e00-\u9fff])", " ", name)
    return (name.replace("_", " ").strip()) or "图片"


def _exif_datetime(path: str) -> str | None:
    """读取图片 EXIF 拍摄时间，格式化为 'MM-DD HH:MM'；无则 None。"""
    try:
        from PIL import Image
        with Image.open(path) as im:
            ex = im._getexif() or {}
        raw = str(ex.get(36867) or ex.get(306) or "").strip()  # DateTimeOriginal / DateTime
        if not raw:
            return None
        parts = raw.replace(":", "-", 2).split(" ")
        date_part = parts[0].split("-") if parts else []
        time_part = parts[1].split(":") if len(parts) > 1 else []
        if len(date_part) == 3 and len(time_part) >= 2:
            return f"{date_part[1]}-{date_part[2]} {time_part[0]}:{time_part[1]}"
        return raw
    except Exception:
        return None


def _render_gallery(slide, sd: SlideData, theme: config.Theme,
                    page_no: int, total: int) -> None:
    """画廊页：每张图放入圆角卡片（白底+细边），图下方配一行说明，形成设计好的网格。"""
    _content_header(slide, theme, sd.title or "图解与图片", page_no, total,
                    kind="gallery")
    imgs = sd.images
    n = len(imgs)
    if n == 0:
        return
    gap = 0.24
    avail_w = CONTENT_W.inches
    avail_h = 4.95
    if n == 1:
        cols, rows = 1, 1
    else:
        cols = n if n <= 3 else (2 if n == 4 else 3)
        rows = math.ceil(n / cols)
    cell_w = (avail_w - gap * (cols - 1)) / cols
    cell_h = (avail_h - gap * (rows - 1)) / rows
    for i, p in enumerate(imgs):
        c, r = i % cols, i // cols
        x = MARGIN.inches + c * (cell_w + gap)
        y = 1.7 + r * (cell_h + gap)
        _draw_gallery_card(slide, p, x, y, cell_w, cell_h, theme)


def _draw_gallery_card(slide, path: str, x: float, y: float,
                       w: float, h: float, theme: config.Theme) -> None:
    """单张图片卡片：圆角白底容器 + 图片 + 底部说明。"""
    pad = 0.14
    cap_h = 0.40
    # 卡片容器
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        int(x * 914400), int(y * 914400), int(w * 914400), int(h * 914400))
    try:
        card.adjustments[0] = 0.045
    except Exception:
        pass
    card.fill.solid()
    card.fill.fore_color.rgb = _rgb("FFFFFF")
    card.line.color.rgb = _rgb("D9DEE7")
    card.line.width = Pt(1)
    card.shadow.inherit = False

    # 图片（卡片内侧，顶部对齐）
    img_w = w - 2 * pad
    img_h = h - pad - cap_h - pad * 0.5
    _place_image(slide, path,
                 Emu(int((x + pad) * 914400)), Emu(int((y + pad) * 914400)),
                 Emu(int(img_w * 914400)), Emu(int(img_h * 914400)),
                 theme, valign="top")

    # 底部说明
    cap = _gallery_caption(path)
    _textbox(slide,
             Emu(int((x + pad) * 914400)),
             Emu(int((y + h - cap_h - 0.02) * 914400)),
             Emu(int(img_w * 914400)), Inches(0.44),
             cap, size=13, color=theme.fg, bold=True, align=PP_ALIGN.CENTER)


def _draw_photo_card(slide, path: str, x: float, y: float, w: float,
                     h: float, theme: config.Theme) -> None:
    """白色圆角卡片：顶部图片 + 底部拍摄时间说明（照片时间轴用）。"""
    x0, y0 = int(x * 914400), int(y * 914400)
    w0, h0 = int(w * 914400), int(h * 914400)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x0, y0, w0, h0)
    try:
        card.adjustments[0] = 0.06
    except Exception:
        pass
    card.fill.solid()
    card.fill.fore_color.rgb = _rgb("FFFFFF")
    card.line.color.rgb = _rgb(theme.secondary)
    card.line.width = Pt(1)
    card.shadow.inherit = False

    pad = 0.10
    cap_h = 0.34
    img_w = w - 2 * pad
    img_h = h - pad - cap_h - pad * 0.6
    _place_image(slide, path,
                 Emu(int((x + pad) * 914400)), Emu(int((y + pad) * 914400)),
                 Emu(int(img_w * 914400)), Emu(int(img_h * 914400)),
                 theme, valign="center")

    cap = _exif_datetime(path) or _gallery_caption(path)
    _textbox(slide, Emu(int((x + pad) * 914400)),
             Emu(int((y + h - cap_h - 0.02) * 914400)),
             Emu(int(img_w * 914400)), Inches(cap_h + 0.06),
             cap, size=12, color=theme.fg, bold=True, align=PP_ALIGN.CENTER)


def _render_photo_timeline(slide, sd: SlideData, theme: config.Theme,
                           page_no: int, total: int) -> None:
    """照片时间轴页：照片沿中央时间轴上下交错，节点+连线+拍摄时间，按时间顺序。"""
    _content_header(slide, theme, sd.title, page_no, total, kind="timeline")
    imgs = [p for p in sd.images if os.path.isfile(p)][:6]
    n = len(imgs)
    if n == 0:
        return
    axis_y = 4.35
    x0, x1 = 0.8, 12.53
    span = x1 - x0
    slot = span / n

    # 主轴
    _bar(slide, Emu(int(x0 * 914400)), Emu(int((axis_y - 0.02) * 914400)),
         Emu(int(span * 914400)), Pt(3), theme.secondary)

    for i, p in enumerate(imgs):
        cx = x0 + slot * i + slot / 2
        above = (i % 2 == 0)
        card_w = min(2.8, slot - 0.35)
        card_h = 2.1
        if above:
            y0 = axis_y - 0.42 - card_h          # ~1.83
        else:
            y0 = axis_y + 0.42                   # ~4.77
        x = cx - card_w / 2
        _draw_photo_card(slide, p, x, y0, card_w, card_h, theme)

        # 竖直连线：卡片边缘 -> 轴线
        line_top = y0 if above else y0 + card_h
        _bar(slide, Emu(int((cx - 0.01) * 914400)),
             Emu(int(line_top * 914400)), Pt(2),
             Emu(int((axis_y - line_top) * 914400)), theme.accent)

        # 轴线节点
        d = 0.28
        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, int((cx - d / 2) * 914400),
            int((axis_y - d / 2) * 914400), int(d * 914400), int(d * 914400))
        circ.fill.solid()
        circ.fill.fore_color.rgb = _rgb(theme.primary)
        circ.line.color.rgb = _rgb(theme.bg)
        circ.line.width = Pt(2)
        circ.shadow.inherit = False


def _place_image(slide, path: str, x, y, max_w, max_h, theme: config.Theme,
                 valign: str = "center") -> None:
    """等比缩放放置图片（水平居中；valign=top 顶部对齐，否则垂直居中）、加描边。

    支持：EXIF 方向自动转置 + rotate.json 逐图旋转（旋转后保持长宽比，
    重编码为直立像素；未旋转的图保持原文件不重压缩）。
    """
    try:
        from PIL import Image, ImageOps
        rotate_deg = int(_ROTATE_MAP.get(os.path.basename(path), 0) or 0) % 360
        if rotate_deg:
            with Image.open(path) as im:
                im = ImageOps.exif_transpose(im)
                im = im.rotate(rotate_deg, expand=True)
                iw, ih = im.size
                buf = io.BytesIO()
                im.convert("RGB").save(buf, format="JPEG", quality=92)
                src = io.BytesIO(buf.getvalue())
        else:
            with Image.open(path) as im:
                im = ImageOps.exif_transpose(im)
                iw, ih = im.size
            src = path
    except Exception:
        return
    ratio = iw / ih if ih else 1.0
    mw, mh = max_w.inches, max_h.inches
    if mw / mh > ratio:
        w, h = mh * ratio, mh
    else:
        w, h = mw, mw / ratio
    # 全程使用 emu（x/y 已是 Length=emu），避免英寸与 emu 混用
    xx = int(x) + int(round((mw - w) / 2 * 914400))
    if valign == "top":
        yy = int(y)
    else:
        yy = int(y) + int(round((mh - h) / 2 * 914400))
    pic = slide.shapes.add_picture(src, xx, yy,
                                   int(round(w * 914400)), int(round(h * 914400)))
    pic.line.color.rgb = _rgb(theme.secondary)
    pic.line.width = Pt(1)


def _render_figure(slide, sd: SlideData, theme: config.Theme,
                   page_no: int, total: int) -> None:
    """图文讲解页：左侧大图，右侧逐条讲解要点。"""
    _content_header(slide, theme, sd.title, page_no, total, kind="figure")
    # 左侧大图（顶部对齐，等比缩放）
    if sd.images:
        _place_image(slide, sd.images[0], MARGIN, Inches(1.75),
                     Inches(7.0), Inches(5.05), theme, valign="top")
    # 右侧说明要点
    items = [b for b in sd.bullets if b.strip()]
    if not items:
        return
    size = _fs(17 if len(items) <= 5 else 15)
    tx = Emu(int(MARGIN + Inches(7.35)))
    tw = Emu(int(CONTENT_W - Inches(7.35)))
    cy = Inches(1.95)
    for b in items:
        h = _est_height(b, size, tw.inches - 0.4)
        _textbox(slide, tx, cy, Inches(0.35), Inches(0.42), "▸",
                 size=size + 1, color=theme.accent, bold=True)
        _textbox(slide, Emu(int(tx + Inches(0.4))), cy,
                 Emu(int(tw - Inches(0.4))), Inches(h), b, size=size,
                 color=theme.fg)
        cy = Emu(int(cy + Inches(h + 0.2)))

    # 一句话结论条（填补右下空白，突出要点）
    if sd.takeaway:
        bar_y = Emu(int(Inches(6.0)))
        bar_h = Inches(0.64)
        bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, int(tx), int(bar_y), int(tw),
            int(bar_h))
        try:
            bar.adjustments[0] = 0.18
        except Exception:
            pass
        bar.fill.solid()
        bar.fill.fore_color.rgb = _rgb(_tint(theme.accent, 0.86))
        bar.line.color.rgb = _rgb(theme.accent)
        bar.line.width = Pt(1)
        bar.shadow.inherit = False
        btf = bar.text_frame
        btf.word_wrap = True
        btf.auto_size = MSO_AUTO_SIZE.NONE
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        btf.margin_left = btf.margin_right = Inches(0.1)
        btf.margin_top = btf.margin_bottom = Inches(0.02)
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.LEFT
        bp.line_spacing = 1.1
        br = bp.add_run()
        br.text = "💡 " + sd.takeaway
        _style_run(br, _fs(13), theme.fg, True)


def _render_summary(slide, sd: SlideData, theme: config.Theme,
                    page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="summary")
    y = Inches(1.8)
    for i, b in enumerate(sd.bullets):
        h = _est_height(b, _fs(18))
        _textbox(slide, MARGIN, y, Inches(0.5), Inches(0.5),
                 "✔", size=_fs(18), color=theme.accent, bold=True)
        _textbox(slide, MARGIN + Inches(0.55), y, CONTENT_W - Inches(0.55),
                 Inches(h), b, size=_fs(18), color=theme.fg)
        y += Inches(h + 0.12)


def _render_thanks(slide, sd: SlideData, theme: config.Theme,
                   plan: DeckPlan) -> None:
    _set_bg(slide, theme.primary)
    _bar(slide, 0, 0, SLIDE_W, Inches(0.18), theme.accent)
    _textbox(slide, MARGIN, Inches(2.9), CONTENT_W, Inches(1.2),
             sd.title or "谢谢", size=48, color="FFFFFF", bold=True,
             align=PP_ALIGN.CENTER)
    _textbox(slide, MARGIN, Inches(4.3), CONTENT_W, Inches(0.7),
             "欢迎提问与交流  |  Q&A", size=18, color="D0D0D0",
             align=PP_ALIGN.CENTER)


# ---------------------------------------------------------------------------
# 主渲染
# ---------------------------------------------------------------------------
def render(plan: DeckPlan, output_path: str, *,
           font_scale: float = 1.0, hpad: float = 0.0,
           gap_mult: float = 1.0, rotate_map: dict | None = None) -> str:
    """渲染 DeckPlan 为 .pptx。

    可选布局参数（供 quality.render_with_quality 自动重排调用）：
      font_scale : 正文字号缩放（0.75~1.2）
      hpad       : 每个文本框额外高度（英寸），解决文本溢出/重叠
      gap_mult   : 行距与卡片间距倍率
      rotate_map : {文件名: 度数}，逐图旋转（配合素材目录 rotate.json）
    """
    global _FONT_SCALE, _HPAD, _GAP_MULT, _ROTATE_MAP
    _FONT_SCALE = font_scale
    _HPAD = hpad
    _GAP_MULT = gap_mult
    _ROTATE_MAP = rotate_map or {}

    theme = config.THEMES.get(plan.theme_name, config.THEMES[config.DEFAULT_THEME])

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    total = len(plan.slides)
    for idx, sd in enumerate(plan.slides):
        slide = prs.slides.add_slide(blank)
        page_no = idx + 1
        kind = sd.kind
        if kind == "title":
            _render_title(slide, sd, theme, plan)
        elif kind == "section":
            _render_section(slide, sd, theme, page_no, total)
        elif kind == "agenda":
            _render_agenda(slide, sd, theme, page_no, total)
        elif kind == "table":
            _render_table(slide, sd, theme, page_no, total)
        elif kind == "chart":
            _render_chart(slide, sd, theme, page_no, total)
        elif kind in ("gallery", "image"):
            _render_gallery(slide, sd, theme, page_no, total)
        elif kind == "photo_timeline":
            _render_photo_timeline(slide, sd, theme, page_no, total)
        elif kind == "cards":
            _render_cards(slide, sd, theme, page_no, total)
        elif kind == "timeline":
            _render_timeline(slide, sd, theme, page_no, total)
        elif kind == "figure":
            _render_figure(slide, sd, theme, page_no, total)
        elif kind == "summary":
            _render_summary(slide, sd, theme, page_no, total)
        elif kind == "thanks":
            _render_thanks(slide, sd, theme, plan)
        else:
            _render_bullets(slide, sd, theme, page_no, total)

        # 备注
        notes = sd.notes or ""
        try:
            if notes:
                slide.notes_slide.notes_text_frame.text = notes
        except Exception:
            pass

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)
    return output_path
