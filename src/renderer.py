"""PPTX 渲染器：用 python-pptx 把 DeckPlan 渲染为标准的、PowerPoint 可打开且可编辑的 .pptx。

所有文本 / 表格 / 图表都是原生 PowerPoint 对象：
- 文本   -> 文本框（可改字）
- 表格   -> 原生表格（可改数）
- 图表   -> 原生图表（双击可编辑数据）

页型清单（kind）：
  结构页: title / section / agenda / thanks
  内容页: bullets / cards / timeline / table / chart
          gallery / image / figure / photo_timeline / summary
  新增页: kpi / quote / process / comparison
"""
from __future__ import annotations

import colorsys
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
from .slide_import import fix_app_properties_slide_count, remove_slide

# kind -> 优先尝试匹配的 layout 名称关键词（小写包含匹配，取第一个命中的）
_KIND_LAYOUT_KEYWORDS = {
    "title":   ["front cover", "cover"],
    "section": ["divider", "back cover"],
    "agenda":  ["agenda"],
    "thanks":  ["back cover", "thank"],
}
_DEFAULT_LAYOUT_KEYWORDS = ["content title", "content title only", "content blank", "blank"]


def _pick_layout(layouts: list, kind: str):
    keywords = _KIND_LAYOUT_KEYWORDS.get(kind, _DEFAULT_LAYOUT_KEYWORDS)
    for kw in keywords:
        for layout in layouts:
            if kw in (layout.name or "").lower():
                return layout
    return layouts[0]


def _scale_new_shapes(slide, inherited: set, scale: float) -> None:
    """把新加形状按 scale 整体缩放位置/尺寸/字号，映射到模板真实画布。"""
    for sh in slide.shapes:
        if sh._element in inherited:
            continue
        if sh.left is not None:
            sh.left = Emu(int(sh.left * scale))
        if sh.top is not None:
            sh.top = Emu(int(sh.top * scale))
        if sh.width is not None:
            sh.width = Emu(int(sh.width * scale))
        if sh.height is not None:
            sh.height = Emu(int(sh.height * scale))
        if sh.has_text_frame:
            _scale_text_frame_fonts(sh.text_frame, scale)
        if getattr(sh, "has_table", False):
            for cell in sh.table.iter_cells():
                _scale_text_frame_fonts(cell.text_frame, scale)


def _scale_text_frame_fonts(text_frame, scale: float) -> None:
    for p in text_frame.paragraphs:
        for r in p.runs:
            if r.font.size is not None:
                r.font.size = Pt(r.font.size.pt * scale)


# ---------------------------------------------------------------------------
# 画布常量
# ---------------------------------------------------------------------------
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN  = Inches(0.6)
CONTENT_W = Emu(SLIDE_W - 2 * MARGIN)

# 布局可调参数（供质检自动重排使用）
_FONT_SCALE = 1.0
_HPAD       = 0.0
_GAP_MULT   = 1.0
_ROTATE_MAP: dict = {}


def _fs(size: int) -> int:
    """按 _FONT_SCALE 缩放字号（不缩小于 8pt）。"""
    return max(8, int(round(size * _FONT_SCALE)))


# ---------------------------------------------------------------------------
# 颜色工具
# ---------------------------------------------------------------------------
def _rgb(hex_str: str) -> RGBColor:
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _tint(hex_str: str, factor: float = 0.88) -> str:
    """颜色向白色混合（生成浅色）。"""
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return "{:02X}{:02X}{:02X}".format(
        int(r + (255 - r) * factor),
        int(g + (255 - g) * factor),
        int(b + (255 - b) * factor),
    )


def _shade(hex_str: str, factor: float = 0.85) -> str:
    """颜色向黑色混合（生成深色）。"""
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return "{:02X}{:02X}{:02X}".format(
        int(r * (1 - factor)),
        int(g * (1 - factor)),
        int(b * (1 - factor)),
    )


def _luminance(hex_str: str) -> float:
    """相对亮度（0=黑, 1=白）。"""
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255
    return 0.299 * r + 0.587 * g + 0.114 * b


def _is_dark(hex_str: str) -> bool:
    return _luminance(hex_str) < 0.45


def _is_dark_theme(theme: config.Theme) -> bool:
    return _is_dark(theme.bg)


def _auto_fg(bg: str, light: str = "FFFFFF", dark: str = "111827") -> str:
    """根据背景色亮度自动选前景文字色。"""
    return light if _is_dark(bg) else dark


def _derive_accent_palette(theme: config.Theme, n: int) -> list[str]:
    """从主题派生 n 个渐变色（图表多系列/卡片/时间轴着色用）。

    深色主题下 primary/secondary 往往过暗，自动改用 accent 系列亮色组。
    """
    if _is_dark_theme(theme):
        # 深色主题：用 accent + success + warning 等亮色，避免深色 primary/secondary
        colors = [theme.accent, theme.success, theme.warning,
                  _tint(theme.accent, 0.25), _tint(theme.success, 0.25)]
    else:
        colors = [theme.primary, theme.secondary, theme.accent]
    if n <= len(colors):
        return colors[:n]
    # 在 HSV 空间均匀采样
    result: list[str] = []
    for i in range(n):
        t = i / n
        src = colors[int(t * (len(colors) - 1))]
        h = src.lstrip("#")
        r, g, b = int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255
        hv, sv, vv = colorsys.rgb_to_hsv(r, g, b)
        hv = (hv + t * 0.15) % 1.0
        r2, g2, b2 = colorsys.hsv_to_rgb(hv, sv, vv)
        result.append("{:02X}{:02X}{:02X}".format(int(r2 * 255), int(g2 * 255), int(b2 * 255)))
    return result


def _card_bg(color: str, theme: config.Theme) -> str:
    """卡片背景色。

    深色主题：如果 color 本身是深色，用微亮版本（shade→深，但要防止更黑）
               如果 color 是亮色，用深色背景 + 亮色边框
    浅色主题：用 color 的淡色版本（tint）
    """
    if theme.bg_card:
        return theme.bg_card
    if _is_dark_theme(theme):
        # 深色主题：卡片用比背景稍亮的色，确保可见度
        # 如果 color 亮度 < 0.3，就 tint 出一个可辨识的深卡片色
        if _luminance(color) < 0.30:
            return _tint(color, 0.15)   # 微亮但不刺眼
        return _shade(color, 0.55)      # 已经是亮色，稍微压暗
    return _tint(color, 0.90)


# ---------------------------------------------------------------------------
# 基础绘图工具
# ---------------------------------------------------------------------------
def _style_run(run, size: int, color: str, bold: bool = False,
               font: str | None = None) -> None:
    """设置 run 字体（含东亚字体 a:ea，保证中文正常渲染）。"""
    if font is None:
        font = config.THEMES.get("modern").font_minor  # 兜底
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


def _textbox(slide, x, y, w, h, text: str,
             size: int = 18, color: str = "222222",
             bold: bool = False, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, font: str | None = None,
             wrap: bool = True, line_spacing: float = 1.2):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    r = p.add_run()
    r.text = text
    _style_run(r, size, color, bold, font)
    return tb


def _bar(slide, x, y, w, h, color: str):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(color)
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _rounded_rect(slide, x, y, w, h, color: str, line_color: str = "",
                  line_pt: float = 0, radius: float = 0.07):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 int(x), int(y), int(w), int(h))
    try:
        shp.adjustments[0] = radius
    except Exception:
        pass
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(color)
    if line_color and line_pt > 0:
        shp.line.color.rgb = _rgb(line_color)
        shp.line.width = Pt(line_pt)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _est_height(text: str, size_pt: int, width_in: float | None = None) -> float:
    w = width_in if width_in is not None else CONTENT_W.inches
    return text_metrics.est_height(text, size_pt, w, pad=_HPAD)


# ---------------------------------------------------------------------------
# 页型图标（右上角装饰）
# ---------------------------------------------------------------------------
_KIND_ICONS = {
    "agenda":     "📋", "bullets":    "📝", "cards":      "🃏",
    "timeline":   "⏱️", "table":      "🗂️", "chart":      "📈",
    "gallery":    "🖼️", "image":      "🖼️", "figure":     "🧭",
    "summary":    "✅", "kpi":        "📊", "quote":      "💬",
    "process":    "🔄", "comparison": "⚖️",
}

_CARD_ICON_RULES = [
    ("工具", "🔧"), ("接口", "🔌"), ("API", "🔌"), ("函数", "⚙️"), ("搜索", "🔍"),
    ("数据", "📊"), ("分析", "📈"), ("图表", "📉"), ("指标", "📉"),
    ("目标", "🎯"), ("规划", "🗺️"), ("步骤", "🗺️"), ("拆解", "🗺️"), ("计划", "🗺️"),
    ("记忆", "🧠"), ("学习", "📚"), ("训练", "📚"), ("知识", "📚"),
    ("感知", "👁️"), ("理解", "👁️"), ("观察", "👀"),
    ("安全", "🛡️"), ("风险", "🛡️"), ("隐私", "🔒"),
    ("发布", "🚀"), ("部署", "🚀"), ("上线", "🚀"),
    ("对话", "💬"), ("沟通", "💬"), ("聊天", "💬"),
    ("复用", "📦"), ("封装", "📦"), ("打包", "📦"),
    ("加载", "⚡"), ("速度", "⚡"), ("效率", "⚡"), ("性能", "⚡"),
    ("完成", "✅"), ("稳定", "✅"), ("交付", "✅"),
    ("代码", "💻"), ("编程", "💻"), ("调试", "🐞"),
    ("反馈", "🔁"), ("循环", "🔁"), ("迭代", "🔁"),
    ("用户", "👤"), ("团队", "👥"), ("时间", "🕐"), ("成本", "💰"),
    ("增长", "📈"), ("下降", "📉"), ("优化", "⚡"), ("创新", "💡"),
]


def _card_icon(text: str) -> str:
    for kw, icon in _CARD_ICON_RULES:
        if kw in text:
            return icon
    return "⭐"


# ---------------------------------------------------------------------------
# 主题自适应颜色工具
# ---------------------------------------------------------------------------
def _title_color(theme: config.Theme) -> str:
    """内容页标题文字色：深色背景用 fg（浅色），浅色背景用 primary（深色）。"""
    return theme.fg if _is_dark_theme(theme) else theme.primary


def _icon_color(theme: config.Theme) -> str:
    """右上角图标色：深色背景用 accent 亮色，浅色背景用 primary 深色。"""
    return theme.accent if _is_dark_theme(theme) else theme.primary


def _subtitle_color(theme: config.Theme) -> str:
    """次级文字（说明/小标题）自适应色。"""
    return _tint(theme.fg, 0.20) if _is_dark_theme(theme) else theme.secondary


# ---------------------------------------------------------------------------
# 通用页头（内容页复用）
# ---------------------------------------------------------------------------
def _content_header(slide, theme: config.Theme, title: str,
                    page_no: int, total: int, kind: str = "") -> None:
    _set_bg(slide, theme.bg)
    title_clr = _title_color(theme)
    icon_clr  = _icon_color(theme)
    # 顶部主题色条
    _bar(slide, 0, 0, SLIDE_W, Inches(0.10), theme.accent)
    # 左侧竖向强调块
    _bar(slide, MARGIN - Inches(0.10), Inches(0.28),
         Inches(0.055), Inches(0.80), theme.accent)
    # 标题
    _textbox(slide, MARGIN, Inches(0.28), CONTENT_W - Inches(1.6), Inches(0.85),
             title, size=theme.size_slide_title, color=title_clr, bold=True,
             font=theme.font_major, line_spacing=1.1)
    # 细分隔线
    _bar(slide, MARGIN, Inches(1.18), CONTENT_W, Pt(2), theme.accent)
    # 页型图标（右上角）
    icon = _KIND_ICONS.get(kind)
    if icon:
        _textbox(slide, SLIDE_W - Inches(1.5), Inches(0.22),
                 Inches(0.9), Inches(0.9), icon, size=32, color=icon_clr,
                 align=PP_ALIGN.RIGHT)
    # 页码
    _textbox(slide, SLIDE_W - Inches(1.6), SLIDE_H - Inches(0.50),
             Inches(1.0), Inches(0.35), f"{page_no} / {total}",
             size=theme.size_footer, color=theme.neutral, align=PP_ALIGN.RIGHT)


def _footer(slide, theme: config.Theme, text: str) -> None:
    _textbox(slide, MARGIN, SLIDE_H - Inches(0.50), CONTENT_W, Inches(0.35),
             text, size=theme.size_footer, color=theme.neutral)


# ---------------------------------------------------------------------------
# ① 封面页
# ---------------------------------------------------------------------------
def _render_title(slide, sd: SlideData, theme: config.Theme,
                  plan: DeckPlan) -> None:
    dark_bg = theme.bg_dark or theme.primary
    _set_bg(slide, dark_bg)

    # 顶部强调色条（较粗）
    _bar(slide, 0, 0, SLIDE_W, Inches(0.22), theme.accent)
    # 底部细线
    _bar(slide, 0, SLIDE_H - Inches(0.18), SLIDE_W, Inches(0.18), theme.accent)

    photo = sd.images[0] if sd.images else ""
    txt_w = CONTENT_W - Inches(6.0) if photo else CONTENT_W

    # 左侧装饰竖条
    _bar(slide, MARGIN - Inches(0.05), Inches(1.8),
         Inches(0.06), Inches(3.2), theme.accent)

    # 主标题
    _textbox(slide, MARGIN + Inches(0.12), Inches(1.9), txt_w, Inches(2.2),
             sd.title or plan.title, size=theme.size_cover_title,
             color="FFFFFF", bold=True, font=theme.font_major,
             line_spacing=1.15, wrap=True)

    # 副标题
    sub = plan.subtitle or ""
    if sub:
        _textbox(slide, MARGIN + Inches(0.12), Inches(4.1), txt_w, Inches(0.8),
                 sub, size=theme.size_cover_sub,
                 color=_tint(theme.accent, 0.30), font=theme.font_minor,
                 line_spacing=1.1)

    # 作者 + 日期
    author = plan.author or "Auto-Generated"
    _textbox(slide, MARGIN + Inches(0.12), SLIDE_H - Inches(1.0),
             txt_w, Inches(0.55),
             f"{author}   ·   {datetime.date.today().isoformat()}",
             size=theme.size_footer + 1,
             color=_tint(theme.neutral, 0.40), font=theme.font_minor)

    # 封面图片（右半区）
    if photo and os.path.isfile(photo):
        _place_image(slide, photo,
                     Emu(int(MARGIN + Inches(7.4))), Emu(int(Inches(0.5))),
                     Inches(4.9), Inches(6.5), theme, valign="center")


# ---------------------------------------------------------------------------
# ② 章节分隔页
# ---------------------------------------------------------------------------
def _render_section(slide, sd: SlideData, theme: config.Theme,
                    page_no: int, total: int) -> None:
    dark_bg = theme.bg_dark or theme.primary
    _set_bg(slide, dark_bg)

    # 左侧彩色宽条（装饰）
    _bar(slide, 0, 0, Inches(0.40), SLIDE_H, theme.accent)
    # 底部装饰细条（替代右侧线，避免超画布右边界）
    _bar(slide, 0, SLIDE_H - Inches(0.12), SLIDE_W, Inches(0.12), theme.secondary)

    # 章节编号背景块
    _rounded_rect(slide, int(Inches(0.9)), int(Inches(2.6)),
                  int(Inches(1.4)), int(Inches(1.0)),
                  theme.accent, radius=0.10)
    _textbox(slide, Inches(0.9), Inches(2.62), Inches(1.4), Inches(0.9),
             f"{page_no:02d}", size=38, color="FFFFFF", bold=True,
             font=theme.font_major, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 章节标题
    _textbox(slide, Inches(2.6), Inches(2.45), CONTENT_W - Inches(2.0), Inches(1.5),
             sd.title, size=theme.size_section, color="FFFFFF", bold=True,
             font=theme.font_major, line_spacing=1.1)

    # 页码
    _textbox(slide, Inches(0.9), SLIDE_H - Inches(0.65),
             Inches(2.0), Inches(0.40),
             f"{page_no} / {total}", size=theme.size_footer,
             color=_tint(theme.neutral, 0.20))


# ---------------------------------------------------------------------------
# ③ 目录页
# ---------------------------------------------------------------------------
def _render_agenda(slide, sd: SlideData, theme: config.Theme,
                   page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="agenda")
    items = sd.bullets[:8]
    n = len(items)
    if n == 0:
        return
    # 两栏排列（≥5 条）
    cols = 2 if n >= 5 else 1
    per_col = math.ceil(n / cols)
    col_w = (CONTENT_W.inches - 0.5) / cols

    for idx, item in enumerate(items):
        col = idx // per_col
        row = idx % per_col
        x = MARGIN.inches + col * (col_w + 0.5)
        y = 1.55 + row * 0.68

        # 编号圆角块（深色主题用 accent 做背景保证可见）
        badge_bg = theme.accent if _is_dark_theme(theme) else theme.primary
        _rounded_rect(slide,
                      int(x * 914400), int(y * 914400),
                      int(0.48 * 914400), int(0.48 * 914400),
                      badge_bg, radius=0.12)
        _textbox(slide, Emu(int(x * 914400)), Emu(int(y * 914400)),
                 Inches(0.48), Inches(0.48),
                 f"{idx + 1:02d}", size=16, color="FFFFFF", bold=True,
                 font=theme.font_major, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # 条目文字
        _textbox(slide,
                 Emu(int((x + 0.58) * 914400)), Emu(int((y + 0.02) * 914400)),
                 Emu(int((col_w - 0.62) * 914400)), Inches(0.50),
                 item, size=_fs(17), color=theme.fg, font=theme.font_minor,
                 anchor=MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------------------
# ④ 要点页
# ---------------------------------------------------------------------------
def _render_bullet_col(slide, items: list[str], x, y, w, size: int,
                       gap_in: float, theme: config.Theme) -> None:
    cy = y
    for b in items:
        h = _est_height(b, size, w - 0.45)
        # 箭头符号
        _textbox(slide, Emu(int(x * 914400)), Emu(int(cy * 914400)),
                 Inches(0.38), Inches(0.42),
                 "›", size=size + 4, color=theme.accent, bold=True)
        _textbox(slide, Emu(int((x + 0.42) * 914400)), Emu(int(cy * 914400)),
                 Emu(int((w - 0.42) * 914400)), Inches(h),
                 b, size=size, color=theme.fg, font=theme.font_minor,
                 line_spacing=theme.line_height)
        cy += h + gap_in * _GAP_MULT


def _render_bullets(slide, sd: SlideData, theme: config.Theme,
                    page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="bullets")
    items = [b for b in sd.bullets if b.strip()]
    n = len(items)
    if n == 0:
        return
    if n <= 3:
        size, start_y, gap, two_col = theme.size_body, 2.2, 0.40, False
    elif n <= 5:
        size, start_y, gap, two_col = theme.size_body, 1.85, 0.30, False
    elif n <= 7:
        size, start_y, gap, two_col = theme.size_body_sm + 2, 1.75, 0.20, True
    else:
        size, start_y, gap, two_col = theme.size_body_sm, 1.70, 0.14, True

    if two_col and any(len(b) > 55 for b in items):
        two_col = False
        size = min(size, theme.size_body_sm + 1)
        start_y = 1.65
        gap = 0.20
    size = _fs(size)

    if two_col:
        half = math.ceil(n / 2)
        col_w = (CONTENT_W.inches - 0.7) / 2
        _render_bullet_col(slide, items[:half], MARGIN.inches, start_y,
                           col_w, size, gap, theme)
        _render_bullet_col(slide, items[half:],
                           MARGIN.inches + col_w + 0.7,
                           start_y, col_w, size, gap, theme)
    else:
        _render_bullet_col(slide, items, MARGIN.inches, start_y,
                           CONTENT_W.inches, size, gap, theme)


# ---------------------------------------------------------------------------
# ⑤ 卡片页（升级版：渐变色边框 + 图标 + 阴影可选）
# ---------------------------------------------------------------------------
def _render_cards(slide, sd: SlideData, theme: config.Theme,
                  page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="cards")
    items = [b for b in sd.bullets if b.strip()]
    n = len(items)
    if n == 0:
        return

    cols = 2 if n <= 4 else 3
    rows = math.ceil(n / cols)
    gap = 0.28 * _GAP_MULT
    top = 1.62
    avail_w = CONTENT_W.inches
    avail_h = 5.5 - top + 1.38
    cell_w = (avail_w - gap * (cols - 1)) / cols
    cell_h = (avail_h - gap * (rows - 1)) / rows
    palette = _derive_accent_palette(theme, max(n, 3))

    for i, item in enumerate(items):
        c, r = i % cols, i // cols
        x = MARGIN.inches + c * (cell_w + gap)
        y = top + r * (cell_h + gap)
        _draw_card(slide, x, y, cell_w - 0.02, cell_h - 0.02,
                   i + 1, item, palette[i % len(palette)], theme)


def _draw_card(slide, x: float, y: float, w: float, h: float,
               idx: int, text: str, color: str,
               theme: config.Theme) -> None:
    """单张卡片：圆角矩形 + 彩色顶条 + 编号徽章 + 图标 + 文本。"""
    x0, y0 = int(x * 914400), int(y * 914400)
    w0, h0 = int(w * 914400), int(h * 914400)
    bg = _card_bg(color, theme)
    fg_text = _auto_fg(bg)

    # 卡片主体
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x0, y0, w0, h0)
    try:
        card.adjustments[0] = theme.card_radius
    except Exception:
        pass
    card.fill.solid()
    card.fill.fore_color.rgb = _rgb(bg)
    card.line.color.rgb = _rgb(color)
    card.line.width = Pt(1.5)
    card.shadow.inherit = False

    # 顶部色条（彩色）
    top_h = int(0.10 * 914400)
    top_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     x0, y0, w0, top_h)
    try:
        top_bar.adjustments[0] = theme.card_radius
    except Exception:
        pass
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = _rgb(color)
    top_bar.line.fill.background()
    top_bar.shadow.inherit = False

    # 编号徽章（左上角）
    bd = 0.38
    bx, by = int((x + 0.14) * 914400), int((y + 0.18) * 914400)
    badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, bx, by,
                                   int(bd * 914400), int(bd * 914400))
    badge.fill.solid()
    badge.fill.fore_color.rgb = _rgb(color)
    badge.line.fill.background()
    badge.shadow.inherit = False
    btf = badge.text_frame
    btf.word_wrap = False
    btf.margin_left = btf.margin_right = btf.margin_top = btf.margin_bottom = 0
    bp = btf.paragraphs[0]
    bp.alignment = PP_ALIGN.CENTER
    br = bp.add_run()
    br.text = str(idx)
    _style_run(br, 14, "FFFFFF", True, theme.font_major)

    # 内容图标（右上角）
    icon = _card_icon(text)
    _textbox(slide,
             Emu(int((x + w - 0.50) * 914400)),
             Emu(int((y + 0.14) * 914400)),
             Inches(0.38), Inches(0.38),
             icon, size=18, color=fg_text, align=PP_ALIGN.RIGHT)

    # 卡片文本
    size = _fs(16 if len(text) <= 30 else 14)
    tb = slide.shapes.add_textbox(
        int((x + 0.16) * 914400), int((y + 0.68) * 914400),
        int((w - 0.32) * 914400), int((h - 0.86) * 914400))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = theme.line_height
    r = p.add_run()
    r.text = text
    _style_run(r, size, fg_text, False, theme.font_minor)


# ---------------------------------------------------------------------------
# ⑥ 时间轴页
# ---------------------------------------------------------------------------
def _render_timeline(slide, sd: SlideData, theme: config.Theme,
                     page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="timeline")
    items = [b for b in sd.bullets if b.strip()]
    n = min(len(items), 8)
    if n == 0:
        return

    _textbox(slide, MARGIN, Inches(1.40), CONTENT_W, Inches(0.38),
             f"▎按阶段推进 · 共 {len(items)} 个关键节点",
             size=12, color=_subtitle_color(theme), font=theme.font_minor)

    line_y = 4.10
    start_x = 1.0
    end_x   = SLIDE_W.inches - 1.0
    span    = end_x - start_x

    # 主轴（双线：底层浅色粗线 + 上层主色细线）
    _bar(slide, Emu(int(start_x * 914400)), Emu(int((line_y - 0.025) * 914400)),
         Emu(int(span * 914400)), Pt(5), _tint(theme.secondary, 0.60))
    _bar(slide, Emu(int(start_x * 914400)), Emu(int((line_y - 0.012) * 914400)),
         Emu(int(span * 914400)), Pt(2.5), theme.primary)

    palette = _derive_accent_palette(theme, n)

    for i in range(n):
        t = i / (n - 1) if n > 1 else 0.5
        nx = start_x + t * span
        color = palette[i % len(palette)]

        # 节点圆
        d = 0.56
        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            int((nx - d / 2) * 914400), int((line_y - d / 2) * 914400),
            int(d * 914400), int(d * 914400))
        circ.fill.solid()
        circ.fill.fore_color.rgb = _rgb(color)
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
        _style_run(cr, 14, "FFFFFF", True, theme.font_major)

        # 节点文本（等宽平铺）
        box_w = span / n - 0.12
        bx = start_x + i * (span / n)
        tb = slide.shapes.add_textbox(
            int(bx * 914400), int((line_y + 0.52) * 914400),
            int(box_w * 914400), int(2.5 * 914400))
        ttf = tb.text_frame
        ttf.word_wrap = True
        ttf.auto_size = MSO_AUTO_SIZE.NONE
        ttf.vertical_anchor = MSO_ANCHOR.TOP
        ttf.margin_left = ttf.margin_right = ttf.margin_top = ttf.margin_bottom = 0
        tp = ttf.paragraphs[0]
        tp.alignment = PP_ALIGN.CENTER
        tp.line_spacing = theme.line_height
        tr = tp.add_run()
        tr.text = items[i]
        _style_run(tr, _fs(13), theme.fg, False, theme.font_minor)

    if n < len(items):
        _footer(slide, theme, f"… 共 {len(items)} 项")


# ---------------------------------------------------------------------------
# ⑦ 表格页
# ---------------------------------------------------------------------------
def _render_table(slide, sd: SlideData, theme: config.Theme,
                  page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="table")
    table = sd.table
    nrows = min(len(table.rows) + 1, 14)
    ncols = max(len(table.headers),
                max((len(r) for r in table.rows), default=1))
    top    = Inches(1.50)
    height = Inches(min(0.44 * nrows, 5.6))

    tbl_shape = slide.shapes.add_table(nrows, ncols, MARGIN, top, CONTENT_W, height)
    tbl = tbl_shape.table
    col_w = int(CONTENT_W / ncols)
    for j in range(ncols):
        tbl.columns[j].width = col_w

    # 表头
    for j in range(ncols):
        hdr = table.headers[j] if j < len(table.headers) else ""
        _set_cell(tbl.cell(0, j), str(hdr), theme.primary, True, theme)

    # 数据行（交替底色）
    even_bg = _tint(theme.primary, 0.93)
    for i, row in enumerate(table.rows[: nrows - 1]):
        for j in range(ncols):
            val = row[j] if j < len(row) else ""
            bg = even_bg if i % 2 == 0 else "FFFFFF"
            _set_cell(tbl.cell(i + 1, j), str(val), bg, False, theme)


def _set_cell(cell, text: str, fill: str, bold: bool,
              theme: config.Theme) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = _rgb(fill)
    cell.margin_left = cell.margin_right = Inches(0.10)
    cell.margin_top  = cell.margin_bottom = Inches(0.04)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    fg = "FFFFFF" if bold else theme.fg
    _style_run(r, _fs(13), fg, bold,
               theme.font_major if bold else theme.font_minor)


# ---------------------------------------------------------------------------
# ⑧ 图表页
# ---------------------------------------------------------------------------
_CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar":    XL_CHART_TYPE.BAR_CLUSTERED,
    "line":   XL_CHART_TYPE.LINE_MARKERS,
    "pie":    XL_CHART_TYPE.PIE,
}
_C_NS = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _color_pie_slices(ch, palette: list[str]) -> None:
    """给饼图每个扇区独立着色。

    python-pptx 对饼图系列的 ser.format.fill 只能设置整个系列色（导致全部扇区同色）。
    正确做法：直接操作 XML，向系列内插入 <c:dPt> 节点，每个对应一个扇区。
    """
    from lxml import etree
    C = _C_NS  # {http://...chart}
    A = _A_NS  # http://...main

    for ser in ch.series:
        ser_el = ser._element  # <c:ser>
        # 找到已有的 dPt 节点，先全部删除
        for old in ser_el.findall(f"{C}dPt"):
            ser_el.remove(old)

        # 找到 <c:dLbls> 的位置（dPt 要插入在 dLbls 之前）
        dLbls = ser_el.find(f"{C}dLbls")
        cat    = ser_el.find(f"{C}cat")
        # 如果找不到 cat，用 series 长度估算
        try:
            n_pts = len(ser.values)
        except Exception:
            n_pts = len(palette)

        insert_before = dLbls  # dPt 必须在 dLbls 前

        for idx in range(n_pts):
            color_hex = palette[idx % len(palette)]
            r, g, b = int(color_hex[0:2], 16), int(color_hex[2:4], 16), int(color_hex[4:6], 16)

            dPt = etree.Element(f"{C}dPt")
            # <c:idx val="idx"/>
            idx_el = etree.SubElement(dPt, f"{C}idx", val=str(idx))
            # <c:invertIfNegative val="0"/>
            etree.SubElement(dPt, f"{C}invertIfNegative", val="0")
            # <c:spPr> → <a:solidFill> → <a:srgbClr>
            spPr = etree.SubElement(dPt, f"{C}spPr")
            solidFill = etree.SubElement(spPr, f"{{{A}}}solidFill")
            srgbClr = etree.SubElement(solidFill, f"{{{A}}}srgbClr",
                                       val=f"{r:02X}{g:02X}{b:02X}")
            # 线条：用白色细线分割扇区
            ln = etree.SubElement(spPr, f"{{{A}}}ln", w="12700")  # 1pt
            etree.SubElement(etree.SubElement(ln, f"{{{A}}}solidFill"),
                             f"{{{A}}}srgbClr", val="FFFFFF")

            if insert_before is not None:
                ser_el.insert(list(ser_el).index(insert_before), dPt)
            else:
                ser_el.append(dPt)


def _fix_negative_axis_ids(chart) -> None:
    """修复 python-pptx 生成的负数 axId（schema 要求正整数）。"""
    mapping: dict[str, str] = {}
    next_id = 100000001
    for el in chart._chartSpace.iter():
        if el.tag not in (_C_NS + "axId", _C_NS + "crossAx"):
            continue
        val = el.get("val")
        if val is None:
            continue
        n = int(val)
        if 0 <= n <= 2147483647:
            continue
        if val not in mapping:
            mapping[val] = str(next_id)
            next_id += 1
        el.set("val", mapping[val])


def _render_chart(slide, sd: SlideData, theme: config.Theme,
                  page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="chart")
    chart = sd.chart
    cdata = CategoryChartData()
    cdata.categories = list(chart.categories)
    for s in chart.series:
        cdata.add_series(s["name"], tuple(s["values"]))
    ctype = _CHART_TYPES.get(chart.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    # pie 图稍小，柱/线图用满高度
    chart_h = Inches(4.55) if chart.chart_type == "pie" else Inches(4.80)
    gf = slide.shapes.add_chart(ctype, MARGIN, Inches(1.42), CONTENT_W, chart_h, cdata)
    ch = gf.chart
    _fix_negative_axis_ids(ch)

    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False

    # 柱宽 / 间距
    try:
        ch.plots[0].gap_width = 70
    except Exception:
        pass

    # 数据标签
    try:
        plot = ch.plots[0]
        plot.has_data_labels = True
        plot.data_labels.font.size = Pt(9)
        plot.data_labels.number_format = "0.##"
        plot.data_labels.number_format_is_linked = False
    except Exception:
        pass

    dark = _is_dark_theme(theme)

    # 坐标轴 / 网格线样式（深色主题保证可读）
    if dark:
        try:
            for ax in (ch.category_axis, ch.value_axis):
                ax.tick_labels.font.color.rgb = _rgb(theme.fg)
                ax.tick_labels.font.size = Pt(10)
                ax.format.line.color.rgb = _rgb(theme.neutral)
            ch.value_axis.has_major_gridlines = True
            ch.value_axis.major_gridlines.format.line.color.rgb = _rgb(
                _shade(theme.neutral, 0.50))
            ch.legend.font.color.rgb = _rgb(theme.fg)
            ch.legend.font.size = Pt(10)
            ch.plots[0].data_labels.font.color.rgb = _rgb(theme.fg)
        except Exception:
            pass

    # 系列/数据点颜色
    n_series = len(ch.series)
    if chart.chart_type == "pie":
        # 饼图：每个扇区需独立着色（按 category 数量派生足够多的颜色）
        n_cats = len(chart.categories) if chart.categories else n_series
        palette = _derive_accent_palette(theme, max(n_cats, 4))
        _color_pie_slices(ch, palette)
    else:
        palette = _derive_accent_palette(theme, max(n_series, 1))
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

    # 底部附加精简数据表
    if sd.table and len(sd.table.rows) <= 6:
        t = sd.table
        nr = len(t.rows) + 1
        nc = max(len(t.headers), max((len(r) for r in t.rows), default=1))
        tiny = slide.shapes.add_table(nr, nc, MARGIN, Inches(6.40),
                                      CONTENT_W, Inches(0.80))
        tbl = tiny.table
        cw = int(CONTENT_W / nc)
        for j in range(nc):
            tbl.columns[j].width = cw
        for j in range(nc):
            _set_cell(tbl.cell(0, j), t.headers[j] if j < len(t.headers) else "",
                      theme.secondary, True, theme)
        for i, row in enumerate(t.rows):
            for j in range(nc):
                bg = _tint(theme.primary, 0.93) if i % 2 == 0 else "FFFFFF"
                _set_cell(tbl.cell(i + 1, j), str(row[j] if j < len(row) else ""),
                          bg, False, theme)


# ---------------------------------------------------------------------------
# ⑨ 画廊页
# ---------------------------------------------------------------------------
def _gallery_caption(path: str) -> str:
    t = _exif_datetime(path)
    if t:
        return t
    name = os.path.splitext(os.path.basename(path))[0]
    name = re.sub(r"^(图_|图-|IMG_|img_|图)", "", name)
    name = re.sub(r"(?<=[\u4e00-\u9fff])(?=[A-Za-z0-9])", " ", name)
    name = re.sub(r"(?<=[A-Za-z0-9])(?=[\u4e00-\u9fff])", " ", name)
    return (name.replace("_", " ").strip()) or "图片"


def _exif_datetime(path: str) -> str | None:
    try:
        from PIL import Image
        with Image.open(path) as im:
            ex = im._getexif() or {}
        raw = str(ex.get(36867) or ex.get(306) or "").strip()
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
    _content_header(slide, theme, sd.title or "图解与图片", page_no, total, kind="gallery")
    imgs = sd.images
    n = len(imgs)
    if n == 0:
        return
    gap = 0.22
    avail_w = CONTENT_W.inches
    avail_h = 5.10
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
        y = 1.58 + r * (cell_h + gap)
        _draw_gallery_card(slide, p, x, y, cell_w, cell_h, theme)


def _draw_gallery_card(slide, path: str, x: float, y: float,
                       w: float, h: float, theme: config.Theme) -> None:
    pad = 0.12
    cap_h = 0.38
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        int(x * 914400), int(y * 914400), int(w * 914400), int(h * 914400))
    try:
        card.adjustments[0] = 0.04
    except Exception:
        pass
    card.fill.solid()
    card.fill.fore_color.rgb = _rgb("FFFFFF")
    card.line.color.rgb = _rgb(_tint(theme.primary, 0.60))
    card.line.width = Pt(1)
    card.shadow.inherit = False

    img_w = w - 2 * pad
    img_h = h - pad - cap_h - pad * 0.5
    _place_image(slide, path,
                 Emu(int((x + pad) * 914400)), Emu(int((y + pad) * 914400)),
                 Emu(int(img_w * 914400)), Emu(int(img_h * 914400)),
                 theme, valign="top")

    cap = _gallery_caption(path)
    _textbox(slide,
             Emu(int((x + pad) * 914400)),
             Emu(int((y + h - cap_h - 0.02) * 914400)),
             Emu(int(img_w * 914400)), Inches(0.40),
             cap, size=_fs(theme.size_caption), color=theme.fg,
             bold=True, align=PP_ALIGN.CENTER, font=theme.font_minor)


# ---------------------------------------------------------------------------
# ⑩ 照片时间轴
# ---------------------------------------------------------------------------
def _draw_photo_card(slide, path: str, x: float, y: float, w: float,
                     h: float, theme: config.Theme) -> None:
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
             cap, size=_fs(12), color=theme.fg, bold=True,
             align=PP_ALIGN.CENTER, font=theme.font_minor)


def _render_photo_timeline(slide, sd: SlideData, theme: config.Theme,
                           page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="timeline")
    imgs = [p for p in sd.images if os.path.isfile(p)][:6]
    n = len(imgs)
    if n == 0:
        return
    axis_y = 4.35
    x0, x1 = 0.8, 12.53
    span = x1 - x0
    slot = span / n

    _bar(slide, Emu(int(x0 * 914400)), Emu(int((axis_y - 0.02) * 914400)),
         Emu(int(span * 914400)), Pt(3), theme.secondary)

    for i, p in enumerate(imgs):
        cx = x0 + slot * i + slot / 2
        above = (i % 2 == 0)
        card_w = min(2.8, slot - 0.35)
        card_h = 2.1
        y0_c = axis_y - 0.42 - card_h if above else axis_y + 0.42
        x_c  = cx - card_w / 2
        _draw_photo_card(slide, p, x_c, y0_c, card_w, card_h, theme)

        if above:
            line_top = y0_c + card_h
            line_h   = axis_y - line_top
        else:
            line_top = axis_y
            line_h   = y0_c - axis_y
        _bar(slide, Emu(int((cx - 0.01) * 914400)),
             Emu(int(line_top * 914400)), Pt(2),
             Emu(int(line_h * 914400)), theme.accent)

        d = 0.28
        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            int((cx - d / 2) * 914400), int((axis_y - d / 2) * 914400),
            int(d * 914400), int(d * 914400))
        circ.fill.solid()
        circ.fill.fore_color.rgb = _rgb(theme.primary)
        circ.line.color.rgb = _rgb(theme.bg)
        circ.line.width = Pt(2)
        circ.shadow.inherit = False


# ---------------------------------------------------------------------------
# ⑪ 图文讲解页
# ---------------------------------------------------------------------------
def _render_figure(slide, sd: SlideData, theme: config.Theme,
                   page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="figure")
    if sd.images:
        _place_image(slide, sd.images[0], MARGIN, Inches(1.55),
                     Inches(6.8), Inches(5.20), theme, valign="top")
    items = [b for b in sd.bullets if b.strip()]
    if not items:
        return
    size = _fs(17 if len(items) <= 5 else 15)
    tx = Emu(int(MARGIN + Inches(7.2)))
    tw = Emu(int(CONTENT_W - Inches(7.2)))
    cy = Inches(1.72)
    for b in items:
        h = _est_height(b, size, tw.inches - 0.42)
        _textbox(slide, tx, cy, Inches(0.35), Inches(0.42),
                 "›", size=size + 4, color=theme.accent, bold=True)
        _textbox(slide, Emu(int(tx + Inches(0.40))), cy,
                 Emu(int(tw - Inches(0.40))), Inches(h),
                 b, size=size, color=theme.fg, font=theme.font_minor,
                 line_spacing=theme.line_height)
        cy = Emu(int(cy + Inches(h + 0.18)))

    if sd.takeaway:
        bar_y = Emu(int(Inches(6.05)))
        bar_h = Inches(0.62)
        bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, int(tx), int(bar_y), int(tw), int(bar_h))
        try:
            bar.adjustments[0] = 0.18
        except Exception:
            pass
        bar.fill.solid()
        bar.fill.fore_color.rgb = _rgb(_tint(theme.accent, 0.85))
        bar.line.color.rgb = _rgb(theme.accent)
        bar.line.width = Pt(1)
        bar.shadow.inherit = False
        btf = bar.text_frame
        btf.word_wrap = True
        btf.auto_size = MSO_AUTO_SIZE.NONE
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        btf.margin_left = btf.margin_right = Inches(0.10)
        btf.margin_top = btf.margin_bottom = Inches(0.02)
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.LEFT
        bp.line_spacing = 1.1
        br = bp.add_run()
        br.text = "💡 " + sd.takeaway
        _style_run(br, _fs(13), theme.fg, True, theme.font_minor)


# ---------------------------------------------------------------------------
# ⑫ 总结页
# ---------------------------------------------------------------------------
def _render_summary(slide, sd: SlideData, theme: config.Theme,
                    page_no: int, total: int) -> None:
    _content_header(slide, theme, sd.title, page_no, total, kind="summary")
    items = [b for b in sd.bullets if b.strip()]
    y = Inches(1.65)
    for b in items:
        h = _est_height(b, _fs(theme.size_body))
        # 绿色勾选圆角块
        _rounded_rect(slide,
                      int(MARGIN), int(y),
                      int(Inches(0.44)), int(Inches(0.44)),
                      theme.success, radius=0.14)
        _textbox(slide, MARGIN, y, Inches(0.44), Inches(0.44),
                 "✓", size=_fs(16), color="FFFFFF", bold=True,
                 font=theme.font_major, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        _textbox(slide, MARGIN + Inches(0.55), y,
                 CONTENT_W - Inches(0.55), Inches(h),
                 b, size=_fs(theme.size_body), color=theme.fg,
                 font=theme.font_minor, line_spacing=theme.line_height)
        y += Inches(h + 0.16)


# ---------------------------------------------------------------------------
# ⑬ 致谢页
# ---------------------------------------------------------------------------
def _render_thanks(slide, sd: SlideData, theme: config.Theme,
                   plan: DeckPlan) -> None:
    dark_bg = theme.bg_dark or theme.primary
    _set_bg(slide, dark_bg)
    _bar(slide, 0, 0, SLIDE_W, Inches(0.22), theme.accent)
    _bar(slide, 0, SLIDE_H - Inches(0.18), SLIDE_W, Inches(0.18), theme.accent)

    # 大字"谢谢"
    _textbox(slide, MARGIN, Inches(2.55), CONTENT_W, Inches(1.5),
             sd.title or "谢谢", size=theme.size_cover_title + 4,
             color="FFFFFF", bold=True, font=theme.font_major,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 副语
    _textbox(slide, MARGIN, Inches(4.15), CONTENT_W, Inches(0.65),
             "欢迎提问与交流  ·  Q & A",
             size=theme.size_cover_sub, color=_tint(theme.accent, 0.30),
             font=theme.font_minor, align=PP_ALIGN.CENTER)

    # 作者
    if plan.author:
        _textbox(slide, MARGIN, SLIDE_H - Inches(0.95), CONTENT_W, Inches(0.50),
                 plan.author, size=theme.size_footer + 1,
                 color=_tint(theme.neutral, 0.30),
                 font=theme.font_minor, align=PP_ALIGN.CENTER)


# ===========================================================================
# 新增页型
# ===========================================================================

# ---------------------------------------------------------------------------
# ⑭ KPI 看板页
# ---------------------------------------------------------------------------
def _render_kpi(slide, sd: SlideData, theme: config.Theme,
                page_no: int, total: int) -> None:
    """KPI 看板：3~6 个大数字指标卡，每卡含指标名/数值/趋势箭头/补充说明。"""
    _content_header(slide, theme, sd.title, page_no, total, kind="kpi")
    metrics = sd.metrics
    n = len(metrics)
    if n == 0:
        return

    cols = min(n, 3)
    rows = math.ceil(n / cols)
    gap  = 0.30
    top  = 1.55
    avail_w = CONTENT_W.inches
    avail_h = 5.60
    cell_w = (avail_w - gap * (cols - 1)) / cols
    cell_h = (avail_h - gap * (rows - 1)) / rows

    trend_icon  = {"up": "▲", "down": "▼", "flat": "─"}
    trend_color = {"up": theme.success, "down": theme.warning, "flat": theme.neutral}

    for i, m in enumerate(metrics):
        c = i % cols
        r = i // cols
        x = MARGIN.inches + c * (cell_w + gap)
        y = top + r * (cell_h + gap)

        # 卡片背景：深色主题用稍亮的深色卡片，浅色主题用淡primary
        kpi_card_base = theme.secondary if _is_dark_theme(theme) else theme.primary
        bg = _card_bg(kpi_card_base, theme)
        kpi_line_color = theme.accent if _is_dark_theme(theme) else theme.primary
        _rounded_rect(slide,
                      int(x * 914400), int(y * 914400),
                      int(cell_w * 914400), int(cell_h * 914400),
                      bg, line_color=kpi_line_color, line_pt=1.2,
                      radius=theme.card_radius)

        # 顶部色条（用 accent 保证深色主题下可见）
        top_h = int(0.09 * 914400)
        top_bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            int(x * 914400), int(y * 914400),
            int(cell_w * 914400), top_h)
        try:
            top_bar.adjustments[0] = theme.card_radius
        except Exception:
            pass
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = _rgb(theme.accent)
        top_bar.line.fill.background()
        top_bar.shadow.inherit = False

        fg = _auto_fg(bg)

        # 指标名（上方小字）
        label_y = y + 0.18
        _textbox(slide,
                 Emu(int((x + 0.18) * 914400)),
                 Emu(int(label_y * 914400)),
                 Emu(int((cell_w - 0.36) * 914400)),
                 Inches(0.42),
                 m.label or "指标",
                 size=_fs(theme.size_kpi_label),
                 color=fg, font=theme.font_minor)

        # 大数值（居中，主角）
        value_y = y + label_y - y + 0.50
        _textbox(slide,
                 Emu(int((x + 0.10) * 914400)),
                 Emu(int(value_y * 914400)),
                 Emu(int((cell_w - 0.20) * 914400)),
                 Inches(1.10),
                 m.value or "—",
                 size=_fs(theme.size_kpi_value),
                 color=m.color or theme.accent,
                 bold=True, font=theme.font_mono,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # 趋势箭头 + 备注（下方）
        t = m.trend or "flat"
        ticon  = trend_icon.get(t, "─")
        tcolor = trend_color.get(t, theme.neutral)
        note   = m.note or ""
        bottom_y = y + cell_h - 0.48
        _textbox(slide,
                 Emu(int((x + 0.18) * 914400)),
                 Emu(int(bottom_y * 914400)),
                 Emu(int(0.36 * 914400)),
                 Inches(0.38),
                 ticon,
                 size=_fs(16), color=tcolor, bold=True,
                 font=theme.font_major)
        if note:
            _textbox(slide,
                     Emu(int((x + 0.56) * 914400)),
                     Emu(int(bottom_y * 914400)),
                     Emu(int((cell_w - 0.72) * 914400)),
                     Inches(0.40),
                     note,
                     size=_fs(11), color=fg,
                     font=theme.font_minor)


# ---------------------------------------------------------------------------
# ⑮ 金句/引语页
# ---------------------------------------------------------------------------
def _render_quote(slide, sd: SlideData, theme: config.Theme,
                  page_no: int, total: int) -> None:
    """金句页：大字引语 + 作者署名，视觉冲击力强，适合放在章节之间提升节奏。"""
    dark_bg = theme.bg_dark or theme.primary
    _set_bg(slide, dark_bg)
    _bar(slide, 0, 0, SLIDE_W, Inches(0.18), theme.accent)

    quote_text = sd.takeaway or (sd.bullets[0] if sd.bullets else sd.title)
    author     = sd.author or ""

    # 巨型引号装饰（左上）
    _textbox(slide, Inches(0.35), Inches(0.55), Inches(2.0), Inches(2.0),
             "\u201c",   # 左双引号
             size=180, color=_tint(theme.accent, 0.20),
             bold=True, font=theme.font_major, wrap=False)

    # 引语正文
    q_len = len(quote_text)
    q_size = 34 if q_len <= 30 else (28 if q_len <= 60 else 22)
    _textbox(slide, Inches(1.2), Inches(2.0),
             SLIDE_W - Inches(2.4), Inches(3.2),
             quote_text,
             size=_fs(q_size), color="FFFFFF", bold=True,
             font=theme.font_major,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.35, wrap=True)

    # 右下角引号装饰
    _textbox(slide, SLIDE_W - Inches(2.2), Inches(4.0),
             Inches(2.0), Inches(2.0),
             "\u201d",   # 右双引号
             size=180, color=_tint(theme.accent, 0.20),
             bold=True, font=theme.font_major, wrap=False)

    # 分隔线 + 作者
    if author:
        _bar(slide,
             Emu(int(SLIDE_W / 2 - Inches(1.5))),
             Inches(5.40), Inches(3.0), Pt(2), theme.accent)
        _textbox(slide, MARGIN, Inches(5.60), CONTENT_W, Inches(0.60),
                 f"—— {author}",
                 size=_fs(16), color=_tint(theme.accent, 0.30),
                 font=theme.font_minor, align=PP_ALIGN.CENTER)

    # 页码
    _textbox(slide, SLIDE_W - Inches(1.6), SLIDE_H - Inches(0.50),
             Inches(1.0), Inches(0.35), f"{page_no} / {total}",
             size=theme.size_footer, color=theme.neutral, align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------------------
# ⑯ 流程图页
# ---------------------------------------------------------------------------
def _render_process(slide, sd: SlideData, theme: config.Theme,
                    page_no: int, total: int) -> None:
    """流程图页：横向 A → B → C → D，带箭头连接，适合步骤/工作流。"""
    _content_header(slide, theme, sd.title, page_no, total, kind="process")
    items = [b for b in sd.bullets if b.strip()]
    n = len(items)
    if n == 0:
        return
    n = min(n, 6)   # 最多 6 步

    gap_arrow = 0.55
    avail_w   = CONTENT_W.inches
    node_w    = (avail_w - gap_arrow * (n - 1)) / n
    node_h    = 2.80
    top_y     = 2.55
    palette   = _derive_accent_palette(theme, n)

    for i in range(n):
        x = MARGIN.inches + i * (node_w + gap_arrow)
        color = palette[i % len(palette)]
        bg    = _card_bg(color, theme)
        fg    = _auto_fg(bg)

        # 节点矩形（圆角）
        _rounded_rect(slide,
                      int(x * 914400), int(top_y * 914400),
                      int(node_w * 914400), int(node_h * 914400),
                      bg, line_color=color, line_pt=1.5,
                      radius=theme.card_radius)

        # 顶部彩色步骤块
        step_h = 0.52
        step_bg = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            int(x * 914400), int(top_y * 914400),
            int(node_w * 914400), int(step_h * 914400))
        try:
            step_bg.adjustments[0] = theme.card_radius
        except Exception:
            pass
        step_bg.fill.solid()
        step_bg.fill.fore_color.rgb = _rgb(color)
        step_bg.line.fill.background()
        step_bg.shadow.inherit = False

        # 步骤编号
        _textbox(slide,
                 Emu(int(x * 914400)), Emu(int(top_y * 914400)),
                 Emu(int(node_w * 914400)), Emu(int(step_h * 914400)),
                 f"Step {i + 1}",
                 size=_fs(13), color="FFFFFF", bold=True,
                 font=theme.font_major,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # 节点内容文字
        txt_y = top_y + step_h + 0.16
        txt_h = node_h - step_h - 0.24
        _textbox(slide,
                 Emu(int((x + 0.12) * 914400)),
                 Emu(int(txt_y * 914400)),
                 Emu(int((node_w - 0.24) * 914400)),
                 Emu(int(txt_h * 914400)),
                 items[i],
                 size=_fs(14 if len(items[i]) <= 20 else 12),
                 color=fg, font=theme.font_minor,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                 line_spacing=theme.line_height, wrap=True)

        # 箭头（节点之间，最后一个不画）
        if i < n - 1:
            ax = x + node_w + 0.06
            ay = top_y + node_h / 2
            # 箭身
            _bar(slide,
                 Emu(int(ax * 914400)), Emu(int((ay - 0.025) * 914400)),
                 Emu(int((gap_arrow - 0.22) * 914400)), Pt(3.5),
                 theme.accent)
            # 箭头三角（用细长矩形模拟）
            arr_x = ax + gap_arrow - 0.22
            arr = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                int(arr_x * 914400), int((ay - 0.14) * 914400),
                int(0.20 * 914400), int(0.28 * 914400))
            arr.fill.solid()
            arr.fill.fore_color.rgb = _rgb(theme.accent)
            arr.line.fill.background()
            arr.shadow.inherit = False


# ---------------------------------------------------------------------------
# ⑰ 对比页
# ---------------------------------------------------------------------------
def _render_comparison(slide, sd: SlideData, theme: config.Theme,
                       page_no: int, total: int) -> None:
    """对比页：左右双栏，两套要点对照，适合 Before/After 或方案对比。"""
    _content_header(slide, theme, sd.title, page_no, total, kind="comparison")

    half_w = (CONTENT_W.inches - 0.40) / 2
    left_x  = MARGIN.inches
    right_x = MARGIN.inches + half_w + 0.40
    top_y   = 1.50

    label_l = sd.label_left  or "现状 / Before"
    label_r = sd.label_right or "目标 / After"
    items_l = [b for b in sd.bullets if b.strip()]
    items_r = [b for b in sd.bullets_right if b.strip()]

    # 深色主题用亮色对比色以保证可见
    dark = _is_dark_theme(theme)
    left_color  = theme.secondary if not dark else _tint(theme.secondary, 0.40)
    right_color = theme.accent

    # 左栏背景
    _rounded_rect(slide,
                  int(left_x * 914400), int(top_y * 914400),
                  int(half_w * 914400), int(5.60 * 914400),
                  _card_bg(left_color, theme),
                  line_color=left_color, line_pt=1.2,
                  radius=theme.card_radius)

    # 右栏背景（强调色，更醒目）
    _rounded_rect(slide,
                  int(right_x * 914400), int(top_y * 914400),
                  int(half_w * 914400), int(5.60 * 914400),
                  _card_bg(right_color, theme),
                  line_color=right_color, line_pt=1.5,
                  radius=theme.card_radius)

    # 中间 VS 分隔圆（用 accent 保证可见）
    cx = MARGIN.inches + half_w + 0.20
    cy = top_y + 2.80
    _rounded_rect(slide,
                  int((cx - 0.30) * 914400), int((cy - 0.30) * 914400),
                  int(0.60 * 914400), int(0.60 * 914400),
                  theme.accent, radius=0.50)
    _textbox(slide,
             Emu(int((cx - 0.30) * 914400)), Emu(int((cy - 0.30) * 914400)),
             Inches(0.60), Inches(0.60),
             "VS",
             size=_fs(14), color="FFFFFF", bold=True,
             font=theme.font_major,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    def _draw_col(label: str, items: list[str], x: float, col_color: str) -> None:
        # 标签头
        _textbox(slide,
                 Emu(int((x + 0.16) * 914400)), Emu(int((top_y + 0.14) * 914400)),
                 Emu(int((half_w - 0.32) * 914400)), Inches(0.48),
                 label,
                 size=_fs(16), color=col_color, bold=True,
                 font=theme.font_major)
        _bar(slide,
             Emu(int((x + 0.12) * 914400)),
             Emu(int((top_y + 0.64) * 914400)),
             Emu(int((half_w - 0.24) * 914400)), Pt(1.5), col_color)
        # 要点
        item_y = top_y + 0.82
        for b in items[:6]:
            h = _est_height(b, _fs(theme.size_body_sm), half_w - 0.52)
            _textbox(slide,
                     Emu(int((x + 0.16) * 914400)), Emu(int(item_y * 914400)),
                     Inches(0.28), Emu(int(h * 914400)),
                     "•", size=_fs(14), color=col_color, bold=True)
            _textbox(slide,
                     Emu(int((x + 0.46) * 914400)), Emu(int(item_y * 914400)),
                     Emu(int((half_w - 0.62) * 914400)), Emu(int(h * 914400)),
                     b, size=_fs(theme.size_body_sm),
                     color=_auto_fg(_card_bg(col_color, theme)),
                     font=theme.font_minor, line_spacing=theme.line_height)
            item_y += h + 0.18

    _draw_col(label_l, items_l, left_x, left_color)
    _draw_col(label_r, items_r, right_x, right_color)


# ---------------------------------------------------------------------------
# 图片放置（等比缩放，支持 EXIF 旋转）
# ---------------------------------------------------------------------------
def _place_image(slide, path: str, x, y, max_w, max_h,
                 theme: config.Theme, valign: str = "center") -> None:
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
    xx = int(x) + int(round((mw - w) / 2 * 914400))
    yy = int(y) if valign == "top" else int(y) + int(round((mh - h) / 2 * 914400))
    pic = slide.shapes.add_picture(src, xx, yy,
                                   int(round(w * 914400)), int(round(h * 914400)))
    pic.line.color.rgb = _rgb(theme.secondary)
    pic.line.width = Pt(1)


# ---------------------------------------------------------------------------
# 主渲染入口
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 字体回退检测（解决 Linux/Mac/Docker 下方块字问题）
# ---------------------------------------------------------------------------
_FONT_FALLBACK_CACHE: dict[str, str] = {}


def _detect_font(preferred: str) -> str:
    """检测字体是否安装，找不到时返回可用备选。"""
    if preferred in _FONT_FALLBACK_CACHE:
        return _FONT_FALLBACK_CACHE[preferred]

    import platform
    result = preferred

    if platform.system() == "Windows":
        # Windows 上 Microsoft YaHei 通常安装，直接信任
        _FONT_FALLBACK_CACHE[preferred] = preferred
        return preferred

    # Linux / Mac 上扫描字体目录
    try:
        import subprocess
        candidates: list[str] = []
        # 中文字体候选顺序
        fallbacks = [
            "Noto Sans CJK SC", "Noto Sans SC",
            "WenQuanYi Micro Hei", "Source Han Sans CN",
            "PingFang SC", "Hiragino Sans GB",
            "Microsoft YaHei", "Arial Unicode MS", "Arial",
        ]
        out = subprocess.run(["fc-list", ":lang=zh"],
                             capture_output=True, text=True, timeout=5)
        installed = out.stdout.lower()
        for fb in fallbacks:
            if fb.lower() in installed:
                result = fb
                break
    except Exception:
        pass

    _FONT_FALLBACK_CACHE[preferred] = result
    return result


def _apply_font_fallback(theme: "config.Theme") -> None:
    """如果在非 Windows 环境，将主题字体映射到已安装的备选。"""
    import platform
    if platform.system() == "Windows":
        return  # Windows 不需要回退
    theme.font_major = _detect_font(theme.font_major)
    theme.font_minor = _detect_font(theme.font_minor)


def render(plan: DeckPlan, output_path: str, *,
           font_scale: float = 1.0, hpad: float = 0.0,
           gap_mult: float = 1.0, rotate_map: dict | None = None,
           template_path: str | None = None,
           animation: bool = False,
           slide_images: dict | None = None) -> str:
    """渲染 DeckPlan 为 .pptx。

    animation=True 时按页型自动注入进入动画（淡入/飞入/缩放等）。
    slide_images: {1based页号: 图片路径} 为分隔页指定照片（editorial 风格时生效）。
    """
    global _FONT_SCALE, _HPAD, _GAP_MULT, _ROTATE_MAP
    _FONT_SCALE = font_scale
    _HPAD       = hpad
    _GAP_MULT   = gap_mult
    _ROTATE_MAP = rotate_map or {}

    theme = config.THEMES.get(plan.theme_name, config.THEMES[config.DEFAULT_THEME])
    _apply_font_fallback(theme)  # 非 Windows 自动回退字体

    # editorial 风格路由
    if getattr(theme, "design_style", "") == "editorial":
        from .renderer_editorial import render_editorial
        return render_editorial(plan, output_path,
                                font_scale=font_scale,
                                rotate_map=rotate_map,
                                template_path=template_path,
                                animation=animation,
                                slide_images=slide_images)

    if template_path:
        prs = Presentation(template_path)
        for idx in range(len(prs.slides) - 1, -1, -1):
            remove_slide(prs, idx)
        template_layouts = [l for m in prs.slide_masters for l in m.slide_layouts]
        canvas_scale = prs.slide_width / SLIDE_W
    else:
        prs = Presentation()
        prs.slide_width  = SLIDE_W
        prs.slide_height = SLIDE_H
        template_layouts = None
        canvas_scale = 1.0

    blank = prs.slide_layouts[6] if template_layouts is None else None

    total = len(plan.slides)
    for idx, sd in enumerate(plan.slides):
        layout    = blank if template_layouts is None else _pick_layout(template_layouts, sd.kind)
        slide     = prs.slides.add_slide(layout)
        inherited = set(slide.shapes._spTree) if template_path else None
        page_no   = idx + 1
        kind      = sd.kind

        # 背景图（封面/章节用主图，内容页用淡版）
        bg_img = theme.bg_image if kind in ("title", "section") else theme.bg_image_soft
        if bg_img and os.path.isfile(bg_img):
            try:
                slide.shapes.add_picture(bg_img, 0, 0, SLIDE_W, SLIDE_H)
            except Exception:
                pass

        # 派发到各页型渲染函数
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
        # ── 新增页型 ──────────────────────────────────
        elif kind == "kpi":
            _render_kpi(slide, sd, theme, page_no, total)
        elif kind == "quote":
            _render_quote(slide, sd, theme, page_no, total)
        elif kind == "process":
            _render_process(slide, sd, theme, page_no, total)
        elif kind == "comparison":
            _render_comparison(slide, sd, theme, page_no, total)
        else:
            _render_bullets(slide, sd, theme, page_no, total)

        # 备注
        try:
            if sd.notes:
                slide.notes_slide.notes_text_frame.text = sd.notes
        except Exception:
            pass

        # 来源标注（页脚小字，结构页除外）
        if sd.source and kind not in ("title", "section", "thanks", "quote"):
            try:
                _textbox(slide, MARGIN, SLIDE_H - Inches(0.48),
                         CONTENT_W - Inches(2.0), Inches(0.33),
                         f"数据来源：{sd.source}",
                         size=theme.size_footer, color=theme.neutral,
                         font=theme.font_minor)
            except Exception:
                pass

        if template_path and canvas_scale != 1.0:
            _scale_new_shapes(slide, inherited, canvas_scale)

    # 动画注入（渲染完成后再加，避免干扰形状 shape_id）
    if animation:
        try:
            from .animation import apply_slide_animation
            for prs_slide, sd in zip(prs.slides, plan.slides):
                apply_slide_animation(prs_slide, kind=sd.kind)
        except Exception as exc:
            print(f"[动画] 注入失败（已跳过）：{exc}")

    fix_app_properties_slide_count(prs)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)
    return output_path
