"""Editorial 风格渲染器 —— 忠实复刻参考文件的设计语言。

设计语言（源自 golden_week_travel2.pptx 逆向分析）：
  色彩：5色体系 primary=#0B4F4A / accent=#F05D5E / secondary=#18A999
         bg=#F7F3EA / fg=#183438 / neutral=#5D7776
  排版：3层页头（标签11pt+横线 / 大标题26pt / 副标题13pt）+ 右下角页码8pt
  分隔页：全屏照片铺底（无照片时渐变）+ 左侧超大数字 + 右侧目的地名
  内容页：米白底，无圆角卡片，竖分割线，编号用色点

支持页型：
  title / section / bullets / cards / timeline / comparison
  process / kpi / summary / thanks / chart / table / quote / agenda / figure
"""
from __future__ import annotations

import os
import math
import datetime

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

# ---------------------------------------------------------------------------
# 画布常量
# ---------------------------------------------------------------------------
W  = Inches(13.333)
H  = Inches(7.5)
ML = Inches(0.62)    # margin left  (与参考文件匹配)
MR = Inches(0.62)    # margin right (保持对称，同时满足 quality 检测边界)
CW = Emu(W - ML - MR)

# 页头高度分配（严格匹配参考）
HDR_LABEL_Y  = Inches(0.38)   # 红色标签 y
HDR_TITLE_Y  = Inches(0.78)   # 大标题 y
HDR_SUB_Y    = Inches(1.36)   # 副标题 y
HDR_LINE_Y   = Inches(1.80)   # 红色横线 y
CONTENT_Y    = Inches(2.20)   # 内容区起始 y
FOOTER_Y     = Inches(7.08)   # 页码 y

# 字体缩放（质检重排用）
_FONT_SCALE = 1.0


def _fs(n: int) -> int:
    return max(8, int(round(n * _FONT_SCALE)))


# ---------------------------------------------------------------------------
# 颜色工具
# ---------------------------------------------------------------------------
def _rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _tint(h: str, f: float) -> str:
    h = h.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return "{:02X}{:02X}{:02X}".format(
        int(r + (255 - r) * f), int(g + (255 - g) * f), int(b + (255 - b) * f))


def _luminance(h: str) -> float:
    h = h.lstrip("#")
    r, g, b = int(h[0:2], 16) / 255, int(h[2:4], 16) / 255, int(h[4:6], 16) / 255
    return 0.299 * r + 0.587 * g + 0.114 * b


# ---------------------------------------------------------------------------
# 基础绘图
# ---------------------------------------------------------------------------
def _run(tf, text: str, size: int, color: str, bold: bool = False,
         font: str | None = None, align=PP_ALIGN.LEFT,
         line_spacing: float = 1.25) -> None:
    """向 text_frame 首段追加 run。"""
    p = tf.paragraphs[0] if tf.paragraphs else tf.add_paragraph()
    p.alignment = align
    try:
        p.line_spacing = line_spacing
    except Exception:
        pass
    r = p.add_run()
    r.text = text
    r.font.size = Pt(_fs(size))
    r.font.bold = bold
    r.font.color.rgb = _rgb(color)
    fname = font or "Microsoft YaHei"
    r.font.name = fname
    rPr = r._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", fname)


def _tb(slide, x, y, w, h, text: str, size: int, color: str,
        bold: bool = False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
        font: str | None = None, line_spacing: float = 1.25,
        wrap: bool = True):
    tb = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    _run(tf, text, size, color, bold, font, align, line_spacing)
    return tb


def _rect(slide, x, y, w, h, fill: str, line: str = "", line_pt: float = 0):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  int(x), int(y), int(w), int(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(fill)
    if line and line_pt > 0:
        shp.line.color.rgb = _rgb(line)
        shp.line.width = Pt(line_pt)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _set_bg(slide, color: str) -> None:
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(color)


def _place_image(slide, path: str, x, y, w, h) -> bool:
    """放置图片（等比居中裁切），返回是否成功。"""
    if not path or not os.path.isfile(path):
        return False
    try:
        from PIL import Image, ImageOps
        with Image.open(path) as im:
            im = ImageOps.exif_transpose(im)
            iw, ih = im.size
        ratio = iw / ih
        tw, th = Emu(w).inches, Emu(h).inches
        if tw / th > ratio:
            fw, fh = tw, tw / ratio
        else:
            fw, fh = th * ratio, th
        ox = int(x) + int((tw - fw) / 2 * 914400)
        oy = int(y) + int((th - fh) / 2 * 914400)
        slide.shapes.add_picture(
            path, ox, oy,
            int(fw * 914400), int(fh * 914400))
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 通用页头（内容页统一结构）
# ---------------------------------------------------------------------------
def _content_header(slide, theme: config.Theme,
                    label: str, title: str, subtitle: str,
                    page_no: int, total: int) -> None:
    """
    内容页页头：
      ① 红色标签（11pt）
      ② 大标题（26pt 深绿）
      ③ 副标题（13pt 灰绿）
      ④ 红色横线（1px）
      ⑤ 右下角页码（8pt 极浅）
    """
    _set_bg(slide, theme.bg)

    # ① 标签
    _tb(slide, ML, HDR_LABEL_Y, CW, Inches(0.30),
        label.upper(), 11, theme.accent, bold=True,
        font=theme.font_minor)

    # ② 大标题
    _tb(slide, ML, HDR_TITLE_Y, CW, Inches(0.62),
        title, _fs(theme.size_slide_title), theme.primary, bold=True,
        font=theme.font_major, line_spacing=1.1)

    # ③ 副标题
    if subtitle:
        _tb(slide, ML, HDR_SUB_Y, CW, Inches(0.40),
            subtitle, _fs(theme.size_caption), theme.neutral,
            font=theme.font_minor)

    # ④ 红色横线（1pt）
    _rect(slide, ML, HDR_LINE_Y, CW, Pt(1.5), theme.accent)

    # ⑤ 页码
    _tb(slide, W - Inches(1.50), FOOTER_Y, Inches(1.38), Inches(0.25),
        f"{page_no:02d}  /  {total:02d}", 8, "8AA5A1",
        align=PP_ALIGN.RIGHT, font=theme.font_minor)


# ---------------------------------------------------------------------------
# ① 封面页
# ---------------------------------------------------------------------------
def _render_title(slide, sd: SlideData, theme: config.Theme,
                  plan: DeckPlan) -> None:
    # 背景：照片 or 深绿渐变
    if theme.bg_image and os.path.isfile(theme.bg_image):
        slide.shapes.add_picture(theme.bg_image, 0, 0, W, H)
        # 半透明深色遮罩
        overlay = _rect(slide, 0, 0, W, H, "0B4F4A")
        overlay.fill.fore_color.rgb = _rgb("0B4F4A")
        # 透明度需要通过 XML 设置
        try:
            from lxml import etree
            sp = overlay._element
            spPr = sp.find(qn("p:spPr"))
            solidFill = spPr.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill")
            if solidFill is not None:
                clr = solidFill.find("{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
                if clr is not None:
                    alpha = etree.SubElement(
                        clr, "{http://schemas.openxmlformats.org/drawingml/2006/main}alpha")
                    alpha.set("val", "72000")  # 72% opacity = 28% transparent
        except Exception:
            pass
    else:
        _set_bg(slide, theme.bg_dark or theme.primary)

    # 顶部细红线
    _rect(slide, 0, 0, W, Inches(0.14), theme.accent)

    left = ML
    # 上方小标签
    _tb(slide, left, Inches(0.70), Inches(8), Inches(0.30),
        "2026 NATIONAL DAY · 7-DAY JOURNEY",
        11, _tint(theme.primary, 0.80), bold=True,
        font=theme.font_minor)

    # 主标题（中文大字）
    _tb(slide, left, Inches(1.40), Inches(8), Inches(1.10),
        plan.title or sd.title, _fs(theme.size_cover_title),
        "FFFFFF", bold=True, font=theme.font_major, line_spacing=1.1)

    # 副标题
    sub = plan.subtitle or ""
    if sub:
        _tb(slide, left, Inches(2.55), Inches(8), Inches(0.60),
            sub, _fs(theme.size_cover_sub),
            "F2D1C4", font=theme.font_minor, line_spacing=1.1)

    # 目的地流向（5站箭头文字）
    route_bullets = [b for b in sd.bullets if b.strip()]
    if route_bullets:
        route_text = "  →  ".join(route_bullets[:5])
    else:
        route_text = "江南水乡  →  竹林山居  →  碧水万岛  →  云海奇峰  →  世遗古村"
    _tb(slide, left, Inches(3.60), Inches(8.5), Inches(1.20),
        route_text, 18, "FFFFFF", font=theme.font_minor,
        line_spacing=1.3, wrap=True)

    # 底部预算/出发信息
    info_parts = []
    if plan.author:
        info_parts.append(plan.author)
    for b in sd.bullets:
        if "预算" in b or "出发" in b or "上海" in b:
            info_parts.append(b)
            break
    info_text = "  ｜  ".join(info_parts) if info_parts else \
        f"上海出发｜高铁 + 包车｜双人预算约 ¥10,200–10,800"
    _tb(slide, left, Inches(5.80), Inches(9), Inches(0.40),
        info_text, 12, _tint(theme.primary, 0.85),
        font=theme.font_minor)

    # 出品人
    author_line = plan.author or "旅行规划团队 · 2026 国庆特辑"
    _tb(slide, left, Inches(6.50), Inches(6), Inches(0.30),
        author_line, 10, _tint(theme.primary, 0.70),
        font=theme.font_minor)


# ---------------------------------------------------------------------------
# ② 章节分隔页（editorial 风：照片全屏 + 大数字 + 右侧标题）
# ---------------------------------------------------------------------------
def _render_section(slide, sd: SlideData, theme: config.Theme,
                    page_no: int, total: int,
                    section_idx: int = 0,
                    day_label: str = "",
                    subtitle_line: str = "",
                    tagline: str = "") -> None:
    # 背景：照片或深绿
    has_photo = False
    if sd.images:
        has_photo = _place_image(slide, sd.images[0], 0, 0, W, H)
    if not has_photo and theme.bg_image and os.path.isfile(theme.bg_image):
        slide.shapes.add_picture(theme.bg_image, 0, 0, W, H)
        has_photo = True
    if not has_photo:
        _set_bg(slide, theme.bg_dark or theme.primary)
    else:
        # 深色遮罩（底部更重，保证文字可读）
        _rect(slide, 0, 0, W, H, "0B4F4A")
        try:
            from lxml import etree
            overlay_el = slide.shapes[-1]._element
            spPr = overlay_el.find(qn("p:spPr"))
            sf = spPr.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill")
            if sf is not None:
                clr = sf.find("{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr")
                if clr is not None:
                    a = etree.SubElement(
                        clr, "{http://schemas.openxmlformats.org/drawingml/2006/main}alpha")
                    a.set("val", "55000")  # 55% opacity = photo visible
        except Exception:
            pass

    left = Inches(0.70)

    # 左上角目的地序号（大红字）
    num_str = f"{section_idx:02d}" if section_idx > 0 else f"{page_no:02d}"
    _tb(slide, left, Inches(0.70), Inches(1.50), Inches(0.80),
        num_str, 40, theme.accent, bold=True, font=theme.font_major)

    # Day 标签
    if day_label:
        _tb(slide, left, Inches(1.60), Inches(5), Inches(0.30),
            day_label.upper(), 12, _tint(theme.primary, 0.88),
            bold=True, font=theme.font_minor)

    # 目的地大标题
    _tb(slide, left, Inches(2.00), Inches(7), Inches(1.00),
        sd.title, 34, "FFFFFF", bold=True,
        font=theme.font_major, line_spacing=1.1)

    # 副标题（一句话描述）
    if subtitle_line:
        _tb(slide, left, Inches(2.90), Inches(7.5), Inches(0.45),
            subtitle_line, 17, _tint(theme.primary, 0.92),
            font=theme.font_minor, line_spacing=1.2)

    # 金句/行动语
    if tagline:
        _tb(slide, left, Inches(3.80), Inches(7.5), Inches(1.10),
            tagline, 19, "FFFFFF",
            font=theme.font_minor, line_spacing=1.3, wrap=True)

    # 右下角页码
    _tb(slide, W - Inches(2.50), FOOTER_Y + Inches(0.10),
        Inches(2.0), Inches(0.25),
        f"{page_no:02d}  /  {total:02d}",
        11, _tint(theme.primary, 0.65),
        align=PP_ALIGN.RIGHT, font=theme.font_minor)


# ---------------------------------------------------------------------------
# 通用内容页 — 「时间线」式布局（多个时段/步骤）
# ---------------------------------------------------------------------------
def _render_timeline_editorial(slide, sd: SlideData, theme: config.Theme,
                                page_no: int, total: int,
                                label: str = "ROUTE") -> None:
    """水平步骤时间线，复刻 P04 ROUTE 和 P10/P12 的步骤布局。"""
    _content_header(slide, theme, label, sd.title,
                    sd.notes or "", page_no, total)

    items = [b for b in sd.bullets if b.strip()]
    n = min(len(items), 7)
    if n == 0:
        return

    # 每个步骤的 x 位置
    avail_w = CW.inches
    slot_w = avail_w / n
    top_num_y = CONTENT_Y.inches + 0.15
    top_title_y = top_num_y + 1.05
    top_desc_y = top_title_y + 0.52

    for i, item in enumerate(items[:n]):
        cx = ML.inches + i * slot_w

        # 编号圆点（深绿色矩形，参考用实心方块）
        _rect(slide,
              Emu(int((cx + 0.02) * 914400)),
              Emu(int(top_num_y * 914400)),
              Inches(0.60), Inches(0.55),
              theme.primary)
        _tb(slide,
            Emu(int((cx + 0.02) * 914400)),
            Emu(int(top_num_y * 914400)),
            Inches(0.60), Inches(0.55),
            f"{i+1:02d}", 10, "FFFFFF", bold=True,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            font=theme.font_major)

        # Day 标签（小青绿标签）
        day_parts = item.split("·")
        day_tag = day_parts[0].strip() if len(day_parts) > 1 else f"D{i+1}"
        content = "·".join(day_parts[1:]).strip() if len(day_parts) > 1 else item

        _tb(slide,
            Emu(int(cx * 914400)),
            Emu(int((top_num_y - 0.42) * 914400)),
            Emu(int(slot_w * 914400)), Inches(0.30),
            day_tag, 11, theme.secondary, bold=True,
            font=theme.font_minor)

        # 目的地名（18pt）
        dest_parts = content.split("：")
        dest = dest_parts[0].strip()[:8]
        desc = dest_parts[1].strip() if len(dest_parts) > 1 else ""

        _tb(slide,
            Emu(int(cx * 914400)),
            Emu(int(top_title_y * 914400)),
            Emu(int(slot_w * 914400)), Inches(0.45),
            dest, 18, theme.primary, bold=True,
            font=theme.font_major)

        # 描述小字
        if desc:
            _tb(slide,
                Emu(int(cx * 914400)),
                Emu(int(top_desc_y * 914400)),
                Emu(int((slot_w - 0.1) * 914400)), Inches(0.60),
                desc, 12, theme.neutral,
                font=theme.font_minor, line_spacing=1.2, wrap=True)

    # 连接线（浅色）
    line_y = top_num_y + 0.275
    _rect(slide,
          Emu(int((ML.inches + 0.65) * 914400)),
          Emu(int((line_y - 0.012) * 914400)),
          Emu(int((avail_w - 0.70) * 914400)),
          Pt(1.5), _tint(theme.neutral, 0.50))

    # 底部备注
    if sd.takeaway:
        _rect(slide, ML, H - Inches(0.92), CW, Inches(0.52), theme.primary)
        _tb(slide, ML + Inches(0.16), H - Inches(0.88),
            CW - Inches(0.32), Inches(0.45),
            sd.takeaway, 14, "FFFFFF", bold=True,
            font=theme.font_minor)


# ---------------------------------------------------------------------------
# 要点列表页（editorial：左侧时段标签 + 右侧内容）
# ---------------------------------------------------------------------------
def _render_bullets_editorial(slide, sd: SlideData, theme: config.Theme,
                               page_no: int, total: int,
                               label: str = "") -> None:
    lbl = label or sd.kind.replace("_", " ").upper()
    _content_header(slide, theme, lbl, sd.title,
                    sd.notes or "", page_no, total)

    items = [b for b in sd.bullets if b.strip()]
    if not items:
        return

    # 竖分割线（参考 P06 左侧 x=1.2）
    vline_x = ML.inches + 0.58
    _rect(slide,
          Emu(int(vline_x * 914400)),
          Emu(int((CONTENT_Y.inches + 0.05) * 914400)),
          Pt(1.0),
          Emu(int((len(items) * 0.82 + 0.2) * 914400)),
          _tint(theme.secondary, 0.55))

    label_x = ML.inches + 0.08
    content_x = ML.inches + 0.78
    content_w = CW.inches - 0.82

    # 判断是否有时段标签（"Day X · 时段"格式）
    use_time_labels = any("·" in b for b in items[:3])

    y = CONTENT_Y.inches + 0.10
    for i, item in enumerate(items):
        # 时段标签（青绿色小字）
        if use_time_labels and "·" in item:
            parts = item.split("·", 1)
            time_tag = parts[0].strip()
            content = parts[1].strip()
        else:
            time_tag = f"{i+1:02d}"
            content = item

        # 左侧编号/时段（用色点+小字）
        _rect(slide,
              Emu(int((label_x + 0.08) * 914400)),
              Emu(int((y + 0.06) * 914400)),
              Inches(0.38), Inches(0.38),
              theme.accent if use_time_labels else theme.primary)
        _tb(slide,
            Emu(int(label_x * 914400)),
            Emu(int((y - 0.01) * 914400)),
            Inches(0.52), Inches(0.42),
            time_tag[:6], 10, "FFFFFF", bold=True,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            font=theme.font_minor)

        # 如果有子标题（"标题：说明"格式）
        if "：" in content or ":" in content:
            sep = "：" if "：" in content else ":"
            sub_title, sub_desc = content.split(sep, 1)
            _tb(slide,
                Emu(int(content_x * 914400)),
                Emu(int(y * 914400)),
                Emu(int(content_w * 914400)), Inches(0.35),
                sub_title.strip(), 18, theme.primary, bold=True,
                font=theme.font_major)
            _tb(slide,
                Emu(int(content_x * 914400)),
                Emu(int((y + 0.38) * 914400)),
                Emu(int(content_w * 914400)), Inches(0.45),
                sub_desc.strip(), 15, theme.neutral,
                font=theme.font_minor, line_spacing=1.2, wrap=True)
            y += 0.88
        else:
            _tb(slide,
                Emu(int(content_x * 914400)),
                Emu(int(y * 914400)),
                Emu(int(content_w * 914400)), Inches(0.42),
                content, 16, theme.fg, bold=False,
                font=theme.font_minor, line_spacing=1.25)
            y += 0.65

    # 底部结论条
    if sd.takeaway:
        _render_bottom_bar(slide, sd.takeaway, theme)


def _render_bottom_bar(slide, text: str, theme: config.Theme) -> None:
    """深绿色底部高亮条（白色加粗文字）。"""
    bar_y = H - Inches(0.92)
    _rect(slide, ML, bar_y, CW, Inches(0.52), theme.primary)
    _tb(slide, ML + Inches(0.16), bar_y + Inches(0.06),
        CW - Inches(0.32), Inches(0.42),
        text, 14, "FFFFFF", bold=True,
        font=theme.font_minor, anchor=MSO_ANCHOR.MIDDLE)


# ---------------------------------------------------------------------------
# 卡片页（editorial：编号横条 + 标题 + 说明，竖向列表）
# ---------------------------------------------------------------------------
def _render_cards_editorial(slide, sd: SlideData, theme: config.Theme,
                             page_no: int, total: int) -> None:
    _content_header(slide, theme, "HIGHLIGHTS", sd.title,
                    sd.notes or "", page_no, total)

    items = [b for b in sd.bullets if b.strip()]
    n = min(len(items), 6)
    if n == 0:
        return

    cols = 2 if n > 3 else 1
    per_col = math.ceil(n / cols)
    col_w = (CW.inches - 0.40) / cols
    start_y = CONTENT_Y.inches + 0.05

    for i, item in enumerate(items[:n]):
        col = i // per_col
        row = i % per_col
        x = ML.inches + col * (col_w + 0.40)
        y = start_y + row * 0.88

        # 编号（红色）
        _tb(slide,
            Emu(int(x * 914400)),
            Emu(int(y * 914400)),
            Inches(0.40), Inches(0.28),
            f"{i+1:02d}", 10,
            theme.accent if col == 0 else theme.secondary,
            bold=True, font=theme.font_minor)

        # 内容（有"：" 则分标题+说明）
        if "：" in item:
            title_part, desc_part = item.split("：", 1)
            _tb(slide,
                Emu(int((x + 0.45) * 914400)),
                Emu(int(y * 914400)),
                Emu(int((col_w - 0.50) * 914400)), Inches(0.32),
                title_part.strip(), 16, theme.primary, bold=False,
                font=theme.font_major)
            _tb(slide,
                Emu(int((x + 0.45) * 914400)),
                Emu(int((y + 0.35) * 914400)),
                Emu(int((col_w - 0.50) * 914400)), Inches(0.48),
                desc_part.strip(), 14, theme.neutral,
                font=theme.font_minor, line_spacing=1.2, wrap=True)
        else:
            _tb(slide,
                Emu(int((x + 0.45) * 914400)),
                Emu(int(y * 914400)),
                Emu(int((col_w - 0.50) * 914400)), Inches(0.75),
                item, 15, theme.fg, font=theme.font_minor,
                line_spacing=1.25, wrap=True)

    if sd.takeaway:
        _render_bottom_bar(slide, sd.takeaway, theme)


# ---------------------------------------------------------------------------
# 对比页（editorial：双栏竖线分割，无圆角框）
# ---------------------------------------------------------------------------
def _render_comparison_editorial(slide, sd: SlideData, theme: config.Theme,
                                  page_no: int, total: int) -> None:
    _content_header(slide, theme, "COMPARISON", sd.title,
                    sd.notes or "", page_no, total)

    half_w = (CW.inches - 0.30) / 2
    left_x  = ML.inches
    right_x = ML.inches + half_w + 0.30
    label_l = sd.label_left  or "选项 A"
    label_r = sd.label_right or "选项 B"
    items_l = [b for b in sd.bullets if b.strip()]
    items_r = [b for b in sd.bullets_right if b.strip()]

    # 竖分割线
    vx = ML.inches + half_w + 0.15
    _rect(slide,
          Emu(int(vx * 914400)),
          Emu(int((CONTENT_Y.inches + 0.05) * 914400)),
          Pt(1.0),
          Emu(int(4.0 * 914400)),
          _tint(theme.neutral, 0.60))

    # 左右标签（大号）
    _tb(slide,
        Emu(int(left_x * 914400)),
        Emu(int((CONTENT_Y.inches + 0.05) * 914400)),
        Emu(int(half_w * 914400)), Inches(0.48),
        label_l, 24, theme.primary, bold=True, font=theme.font_major)
    _tb(slide,
        Emu(int(right_x * 914400)),
        Emu(int((CONTENT_Y.inches + 0.05) * 914400)),
        Emu(int(half_w * 914400)), Inches(0.48),
        label_r, 24, theme.accent, bold=True, font=theme.font_major)

    # 价格/小标签
    for bl_note, col_x, col_color in [
        (sd.notes, left_x, theme.secondary),
        ("", right_x, theme.primary),
    ]:
        if bl_note:
            _tb(slide,
                Emu(int(col_x * 914400)),
                Emu(int((CONTENT_Y.inches + 0.58) * 914400)),
                Emu(int(half_w * 914400)), Inches(0.30),
                bl_note, 13, col_color, bold=True, font=theme.font_minor)

    # 各栏要点
    def _draw_col(items, x, num_color):
        y = CONTENT_Y.inches + 0.98
        for idx, item in enumerate(items[:5]):
            _tb(slide,
                Emu(int(x * 914400)),
                Emu(int(y * 914400)),
                Inches(0.40), Inches(0.28),
                f"{idx+1:02d}", 10, num_color,
                bold=True, font=theme.font_minor)
            _tb(slide,
                Emu(int((x + 0.45) * 914400)),
                Emu(int(y * 914400)),
                Emu(int((half_w - 0.50) * 914400)), Inches(0.45),
                item, 16, theme.fg,
                font=theme.font_minor, line_spacing=1.2, wrap=True)
            y += 0.65

    _draw_col(items_l, left_x, theme.accent)
    _draw_col(items_r, right_x, theme.secondary)

    # 底部建议条
    if sd.takeaway:
        _render_bottom_bar(slide, sd.takeaway, theme)


# ---------------------------------------------------------------------------
# 流程页（横向步骤，大编号）
# ---------------------------------------------------------------------------
def _render_process_editorial(slide, sd: SlideData, theme: config.Theme,
                               page_no: int, total: int) -> None:
    _content_header(slide, theme, "PROCESS", sd.title,
                    sd.notes or "", page_no, total)

    items = [b for b in sd.bullets if b.strip()]
    n = min(len(items), 5)
    if n == 0:
        return

    slot_w = CW.inches / n
    num_y   = CONTENT_Y.inches + 0.12
    title_y = num_y + 1.10
    desc_y  = title_y + 0.52

    # 连接线
    line_y = num_y + 0.275
    _rect(slide,
          Emu(int((ML.inches + slot_w * 0.5) * 914400)),
          Emu(int((line_y - 0.008) * 914400)),
          Emu(int((CW.inches - slot_w) * 914400)),
          Pt(1.5), _tint(theme.neutral, 0.55))

    for i, item in enumerate(items[:n]):
        x = ML.inches + i * slot_w

        # 编号（大红字）
        _tb(slide,
            Emu(int(x * 914400)),
            Emu(int(num_y * 914400)),
            Emu(int(slot_w * 914400)), Inches(0.65),
            f"{i+1:02d}", 28, theme.accent, bold=True,
            font=theme.font_major)

        # 标题/说明分割
        if "：" in item or "·" in item:
            sep = "：" if "：" in item else "·"
            parts = item.split(sep, 1)
            step_title = parts[0].strip()
            step_desc  = parts[1].strip()
        else:
            words = item.split()
            step_title = " ".join(words[:3])
            step_desc  = " ".join(words[3:])

        _tb(slide,
            Emu(int(x * 914400)),
            Emu(int(title_y * 914400)),
            Emu(int((slot_w - 0.10) * 914400)), Inches(0.45),
            step_title, 18, theme.primary, bold=True,
            font=theme.font_major)

        if step_desc:
            _tb(slide,
                Emu(int(x * 914400)),
                Emu(int(desc_y * 914400)),
                Emu(int((slot_w - 0.10) * 914400)), Inches(0.95),
                step_desc, 14, theme.neutral,
                font=theme.font_minor, line_spacing=1.25, wrap=True)

        # 下划线（每步骤）
        _rect(slide,
              Emu(int(x * 914400)),
              Emu(int((title_y - 0.06) * 914400)),
              Emu(int((slot_w * 0.80) * 914400)),
              Pt(1.5), theme.secondary)

    if sd.takeaway:
        _render_bottom_bar(slide, sd.takeaway, theme)


# ---------------------------------------------------------------------------
# KPI 页（横排大数字 + 单位 + 说明）
# ---------------------------------------------------------------------------
def _render_kpi_editorial(slide, sd: SlideData, theme: config.Theme,
                           page_no: int, total: int) -> None:
    _content_header(slide, theme, "AT A GLANCE", sd.title,
                    sd.notes or "", page_no, total)

    metrics = sd.metrics
    n = len(metrics)
    if n == 0:
        return

    cols = min(n, 5)
    slot_w = CW.inches / cols
    num_y   = CONTENT_Y.inches + 0.18
    unit_y  = num_y + 0.82
    desc_y  = unit_y + 0.40

    trend_color = {"up": theme.accent, "down": theme.secondary, "flat": theme.neutral}

    for i, m in enumerate(metrics[:cols]):
        x = ML.inches + i * slot_w

        # 大数字（交替红/深绿）
        num_color = theme.accent if i % 2 == 0 else theme.primary
        _tb(slide,
            Emu(int(x * 914400)),
            Emu(int(num_y * 914400)),
            Emu(int(slot_w * 914400)), Inches(0.95),
            m.value or "—", _fs(34), num_color, bold=True,
            font=theme.font_major)

        # 单位/指标名（青绿 bold）
        _tb(slide,
            Emu(int(x * 914400)),
            Emu(int(unit_y * 914400)),
            Emu(int(slot_w * 914400)), Inches(0.35),
            m.label or "", _fs(13), theme.secondary, bold=True,
            font=theme.font_minor)

        # 说明（灰绿小字，多行）
        if m.note:
            _tb(slide,
                Emu(int(x * 914400)),
                Emu(int(desc_y * 914400)),
                Emu(int((slot_w - 0.10) * 914400)), Inches(1.10),
                m.note, _fs(14), theme.neutral,
                font=theme.font_minor, line_spacing=1.25, wrap=True)

    # 底部预算口径说明
    if sd.takeaway:
        _tb(slide,
            Emu(int((ML.inches + 0.20) * 914400)),
            Emu(int((H.inches - 1.28) * 914400)),
            CW - Inches(0.40), Inches(0.40),
            sd.takeaway, 14, theme.primary, bold=True,
            font=theme.font_minor)


# ---------------------------------------------------------------------------
# 总结页（3列大数字+标题+说明）
# ---------------------------------------------------------------------------
def _render_summary_editorial(slide, sd: SlideData, theme: config.Theme,
                               page_no: int, total: int) -> None:
    # 深色背景
    _set_bg(slide, theme.bg_dark or theme.primary)
    if theme.bg_image and os.path.isfile(theme.bg_image):
        slide.shapes.add_picture(theme.bg_image, 0, 0, W, H)

    # 页头（白色版本）
    _tb(slide, ML, HDR_LABEL_Y, CW, Inches(0.30),
        "WHY THIS ROUTE", 11, theme.accent, bold=True,
        font=theme.font_minor)
    _tb(slide, ML, HDR_TITLE_Y, CW, Inches(0.62),
        sd.title, _fs(theme.size_slide_title), "FFFFFF", bold=True,
        font=theme.font_major)
    if sd.notes:
        _tb(slide, ML, HDR_SUB_Y, CW, Inches(0.40),
            sd.notes, _fs(theme.size_caption), _tint(theme.secondary, 0.70),
            font=theme.font_minor)
    _rect(slide, ML, HDR_LINE_Y, CW, Pt(1.5), theme.accent)

    items = [b for b in sd.bullets if b.strip()]
    n = min(len(items), 3)
    if n == 0:
        return

    cols = n
    slot_w = CW.inches / cols
    num_y   = CONTENT_Y.inches + 0.18
    title_y = num_y + 0.50
    desc_y  = title_y + 0.62

    # 竖分割线
    for j in range(1, cols):
        vx = ML.inches + j * slot_w
        _rect(slide,
              Emu(int((vx - 0.01) * 914400)),
              Emu(int(num_y * 914400)),
              Pt(1.0),
              Emu(int(3.60 * 914400)),
              "3C716C")

    for i, item in enumerate(items[:n]):
        x = ML.inches + i * slot_w + 0.04

        # 编号（红色）
        _tb(slide,
            Emu(int(x * 914400)),
            Emu(int(num_y * 914400)),
            Emu(int(0.80 * 914400)), Inches(0.38),
            f"{i+1:02d}", 16, theme.accent, bold=True,
            font=theme.font_minor)

        # 主题词（25pt 白色）
        if "：" in item:
            title_part, desc_part = item.split("：", 1)
        elif "—" in item:
            title_part, desc_part = item.split("—", 1)
        else:
            words = item.split()
            mid = max(2, len(words) // 3)
            title_part = " ".join(words[:mid])
            desc_part  = " ".join(words[mid:])

        _tb(slide,
            Emu(int(x * 914400)),
            Emu(int(title_y * 914400)),
            Emu(int((slot_w - 0.15) * 914400)), Inches(0.60),
            title_part.strip(), 25, "FFFFFF", bold=True,
            font=theme.font_major)

        _tb(slide,
            Emu(int(x * 914400)),
            Emu(int(desc_y * 914400)),
            Emu(int((slot_w - 0.15) * 914400)), Inches(1.20),
            desc_part.strip(), 16, _tint(theme.secondary, 0.60),
            font=theme.font_minor, line_spacing=1.3, wrap=True)

    # 底部金句
    if sd.takeaway:
        _tb(slide, ML, H - Inches(1.28), CW, Inches(0.48),
            sd.takeaway, 19, "FFFFFF", bold=True,
            font=theme.font_minor)

    # 页码
    _tb(slide, W - Inches(1.50), FOOTER_Y, Inches(1.38), Inches(0.25),
        f"{page_no:02d}  /  {total:02d}", 8, _tint(theme.secondary, 0.55),
        align=PP_ALIGN.RIGHT, font=theme.font_minor)


# ---------------------------------------------------------------------------
# 致谢/行动页
# ---------------------------------------------------------------------------
def _render_thanks_editorial(slide, sd: SlideData, theme: config.Theme,
                              plan: DeckPlan) -> None:
    # 背景：照片或深绿
    if theme.bg_image and os.path.isfile(theme.bg_image):
        slide.shapes.add_picture(theme.bg_image, 0, 0, W, H)
    else:
        _set_bg(slide, theme.bg_dark or theme.primary)

    # 顶部红线
    _rect(slide, 0, 0, W, Inches(0.14), theme.accent)

    left = ML

    # 标签
    _tb(slide, left, Inches(0.72), Inches(4), Inches(0.28),
        "NEXT STEP", 11, theme.accent, bold=True, font=theme.font_minor)

    # 主标题
    _tb(slide, left, Inches(1.32), Inches(7), Inches(0.80),
        sd.title or "下一步，确认四件事", 34, "FFFFFF", bold=True,
        font=theme.font_major, line_spacing=1.1)

    # 4件事列表
    action_items = [b for b in sd.bullets if b.strip()] or [
        "出发日期", "出行人数", "住宿档次", "预订授权"]
    for i, item in enumerate(action_items[:4]):
        iy = Inches(2.44 + i * 0.62)
        _tb(slide, left, iy, Inches(0.45), Inches(0.28),
            f"{i+1:02d}", 10, theme.accent, bold=True, font=theme.font_minor)
        _tb(slide, left + Inches(0.57), iy, Inches(5), Inches(0.36),
            item, 18, "FFFFFF", font=theme.font_minor)

    # Q & A
    _tb(slide, left, Inches(5.62), Inches(3), Inches(0.50),
        "Q & A", 23, "FFFFFF", bold=True, font=theme.font_major)

    # 底部金句
    _tb(slide, left, Inches(6.35), Inches(5), Inches(0.36),
        sd.takeaway or "祝一路山海皆风景", 16,
        _tint(theme.secondary, 0.70), font=theme.font_minor)

    # 出品信息
    author = plan.author or "旅行规划团队 · 2026 国庆特辑"
    _tb(slide, W - Inches(4.5), H - Inches(0.42), Inches(4.2), Inches(0.28),
        author, 10, _tint(theme.secondary, 0.70),
        align=PP_ALIGN.RIGHT, font=theme.font_minor)


# ---------------------------------------------------------------------------
# 图表页（复用 default renderer）
# ---------------------------------------------------------------------------
def _render_chart_editorial(slide, sd: SlideData, theme: config.Theme,
                              page_no: int, total: int) -> None:
    """图表页用 default renderer 的逻辑，但应用 editorial 页头。"""
    from .renderer import _render_chart as _default_chart, _content_header as _dc_hdr
    # 先画 editorial 页头
    _content_header(slide, theme, "BUDGET", sd.title,
                    sd.notes or "", page_no, total)
    # 复用图表绘制（跳过页头）
    chart = sd.chart
    if not chart:
        return
    from pptx.chart.data import CategoryChartData
    cdata = CategoryChartData()
    cdata.categories = list(chart.categories)
    for s in chart.series:
        cdata.add_series(s["name"], tuple(s["values"]))
    from pptx.enum.chart import XL_CHART_TYPE
    _CHART_TYPES = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED,
                    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                    "line": XL_CHART_TYPE.LINE_MARKERS,
                    "pie": XL_CHART_TYPE.PIE}
    ctype = _CHART_TYPES.get(chart.chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)
    chart_h = Inches(4.40) if chart.chart_type == "pie" else Inches(4.60)
    gf = slide.shapes.add_chart(ctype, ML, Inches(2.08), CW, chart_h, cdata)
    ch = gf.chart
    from .renderer import _fix_negative_axis_ids, _derive_accent_palette, _color_pie_slices
    _fix_negative_axis_ids(ch)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    # 颜色
    n_cats = len(chart.categories) if chart.categories else len(ch.series)
    palette = _derive_accent_palette(theme, max(n_cats, 4))
    if chart.chart_type == "pie":
        _color_pie_slices(ch, palette)
    else:
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
    # 页码
    _tb(slide, W - Inches(1.50), FOOTER_Y, Inches(1.38), Inches(0.25),
        f"{page_no:02d}  /  {total:02d}", 8, "8AA5A1",
        align=PP_ALIGN.RIGHT, font=theme.font_minor)


# ---------------------------------------------------------------------------
# 表格页
# ---------------------------------------------------------------------------
def _render_table_editorial(slide, sd: SlideData, theme: config.Theme,
                              page_no: int, total: int) -> None:
    _content_header(slide, theme, "BOOKING LIST", sd.title,
                    sd.notes or "", page_no, total)
    tbl_data = sd.table
    if not tbl_data:
        return
    nrows = min(len(tbl_data.rows) + 1, 10)
    ncols = max(len(tbl_data.headers),
                max((len(r) for r in tbl_data.rows), default=1))
    top = Inches(2.05)
    height = Inches(min(0.42 * nrows, 5.0))
    tbl_shape = slide.shapes.add_table(nrows, ncols, ML, top, CW, height)
    tbl = tbl_shape.table
    col_w = int(CW / ncols)
    for j in range(ncols):
        tbl.columns[j].width = col_w
    # 表头
    for j in range(ncols):
        hdr = tbl_data.headers[j] if j < len(tbl_data.headers) else ""
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(theme.primary)
        cell.margin_left = cell.margin_right = Inches(0.08)
        cell.margin_top = cell.margin_bottom = Inches(0.03)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = str(hdr)
        r.font.size = Pt(_fs(12))
        r.font.bold = True
        r.font.color.rgb = _rgb("FFFFFF")
        r.font.name = theme.font_minor
    # 数据行
    even_bg = _tint(theme.primary, 0.94)
    for i, row in enumerate(tbl_data.rows[:nrows-1]):
        for j in range(ncols):
            val = row[j] if j < len(row) else ""
            cell = tbl.cell(i+1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(even_bg if i % 2 == 0 else "FFFFFF")
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = str(val)
            r.font.size = Pt(_fs(12))
            r.font.color.rgb = _rgb(theme.fg)
            r.font.name = theme.font_minor
    _tb(slide, W - Inches(1.50), FOOTER_Y, Inches(1.38), Inches(0.25),
        f"{page_no:02d}  /  {total:02d}", 8, "8AA5A1",
        align=PP_ALIGN.RIGHT, font=theme.font_minor)


# ---------------------------------------------------------------------------
# 主渲染入口
# ---------------------------------------------------------------------------
# section 页面的辅助信息（subtitle + tagline），按目的地索引
_SECTION_META: list[tuple[str, str, str]] = [
    # (day_label, subtitle, tagline)
    ("DAY 1–2", "古镇夜色与清晨，是这趟旅程最温柔的开场。", "住进西栅，把清晨留给自己。"),
    ("DAY 3",   "从古镇切换到山林，把这一晚留给住宿本身。", "预算不追奢华，体验要住进竹林。"),
    ("DAY 4",   "从山林到湖面，把速度放慢，换一段开阔的视野。", "重点不在赶岛，而在把游船、观景与晚餐串成顺畅的一天。"),
    ("DAY 5–6", "全程的高潮，也是一段需要尊重天气与体力的旅程。", "Day 5 山脚休整，Day 6 一日完成经典主线。"),
    ("DAY 7",   "用一段水墨般的清晨，为七天旅程收束。", "先看南湖晨雾，再把时间交给返程。"),
]
_section_counter = [0]


def render_editorial(plan: DeckPlan, output_path: str,
                     font_scale: float = 1.0,
                     rotate_map: dict | None = None,
                     template_path: str | None = None,
                     animation: bool = False,
                     slide_images: dict | None = None) -> str:
    """渲染 editorial 风格 PPTX。

    slide_images: {page_1_based_index: image_path} 为分隔页指定照片。
    """
    global _FONT_SCALE
    _FONT_SCALE = font_scale
    _section_counter[0] = 0

    theme = config.THEMES.get(plan.theme_name, config.THEMES["travel"])

    if template_path:
        prs = Presentation(template_path)
        for idx in range(len(prs.slides) - 1, -1, -1):
            remove_slide(prs, idx)
    else:
        prs = Presentation()
        prs.slide_width  = W
        prs.slide_height = H

    blank_layout = prs.slide_layouts[6]
    total = len(plan.slides)
    slide_images = slide_images or {}

    for idx, sd in enumerate(plan.slides):
        slide   = prs.slides.add_slide(blank_layout)
        page_no = idx + 1
        kind    = sd.kind

        # 为分隔页指定照片
        if kind == "section" and page_no in slide_images:
            sd = type(sd)(**{
                **sd.__dict__,
                "images": [slide_images[page_no]] + list(sd.images),
            })

        if kind == "title":
            _render_title(slide, sd, theme, plan)

        elif kind == "section":
            _section_counter[0] += 1
            meta = _SECTION_META[min(_section_counter[0]-1,
                                     len(_SECTION_META)-1)]
            # 从 sd.notes 覆盖 tagline
            tagline = sd.notes.split("|")[0].strip() if sd.notes else meta[2]
            sub     = meta[1]
            _render_section(slide, sd, theme, page_no, total,
                            section_idx=_section_counter[0],
                            day_label=meta[0],
                            subtitle_line=sub,
                            tagline=tagline)

        elif kind in ("bullets", "agenda"):
            lbl = "WUZHEN · 24 HOURS" if "乌镇" in sd.title else \
                  "QIANDAO LAKE · ONE DAY" if "千岛湖" in sd.title else \
                  "FIELD NOTES" if "黄山" in sd.title and "路线" not in sd.title else \
                  "STAY" if "住宿" in sd.title or "民宿" in sd.title else \
                  sd.kind.upper()
            _render_bullets_editorial(slide, sd, theme, page_no, total, lbl)

        elif kind == "cards":
            _render_cards_editorial(slide, sd, theme, page_no, total)

        elif kind == "timeline":
            lbl = "ROUTE" if "总览" in sd.title or "行程" in sd.title else "TIMELINE"
            _render_timeline_editorial(slide, sd, theme, page_no, total, lbl)

        elif kind == "comparison":
            _render_comparison_editorial(slide, sd, theme, page_no, total)

        elif kind == "process":
            _render_process_editorial(slide, sd, theme, page_no, total)

        elif kind == "kpi":
            _render_kpi_editorial(slide, sd, theme, page_no, total)

        elif kind == "summary":
            _render_summary_editorial(slide, sd, theme, page_no, total)

        elif kind == "thanks":
            _render_thanks_editorial(slide, sd, theme, plan)

        elif kind == "chart":
            _render_chart_editorial(slide, sd, theme, page_no, total)

        elif kind == "table":
            _render_table_editorial(slide, sd, theme, page_no, total)

        else:
            # 兜底用 bullets
            _render_bullets_editorial(slide, sd, theme, page_no, total)

        # 备注
        try:
            if sd.notes:
                slide.notes_slide.notes_text_frame.text = sd.notes
        except Exception:
            pass

    # 动画
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
