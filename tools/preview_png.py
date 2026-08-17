"""轻量 PPTX -> PNG 预览渲染器（无需 LibreOffice）。

按每页真实几何用 PIL 重绘 1280x720 缩略图：背景色、色块/卡片、标题与要点文本
（自动换行、按 run 字号/颜色）、真实图片、表格、图表（柱状近似）。

输出：
    <out-dir>/preview_<n>.png   每页一张
    <out-dir>/contact.png       总览拼图

用法：
    python tools/preview_png.py output/demo_final.pptx [--out-dir preview]
"""
from __future__ import annotations

import argparse
import io
import os

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

FONT_REGULAR = r"C:\Windows\Fonts\msyh.ttc"
FONT_BOLD = r"C:\Windows\Fonts\msyhbd.ttc"
W, H = 1280, 720
EMU_PER_IN = 914400

# 会被从文本中剔除的装饰性 emoji 区间（预览字体不支持彩色 emoji）
_EMOJI_RANGES = (
    range(0x1F000, 0x1FAFF), range(0x2600, 0x27BF),
    range(0x2B00, 0x2BFF), range(0xFE00, 0xFE0F),
)


def _strip_emoji(text: str) -> str:
    return "".join(ch for ch in text
                   if not any(ch in r for r in _EMOJI_RANGES))


def _load_font(px: int, bold: bool = False):
    try:
        return ImageFont.truetype(FONT_BOLD if bold else FONT_REGULAR, px)
    except Exception:
        return ImageFont.load_default()


def _pt_to_px(pt) -> int:
    return max(8, int(round((pt or 12) * 96 / 72)))


def _wrapped(draw, text: str, font, max_w: int) -> list[str]:
    lines: list[str] = []
    for raw in text.split("\n"):
        if not raw:
            lines.append("")
            continue
        cur = ""
        for ch in raw:
            if draw.textlength(cur + ch, font=font) <= max_w:
                cur += ch
            else:
                lines.append(cur)
                cur = ch
        lines.append(cur)
    return lines


def _solid_rgb(color_obj):
    try:
        rgb = color_obj.rgb
        return (rgb[0], rgb[1], rgb[2])
    except Exception:
        return None


def _slide_bg(slide):
    try:
        if slide.background.fill.type == 1:  # SOLID
            c = _solid_rgb(slide.background.fill.fore_color)
            if c:
                return c
    except Exception:
        pass
    return (255, 255, 255)


def _run_style(p):
    """取段落首个 run 的字号/颜色/粗体。"""
    r0 = p.runs[0] if p.runs else None
    pt = r0.font.size.pt if (r0 and r0.font.size) else 12
    color = _solid_rgb(r0.font.color) if (r0 and r0.font.color and
                                          r0.font.color.type is not None) else None
    if color is None:
        color = (30, 30, 30)
    return pt, color, bool(r0 and r0.font.bold)


def _draw_text(img, draw, sh, sx):
    if not sh.has_text_frame:
        return
    x0, y0 = sh.left * sx, sh.top * sx
    w, h = sh.width * sx, sh.height * sx
    y = y0
    for p in sh.text_frame.paragraphs:
        txt = _strip_emoji("".join(r.text for r in p.runs) or "")
        if not txt.strip():
            y += _pt_to_px(12) * 0.8
            continue
        pt, color, bold = _run_style(p)
        px = _pt_to_px(pt)
        font = _load_font(px, bold)
        lines = _wrapped(draw, txt, font, max(8, w - 4))
        line_h = px * 1.35
        for ln in lines:
            draw.text((x0 + 2, y), ln, font=font, fill=color)
            y += line_h
    # 文本框过高时截断，避免画出界
    return


def _draw_picture(img, draw, sh, sx):
    try:
        data = sh.image.blob
        im = Image.open(io.BytesIO(data)).convert("RGB")
    except Exception:
        return
    x0, y0 = sh.left * sx, sh.top * sx
    w, h = sh.width * sx, sh.height * sx
    im.thumbnail((int(w), int(h)))
    img.paste(im, (int(x0 + (w - im.width) / 2), int(y0 + (h - im.height) / 2)))


def _draw_table(img, draw, sh, sx):
    if not sh.has_table:
        return
    tbl = sh.table
    x0, y0 = sh.left * sx, sh.top * sx
    w, h = sh.width * sx, sh.height * sx
    nrows, ncols = len(tbl.rows), len(tbl.columns)
    rw, cw = h / nrows, w / ncols
    for r in range(nrows):
        for c in range(ncols):
            cell = tbl.cell(r, c)
            fill = _solid_rgb(cell.fill.fore_color) or (255, 255, 255)
            cx, cy = x0 + c * cw, y0 + r * rw
            draw.rectangle([cx, cy, cx + cw, cy + rw], fill=fill,
                           outline=(170, 170, 170))
            txt = _strip_emoji(cell.text or "")
            if not txt:
                continue
            font = _load_font(13, bold=(r == 0))
            fg = (255, 255, 255) if (r == 0 and fill != (255, 255, 255)) else (30, 30, 30)
            ty = cy + 3
            for ln in _wrapped(draw, txt, font, cw - 8)[:2]:
                draw.text((cx + 4, ty), ln, font=font, fill=fg)
                ty += 17


def _draw_chart(img, draw, sh, sx):
    if not sh.has_chart:
        return
    x0, y0 = sh.left * sx, sh.top * sx
    w, h = sh.width * sx, sh.height * sx
    try:
        plot = sh.chart.plots[0]
        cats = [str(c) for c in plot.categories]
        series = [(s.name or "", list(s.values)) for s in plot.series]
    except Exception:
        series, cats = [], []
    if not series:
        return
    pad = 40
    aw, ah = w - 2 * pad, h - 2 * pad
    ncats = max(len(cats), 1)
    nser = len(series)
    all_vals = [v for _, vals in series for v in vals if isinstance(v, (int, float))]
    vmax = max(all_vals) if all_vals else 1
    col_w = aw / ncats
    ser_w = col_w * 0.7 / max(nser, 1)
    palette = [(31, 78, 121), (0, 176, 240), (255, 192, 0),
               (31, 56, 100), (191, 144, 0), (192, 0, 0)]
    draw.rectangle([x0 + pad, y0 + pad, x0 + w - pad, y0 + h - pad],
                   outline=(200, 200, 200))
    for i, cat in enumerate(cats):
        cx = x0 + pad + i * col_w + (col_w - ser_w * nser) / 2
        for j, (name, vals) in enumerate(series):
            v = vals[i] if i < len(vals) else 0
            if not isinstance(v, (int, float)):
                v = 0
            bh = (v / vmax) * (ah - 40)
            bx = cx + j * ser_w
            draw.rectangle([bx, y0 + pad + ah - bh, bx + ser_w,
                            y0 + pad + ah - 4],
                           fill=palette[j % len(palette)])
        f = _load_font(12)
        draw.text((cx, y0 + h - pad + 8), cat[:6], font=f, fill=(60, 60, 60))


def _draw_shape(img, draw, sh, sx):
    if sh.left is None or sh.top is None:
        return
    st = sh.shape_type
    if st == MSO_SHAPE_TYPE.PICTURE:
        _draw_picture(img, draw, sh, sx)
        return
    if st == MSO_SHAPE_TYPE.TABLE:
        _draw_table(img, draw, sh, sx)
        return
    if sh.has_chart:
        _draw_chart(img, draw, sh, sx)
        return
    if st == MSO_SHAPE_TYPE.GROUP:
        return

    x0, y0 = sh.left * sx, sh.top * sx
    w, h = sh.width * sx, sh.height * sx

    # 有填充的自动形状 -> 画色块/圆角卡片
    try:
        is_solid = sh.fill.type == 1  # SOLID
    except Exception:
        is_solid = False
    if is_solid and w > 1 and h > 1:
        fill = _solid_rgb(sh.fill.fore_color) or (200, 200, 200)
        rounded = False
        try:
            from pptx.enum.shapes import MSO_SHAPE
            rounded = (st == MSO_SHAPE_TYPE.AUTO_SHAPE
                       and sh.auto_shape_type == MSO_SHAPE.ROUNDED_RECTANGLE)
        except Exception:
            pass
        if rounded:
            draw.rounded_rectangle([x0, y0, x0 + w, y0 + h], radius=min(w, h) * 0.06,
                                   fill=fill)
        else:
            draw.rectangle([x0, y0, x0 + w, y0 + h], fill=fill)

    _draw_text(img, draw, sh, sx)


def render_deck(path: str, out_dir: str = "preview") -> list[str]:
    prs = Presentation(path)
    sx = W / prs.slide_width   # px per EMU
    os.makedirs(out_dir, exist_ok=True)
    pages: list[str] = []
    for idx, slide in enumerate(prs.slides, 1):
        img = Image.new("RGB", (W, H), _slide_bg(slide))
        draw = ImageDraw.Draw(img)
        for sh in slide.shapes:
            try:
                _draw_shape(img, draw, sh, sx)
            except Exception:
                continue
        out = os.path.join(out_dir, f"preview_{idx:02d}.png")
        img.save(out)
        pages.append(out)
    _contact_sheet(pages, out_dir)
    return pages


def _contact_sheet(pages: list[str], out_dir: str) -> str:
    thumbs = []
    tw = 320
    for p in pages:
        im = Image.open(p)
        ratio = tw / im.width
        thumbs.append(im.resize((tw, int(im.height * ratio)), Image.LANCZOS))
    if not thumbs:
        return ""
    per_row = 4
    rows = (len(thumbs) + per_row - 1) // per_row
    th = thumbs[0].height
    pad = 12
    sheet = Image.new("RGB", (per_row * tw + (per_row + 1) * pad,
                             rows * (th + 26) + (rows + 1) * pad), (245, 245, 245))
    d = ImageDraw.Draw(sheet)
    font = _load_font(16, bold=True)
    for i, t in enumerate(thumbs):
        r, c = divmod(i, per_row)
        x = pad + c * (tw + pad)
        y = pad + r * (th + 26)
        sheet.paste(t, (x, y))
        d.text((x + 2, y + th + 2), f"{i + 1:02d}", font=font, fill=(40, 40, 40))
    out = os.path.join(out_dir, "contact.png")
    sheet.save(out)
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="PPTX -> PNG 预览（无 LibreOffice）")
    ap.add_argument("deck", help="输入 .pptx 路径")
    ap.add_argument("--out-dir", default="preview", help="输出目录")
    args = ap.parse_args()
    pages = render_deck(args.deck, args.out_dir)
    print(f"[预览] 已渲染 {len(pages)} 页 -> {os.path.abspath(args.out_dir)}")
    print(f"[预览] 总览图：{os.path.join(os.path.abspath(args.out_dir), 'contact.png')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
