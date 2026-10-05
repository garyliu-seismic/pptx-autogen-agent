"""为每个主题生成专属背景图（封面版 + 内容淡版）。

使用 Pillow 纯 Python 生成，无外部依赖。
生成策略：
  - 封面/章节（bg_image）：深色渐变底 + 几何装饰纹理 + 主题色光晕
  - 内容页（bg_image_soft）：极淡色调，透明度约 8-12%，确保文字可读

运行方式：
    python tools/make_theme_bg.py              # 生成所有主题
    python tools/make_theme_bg.py fresh warm   # 只生成指定主题
"""
from __future__ import annotations

import math
import os
import sys

from PIL import Image, ImageDraw, ImageFilter

# 输出目录
ASSETS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                           "assets", "backgrounds")
os.makedirs(ASSETS_DIR, exist_ok=True)

W, H = 1920, 1080  # 16:9 全高清


# ---------------------------------------------------------------------------
# 颜色工具
# ---------------------------------------------------------------------------
def _hex(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _blend(c1: tuple, c2: tuple, t: float) -> tuple[int, int, int]:
    return tuple(int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3))


def _tint(c: tuple, factor: float) -> tuple[int, int, int]:
    """向白色混合。"""
    return tuple(int(c[i] + (255 - c[i]) * factor) for i in range(3))


def _shade(c: tuple, factor: float) -> tuple[int, int, int]:
    """向黑色压暗。"""
    return tuple(int(c[i] * (1 - factor)) for i in range(3))


def _alpha(c: tuple, a: int) -> tuple[int, int, int, int]:
    return (*c, a)


# ---------------------------------------------------------------------------
# 背景生成函数
# ---------------------------------------------------------------------------
def make_gradient(c1: tuple, c2: tuple, direction: str = "diagonal") -> Image.Image:
    """生成双色渐变底图。"""
    img = Image.new("RGB", (W, H))
    pixels = img.load()
    for y in range(H):
        for x in range(W):
            if direction == "diagonal":
                t = (x / W * 0.6 + y / H * 0.4)
            elif direction == "radial":
                dx = (x - W * 0.3) / W
                dy = (y - H * 0.4) / H
                t = min(1.0, math.sqrt(dx * dx + dy * dy) * 1.4)
            else:  # vertical
                t = y / H
            pixels[x, y] = _blend(c1, c2, t)
    return img


def add_geometric_overlay(img: Image.Image, accent: tuple,
                           style: str = "lines") -> Image.Image:
    """在底图上叠加几何装饰。"""
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    if style == "lines":
        # 右侧斜线组
        for i in range(0, 22):
            x0 = W * 0.50 + i * 50
            draw.line([(x0, 0), (x0 - H * 0.75, H)],
                      fill=(*accent, 35), width=1)
        # 左上角水平装饰线
        for i in range(5):
            off = i * 35
            draw.line([(0, off + 20), (int(W * 0.45), off + 20)],
                      fill=(*accent, 40), width=1)

    elif style == "circles":
        # 大光晕圆（右侧）
        for r, a in [(520, 20), (390, 35), (270, 50), (160, 40)]:
            cx, cy = int(W * 0.80), int(H * 0.35)
            draw.ellipse([cx - r, cy - r, cx + r, cy + r],
                         outline=(*accent, a), width=2)
        # 小装饰圆点
        for cx, cy, r in [(120, 80, 8), (200, 200, 5), (W - 150, H - 80, 6),
                          (400, 50, 4), (W - 300, 150, 5)]:
            draw.ellipse([cx - r, cy - r, cx + r, cy + r],
                         fill=(*accent, 90))

    elif style == "grid":
        # 网格
        for x in range(0, W, 80):
            draw.line([(x, 0), (x, H)], fill=(*accent, 22), width=1)
        for y in range(0, H, 80):
            draw.line([(0, y), (W, y)], fill=(*accent, 22), width=1)
        # 右下角装饰三角
        pts = [(int(W * 0.7), H), (W, int(H * 0.4)), (W, H)]
        draw.polygon(pts, fill=(*accent, 28))

    elif style == "dots":
        # 点阵
        for x in range(60, W, 80):
            for y in range(60, H, 80):
                r = 3
                draw.ellipse([x - r, y - r, x + r, y + r],
                              fill=(*accent, 60))

    elif style == "waves":
        # 弧线组
        for i in range(6):
            y_off = H * 0.5 + i * 90
            pts = []
            for x in range(0, W + 20, 20):
                y = y_off + math.sin(x / 200 + i * 0.5) * 60
                pts.append((x, y))
            if len(pts) >= 2:
                draw.line(pts, fill=(*accent, 20 - i * 2), width=2)

    elif style == "triangles":
        # 左侧装饰三角形
        draw.polygon([(0, 0), (W * 0.45, 0), (0, H * 0.65)],
                     fill=(*accent, 20))
        draw.polygon([(W, H), (W * 0.55, H), (W, H * 0.35)],
                     fill=(*accent, 15))

    img = img.convert("RGBA")
    img = Image.alpha_composite(img, overlay)
    return img.convert("RGB")


def add_bottom_bar(img: Image.Image, accent: tuple) -> Image.Image:
    """底部渐变色条（增强封面质感）。"""
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    bar_h = 6
    for x in range(W):
        t = x / W
        c = _blend(accent, _tint(accent, 0.4), t)
        draw.line([(x, H - bar_h), (x, H)], fill=(*c, 220))
    img = img.convert("RGBA")
    img = Image.alpha_composite(img, overlay)
    return img.convert("RGB")


def make_soft_version(cover_img: Image.Image, bg_color: tuple,
                      opacity: float = 0.10) -> Image.Image:
    """生成内容页淡版：将封面图极度降低不透明度，叠加在页面背景色上。"""
    base = Image.new("RGB", (W, H), bg_color)
    # 将封面图降透明度叠加
    cover_rgba = cover_img.convert("RGBA")
    r, g, b, a_ch = cover_rgba.split()
    # 降低 opacity
    a_ch = a_ch.point(lambda x: int(x * opacity))
    cover_rgba = Image.merge("RGBA", (r, g, b, a_ch))
    base_rgba = base.convert("RGBA")
    result = Image.alpha_composite(base_rgba, cover_rgba)
    return result.convert("RGB")


# ---------------------------------------------------------------------------
# 各主题背景图生成规则
# ---------------------------------------------------------------------------
THEME_BG_RECIPES: dict[str, dict] = {
    "tech": {
        "grad_c1": "0D2137", "grad_c2": "1F4E79",
        "accent": "00B0F0", "style": "grid", "direction": "diagonal",
    },
    "business": {
        "grad_c1": "0F1F3C", "grad_c2": "1F3864",
        "accent": "C9A84C", "style": "lines", "direction": "diagonal",
    },
    "finance": {
        "grad_c1": "061A2B", "grad_c2": "0B3C5D",
        "accent": "D9B310", "style": "lines", "direction": "diagonal",
    },
    "health": {
        "grad_c1": "083D2C", "grad_c2": "0F6C4F",
        "accent": "FF8A5C", "style": "circles", "direction": "radial",
    },
    "education": {
        "grad_c1": "5C2800", "grad_c2": "92400E",
        "accent": "7C3AED", "style": "dots", "direction": "diagonal",
    },
    "nature": {
        "grad_c1": "073747", "grad_c2": "0E7490",
        "accent": "84CC16", "style": "waves", "direction": "vertical",
    },
    "elegant": {
        "grad_c1": "111827", "grad_c2": "1F2937",
        "accent": "3B82F6", "style": "lines", "direction": "diagonal",
    },
    "modern": {
        "grad_c1": "1E1B4B", "grad_c2": "312E81",
        "accent": "F59E0B", "style": "circles", "direction": "radial",
    },
    "dark_pro": {
        "grad_c1": "020617", "grad_c2": "0F172A",
        "accent": "38BDF8", "style": "grid", "direction": "diagonal",
    },
    "corporate": {
        "grad_c1": "0C1D35", "grad_c2": "1E3A5F",
        "accent": "60A5FA", "style": "lines", "direction": "diagonal",
    },
    "minimal": {
        "grad_c1": "111827", "grad_c2": "374151",
        "accent": "6366F1", "style": "dots", "direction": "vertical",
    },
    "warm": {
        "grad_c1": "431407", "grad_c2": "7C2D12",
        "accent": "FCD34D", "style": "triangles", "direction": "diagonal",
    },
    "medical": {
        "grad_c1": "003A5C", "grad_c2": "005B8E",
        "accent": "00C5DC", "style": "circles", "direction": "radial",
    },
    "luxury": {
        "grad_c1": "000000", "grad_c2": "0A0A0A",
        "accent": "FFD700", "style": "lines", "direction": "diagonal",
    },
    "fresh": {
        "grad_c1": "004D40", "grad_c2": "00897B",
        "accent": "FF4081", "style": "circles", "direction": "radial",
    },
    "academic": {
        "grad_c1": "1A252F", "grad_c2": "2C3E50",
        "accent": "2980B9", "style": "dots", "direction": "vertical",
    },
    "government": {
        "grad_c1": "7B1414", "grad_c2": "C0392B",
        "accent": "F39C12", "style": "triangles", "direction": "diagonal",
    },
}

# bg_color 用于 soft 版叠加基底（浅色主题用主题的 bg 色）
THEME_BG_COLORS: dict[str, str] = {
    "tech": "F0F7FF", "business": "FAFAFA", "finance": "F8FAFB",
    "health": "F2FAF6", "education": "FFFDF5", "nature": "F0FDFA",
    "elegant": "FFFFFF", "modern": "FAFAFE", "dark_pro": "0F172A",
    "corporate": "F8FAFC", "minimal": "FFFFFF", "warm": "FFFBF0",
    "medical": "FFFFFF", "luxury": "0A0A0A", "fresh": "F0FFFC",
    "academic": "FDFDFD", "government": "FAFAFA",
}


def generate_theme_bg(theme_name: str, recipe: dict) -> tuple[str, str]:
    """生成一个主题的封面图和内容淡图，返回 (cover_path, soft_path)。"""
    c1 = _hex(recipe["grad_c1"])
    c2 = _hex(recipe["grad_c2"])
    accent = _hex(recipe["accent"])
    style = recipe["style"]
    direction = recipe.get("direction", "diagonal")

    # 1. 渐变底
    img = make_gradient(c1, c2, direction)

    # 2. 几何纹理
    img = add_geometric_overlay(img, accent, style)

    # 3. 底部色条
    img = add_bottom_bar(img, accent)

    # 4. 极轻微高斯模糊（仅消除锯齿，保留纹理细节）
    img = img.filter(ImageFilter.GaussianBlur(radius=0.3))

    # 保存封面图
    cover_path = os.path.join(ASSETS_DIR, f"{theme_name}_bg.png")
    img.save(cover_path, "PNG", optimize=True)

    # 5. 生成内容淡版
    bg_color = _hex(THEME_BG_COLORS.get(theme_name, "FFFFFF"))
    is_dark = sum(bg_color) < 400
    opacity = 0.18 if is_dark else 0.10
    soft = make_soft_version(img, bg_color, opacity=opacity)
    soft_path = os.path.join(ASSETS_DIR, f"{theme_name}_bg_soft.png")
    soft.save(soft_path, "PNG", optimize=True)

    return cover_path, soft_path


def main(themes: list[str] | None = None) -> None:
    import config
    target = themes or list(THEME_BG_RECIPES.keys())
    print(f"生成 {len(target)} 个主题背景图 -> {ASSETS_DIR}")
    print()
    for name in target:
        if name not in THEME_BG_RECIPES:
            print(f"  [跳过] {name}（无配方）")
            continue
        print(f"  [{name}] 生成中...", end=" ", flush=True)
        cover, soft = generate_theme_bg(name, THEME_BG_RECIPES[name])
        print(f"OK  ({os.path.basename(cover)}, {os.path.basename(soft)})")

    # 自动更新 config.py 里的主题 bg_image 路径
    _patch_config(target)
    print()
    print("完成！已更新 config.py 中的 bg_image 路径。")


def _patch_config(themes: list[str]) -> None:
    """自动将生成的图片路径写入 config.py 各主题的 bg_image/bg_image_soft。"""
    config_path = os.path.join(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))), "config.py")
    with open(config_path, "r", encoding="utf-8") as f:
        src = f.read()

    import re
    for name in themes:
        if name not in THEME_BG_RECIPES:
            continue
        cover_rel = f"assets/backgrounds/{name}_bg.png"
        soft_rel = f"assets/backgrounds/{name}_bg_soft.png"

        # 找到对应主题的 bg_image / bg_image_soft 字段并替换
        # 模式：在 "name": Theme( ... ) 块内替换
        def replace_bg(m):
            block = m.group(0)
            # 替换 bg_image= 行
            block = re.sub(r'bg_image="[^"]*"',
                           f'bg_image="{cover_rel}"', block)
            block = re.sub(r'bg_image_soft="[^"]*"',
                           f'bg_image_soft="{soft_rel}"', block)
            # 如果原来没有这两个字段，在 card_radius 之前插入
            if 'bg_image=' not in block:
                block = re.sub(r'(card_radius=)',
                               f'bg_image="{cover_rel}",\n        '
                               f'bg_image_soft="{soft_rel}",\n        \\1',
                               block, count=1)
            return block

        # 匹配整个主题 Theme(...) 块
        pattern = rf'"{re.escape(name)}":\s*Theme\([^)]+\)'
        src = re.sub(pattern, replace_bg, src, flags=re.DOTALL)

    with open(config_path, "w", encoding="utf-8") as f:
        f.write(src)


if __name__ == "__main__":
    # 确保项目根目录在 sys.path
    _root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if _root not in sys.path:
        sys.path.insert(0, _root)
    args = sys.argv[1:] if len(sys.argv) > 1 else None
    main(args)
