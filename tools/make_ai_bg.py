"""生成 AI 科技主题背景图（PIL，无需外部素材）。

输出 assets/ai_bg.png        —— 封面/章节：深色渐变 + 透视网格 + 电路光点 + 机器人头像
     assets/ai_bg_soft.png   —— 内容页：同上但更暗、机器人更淡（保证文字可读）

用法：python tools/make_ai_bg.py
"""
from __future__ import annotations

import os
import random

from PIL import Image, ImageDraw, ImageFilter

W, H = 1920, 1080
OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")

# 配色
C_BG_TOP = (11, 16, 32)
C_BG_BOT = (22, 30, 58)
C_CYAN = (34, 211, 238)
C_VIOLET = (168, 85, 247)
C_INDIGO = (99, 102, 241)
C_LINE = (56, 100, 160)
C_GRID = (70, 100, 170)


def _gradient(w: int, h: int, top, bot) -> Image.Image:
    img = Image.new("RGB", (w, h))
    d = ImageDraw.Draw(img)
    for y in range(h):
        t = y / h
        col = tuple(int(top[i] + (bot[i] - top[i]) * t) for i in range(3))
        d.line([(0, y), (w, y)], fill=col)
    return img


def _grid(draw: ImageDraw.ImageDraw) -> None:
    horizon, vx = 560, W / 2
    # 底部纵向透视线
    for i in range(-14, 15):
        bx = vx + i * 90
        draw.line([(bx, H), (vx, horizon)], fill=C_GRID, width=1)
    # 横向地平线（越靠近地平线越密）
    y = H
    f = 0.88
    while y > horizon + 6:
        draw.line([(0, y), (W, y)], fill=C_GRID, width=1)
        y = horizon + (y - horizon) * f
    draw.line([(0, horizon), (W, horizon)], fill=C_LINE, width=2)


def _circuit(draw: ImageDraw.ImageDraw, rng: random.Random) -> None:
    for _ in range(26):
        x, y = rng.randint(0, W), rng.randint(80, 640)
        col = C_CYAN if rng.random() < 0.5 else C_VIOLET
        segs = rng.randint(3, 6)
        pts = [(x, y)]
        for k in range(segs):
            nx, ny = pts[-1]
            if k % 2 == 0:
                nx = max(20, min(W - 20, nx + rng.choice([-1, 1]) * rng.randint(80, 260)))
            else:
                ny = max(60, min(700, ny + rng.choice([-1, 1]) * rng.randint(40, 160)))
            pts.append((nx, ny))
        draw.line(pts, fill=col + (110,), width=2, joint="curve")
        # 端点光点
        ex, ey = pts[-1]
        r = rng.randint(4, 8)
        draw.ellipse([ex - r, ey - r, ex + r, ey + r], fill=col)


def _robot_head(draw: ImageDraw.ImageDraw, cx: float, cy: float, s: float) -> None:
    hw, hh = int(175 * s), int(155 * s)
    # 头
    draw.rounded_rectangle([cx - hw, cy - hh, cx + hw, cy + hh],
                           radius=int(36 * s), fill=(30, 41, 59, 150),
                           outline=(56, 189, 248, 210), width=5)
    # 眼睛（发光青色）
    eye_w, eye_h = int(48 * s), int(26 * s)
    ey = cy - int(22 * s)
    for ex in (cx - hw + int(32 * s), cx + hw - int(32 * s) - eye_w):
        draw.rounded_rectangle([ex, ey, ex + eye_w, ey + eye_h], radius=9,
                               fill=(34, 211, 238, 235))
    # 天线 + 发光顶端
    ay = cy - hh - int(55 * s)
    draw.line([cx, cy - hh, cx, ay], fill=(56, 189, 248, 200), width=4)
    draw.ellipse([cx - 9 * s, ay - 22 * s, cx + 9 * s, ay], fill=(168, 85, 247, 255))
    # 耳
    ear_w, ear_h = int(22 * s), int(48 * s)
    draw.rectangle([cx - hw - int(22 * s), cy - int(26 * s),
                    cx - hw, cy - int(26 * s) + ear_h], fill=(51, 65, 85, 210))
    draw.rectangle([cx + hw, cy - int(26 * s),
                    cx + hw + int(22 * s), cy - int(26 * s) + ear_h], fill=(51, 65, 85, 210))
    # 嘴线
    draw.line([cx - int(42 * s), cy + int(42 * s), cx + int(42 * s), cy + int(42 * s)],
              fill=(56, 189, 248, 170), width=4)


def build(dimmed: bool = False, robot: bool = True, robot_alpha: float = 1.0) -> Image.Image:
    base = _gradient(W, H, C_BG_TOP, C_BG_BOT)
    if dimmed:
        base = base.point(lambda p: int(p * 0.55))
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    _grid(od)
    _circuit(od, random.Random(42))
    # 弱光晕（增强科技感）
    glow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    gd = ImageDraw.Draw(glow)
    gd.ellipse([W * 0.58, -100, W * 0.92, 420], fill=(99, 102, 241, 60))
    gd.ellipse([W * 0.02, H * 0.5, W * 0.3, H * 1.1], fill=(34, 211, 238, 40))
    glow = glow.filter(ImageFilter.GaussianBlur(90))
    overlay.alpha_composite(glow)

    if robot:
        if robot_alpha >= 1.0:
            _robot_head(od, W * 0.72, H * 0.40, 1.55)
        else:
            # 机器人淡版：先画在独立层再整层降透明度
            alpha_layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
            _robot_head(ImageDraw.Draw(alpha_layer), W * 0.72, H * 0.40, 1.55)
            alpha_layer = alpha_layer.point(lambda p: int(p * robot_alpha))
            overlay.alpha_composite(alpha_layer)

    base = base.convert("RGBA")
    base.alpha_composite(overlay)
    return base.convert("RGB")


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)
    build(dimmed=False, robot=True).save(os.path.join(OUT_DIR, "ai_bg.png"), quality=90)
    build(dimmed=True, robot=True, robot_alpha=0.45).save(
        os.path.join(OUT_DIR, "ai_bg_soft.png"), quality=90)
    print(f"[AI 背景] 已生成 {OUT_DIR}\\ai_bg.png / ai_bg_soft.png")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
