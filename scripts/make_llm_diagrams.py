"""生成 LLM 科普所需的 3 张原理图 PNG（供 pptx agent 作为图片素材使用）。"""
import math
import os

from PIL import Image, ImageDraw, ImageFont

FONT_PATH = "C:/Windows/Fonts/msyh.ttc"
OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "examples", "llm_intro")


def _font(size: int):
    try:
        return ImageFont.truetype(FONT_PATH, size)
    except Exception:
        return ImageFont.load_default()


def _text_center(d, cx, cy, text, font, fill=(255, 255, 255)):
    bbox = d.textbbox((0, 0), text, font=font)
    w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text((cx - w / 2 - bbox[0], cy - h / 2 - bbox[1]), text, font=font, fill=fill)


def _rounded(d, box, r, fill, outline=None, width=3):
    d.rounded_rectangle(box, radius=r, fill=fill, outline=outline, width=width)


def _arrow(d, x1, y1, x2, y2, color="#93C5FD", width=5):
    d.line([x1, y1, x2, y2], fill=color, width=width)
    ang = math.atan2(y2 - y1, x2 - x1)
    L = 16
    for da in (math.radians(150), math.radians(-150)):
        d.line([x2, y2,
                x2 + L * math.cos(ang + da), y2 + L * math.sin(ang + da)],
               fill=color, width=width)


# ---------------------------------------------------------------------------
# 1) Agent ReAct 循环
# ---------------------------------------------------------------------------
def agent_loop(path: str) -> None:
    W, H = 1240, 780
    img = Image.new("RGB", (W, H), "#0E1E3A")
    d = ImageDraw.Draw(img)
    _text_center(d, W / 2, 64, "Agent 的 ReAct 工作循环", _font(46), "#FFFFFF")
    _text_center(d, W / 2, 118, "Thought → Action → Observation，循环直到任务完成",
                 _font(26), "#93C5FD")

    cx0, cy0, R = W / 2, H / 2 + 30, 235
    nodes = [
        ("思考 Thought", "分析问题 · 拆解任务", "#2563EB"),
        ("行动 Action", "调用工具 · 执行动作", "#16A34A"),
        ("观察 Observation", "读取结果 · 修正计划", "#D97706"),
    ]
    pts = []
    for i, (name, desc, color) in enumerate(nodes):
        ang = math.radians(90 - i * 120)
        pts.append((cx0 + R * math.cos(ang), cy0 - R * math.sin(ang)))

    # 先画箭头三角（在中心圆之下）
    for i in range(3):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % 3]
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        dx, dy = mx - cx0, my - cy0
        nd = math.hypot(dx, dy) or 1
        off = 95
        a = (mx - cx0 - dx / nd * off, my - cy0 - dy / nd * off)
        _arrow(d, x1, y1, a[0], a[1], "#60A5FA", 6)

    # 中心圆
    _rounded(d, [cx0 - 100, cy0 - 100, cx0 + 100, cy0 + 100], 50, "#F59E0B")
    _text_center(d, cx0, cy0, "Agent", _font(38), "#FFFFFF")

    # 节点
    for (name, desc, color), (px, py) in zip(nodes, pts):
        _rounded(d, [px - 175, py - 85, px + 175, py + 85], 26, color)
        _text_center(d, px, py - 22, name, _font(32), "#FFFFFF")
        _text_center(d, px, py + 34, desc, _font(24), "#E5E7EB")

    img.save(path)
    print("[图] agent-loop.png")


# ---------------------------------------------------------------------------
# 2) LLM 训练-应用流水线
# ---------------------------------------------------------------------------
def llm_pipeline(path: str) -> None:
    W, H = 1240, 520
    img = Image.new("RGB", (W, H), "#0F2A43")
    d = ImageDraw.Draw(img)
    _text_center(d, W / 2, 60, "LLM 训练与落地流水线", _font(44), "#FFFFFF")

    stages = [
        ("预训练", "海量语料学习语言", "#1D4ED8"),
        ("指令微调 SFT", "对齐人类指令", "#0D9488"),
        ("RLHF 对齐", "符合偏好与安全", "#9333EA"),
        ("推理应用", "对话 / Agent / 工具", "#EA580C"),
    ]
    n = len(stages)
    box_w, box_h, gap = 250, 150, 60
    total_w = n * box_w + (n - 1) * gap
    x0 = (W - total_w) / 2
    y0 = 150
    for i, (name, desc, color) in enumerate(stages):
        x = x0 + i * (box_w + gap)
        _rounded(d, [x, y0, x + box_w, y0 + box_h], 22, color)
        _text_center(d, x + box_w / 2, y0 + 45, name, _font(30), "#FFFFFF")
        _text_center(d, x + box_w / 2, y0 + 100, desc, _font(22), "#E5E7EB")
        if i < n - 1:
            _arrow(d, x + box_w, y0 + box_h / 2,
                   x + box_w + gap, y0 + box_h / 2, "#7DD3FC", 6)

    img.save(path)
    print("[图] llm-pipeline.png")


# ---------------------------------------------------------------------------
# 3) Skill 构成
# ---------------------------------------------------------------------------
def skill_stack(path: str) -> None:
    W, H = 1240, 640
    img = Image.new("RGB", (W, H), "#14213D")
    d = ImageDraw.Draw(img)
    _text_center(d, W / 2, 56, "Skill = 能力封装，按需复用", _font(42), "#FFFFFF")

    # 外层 Skill 容器
    outer = [80, 120, W - 80, H - 90]
    _rounded(d, outer, 26, "#1E3A5F", "#60A5FA", 4)
    _text_center(d, W / 2, outer[1] + 45, "Skill（标准作业程序）", _font(30), "#93C5FD")

    items = [
        ("指令", "目标 · 约束 · 边界", "#2563EB"),
        ("示例", "输入输出示范", "#0D9488"),
        ("工具", "函数 · 搜索 · API", "#16A34A"),
        ("触发条件", "何时调用 · 何时停止", "#D97706"),
    ]
    inner_top = outer[1] + 95
    inner_h = outer[3] - inner_top - 60
    box_w = (outer[2] - outer[0] - 3 * 30) / 4
    for i, (name, desc, color) in enumerate(items):
        x = outer[0] + 30 + i * (box_w + 30)
        _rounded(d, [x, inner_top, x + box_w, inner_top + inner_h], 18, color)
        _text_center(d, x + box_w / 2, inner_top + 42, name, _font(30), "#FFFFFF")
        _text_center(d, x + box_w / 2, inner_top + inner_h - 45, desc,
                     _font(21), "#E5E7EB")

    # 底座：模型 + 上下文
    base_y = outer[3] - 38
    _rounded(d, [outer[0] + 30, base_y, outer[2] - 30, base_y + 60], 12,
             "#F59E0B")
    _text_center(d, W / 2, base_y + 30, "底层：大模型 + 上下文（模型执行指令）",
                 _font(22), "#FFFFFF")

    img.save(path)
    print("[图] skill-stack.png")


def main() -> None:
    os.makedirs(OUT_DIR, exist_ok=True)
    agent_loop(os.path.join(OUT_DIR, "图_Agent循环.png"))
    llm_pipeline(os.path.join(OUT_DIR, "图_LLM流水线.png"))
    skill_stack(os.path.join(OUT_DIR, "图_Skill构成.png"))
    print("原理图已生成：" + OUT_DIR)


if __name__ == "__main__":
    main()
