"""全局配置：默认路径与主题预设。"""
from __future__ import annotations

from dataclasses import dataclass

# ---------------------------------------------------------------------------
# 默认路径
# ---------------------------------------------------------------------------
DEFAULT_INPUT_DIR = "input"                # 素材目录（txt / csv / pdf / png）
DEFAULT_OUTPUT_PATH = "output/deck.pptx"   # 输出 pptx 路径
DEFAULT_PLAN_PATH = "output/plan.json"     # --plan-only 输出的计划 JSON 路径


# ---------------------------------------------------------------------------
# 主题预设：配色 + 字体（可被内容关键词自动命中，也可用 --theme 强制指定）
# ---------------------------------------------------------------------------
@dataclass
class Theme:
    name: str        # 主题名
    primary: str     # 主色
    secondary: str   # 辅色
    accent: str      # 强调色
    bg: str          # 内容页背景
    fg: str          # 内容页文字色
    font: str = "Microsoft YaHei"  # 中英文字体（PowerPoint/Windows 通用）


THEMES: dict[str, Theme] = {
    "tech":       Theme("科技蓝", "1F4E79", "00B0F0", "FFC000", "FFFFFF", "1A1A2E"),
    "business":   Theme("商务金", "1F3864", "BF9000", "C00000", "FFFFFF", "222222"),
    "finance":    Theme("金融墨", "0B3C5D", "328CC1", "D9B310", "FFFFFF", "1C2833"),
    "health":     Theme("健康绿", "0F6C4F", "2E9E6B", "FF8A5C", "F7FBF9", "1F2937"),
    "education":  Theme("教育橙", "B45309", "F59E0B", "7C3AED", "FFFDF5", "292524"),
    "nature":     Theme("自然青", "0E7490", "06B6D4", "84CC16", "F0FDFA", "164E63"),
    "elegant":    Theme("简约灰", "374151", "6B7280", "111827", "FFFFFF", "1F2937"),
    "modern":     Theme("现代靛", "312E81", "6366F1", "F59E0B", "FFFFFF", "1E1B4B"),
}
DEFAULT_THEME = "modern"
