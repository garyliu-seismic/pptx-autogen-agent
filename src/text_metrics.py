"""中英文混排文本度量：估算字符宽度与所需高度（单位：英寸）。

排版(renderer)与质检(quality)共用同一套模型，避免“排版盒子”与“质检判断”不一致。
- 全角字符(CJK 等)按 1 em 计宽
- 半角字符(ASCII)按 0.55 em 计宽
- 1 em = size(pt) / 72 英寸
"""
from __future__ import annotations

import math

CJK_MIN = 0x2E80     # 中日韩统一表意文字起点（近似覆盖常用中文/日文/韩文）
ASCII_WEIGHT = 0.55  # 半角字符相对全角的宽度权重


def em(size_pt: float) -> float:
    return size_pt / 72.0


def units(text: str) -> float:
    """把文本换算成“全角字符等价单位数”。"""
    u = 0.0
    for ch in text:
        u += 1.0 if ord(ch) >= CJK_MIN else ASCII_WEIGHT
    return u


def lines_needed(text: str, size_pt: float, width_in: float) -> int:
    """估算文本在给定宽度(英寸)与字号下需要几行。"""
    usable = max(2.0, width_in - 0.15)
    per_line = max(1, int(usable / em(size_pt)))
    return max(1, math.ceil(units(text) / per_line))


def est_height(text: str, size_pt: float, width_in: float,
               line_factor: float = 0.024, base: float = 0.16,
               pad: float = 0.0) -> float:
    """估算文本高度（英寸）。line_factor 为行高系数。"""
    n = lines_needed(text, size_pt, width_in)
    return n * (size_pt * line_factor) + base + pad
