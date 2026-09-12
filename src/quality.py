"""渲染质量检测 + 自动重排控制器。

流程：render -> 打开生成的 pptx 做几何检测 -> 若不理想，自动调整布局参数重排
      -> 再次检测，直到达标或达到最大尝试次数。

检测规则（几何启发式）：
  overflow_x     : 形状超出画布右/左边界
  overflow_y     : 形状超出画布下边界
  overlap        : 文本框两两重叠超过阈值
  text_overflow  : 文本框内文字“偏保守估算高度”超过框高（大概率文字溢出）

重排策略：
  发现上述问题 -> 缩小字号(font_scale↓) + 加大文本框留白(hpad↑) + 收紧间距(gap↓)，
  重新渲染，循环至多 max_tries 次。
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

from pptx import Presentation

from . import text_metrics
from .models import DeckPlan

# 阈值（英寸）
MIN_SEP_IN = 0.10        # 文本框允许的最小重叠
EDGE_MARGIN_IN = 0.20    # 距离左右边界的最小边距
BOTTOM_MARGIN_IN = 0.35  # 距离下边界的最小边距
SPARSE_RATIO = 0.32      # 正文底部低于该比例视为“太稀疏”（底部大块空白）


@dataclass
class Issue:
    slide: int
    kind: str            # overflow_x / overflow_y / overlap / text_overflow / sparse
    value: float = 0.0
    detail: str = ""


class LayoutConstraintError(RuntimeError):
    """hard_fail=True 时，重排+LLM 修复用尽仍有严重版式问题，拒绝「将就」输出。"""

    def __init__(self, issues: list[Issue]):
        self.issues = issues
        super().__init__(summary(issues))


def _max_run_size(sh) -> float:
    size = 0.0
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size is not None:
                size = max(size, r.font.size.pt)
    return size


def _est_text_height(sh) -> float:
    """按文本框实际宽度与字号估算其文字需要的高度（与排版同模型，略保守）。"""
    size = _max_run_size(sh) or 12.0
    w_in = sh.width / 914400 if sh.width else 12.0
    total_lines = 0
    for para in sh.text_frame.paragraphs:
        txt = "".join(r.text for r in para.runs) or para.text or ""
        if txt:
            total_lines += text_metrics.lines_needed(txt, size, w_in)
    return max(1, total_lines) * (size * 0.026) + 0.1


def _overlap_xy(a, b) -> float:
    ox = min(a[2], b[2]) - max(a[0], b[0])
    oy = min(a[3], b[3]) - max(a[1], b[1])
    return min(ox, oy) if ox > 0 and oy > 0 else 0.0


def analyze(path: str) -> list[Issue]:
    """打开生成的 pptx，返回所有排版问题（已排除装饰条/页脚等误报）。"""
    prs = Presentation(path)
    W = prs.slide_width / 914400
    H = prs.slide_height / 914400
    issues: list[Issue] = []

    for idx, slide in enumerate(prs.slides, 1):
        texts: list[tuple[float, float, float, float]] = []
        bottom_text = 0.0

        for sh in slide.shapes:
            if sh.left is None or sh.top is None:
                continue
            # 非法几何：宽/高为负（OOXML 无效，PowerPoint 会提示“修复”）
            if (sh.width is not None and sh.width < 0) \
                    or (sh.height is not None and sh.height < 0):
                bad = min(sh.width or 0, sh.height or 0) / 914400
                issues.append(Issue(idx, "negative_geometry", abs(bad),
                                    f"{sh.shape_type} {sh.name} 宽/高为负"))
                continue
            l = sh.left / 914400
            t = sh.top / 914400
            w = sh.width / 914400 if sh.width else 0
            h = sh.height / 914400 if sh.height else 0
            r, b = l + w, t + h

            # 排除装饰：通栏色条 / 细线 / 极小元素
            if w >= W - 0.5 or h <= 0.07 or (w < 0.2 and h < 0.2):
                continue

            # 越界检测（横向）
            if w > 0.3:
                if r > W - EDGE_MARGIN_IN:
                    issues.append(Issue(idx, "overflow_x", r - (W - EDGE_MARGIN_IN),
                                        f"{sh.shape_type} {sh.name}"))
                if l < EDGE_MARGIN_IN - 0.05:
                    issues.append(Issue(idx, "overflow_x", EDGE_MARGIN_IN - l,
                                        f"{sh.shape_type} {sh.name}"))
            # 越界检测（纵向）：跳过页脚区（底部 0.6in）的小元素（页码/页脚/小表）
            in_footer_band = b > H - 0.6 and h < 0.9
            if h > 0.2 and t < H - 0.6 and b > H - BOTTOM_MARGIN_IN and not in_footer_band:
                issues.append(Issue(idx, "overflow_y", b - (H - BOTTOM_MARGIN_IN),
                                    f"{sh.shape_type} {sh.name}"))

            # 文本框：重叠检测 + 文本溢出风险（仅内容文本框）
            if sh.shape_type == 17 and sh.has_text_frame and w >= 1.5 and h >= 0.3:
                text = sh.text_frame.text.strip()
                max_size = _max_run_size(sh)
                if (text and t < H - 0.6 and max_size < 36   # 排除页脚/大标题
                        and not in_footer_band):
                    est = _est_text_height(sh)
                    if est > h * 1.25 + 0.08:
                        issues.append(Issue(idx, "text_overflow", est - h, sh.name))
                    texts.append((l, t, r, b))
                    bottom_text = max(bottom_text, b)

        # 文本框两两重叠
        n = len(texts)
        for i in range(n):
            for j in range(i + 1, n):
                ov = _overlap_xy(texts[i], texts[j])
                if ov > MIN_SEP_IN:
                    issues.append(Issue(idx, "overlap", ov,
                                        f"text boxes {i}-{j}"))

        # 稀疏检测：有正文且底部远未用满 -> 底部大块空白
        if texts and len(texts) >= 2 and bottom_text < H * SPARSE_RATIO:
            issues.append(Issue(idx, "sparse", H * SPARSE_RATIO - bottom_text,
                                "too much bottom whitespace"))

    return issues


def summary(issues: list[Issue]) -> str:
    if not issues:
        return "质检通过：无溢出/重叠/文本溢出问题。"
    from collections import Counter
    cnt = Counter(i.kind for i in issues)
    return "质检发现: " + ", ".join(f"{k}×{v}" for k, v in cnt.items())


def render_with_quality(plan: DeckPlan, output_path: str,
                        max_tries: int = 4, quiet: bool = False,
                        repairer=None, repair_rounds: int = 0,
                        max_repair_rounds: int = 2,
                        rotate_map: dict | None = None,
                        hard_fail: bool = False) -> tuple[str, list[Issue], int]:
    """渲染 + 质检 + 自动重排，返回 (输出路径, 最终问题列表, 尝试次数)。

    快路径：确定性参数调节（font/hpad/gap），最多 max_tries 次。
    慢路径：参数调节仍无法消除严重问题、且传入 repairer 时，把具体 Issue
    交给 repairer(plan, issues)（通常是 LLM 修复，见 llm_repair.py）返回新计划
    后重新渲染；最多 max_repair_rounds 轮。
    rotate_map：{文件名: 度数}，逐图旋转（透传给 renderer）。
    hard_fail：两条路都用尽仍有严重问题时，不再「将就」返回，而是抛
    LayoutConstraintError。
    """
    from .renderer import render

    state = {"font": 1.0, "hpad": 0.0, "gap": 1.0}
    seen: set[tuple] = set()
    tightened = False
    final_issues: list[Issue] = []

    for attempt in range(1, max_tries + 1):
        key = (round(state["font"], 2), round(state["hpad"], 2), round(state["gap"], 2))
        if key in seen:
            break
        seen.add(key)

        render(plan, output_path,
               font_scale=state["font"], hpad=state["hpad"], gap_mult=state["gap"],
               rotate_map=rotate_map)
        final_issues = analyze(output_path)

        critical = [i for i in final_issues
                    if i.kind in ("overflow_x", "overflow_y", "overlap",
                                  "text_overflow", "negative_geometry")]
        sparse = [i for i in final_issues if i.kind == "sparse"]
        if not quiet:
            print(f"  [质检] 第 {attempt} 次: {summary(final_issues)}")

        if not critical:
            if sparse and not tightened and state["font"] < 1.12 and attempt < max_tries:
                # 底部太空且从未收紧 -> 放大字号 + 加行距，填充版面
                state = {"font": 1.12, "hpad": max(0.0, state["hpad"] * 0.5),
                         "gap": 1.2}
                continue
            return output_path, final_issues, attempt

        # 有问题 -> 收紧：缩小字号、加大留白防溢出、压缩间距
        tightened = True
        state = {"font": max(0.72, state["font"] - 0.06),
                 "hpad": state["hpad"] + 0.12,
                 "gap": max(0.6, state["gap"] - 0.12)}

    # 慢路径：确定性参数调节用尽仍有严重问题 -> 交给 repairer（通常为 LLM）
    if repairer is not None and repair_rounds < max_repair_rounds:
        critical = [i for i in final_issues
                    if i.kind in ("overflow_x", "overflow_y", "overlap",
                                  "text_overflow", "negative_geometry")]
        if critical:
            if not quiet:
                print(f"  [质检] 参数调节未消除排版问题，调用 LLM 修复"
                      f"（第 {repair_rounds + 1}/{max_repair_rounds} 轮）…")
            new_plan = repairer(plan, final_issues)
            if new_plan is not None:
                if not quiet:
                    print("  [质检] 已采用修复后的计划，重新渲染…")
                return render_with_quality(new_plan, output_path,
                                           max_tries=max_tries, quiet=quiet,
                                           repairer=repairer,
                                           repair_rounds=repair_rounds + 1,
                                           max_repair_rounds=max_repair_rounds,
                                           rotate_map=rotate_map,
                                           hard_fail=hard_fail)

    if hard_fail:
        critical = [i for i in final_issues
                    if i.kind in ("overflow_x", "overflow_y", "overlap",
                                  "text_overflow", "negative_geometry")]
        if critical:
            raise LayoutConstraintError(critical)

    if not quiet:
        print(f"  [质检] 已达最大尝试 {max_tries} 次，采用当前结果。")
    return output_path, final_issues, max_tries
