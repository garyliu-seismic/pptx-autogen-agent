"""规划前校验与修复 —— Phase B（对应 ppt master 的 validateService 语义）。

在渲染前校验 DeckPlan 的可渲染性，避免把明显有问题的计划送进渲染器：

  - 主题名是否在预设内
  - 总页数上限
  - 未知页型
  - 空内容页
  - 每页 bullets 条数 / 单条长度（过长易溢出）
  - 图表 / 表格数据是否为空
  - 图片文件是否存在

修复分两级：
  1. ``repair_deterministic`` —— 无需 LLM 的安全修复（拆分/截断/删空页/换页型）
  2. ``repair_with_llm`` —— 把校验问题回灌 LLM，仅对问题页重新规划（见 llm_repair）

原则：无 OPENAI_API_KEY 时仍可运行（确定性修复兜底），绝不破坏现有启发式流程。
"""
from __future__ import annotations

import copy
import os
from dataclasses import dataclass

import config
from .analyzer import _split_long
from .models import DeckPlan, SlideData

# 与 analyzer.BULLETS_PER_SLIDE 保持一致
MAX_BULLETS_PER_SLIDE = 6
# 超过此长度的要点会被确定性拆分（renderer 会换行，但过长仍易溢出）
MAX_BULLET_LEN = 80
MAX_SLIDES = 40

CONTENT_KINDS = {"agenda", "bullets", "cards", "timeline", "table", "chart",
                 "gallery", "image", "figure", "summary"}
ALL_KINDS = CONTENT_KINDS | {"title", "section", "thanks"}


@dataclass
class PlanIssue:
    """单条规划问题。slide=0 表示整份计划级问题，否则为 1-based 页号。"""
    slide: int
    kind: str          # bad_theme/too_many_slides/unknown_kind/empty_slide/too_many_bullets/long_bullet/bad_chart/bad_table/missing_image
    message: str
    severity: str = "error"   # error / warning


def _is_empty(sd: SlideData) -> bool:
    """判断内容页是否「空壳」（标题页/章节页/结尾页除外）。"""
    if sd.kind in ("title", "section", "thanks"):
        return False
    if sd.kind in ("bullets", "cards", "timeline", "agenda", "summary"):
        return not any(x.strip() for x in sd.bullets)
    if sd.kind == "table":
        return sd.table is None or (not sd.table.headers and not sd.table.rows)
    if sd.kind == "chart":
        return sd.chart is None or (not sd.chart.categories or not sd.chart.series)
    if sd.kind in ("gallery", "image"):
        return not sd.images
    if sd.kind == "figure":
        return not sd.images and not any(x.strip() for x in sd.bullets)
    return False


def validate_plan(plan: DeckPlan) -> list[PlanIssue]:
    """校验 DeckPlan，返回所有问题（可多页重复）。"""
    issues: list[PlanIssue] = []

    if plan.theme_name not in config.THEMES:
        issues.append(PlanIssue(0, "bad_theme",
                                f"主题「{plan.theme_name}」不在预设内，将使用默认主题",
                                "error"))
    if len(plan.slides) > MAX_SLIDES:
        issues.append(PlanIssue(0, "too_many_slides",
                                f"共 {len(plan.slides)} 页，超过上限 {MAX_SLIDES}",
                                "warning"))

    for i, sd in enumerate(plan.slides, 1):
        if sd.kind not in ALL_KINDS:
            issues.append(PlanIssue(i, "unknown_kind",
                                    f"页 {i}「{sd.title}」未知页型「{sd.kind}」",
                                    "error"))
        if _is_empty(sd):
            issues.append(PlanIssue(i, "empty_slide",
                                    f"页 {i}「{sd.title}」无内容，将被删除",
                                    "warning"))

        if sd.bullets:
            if sd.kind in ("bullets", "cards") \
                    and len(sd.bullets) > MAX_BULLETS_PER_SLIDE:
                issues.append(PlanIssue(
                    i, "too_many_bullets",
                    f"页 {i}「{sd.title}」要点 {len(sd.bullets)} 条 > "
                    f"{MAX_BULLETS_PER_SLIDE}，将被拆分",
                    "warning"))
            if any(len(b) > MAX_BULLET_LEN for b in sd.bullets):
                issues.append(PlanIssue(
                    i, "long_bullet",
                    f"页 {i}「{sd.title}」存在超长要点(>{MAX_BULLET_LEN}字)，将被拆分",
                    "warning"))

        if sd.chart and (not sd.chart.categories or not sd.chart.series):
            issues.append(PlanIssue(i, "bad_chart",
                                    f"页 {i}「{sd.title}」图表数据为空", "error"))
        if sd.table and not sd.table.headers and not sd.table.rows:
            issues.append(PlanIssue(i, "bad_table",
                                    f"页 {i}「{sd.title}」表格为空", "error"))
        for p in sd.images:
            if not os.path.isfile(p):
                issues.append(PlanIssue(
                    i, "missing_image",
                    f"页 {i} 图片不存在：{os.path.basename(p)}", "warning"))

    return issues


def repair_deterministic(plan: DeckPlan) -> tuple[DeckPlan, int]:
    """无需 LLM 的安全修复；返回 (新 plan, 修复项数)。输入 plan 不会被修改。"""
    plan = copy.deepcopy(plan)
    fixed = 0

    if plan.theme_name not in config.THEMES:
        plan.theme_name = config.DEFAULT_THEME
        fixed += 1

    out: list[SlideData] = []
    for sd in plan.slides:
        if sd.kind not in ALL_KINDS:
            sd.kind = "bullets"
            fixed += 1

        # 空内容页删除（结构页保留）
        if _is_empty(sd) and sd.kind not in ("title", "section", "thanks"):
            fixed += 1
            continue

        # 图片不存在 -> 移除；画廊因此变空 -> 删除该页
        kept = [p for p in sd.images if os.path.isfile(p)]
        if len(kept) != len(sd.images):
            sd.images = kept
            fixed += 1
        if sd.kind in ("gallery", "image") and not sd.images:
            fixed += 1
            continue

        # 图表空 -> 退化成本地表格；无表可退 -> 删除
        if sd.kind == "chart" and (sd.chart is None
                                   or not sd.chart.categories
                                   or not sd.chart.series):
            if sd.table is not None and sd.table.headers:
                sd.kind = "table"
                fixed += 1
            else:
                fixed += 1
                continue

        # 表格空 -> 删除
        if sd.kind == "table" and (sd.table is None
                                   or (not sd.table.headers and not sd.table.rows)):
            fixed += 1
            continue

        # 长要点拆分（文本类页型）
        if sd.kind in ("bullets", "cards", "timeline", "agenda",
                       "summary", "figure"):
            pieces: list[str] = []
            bullet_changed = False
            for b in sd.bullets:
                sub = [_p if len(_p) <= MAX_BULLET_LEN
                       else _p[:MAX_BULLET_LEN - 1].rstrip() + "…"
                       for _p in _split_long(b, max_len=MAX_BULLET_LEN)]
                if len(sub) != 1 or sub[0] != b:
                    bullet_changed = True
                pieces.extend(sub)
            if bullet_changed:
                fixed += 1
            sd.bullets = pieces

        # 条数上限：bullets/cards 拆页；summary 截断
        if sd.kind in ("bullets", "cards") \
                and len(sd.bullets) > MAX_BULLETS_PER_SLIDE:
            chunks = [sd.bullets[i:i + MAX_BULLETS_PER_SLIDE]
                      for i in range(0, len(sd.bullets), MAX_BULLETS_PER_SLIDE)]
            for k, chunk in enumerate(chunks):
                ns = copy.deepcopy(sd)
                ns.bullets = chunk
                if k > 0:
                    ns.title = (sd.title or "要点") + f"（{k + 1}）"
                out.append(ns)
            fixed += 1
            continue
        elif sd.kind == "summary" and len(sd.bullets) > MAX_BULLETS_PER_SLIDE:
            sd.bullets = sd.bullets[:MAX_BULLETS_PER_SLIDE]
            fixed += 1
        # timeline/agenda 渲染器内部已有截断（≤6 / ≤8），无需处理

        out.append(sd)

    # 兜底：至少保留 2 页（封面 + 结尾），避免整份被删空
    if len(out) < 2:
        if not out:
            out = [SlideData(title=plan.title or "演示文稿", kind="title")]
        if not any(s.kind == "thanks" for s in out):
            out.append(SlideData(title="谢谢 / Thank You", kind="thanks"))
        fixed += 1

    plan.slides = out
    return plan, fixed
