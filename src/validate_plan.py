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
from .constraints import Constraints
from .models import DeckPlan, SlideData

# 与 analyzer.BULLETS_PER_SLIDE 保持一致
MAX_BULLETS_PER_SLIDE = 6
# 超过此长度的要点会被确定性拆分（renderer 会换行，但过长仍易溢出）
MAX_BULLET_LEN = 80
MAX_SLIDES = 40

CONTENT_KINDS = {"agenda", "bullets", "cards", "timeline", "table", "chart",
                 "gallery", "image", "figure", "summary", "photo_timeline"}
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
    if sd.kind in ("gallery", "image", "photo_timeline"):
        return not sd.images
    if sd.kind == "figure":
        return not sd.images and not any(x.strip() for x in sd.bullets)
    return False


def validate_plan(plan: DeckPlan, constraints: Constraints | None = None,
                  known_sources: set[str] | None = None) -> list[PlanIssue]:
    """校验 DeckPlan，返回所有问题（可多页重复）。

    constraints 非空时额外做约束档校验（主题锁定/数据来源/结构要求），
    这些问题一律标 error，交由 repair_deterministic 强制处理，不只是警告。
    known_sources 是当前素材的 block.source 集合，用于核对 source_ref 是否真实存在
    （防止 LLM 编一个看起来像来源的字符串）。
    """
    issues: list[PlanIssue] = []
    locked_theme = constraints.locked_theme if constraints else None
    max_slides = (constraints.max_slides if constraints and constraints.max_slides
                  else MAX_SLIDES)

    if plan.theme_name not in config.THEMES:
        issues.append(PlanIssue(0, "bad_theme",
                                f"主题「{plan.theme_name}」不在预设内，将使用默认主题",
                                "error"))
    elif locked_theme and plan.theme_name != locked_theme:
        issues.append(PlanIssue(0, "bad_theme",
                                f"主题「{plan.theme_name}」与锁定主题「{locked_theme}」不符",
                                "error"))
    if len(plan.slides) > max_slides:
        issues.append(PlanIssue(0, "too_many_slides",
                                f"共 {len(plan.slides)} 页，超过上限 {max_slides}",
                                "warning"))
    if constraints:
        if constraints.min_slides and len(plan.slides) < constraints.min_slides:
            issues.append(PlanIssue(
                0, "slide_count_out_of_range",
                f"共 {len(plan.slides)} 页，少于要求的下限 {constraints.min_slides}",
                "error"))
        if constraints.max_slides and len(plan.slides) > constraints.max_slides:
            issues.append(PlanIssue(
                0, "slide_count_out_of_range",
                f"共 {len(plan.slides)} 页，超过要求的上限 {constraints.max_slides}",
                "error"))
        for kind in constraints.required_kinds:
            if not any(sd.kind == kind for sd in plan.slides):
                issues.append(PlanIssue(
                    0, "missing_required_kind", f"缺少必需页型「{kind}」", "error"))
        for kind, cap in constraints.max_kind_counts.items():
            count = sum(1 for sd in plan.slides if sd.kind == kind)
            if count > cap:
                issues.append(PlanIssue(
                    0, "too_many_of_kind",
                    f"页型「{kind}」共 {count} 页，超过上限 {cap}", "error"))

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
        if constraints and constraints.require_data_source:
            for data, kind_name in ((sd.chart, "图表"), (sd.table, "表格")):
                if data is None:
                    continue
                ref = data.source_ref.strip() if data.source_ref else ""
                if not ref or (known_sources is not None and ref not in known_sources):
                    issues.append(PlanIssue(
                        i, "unsourced_data",
                        f"页 {i}「{sd.title}」{kind_name}未标注真实来源，将被剔除",
                        "error"))
        for p in sd.images:
            if not os.path.isfile(p):
                issues.append(PlanIssue(
                    i, "missing_image",
                    f"页 {i} 图片不存在：{os.path.basename(p)}", "warning"))

    return issues


def repair_deterministic(plan: DeckPlan,
                         constraints: Constraints | None = None,
                         known_sources: set[str] | None = None) -> tuple[DeckPlan, int]:
    """无需 LLM 的安全修复；返回 (新 plan, 修复项数)。输入 plan 不会被修改。"""
    plan = copy.deepcopy(plan)
    fixed = 0
    locked_theme = constraints.locked_theme if constraints else None

    if plan.theme_name not in config.THEMES or (locked_theme and plan.theme_name != locked_theme):
        plan.theme_name = locked_theme or config.DEFAULT_THEME
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

        # 数据真实性约束：图表/表格没有可核实来源 -> 剔除该页
        if constraints and constraints.require_data_source and sd.kind in ("chart", "table"):
            data = sd.chart if sd.kind == "chart" else sd.table
            ref = data.source_ref.strip() if data is not None and data.source_ref else ""
            if not ref or (known_sources is not None and ref not in known_sources):
                fixed += 1
                continue

        # 图片不存在 -> 移除；画廊因此变空 -> 删除该页
        kept = [p for p in sd.images if os.path.isfile(p)]
        if len(kept) != len(sd.images):
            sd.images = kept
            fixed += 1
        if sd.kind in ("gallery", "image", "photo_timeline") and not sd.images:
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

    if constraints:
        # 结构约束：每类页型数量上限 -> 保留前 N 页，超出的丢弃
        for kind, cap in constraints.max_kind_counts.items():
            kept_count = 0
            trimmed: list[SlideData] = []
            for sd in out:
                if sd.kind == kind:
                    kept_count += 1
                    if kept_count > cap:
                        fixed += 1
                        continue
                trimmed.append(sd)
            out = trimmed

        # 结构约束：必需页型缺失 -> 补一个最小化默认页
        for kind in constraints.required_kinds:
            if not any(sd.kind == kind for sd in out):
                out.append(_default_slide_for_kind(kind, plan))
                fixed += 1

    plan.slides = out
    return plan, fixed


def _default_slide_for_kind(kind: str, plan: DeckPlan) -> SlideData:
    """为缺失的必需页型生成一个最小化的默认页。"""
    if kind == "title":
        return SlideData(title=plan.title or "演示文稿", kind="title")
    if kind == "thanks":
        return SlideData(title="谢谢 / Thank You", kind="thanks")
    if kind == "section":
        return SlideData(title="章节", kind="section")
    if kind in ("bullets", "cards", "summary"):
        return SlideData(title="要点", kind=kind, bullets=["（待补充）"])
    return SlideData(title=kind, kind=kind)
