"""LLM 修复 —— 把「规划校验问题 / 渲染质检问题」回灌 LLM，仅对问题页重新规划。

对应 ppt master 的 repair loop（错误注入 conversation -> LLM 自动修正）。

两种入口：
  - ``repair_plan_with_llm(plan, plan_issues)``   Phase B：规划结构问题
  - ``repair_render_with_llm(plan, render_issues)``  Phase C：渲染几何问题（quality.Issue）

两者共用：
  - ``serialize_slide`` / ``slide_from_dict``：SlideData 与 JSON 互转
  - ``apply_replacements``：按 1-based 页号替换幻灯片（支持一页拆多页）

失败（无 Key / 超时 / 解析失败）一律返回 None，调用方降级到确定性修复。
"""
from __future__ import annotations

import copy
import json
import os

from .llm_advisor import chat_json
from .models import DeckPlan, SlideData

_ALLOWED_KINDS = "title/section/agenda/bullets/cards/timeline/figure/summary/thanks"

_SYSTEM = (
    "你是演示文稿版面修复助手。用户给出若干有问题的幻灯片及其问题，请仅修复这些页。\n"
    "规则：\n"
    "1. 每页要点(bullets)最多 6 条，每条尽量不超过 80 字；超过就拆分或精简。\n"
    "2. 要点过多时，可以把一页拆成多页（slides 数组放多张，标题可加序号）。\n"
    "3. 文本溢出/重叠时：精简要点、减少条数，或把一页拆成多页；不要在 slides 里放表格/图表/图片数据。\n"
    f"4. kind 只能是 {_ALLOWED_KINDS}。\n"
    '5. 只输出 JSON，不要输出任何其它文字或 markdown：{"replacements":[{"index":页号,"slides":[{"kind":"bullets","title":"...","bullets":["..."]}, ...]}, ...]}\n'
    "6. 只返回需要替换的页；其余页不要出现在输出里。"
)


def serialize_slide(idx: int, sd: SlideData) -> dict:
    """SlideData -> JSON（供 LLM 阅读；不携带大数据）。"""
    d = {"index": idx, "kind": sd.kind, "title": sd.title}
    if sd.bullets:
        d["bullets"] = sd.bullets
    if sd.takeaway:
        d["takeaway"] = sd.takeaway
    if sd.notes:
        d["notes"] = sd.notes
    if sd.table is not None:
        d["table_headers"] = list(sd.table.headers)
    if sd.images:
        d["images"] = [os.path.basename(p) for p in sd.images]
    return d


def slide_from_dict(d: dict) -> SlideData:
    """JSON -> SlideData；只取文本类字段，忽略数据字段。"""
    return SlideData(
        title=str(d.get("title") or ""),
        kind=str(d.get("kind") or "bullets"),
        bullets=[str(b) for b in (d.get("bullets") or [])],
        takeaway=str(d.get("takeaway") or ""),
        notes=str(d.get("notes") or ""),
    )


def apply_replacements(plan: DeckPlan, replacements: list[dict]) -> DeckPlan | None:
    """按 1-based 页号替换幻灯片；每个条目可含多张（用于拆分）。"""
    if not replacements:
        return None
    idx_map: dict[int, list[SlideData]] = {}
    for r in replacements:
        if not isinstance(r, dict):
            continue
        try:
            idx = int(r.get("index"))
        except (TypeError, ValueError):
            continue
        slides = r.get("slides") or r.get("slide")
        if isinstance(slides, dict):
            slides = [slides]
        if not isinstance(slides, list) or not slides:
            continue
        idx_map[idx] = [slide_from_dict(s) if isinstance(s, dict) else s
                        for s in slides if isinstance(s, dict)]
    if not idx_map:
        return None

    new_plan = copy.deepcopy(plan)
    out: list[SlideData] = []
    for i, sd in enumerate(new_plan.slides, 1):
        if i in idx_map:
            out.extend(idx_map[i])
        else:
            out.append(sd)
    new_plan.slides = out
    return new_plan


def _repair(plan: DeckPlan, slides_json: list[dict],
            issue_lines: list[str]) -> DeckPlan | None:
    """通用修复：把问题页 + 问题描述发给 LLM，返回应用替换后的 plan 或 None。"""
    if not slides_json or not issue_lines:
        return None
    user = json.dumps({"slides": slides_json, "issues": issue_lines},
                      ensure_ascii=False)
    data = chat_json(_SYSTEM, user)
    if not data:
        return None
    return apply_replacements(plan, data.get("replacements") or [])


def repair_plan_with_llm(plan: DeckPlan, plan_issues: list) -> DeckPlan | None:
    """Phase B：根据规划校验问题，让 LLM 仅修复问题页。"""
    targets = sorted({it.slide for it in plan_issues if it.slide >= 1})
    slides_json = [serialize_slide(i, plan.slides[i - 1])
                   for i in targets if i <= len(plan.slides)]
    issue_lines = [f"- 页{it.slide}: [{it.kind}] {it.message}"
                   for it in plan_issues]
    return _repair(plan, slides_json, issue_lines)


def repair_render_with_llm(plan: DeckPlan, render_issues: list) -> DeckPlan | None:
    """Phase C：根据渲染质检问题（quality.Issue），让 LLM 仅修复问题页。"""
    targets = sorted({it.slide for it in render_issues if getattr(it, "slide", 0) >= 1})
    slides_json = [serialize_slide(i, plan.slides[i - 1])
                   for i in targets if i <= len(plan.slides)]
    issue_lines = [
        f"- 页{it.slide}: [{it.kind} 值={getattr(it, 'value', 0):.2f}] "
        f"{getattr(it, 'detail', '')}"
        for it in render_issues
    ]
    return _repair(plan, slides_json, issue_lines)
