"""DeckPlan <-> JSON（P2-E：dry-run 计划确认）。

`--plan-only` 把当前计划写成可读、可编辑的 JSON（outline），供人工预览/改；
`--plan plan.json` 载入确认过的计划再渲染。与 LLM 修复的 JSON 无关，
这里是完整 DeckPlan（含表格/图表/图片路径）的往返序列化。
"""
from __future__ import annotations

import json
import os

from .models import ChartData, DeckPlan, SlideData, TableData

PLAN_VERSION = 1


def plan_to_dict(plan: DeckPlan) -> dict:
    """DeckPlan -> 可读 JSON 字典。"""
    slides: list[dict] = []
    for s in plan.slides:
        d: dict = {"kind": s.kind}
        if s.title:
            d["title"] = s.title
        if s.bullets:
            d["bullets"] = s.bullets
        if s.takeaway:
            d["takeaway"] = s.takeaway
        if s.notes:
            d["notes"] = s.notes
        if s.table is not None:
            d["headers"] = list(s.table.headers)
            d["rows"] = [[str(c) for c in r] for r in s.table.rows]
        if s.chart is not None:
            d["chart_type"] = s.chart.chart_type
            d["categories"] = list(s.chart.categories)
            d["series"] = s.chart.series
            if s.chart.title:
                d["chart_title"] = s.chart.title
        if s.images:
            d["images"] = list(s.images)
        slides.append(d)
    return {
        "version": PLAN_VERSION,
        "title": plan.title,
        "subtitle": plan.subtitle,
        "author": plan.author,
        "theme": plan.theme_name,
        "slides": slides,
    }


def _slide_from_dict(d: dict) -> SlideData:
    sd = SlideData(
        title=str(d.get("title") or ""),
        kind=str(d.get("kind") or "bullets"),
        bullets=[str(x) for x in (d.get("bullets") or [])],
        takeaway=str(d.get("takeaway") or ""),
        notes=str(d.get("notes") or ""),
        images=[str(x) for x in (d.get("images") or [])],
    )
    if "headers" in d or "rows" in d:
        sd.table = TableData(
            headers=[str(h) for h in (d.get("headers") or [])],
            rows=[[str(c) for c in r] for r in (d.get("rows") or [])])
    if "chart_type" in d or "categories" in d or "series" in d:
        sd.chart = ChartData(
            chart_type=str(d.get("chart_type") or "column"),
            categories=[str(c) for c in (d.get("categories") or [])],
            series=d.get("series") or [],
            title=str(d.get("chart_title") or ""),
        )
    return sd


def plan_from_dict(d: dict) -> DeckPlan:
    return DeckPlan(
        title=str(d.get("title") or ""),
        subtitle=str(d.get("subtitle") or ""),
        author=str(d.get("author") or ""),
        theme_name=str(d.get("theme") or ""),
        slides=[_slide_from_dict(s) for s in (d.get("slides") or [])
                if isinstance(s, dict)],
    )


def save_plan(plan: DeckPlan, path: str) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(plan_to_dict(plan), f, ensure_ascii=False, indent=2)


def load_plan(path: str) -> DeckPlan:
    with open(path, "r", encoding="utf-8") as f:
        return plan_from_dict(json.load(f))
