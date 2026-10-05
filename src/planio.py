"""DeckPlan <-> JSON（P2-E：dry-run 计划确认）。

`--plan-only` 把当前计划写成可读、可编辑的 JSON（outline），供人工预览/改；
`--plan plan.json` 载入确认过的计划再渲染。与 LLM 修复的 JSON 无关，
这里是完整 DeckPlan（含表格/图表/图片路径）的往返序列化。
"""
from __future__ import annotations

import json
import os

from .models import ChartData, DeckPlan, MetricData, SlideData, TableData

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
        if s.bullets_right:
            d["bullets_right"] = s.bullets_right
        if s.takeaway:
            d["takeaway"] = s.takeaway
        if s.author:
            d["author"] = s.author
        if s.label_left:
            d["label_left"] = s.label_left
        if s.label_right:
            d["label_right"] = s.label_right
        if s.notes:
            d["notes"] = s.notes
        if s.source:
            d["source"] = s.source
        if s.metrics:
            d["metrics"] = [
                {"label": m.label, "value": m.value,
                 "trend": m.trend, "note": m.note, "color": m.color}
                for m in s.metrics
            ]
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
    # 兼容 cards 用 "items" 键写列表（与 bullets 等价）
    bullets_raw = d.get("bullets")
    if bullets_raw is None:
        bullets_raw = d.get("items")
    sd = SlideData(
        title=str(d.get("title") or ""),
        kind=str(d.get("kind") or "bullets"),
        bullets=[str(x) for x in (bullets_raw or [])],
        bullets_right=[str(x) for x in (d.get("bullets_right") or [])],
        takeaway=str(d.get("takeaway") or ""),
        author=str(d.get("author") or ""),
        label_left=str(d.get("label_left") or ""),
        label_right=str(d.get("label_right") or ""),
        notes=str(d.get("notes") or ""),
        source=str(d.get("source") or ""),
        images=[str(x) for x in (d.get("images") or [])],
    )
    # KPI 指标
    if d.get("metrics"):
        sd.metrics = [
            MetricData(
                label=str(m.get("label") or ""),
                value=str(m.get("value") or ""),
                trend=str(m.get("trend") or "flat"),
                note=str(m.get("note") or ""),
                color=str(m.get("color") or ""),
            )
            for m in d["metrics"] if isinstance(m, dict)
        ]
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
