"""planner：execute_tool 校验、plan_with_llm 多轮循环（mock chat_tools）。"""
import json

import src.planner as P
from src.models import DeckPlan


def _mk(tool_calls):
    return {"content": "", "tool_calls": [
        {"id": f"call_{i}", "type": "function",
         "function": {"name": c["tool"],
                      "arguments": json.dumps(c["args"], ensure_ascii=False)}}
        for i, c in enumerate(tool_calls)]}


def test_execute_tool_rejects_invalid():
    plan = DeckPlan()
    assert P.execute_tool(plan, "add_bullets",
                          {"title": "x", "bullets": ["1"] * 7}).startswith("ERROR")
    assert P.execute_tool(plan, "set_theme", {"theme": "nope"}).startswith("ERROR")
    assert P.execute_tool(plan, "add_chart",
                          {"title": "c", "chart_type": "bad",
                           "categories": [], "series": []}).startswith("ERROR")
    assert len(plan.slides) == 0


def test_execute_tool_builds_slides():
    plan = DeckPlan()
    assert P.execute_tool(plan, "set_title", {"title": "T"}).startswith("OK")
    assert P.execute_tool(plan, "set_theme", {"theme": "business"}).startswith("OK")
    assert P.execute_tool(plan, "add_bullets",
                          {"title": "B", "bullets": ["a", "b"]}).startswith("OK")
    assert P.execute_tool(plan, "add_chart", {
        "title": "C", "chart_type": "pie", "categories": ["A"],
        "series": [{"name": "S", "values": [1]}]}).startswith("OK")
    assert [s.kind for s in plan.slides] == ["bullets", "chart"]


def test_plan_with_llm_loop_web_search_then_build(monkeypatch):
    calls = iter([
        [{"tool": "web_search", "args": {"query": "背景"}}],
        [{"tool": "add_bullets", "args": {"title": "页", "bullets": ["1", "2"]}},
         {"tool": "done", "args": {}}],
    ])
    monkeypatch.setattr(P, "chat_tools", lambda m, t, **kw: _mk(next(calls)))
    monkeypatch.setattr(P, "_web_search", lambda q, **kw: "[来源] 背景内容")
    plan = P.plan_with_llm([])
    assert plan is not None
    assert plan.slides[0].kind == "title"
    assert plan.slides[-1].kind == "thanks"
    assert any(s.kind == "bullets" and s.bullets == ["1", "2"] for s in plan.slides)


def test_plan_with_llm_stops_on_stall(monkeypatch):
    calls = iter([
        [{"tool": "set_theme", "args": {"theme": "tech"}}],
        [{"tool": "add_bullets", "args": {"title": "页", "bullets": ["1"]}},
         {"tool": "done", "args": {}}],
    ])
    monkeypatch.setattr(P, "chat_tools", lambda m, t, **kw: _mk(next(calls)))
    plan = P.plan_with_llm([])
    assert plan is not None and len(plan.slides) >= 2


def test_finalize_guarantees_cover_and_thanks():
    plan = DeckPlan()
    plan.slides.append(P.SlideData(kind="bullets", title="页", bullets=["a"]))
    assert P.finalize(plan)
    assert plan.slides[0].kind == "title"
    assert plan.slides[-1].kind == "thanks"
