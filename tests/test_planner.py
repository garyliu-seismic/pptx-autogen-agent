"""planner：execute_tool 校验、plan_with_llm 多轮循环（mock chat_tools）。"""
import json

import src.planner as P
from src.constraints import Constraints
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


def test_execute_tool_stores_source_ref():
    plan = DeckPlan()
    P.execute_tool(plan, "add_table", {
        "title": "T", "headers": ["h"], "rows": [["1"]], "source_ref": "data.csv"})
    P.execute_tool(plan, "add_chart", {
        "title": "C", "chart_type": "pie", "categories": ["A"],
        "series": [{"name": "S", "values": [1]}], "source_ref": "data.csv"})
    assert plan.slides[0].table.source_ref == "data.csv"
    assert plan.slides[1].chart.source_ref == "data.csv"


def test_execute_tool_rejects_set_theme_when_locked():
    plan = DeckPlan()
    result = P.execute_tool(plan, "set_theme", {"theme": "tech"}, locked_theme="business")
    assert result.startswith("ERROR")
    assert plan.theme_name != "tech"


def test_plan_with_llm_excludes_set_theme_tool_when_locked(monkeypatch):
    seen_tools: list[list[dict]] = []

    def fake_chat_tools(messages, tools, **kw):
        seen_tools.append(tools)
        return _mk([{"tool": "add_bullets", "args": {"title": "页", "bullets": ["1"]}},
                    {"tool": "done", "args": {}}])

    monkeypatch.setattr(P, "chat_tools", fake_chat_tools)
    plan = P.plan_with_llm([], constraints=Constraints(locked_theme="business"))
    assert plan is not None
    assert plan.theme_name == "business"
    assert all(t["function"]["name"] != "set_theme" for t in seen_tools[0])


def test_finalize_guarantees_cover_and_thanks():
    plan = DeckPlan()
    plan.slides.append(P.SlideData(kind="bullets", title="页", bullets=["a"]))
    assert P.finalize(plan)
    assert plan.slides[0].kind == "title"
    assert plan.slides[-1].kind == "thanks"


# ---- web_search：百度百科优先、维基备选、全失败降级 ----

class _FakeResp:
    def __init__(self, payload):
        self._payload = payload

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def read(self):
        return json.dumps(self._payload).encode("utf-8")


def test_baike_search_parses_abstract(monkeypatch):
    import urllib.request
    payload = {"title": "人工智能", "desc": "短描述",
               "abstract": "人工智能（AI）是<a href=x>模拟</a>智能的学科。"}
    monkeypatch.setattr(urllib.request, "urlopen",
                        lambda *a, **k: _FakeResp(payload))
    r = P._baike_search("人工智能")
    assert r and r.startswith("[来源: 百度百科] 人工智能")
    assert "模拟" in r and "<a" not in r          # 已去 HTML 标签


def test_baike_search_uses_desc_when_no_abstract(monkeypatch):
    import urllib.request
    monkeypatch.setattr(urllib.request, "urlopen",
                        lambda *a, **k: _FakeResp({"title": "t", "desc": "仅有描述"}))
    r = P._baike_search("t")
    assert r and "仅有描述" in r


def test_baike_search_returns_none_on_error(monkeypatch):
    import urllib.request
    monkeypatch.setattr(urllib.request, "urlopen", lambda *a, **k: 1 / 0)
    assert P._baike_search("x") is None


def test_web_search_prefers_baike(monkeypatch):
    monkeypatch.setattr(P, "_baike_search", lambda q, **kw: "[来源: 百度百科] 内容")
    monkeypatch.setattr(P, "_wikipedia_zh_search", lambda q, **kw: "不应被调用")
    assert P._web_search("q").startswith("[来源: 百度百科]")


def test_web_search_falls_back_to_wikipedia(monkeypatch):
    monkeypatch.setattr(P, "_baike_search", lambda q, **kw: None)
    monkeypatch.setattr(P, "_wikipedia_zh_search",
                        lambda q, **kw: "[来源: Wikipedia 中文] 内容")
    assert P._web_search("q").startswith("[来源: Wikipedia 中文]")


def test_web_search_all_fail_degrades(monkeypatch):
    monkeypatch.setattr(P, "_baike_search", lambda q, **kw: None)
    monkeypatch.setattr(P, "_wikipedia_zh_search", lambda q, **kw: None)
    r = P._web_search("q")
    assert "无结果或网络不通" in r


def test_system_prompt_has_grounding_rule():
    sp = P.build_system_prompt()
    assert "GROUNDING RULE" in sp
    assert "严禁编造数字" in sp and "source_ref" in sp
    assert "Markdown/HTML" in sp and "3~7" in sp


def test_tool_schemas_reject_additional_properties():
    def walk(node):
        if isinstance(node, dict):
            if node.get("type") == "object" and "properties" in node:
                assert node.get("additionalProperties") is False, node
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for v in node:
                walk(v)
    for t in P.PLAN_TOOLS:
        walk(t["function"]["parameters"])
    names = {t["function"]["name"] for t in P.PLAN_TOOLS}
    assert {"add_table", "add_chart", "add_kpi"} <= names
