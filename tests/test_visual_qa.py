"""visual_qa：判定 schema/解析、checker 注入、render_with_quality 视觉回灌循环（全 mock，无网络）。"""
import copy

import src.quality as q
import src.renderer as rd
import src.visual_qa as vq
from src.models import DeckPlan, SlideData


def _plan():
    return DeckPlan(theme_name="business", slides=[
        SlideData(kind="title", title="T"),
        SlideData(kind="bullets", title="B", bullets=["a", "b"]),
        SlideData(kind="thanks")])


def test_verdict_schema_is_strict():
    s = vq.VERDICT_SCHEMA
    assert s["additionalProperties"] is False
    assert set(s["required"]) == set(s["properties"])


def test_parse_verdict():
    assert vq.parse_verdict({"passed": True, "issues": [], "suggestions": ""})["passed"]
    v = vq.parse_verdict({"passed": False, "issues": ["overflow"], "suggestions": "shorter"})
    assert v == {"passed": False, "issues": ["overflow"], "suggestions": "shorter"}
    assert vq.parse_verdict({"passed": "no"}) is None
    assert vq.parse_verdict(None) is None


def test_inspect_slide_uses_vision_call(monkeypatch):
    seen = {}

    def fake(system, user, image, **kw):
        seen.update(system=system, schema=kw.get("schema"))
        return {"passed": False, "issues": ["x"], "suggestions": "y"}

    monkeypatch.setattr(vq.llm_advisor, "chat_vision_json", fake)
    assert vq.inspect_slide("a.png", "T")["passed"] is False
    assert seen["schema"] is vq.VERDICT_SCHEMA
    # 无 Key/失败 -> None（视为通过）
    monkeypatch.setattr(vq.llm_advisor, "chat_vision_json", lambda *a, **k: None)
    assert vq.inspect_slide("a.png") is None


def test_checker_reports_failed_slides_only():
    verdicts = {"p1": {"passed": True, "issues": [], "suggestions": ""},
                "p2": {"passed": False, "issues": ["overflow"], "suggestions": "shorten"}}
    check = vq.make_visual_checker(thumbnailer=lambda path: ["p1", "p2"],
                                   inspector=lambda img, title: verdicts[img])
    issues = check("x.pptx", _plan())
    assert [(i.slide, i.kind) for i in issues] == [(2, "visual")]
    assert "overflow" in issues[0].detail and "shorten" in issues[0].detail


def _patch(monkeypatch):
    state = {"renders": 0}

    def fake_render(p, path, **kw):
        state["renders"] += 1
        return path

    monkeypatch.setattr(rd, "render", fake_render)
    monkeypatch.setattr(q, "analyze", lambda path: [])
    return state


def test_visual_failure_feeds_repairer_and_is_bounded(monkeypatch, tmp_path):
    state = _patch(monkeypatch)
    got = []

    def repairer(p, issues):
        got.append(issues)
        return copy.deepcopy(p)

    always_bad = lambda path, plan: [q.Issue(2, "visual", 1.0, "overflow | 建议: shorten")]
    out, issues, tries = q.render_with_quality(
        _plan(), str(tmp_path / "x.pptx"), quiet=True, repairer=repairer,
        visual_checker=always_bad, max_visual_rounds=2)
    assert len(got) == 2                 # 有界：最多 2 轮
    assert got[0][0].kind == "visual" and "shorten" in got[0][0].detail
    assert state["renders"] == 3         # 初次 + 2 次重排
    assert any(i.kind == "visual" for i in issues)   # 用尽后只提示


def test_visual_pass_no_repair(monkeypatch, tmp_path):
    state = _patch(monkeypatch)
    out, issues, tries = q.render_with_quality(
        _plan(), str(tmp_path / "x.pptx"), quiet=True,
        repairer=lambda p, i: 1 / 0, visual_checker=lambda path, plan: [])
    assert state["renders"] == 1 and issues == []


def test_visual_checker_default_off(monkeypatch, tmp_path):
    state = _patch(monkeypatch)
    q.render_with_quality(_plan(), str(tmp_path / "x.pptx"), quiet=True)
    assert state["renders"] == 1
