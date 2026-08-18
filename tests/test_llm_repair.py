"""llm_repair：替换/拆分、SlideData 序列化、LLM 修复（mock chat_json）。"""
import src.llm_repair as lr
from src.models import DeckPlan, SlideData
from src.validate_plan import validate_plan


def _plan():
    return DeckPlan(theme_name="business", slides=[
        SlideData(kind="title", title="封面"),
        SlideData(kind="bullets", title="问题页", bullets=["1"] * 8),
        SlideData(kind="thanks"),
    ])


def test_apply_replacements_split_one_into_two():
    plan = DeckPlan(slides=[
        SlideData(kind="title"),
        SlideData(kind="bullets", title="x", bullets=["1", "2"]),
        SlideData(kind="thanks"),
    ])
    newp = lr.apply_replacements(plan, [{"index": 2, "slides": [
        {"kind": "bullets", "title": "x(1)", "bullets": ["1"]},
        {"kind": "bullets", "title": "x(2)", "bullets": ["2"]},
    ]}])
    assert len(newp.slides) == 4
    assert newp.slides[1].title == "x(1)"
    assert newp.slides[2].title == "x(2)"
    assert newp.slides[3].kind == "thanks"


def test_apply_replacements_ignores_bad_index():
    plan = DeckPlan(slides=[SlideData(kind="title"), SlideData(kind="thanks")])
    newp = lr.apply_replacements(plan, [{"index": "not-a-number", "slides": [{"kind": "bullets"}]}])
    assert newp is None


def test_serialize_slide_roundtrip():
    d = lr.serialize_slide(2, SlideData(kind="bullets", title="t",
                                        bullets=["a", "b"], takeaway="c"))
    assert d["index"] == 2
    sd = lr.slide_from_dict(d)
    assert sd.kind == "bullets"
    assert sd.bullets == ["a", "b"]
    assert sd.takeaway == "c"


def test_repair_plan_with_llm_mocked(monkeypatch):
    plan = _plan()
    issues = validate_plan(plan)
    assert any(i.kind == "too_many_bullets" for i in issues)

    def fake_chat_json(system, user, **kw):
        return {"replacements": [{"index": 2, "slides": [
            {"kind": "bullets", "title": "问题页(1)", "bullets": ["1"] * 6},
            {"kind": "bullets", "title": "问题页(2)", "bullets": ["1"] * 2},
        ]}]}

    monkeypatch.setattr(lr, "chat_json", fake_chat_json)
    newp = lr.repair_plan_with_llm(plan, issues)
    assert newp is not None
    assert all(len(s.bullets) <= 6 for s in newp.slides)
    assert not validate_plan(newp)
