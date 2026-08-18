"""validate_plan：规划前校验 + 确定性修复。"""
from src.models import DeckPlan, SlideData
from src.validate_plan import repair_deterministic, validate_plan


def _plan(*slides, theme="business"):
    return DeckPlan(theme_name=theme, slides=list(slides) or [SlideData(kind="title")])


def test_validate_catches_bad_theme():
    plan = _plan(theme="nope")
    assert any(i.kind == "bad_theme" for i in validate_plan(plan))


def test_validate_catches_too_many_bullets():
    plan = _plan(SlideData(kind="bullets", title="多", bullets=["1"] * 7))
    assert any(i.kind == "too_many_bullets" for i in validate_plan(plan))


def test_repair_deterministic_splits_bullets_and_theme():
    plan = _plan(
        SlideData(kind="title", title="封"),
        SlideData(kind="bullets", title="多", bullets=[str(i) for i in range(7)]),
        SlideData(kind="thanks"),
        theme="nope",
    )
    newp, n = repair_deterministic(plan)
    bullets = [s for s in newp.slides if s.kind == "bullets"]
    assert len(bullets) == 2
    assert all(len(s.bullets) <= 6 for s in bullets)
    assert newp.theme_name == "modern"
    assert n >= 1


def test_repair_deterministic_drops_empty_slide():
    plan = _plan(
        SlideData(kind="title"),
        SlideData(kind="bullets", title="空", bullets=[]),
        SlideData(kind="thanks"),
    )
    newp, _ = repair_deterministic(plan)
    assert not any(s.kind == "bullets" for s in newp.slides)


def test_repair_deterministic_keeps_min_two_slides():
    plan = _plan(SlideData(kind="bullets", title="空", bullets=[]), theme="nope")
    newp, _ = repair_deterministic(plan)
    assert len(newp.slides) >= 2
