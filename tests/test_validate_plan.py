"""validate_plan：规划前校验 + 确定性修复。"""
from src.constraints import Constraints
from src.models import ChartData, DeckPlan, SlideData, TableData
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


# ---------------------------------------------------------------------------
# 约束档：数据真实性
# ---------------------------------------------------------------------------
def test_validate_catches_unsourced_chart():
    plan = _plan(
        SlideData(kind="title"),
        SlideData(kind="chart", title="图",
                  chart=ChartData(categories=["a"], series=[{"name": "x", "values": [1]}])),
        SlideData(kind="thanks"),
    )
    c = Constraints(require_data_source=True)
    issues = validate_plan(plan, constraints=c, known_sources={"data.csv"})
    assert any(i.kind == "unsourced_data" for i in issues)


def test_validate_accepts_chart_with_known_source():
    plan = _plan(
        SlideData(kind="title"),
        SlideData(kind="chart", title="图",
                  chart=ChartData(categories=["a"], series=[{"name": "x", "values": [1]}],
                                  source_ref="data.csv")),
        SlideData(kind="thanks"),
    )
    c = Constraints(require_data_source=True)
    issues = validate_plan(plan, constraints=c, known_sources={"data.csv"})
    assert not any(i.kind == "unsourced_data" for i in issues)


def test_repair_deterministic_drops_unsourced_table():
    plan = _plan(
        SlideData(kind="title"),
        SlideData(kind="table", title="表",
                  table=TableData(headers=["h"], rows=[["1"]])),
        SlideData(kind="thanks"),
    )
    c = Constraints(require_data_source=True)
    newp, fixed = repair_deterministic(plan, constraints=c, known_sources={"data.csv"})
    assert not any(s.kind == "table" for s in newp.slides)
    assert fixed >= 1


# ---------------------------------------------------------------------------
# 约束档：品牌/主题锁定
# ---------------------------------------------------------------------------
def test_validate_catches_locked_theme_mismatch():
    plan = _plan(theme="modern")
    c = Constraints(locked_theme="business")
    issues = validate_plan(plan, constraints=c)
    assert any(i.kind == "bad_theme" and i.severity == "error" for i in issues)


def test_repair_deterministic_forces_locked_theme():
    plan = _plan(theme="modern")
    c = Constraints(locked_theme="business")
    newp, _ = repair_deterministic(plan, constraints=c)
    assert newp.theme_name == "business"


# ---------------------------------------------------------------------------
# 约束档：结构
# ---------------------------------------------------------------------------
def test_validate_catches_missing_required_kind():
    plan = _plan(SlideData(kind="title"), SlideData(kind="bullets", bullets=["a"]))
    c = Constraints(required_kinds=["thanks"])
    issues = validate_plan(plan, constraints=c)
    assert any(i.kind == "missing_required_kind" for i in issues)


def test_repair_deterministic_adds_missing_required_kind():
    plan = _plan(SlideData(kind="title"), SlideData(kind="bullets", bullets=["a"]))
    c = Constraints(required_kinds=["thanks"])
    newp, fixed = repair_deterministic(plan, constraints=c)
    assert any(s.kind == "thanks" for s in newp.slides)
    assert fixed >= 1


def test_validate_catches_too_many_of_kind():
    plan = _plan(
        SlideData(kind="title"),
        SlideData(kind="chart", chart=ChartData(categories=["a"],
                                                series=[{"name": "x", "values": [1]}])),
        SlideData(kind="chart", chart=ChartData(categories=["a"],
                                                series=[{"name": "x", "values": [1]}])),
        SlideData(kind="thanks"),
    )
    c = Constraints(max_kind_counts={"chart": 1})
    issues = validate_plan(plan, constraints=c)
    assert any(i.kind == "too_many_of_kind" for i in issues)


def test_repair_deterministic_trims_excess_kind():
    plan = _plan(
        SlideData(kind="title"),
        SlideData(kind="chart", title="c1",
                  chart=ChartData(categories=["a"], series=[{"name": "x", "values": [1]}])),
        SlideData(kind="chart", title="c2",
                  chart=ChartData(categories=["a"], series=[{"name": "x", "values": [1]}])),
        SlideData(kind="thanks"),
    )
    c = Constraints(max_kind_counts={"chart": 1})
    newp, fixed = repair_deterministic(plan, constraints=c)
    assert sum(1 for s in newp.slides if s.kind == "chart") == 1
    assert fixed >= 1


def test_validate_catches_slide_count_out_of_range():
    plan = _plan(SlideData(kind="title"), SlideData(kind="thanks"))
    c = Constraints(min_slides=5)
    issues = validate_plan(plan, constraints=c)
    assert any(i.kind == "slide_count_out_of_range" for i in issues)


def test_validate_catches_slide_count_mismatch_against_reference():
    plan = _plan(SlideData(kind="title"), SlideData(kind="thanks"))
    c = Constraints(exact_slides=5)
    issues = validate_plan(plan, constraints=c)
    assert any(i.kind == "slide_count_mismatch" for i in issues)


def test_repair_deterministic_pads_to_exact_slide_count():
    plan = _plan(SlideData(kind="title"), SlideData(kind="thanks"))
    c = Constraints(exact_slides=5)
    newp, fixed = repair_deterministic(plan, constraints=c)
    assert len(newp.slides) == 5
    # 结尾页仍然是结尾页，补的分隔页插在它前面
    assert newp.slides[-1].kind == "thanks"
    assert fixed >= 1


def test_repair_deterministic_trims_to_exact_slide_count():
    plan = _plan(
        SlideData(kind="title"),
        SlideData(kind="bullets", title="a", bullets=["1"]),
        SlideData(kind="bullets", title="b", bullets=["1"]),
        SlideData(kind="bullets", title="c", bullets=["1"]),
        SlideData(kind="thanks"),
    )
    c = Constraints(exact_slides=2)
    newp, fixed = repair_deterministic(plan, constraints=c)
    assert len(newp.slides) == 2
    # 结构页（title/thanks）优先于内容页保留
    assert {s.kind for s in newp.slides} == {"title", "thanks"}
    assert fixed >= 1


def test_repair_deterministic_noop_when_already_exact():
    plan = _plan(SlideData(kind="title"), SlideData(kind="thanks"))
    c = Constraints(exact_slides=2)
    newp, fixed = repair_deterministic(plan, constraints=c)
    assert len(newp.slides) == 2
