"""readback：读回内容 QA（计划 vs 渲染）+ 空壳页自愈。"""
import src.readback as rb
from src.models import ChartData, DeckPlan, SlideData, TableData
from src.renderer import render


def _plan():
    return DeckPlan(theme_name="business", slides=[
        SlideData(kind="title", title="封面"),
        SlideData(kind="bullets", title="要点", bullets=["a", "b"]),
        SlideData(kind="table", title="表",
                  table=TableData(headers=["h"], rows=[["1"]])),
        SlideData(kind="chart", title="图",
                  chart=ChartData(chart_type="column", categories=["A"],
                                  series=[{"name": "S", "values": [1]}])),
        SlideData(kind="thanks"),
    ])


def test_compare_plan_render_passes(tmp_path):
    out = str(tmp_path / "d.pptx")
    render(_plan(), out)
    assert rb.compare_plan_render(_plan(), out) == []


def test_auto_fix_drops_content_slides_keeps_structure():
    plan = DeckPlan(theme_name="business", slides=[
        SlideData(kind="title"),
        SlideData(kind="bullets", title="a", bullets=["1"]),   # 要剔除
        SlideData(kind="bullets", title="b", bullets=["2"]),   # 要剔除
        SlideData(kind="bullets", title="c", bullets=["3"]),   # 保留（还有内容）
        SlideData(kind="thanks"),
    ])
    issues = [rb.QAIssue(2, "empty_render", "x"), rb.QAIssue(3, "missing_object", "y")]
    newp, n = rb.auto_fix(plan, issues)
    assert n == 2
    assert [s.kind for s in newp.slides] == ["title", "bullets", "thanks"]


def test_auto_fix_keeps_min_two_slides():
    plan = DeckPlan(theme_name="business", slides=[
        SlideData(kind="title"),
        SlideData(kind="bullets", title="a", bullets=["1"]),
        SlideData(kind="thanks"),
    ])
    # 剔除唯一内容页会让整份只剩封面+致谢（无内容）-> 不冒险，保留原计划
    newp, n = rb.auto_fix(plan, [rb.QAIssue(2, "empty_render", "x")])
    assert n == 0
    assert len(newp.slides) == 3
