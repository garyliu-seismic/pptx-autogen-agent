"""planio：DeckPlan <-> JSON 往返序列化。"""
from src.models import ChartData, DeckPlan, SlideData, TableData
from src.planio import load_plan, plan_to_dict, save_plan


def _plan():
    return DeckPlan(title="标题", subtitle="副", author="作者", theme_name="business", slides=[
        SlideData(kind="title", title="封面"),
        SlideData(kind="bullets", title="要点", bullets=["x", "y"], takeaway="z", notes="备注"),
        SlideData(kind="table", title="表", table=TableData(headers=["h1", "h2"], rows=[["1", "2"]])),
        SlideData(kind="chart", title="图",
                  chart=ChartData(chart_type="column", categories=["A", "B"],
                                  series=[{"name": "S", "values": [1, 2]}])),
        SlideData(kind="gallery", title="相册", images=["C:\\imgs\\a.jpg"]),
        SlideData(kind="thanks"),
    ])


def test_roundtrip(tmp_path):
    p = str(tmp_path / "plan.json")
    save_plan(_plan(), p)
    p2 = load_plan(p)
    plan = _plan()
    assert p2.title == plan.title
    assert p2.subtitle == plan.subtitle
    assert p2.author == plan.author
    assert p2.theme_name == plan.theme_name
    assert [s.kind for s in p2.slides] == [s.kind for s in plan.slides]
    assert p2.slides[1].bullets == ["x", "y"]
    assert p2.slides[1].takeaway == "z"
    assert p2.slides[2].table.rows == [["1", "2"]]
    assert p2.slides[3].chart.categories == ["A", "B"]
    assert p2.slides[3].chart.series[0]["values"] == [1, 2]
    assert p2.slides[4].images == ["C:\\imgs\\a.jpg"]


def test_dict_json_roundtrip_serializable():
    d = plan_to_dict(_plan())
    import json
    s = json.dumps(d, ensure_ascii=False)
    assert '"theme": "business"' in s
    assert '"kind": "chart"' in s


def test_slide_from_dict_cards_items_alias():
    from src.planio import plan_from_dict
    plan = plan_from_dict({"theme": "tech", "slides": [
        {"kind": "cards", "title": "能力", "items": ["a", "b", "c"]}]})
    assert plan.slides[0].kind == "cards"
    assert plan.slides[0].bullets == ["a", "b", "c"]


def test_source_field_roundtrip(tmp_path):
    p = str(tmp_path / "src.json")
    save_plan(_plan(), p)
    p2 = load_plan(p)
    # 原计划无 source -> 补一个验证序列化
    plan = _plan()
    plan.slides[1].source = "某来源"
    d = plan_to_dict(plan)
    assert d["slides"][1]["source"] == "某来源"
