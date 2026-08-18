"""renderer：渲染出有效 pptx、无负几何、相册/时间轴回归。"""
from pptx import Presentation

from src.models import ChartData, DeckPlan, SlideData, TableData
from src.renderer import render


def _base_plan():
    return DeckPlan(title="测试", theme_name="business", slides=[
        SlideData(kind="title", title="封面"),
        SlideData(kind="bullets", title="要点", bullets=["甲", "乙", "丙"]),
        SlideData(kind="table", title="表格",
                  table=TableData(headers=["h"], rows=[["1"]])),
        SlideData(kind="chart", title="图表",
                  chart=ChartData(chart_type="column", categories=["Q1"],
                                  series=[{"name": "A", "values": [1]}])),
        SlideData(kind="summary", title="总结", bullets=["要点1"]),
        SlideData(kind="thanks"),
    ])


def _assert_no_negative_geometry(prs):
    for slide in prs.slides:
        for sh in slide.shapes:
            assert sh.width is None or sh.width >= 0
            assert sh.height is None or sh.height >= 0


def test_render_base_deck(tmp_path):
    out = str(tmp_path / "deck.pptx")
    render(_base_plan(), out)
    prs = Presentation(out)
    assert len(prs.slides) == 6
    _assert_no_negative_geometry(prs)


def test_render_photo_timeline_no_negative(tmp_path, sample_img_with_exif):
    # 回归：时间轴下方卡片的连线高度曾为负 -> 损坏 pptx
    plan = DeckPlan(title="t", theme_name="education", slides=[
        SlideData(kind="title", title="封面"),
        SlideData(kind="photo_timeline", title="时间轴",
                  images=[sample_img_with_exif] * 4),
        SlideData(kind="thanks"),
    ])
    out = str(tmp_path / "tl.pptx")
    render(plan, out)
    prs = Presentation(out)
    assert len(prs.slides) == 3
    _assert_no_negative_geometry(prs)


def test_render_gallery_with_plain_images(tmp_path, sample_img):
    plan = DeckPlan(title="t", theme_name="business", slides=[
        SlideData(kind="title", title="封面"),
        SlideData(kind="gallery", title="相册", images=[sample_img] * 3),
        SlideData(kind="thanks"),
    ])
    out = str(tmp_path / "g.pptx")
    render(plan, out)
    prs = Presentation(out)
    _assert_no_negative_geometry(prs)
