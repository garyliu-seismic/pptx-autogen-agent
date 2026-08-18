"""analyzer：build_deck / _plan_content 结构、画廊分页、时间轴识别、相册去目录。"""
from src.analyzer import MAX_GALLERY_IMAGES, _plan_content, build_deck
from src.models import Block


def _img_blocks(paths):
    return [Block(kind="image", image_path=p, source=p) for p in paths]


def test_pure_image_album_no_agenda(sample_img):
    blocks = _img_blocks([sample_img] * 7)
    plan = build_deck(blocks, title="相册")
    kinds = [s.kind for s in plan.slides]
    assert "agenda" not in kinds                      # 纯图片不生成目录
    imgs = [s for s in plan.slides if s.kind in ("gallery", "photo_timeline")]
    assert len(imgs) == 2                             # 7 张 -> 6 + 1 两页
    assert all(len(s.images) <= MAX_GALLERY_IMAGES for s in imgs)
    assert plan.slides[0].kind == "title"
    assert plan.slides[-1].kind == "thanks"


def test_gallery_pagination(sample_img):
    slides = _plan_content(_img_blocks([sample_img] * 13), False)
    galls = [s for s in slides if s.kind in ("gallery", "photo_timeline")]
    assert len(galls) == 3                            # 13 -> 6 + 6 + 1
    assert sum(len(s.images) for s in galls) == 13
    assert all(len(s.images) <= MAX_GALLERY_IMAGES for s in galls)


def test_photo_timeline_auto_detect(sample_img, sample_img_with_exif):
    plain = _plan_content(_img_blocks([sample_img]), False)
    timed = _plan_content(_img_blocks([sample_img_with_exif]), False)
    assert all(s.kind == "gallery" for s in plain)          # 无 EXIF -> 普通画廊
    assert all(s.kind == "photo_timeline" for s in timed)   # 有 EXIF -> 时间轴


def test_text_album_keeps_agenda(sample_img):
    # 有文本内容时仍应生成目录（内容页 >= 3）
    blocks = [
        Block(kind="heading", text="第一章", level=1),
        Block(kind="bullets", bullets=["a", "b"]),
        Block(kind="bullets", bullets=["c", "d"]),
        Block(kind="bullets", bullets=["e", "f"]),
        Block(kind="bullets", bullets=["g", "h"]),
        Block(kind="image", image_path=sample_img, source=sample_img),
    ]
    plan = build_deck(blocks, title="报告")
    kinds = [s.kind for s in plan.slides]
    assert "agenda" in kinds
