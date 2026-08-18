"""pytest 共享夹具：确保项目根目录在 sys.path（import config / from src...）。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest


@pytest.fixture
def sample_img(tmp_path):
    """生成一张无 EXIF 的小图（jpg）。"""
    from PIL import Image
    p = tmp_path / "plain.jpg"
    Image.new("RGB", (100, 80), "skyblue").save(p)
    return str(p)


@pytest.fixture
def sample_img_with_exif(tmp_path):
    """生成一张带 EXIF 拍摄时间的小图（jpg）。"""
    from PIL import Image
    p = tmp_path / "timed.jpg"
    im = Image.new("RGB", (100, 80), "skyblue")
    ex = Image.Exif()
    ex[36867] = "2003:12:16 09:45:50"
    im.save(p, exif=ex)
    return str(p)
