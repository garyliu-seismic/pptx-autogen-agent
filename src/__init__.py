"""pptx-autogen-agent: 根据 txt/csv/pdf/png 素材自动生成可编辑的 PPTX。"""

from .parsers import parse_file, SUPPORTED_EXT
from .analyzer import build_deck
from .renderer import render

__all__ = ["parse_file", "SUPPORTED_EXT", "build_deck", "render"]
__version__ = "0.1.0"
