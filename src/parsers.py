"""素材解析器：txt / csv / pdf / png -> 统一的 Block 列表。

- txt  : 支持 markdown 标题(#)、项目符号(-/*/•)、编号(1.)、普通段落，并做轻量标题启发式
- csv  : 读取为表格；若含数值列则同时生成原生图表数据
- pdf  : 提取文字(标题/段落/要点) + 表格 + 内嵌图片
- png  : 记录图片路径（可选 OCR，需自行安装 pytesseract）
"""
from __future__ import annotations

import csv
import os
import re

from .models import Block, TableData

SUPPORTED_EXT = {".txt", ".csv", ".pdf", ".png", ".jpg", ".jpeg"}

_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_BULLET_RE = re.compile(r"^(\s*)[-*•▪◦·]\s+(.*)$")
_NUMBER_BULLET_RE = re.compile(r"^(\s*)\d+[.、)）]\s+(.*)$")
_END_PUNCT = (".", "。", "!", "？", "，", ",", "；", ";", "：", ":", "…")


def parse_file(path: str) -> list[Block]:
    """根据扩展名解析单个文件。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".txt":
        return _parse_txt(path)
    if ext == ".csv":
        return _parse_csv(path)
    if ext == ".pdf":
        return _parse_pdf(path)
    if ext in (".png", ".jpg", ".jpeg"):
        return _parse_image(path)
    return []


# ---------------------------------------------------------------------------
# 通用：文本 -> Blocks（txt 与 pdf 共用）
# ---------------------------------------------------------------------------
def _text_to_blocks(text: str, source: str) -> list[Block]:
    lines = [ln.rstrip("\n") for ln in text.splitlines()]
    blocks: list[Block] = []
    i, n = 0, len(lines)
    while i < n:
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        # markdown 标题
        m = _HEADING_RE.match(line)
        if m:
            blocks.append(Block(source=source, kind="heading",
                                text=m.group(2).strip(), level=len(m.group(1))))
            i += 1
            continue

        # 轻量标题启发式：短行 + 下一行空白 + 不以句读结尾
        if (len(line) <= 24 and not line.endswith(_END_PUNCT)
                and (i + 1 < n and not lines[i + 1].strip())):
            blocks.append(Block(source=source, kind="heading", text=line, level=1))
            i += 1
            continue

        # 项目符号 / 编号列表
        if _BULLET_RE.match(line) or _NUMBER_BULLET_RE.match(line):
            bullets: list[str] = []
            while i < n and lines[i].strip():
                m2 = _BULLET_RE.match(lines[i]) or _NUMBER_BULLET_RE.match(lines[i])
                if m2:
                    bullets.append(m2.group(2).strip())
                    i += 1
                else:
                    break
            blocks.append(Block(source=source, kind="bullets", bullets=bullets))
            continue

        # 普通段落：收集到下一个空行为止
        paras = [line]
        i += 1
        while i < n and lines[i].strip():
            paras.append(lines[i].strip())
            i += 1
        blocks.append(Block(source=source, kind="paragraph", text=" ".join(paras)))
    return blocks


# ---------------------------------------------------------------------------
# txt
# ---------------------------------------------------------------------------
def _parse_txt(path: str) -> list[Block]:
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        return _text_to_blocks(f.read(), path)


# ---------------------------------------------------------------------------
# csv
# ---------------------------------------------------------------------------
def _csv_title(base: str) -> str:
    """把文件名转成友好标题：去掉开头数字序号/下划线。"""
    name = re.sub(r"^\d+[_\-]?\s*", "", base).strip()
    return name or base


def _parse_csv(path: str) -> list[Block]:
    with open(path, "r", encoding="utf-8-sig", errors="ignore", newline="") as f:
        rows = [r for r in csv.reader(f) if any(c.strip() for c in r)]
    if not rows:
        return []
    table = TableData(headers=[c.strip() for c in rows[0]],
                      rows=[r for r in rows[1:]])
    base = os.path.splitext(os.path.basename(path))[0]
    # 单一块同时携带表格与(可能的)图表数据：
    # 含数值列 -> 分析器展示为“原生图表页(下方附数据表)”；纯文本 -> 表格页
    return [Block(source=path, kind="table", table=table, chart=table.to_chart(),
                  text=_csv_title(base))]


# ---------------------------------------------------------------------------
# pdf
# ---------------------------------------------------------------------------
def _parse_pdf(path: str) -> list[Block]:
    blocks: list[Block] = []
    try:
        import pdfplumber
    except ImportError:
        return blocks
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                txt = page.extract_text() or ""
                blocks.extend(_text_to_blocks(txt, path))
                for tbl in page.extract_tables() or []:
                    if not tbl or len(tbl[0]) < 2:
                        continue  # 单列/空表视为文本，忽略
                    table = TableData(
                        headers=[str(c or "").strip() for c in tbl[0]],
                        rows=[[str(c or "").strip() for c in row] for row in tbl[1:]])
                    chart = table.to_chart()
                    blocks.append(Block(source=path, kind="table", table=table,
                                        chart=chart))
    except Exception:
        pass
    _extract_pdf_images(path, blocks)
    return blocks


def _extract_pdf_images(path: str, blocks: list[Block], limit: int = 6) -> None:
    """用 PyMuPDF 抽取 PDF 内嵌图片，另存为临时 png 供幻灯片使用。"""
    try:
        try:
            import pymupdf as fitz  # PyMuPDF 新版推荐入口
        except ImportError:
            import fitz
    except ImportError:
        return
    out_dir = os.path.join(os.path.dirname(path), ".ppt_agent_images")
    os.makedirs(out_dir, exist_ok=True)
    try:
        doc = fitz.open(path)
        count = 0
        for pno in range(len(doc)):
            for img in doc[pno].get_images(full=True):
                if count >= limit:
                    break
                xref = img[0]
                pix = fitz.Pixmap(doc, xref)
                if pix.n - pix.alpha > 3:          # CMYK -> RGB
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                out = os.path.join(out_dir,
                                   f"{os.path.splitext(os.path.basename(path))[0]}_p{pno}_{xref}.png")
                pix.save(out)
                blocks.append(Block(source=path, kind="image", image_path=out))
                count += 1
            if count >= limit:
                break
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 图片
# ---------------------------------------------------------------------------
def _parse_image(path: str) -> list[Block]:
    blocks = [Block(source=path, kind="image", image_path=path)]
    # 可选 OCR：安装了 pytesseract + tesseract 才会启用
    try:
        import pytesseract
        from PIL import Image
        txt = pytesseract.image_to_string(Image.open(path), lang="chi_sim+eng")
        txt = re.sub(r"\n{2,}", "\n", txt).strip()
        if txt:
            blocks.extend(_text_to_blocks(txt, path))
    except ImportError:
        pass
    except Exception:
        pass
    return blocks
