"""LibreOffice 视觉回归：把 pptx 渲染成每页 PNG 预览图（真实验证排版效果）。

流程：soffice --headless 转 PDF -> PyMuPDF 逐页导出 PNG -> 可选拼成 contact sheet。

用法：
    python tools/render_preview.py output/llm_intro_v6.pptx --outdir preview
    python tools/render_preview.py output/llm_intro_v6.pptx --contact   # 顺便出缩略总览图

依赖：需要安装 LibreOffice（否则仅输出提示与安装方法）。
- Windows: winget install TheDocumentFoundation.LibreOffice
- 也可用环境变量 LIBREOFFICE_PATH 指向 soffice(.exe)
其余依赖（pymupdf、Pillow）已在 requirements.txt。
"""
from __future__ import annotations

import argparse
import glob
import math
import os
import shutil
import subprocess
import sys
import tempfile

_PREVIEW_DIR = "preview"


def find_soffice() -> str | None:
    env = os.environ.get("LIBREOFFICE_PATH")
    if env and os.path.exists(env):
        return env
    cands: list[str] = []
    if sys.platform == "win32":
        cands = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
    elif sys.platform == "darwin":
        cands = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    else:
        cands = ["/usr/bin/soffice", "/usr/local/bin/soffice"]
    for c in cands:
        if os.path.exists(c):
            return c
    return shutil.which("soffice")


def print_install_hint() -> None:
    print("[提示] 未找到 LibreOffice，无法渲染视觉预览。安装方式（任选其一）：")
    if sys.platform == "win32":
        print("   1) winget install TheDocumentFoundation.LibreOffice")
        print("   2) 从 https://www.libreoffice.org/download 安装后设置环境变量：")
        print("      set LIBREOFFICE_PATH=C:\\Program Files\\LibreOffice\\program\\soffice.exe")
    else:
        print("   1) sudo apt install libreoffice   (Debian/Ubuntu)")
        print("   2) brew install --cask libreoffice (macOS)")


def pptx_to_pdf(soffice: str, pptx: str, outdir: str) -> str:
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf",
         "--outdir", outdir, os.path.abspath(pptx)],
        check=True, capture_output=True, timeout=600)
    base = os.path.splitext(os.path.basename(pptx))[0] + ".pdf"
    pdf = os.path.join(outdir, base)
    if not os.path.exists(pdf):
        pdfs = sorted(glob.glob(os.path.join(outdir, "*.pdf")))
        if not pdfs:
            raise RuntimeError("LibreOffice 未生成 PDF")
        pdf = pdfs[0]
    return pdf


def pdf_to_pngs(pdf: str, outdir: str, scale: float = 1.5) -> list[str]:
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    doc = fitz.open(pdf)
    files: list[str] = []
    for i, page in enumerate(doc, 1):
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
        out = os.path.join(outdir, f"page-{i:03d}.png")
        pix.save(out)
        files.append(out)
    return files


def make_contact(pages: list[str], out_path: str, cols: int = 4) -> str:
    from PIL import Image
    imgs = [Image.open(p) for p in pages]
    w = max(im.width for im in imgs)
    h = max(im.height for im in imgs)
    rows = math.ceil(len(imgs) / cols)
    sheet = Image.new("RGB", (cols * w, rows * h), "white")
    for i, im in enumerate(imgs):
        im = im.resize((w, h))
        sheet.paste(im, ((i % cols) * w, (i // cols) * h))
    sheet.save(out_path)
    return out_path


def render_pipeline(pptx: str, outdir: str | None = None,
                    scale: float = 1.5, contact: bool = False) -> tuple[list[str], str | None]:
    """渲染 pptx 为每页 PNG；返回 (页面列表, contact sheet 路径或 None)。"""
    soffice = find_soffice()
    if not soffice:
        print_install_hint()
        raise RuntimeError("LibreOffice 未安装")
    outdir = outdir or os.path.join(os.path.dirname(os.path.abspath(pptx)), _PREVIEW_DIR)
    os.makedirs(outdir, exist_ok=True)
    pdf = pptx_to_pdf(soffice, pptx, outdir)
    pages = pdf_to_pngs(pdf, outdir, scale)
    contact_path = None
    if contact and pages:
        contact_path = make_contact(pages, os.path.join(outdir, "contact.png"))
    return pages, contact_path


def main() -> int:
    ap = argparse.ArgumentParser(description="pptx -> PNG 视觉预览（LibreOffice 渲染）")
    ap.add_argument("pptx", help="要渲染的 .pptx 路径")
    ap.add_argument("--outdir", default=None, help="输出目录（默认 <pptx目录>/preview）")
    ap.add_argument("--scale", type=float, default=1.5, help="渲染缩放倍率（默认 1.5）")
    ap.add_argument("--contact", action="store_true", help="同时生成 contact sheet 总览图")
    args = ap.parse_args()

    if not os.path.exists(args.pptx):
        print(f"[错误] 文件不存在：{args.pptx}")
        return 1
    try:
        pages, contact = render_pipeline(args.pptx, args.outdir,
                                         args.scale, args.contact)
    except RuntimeError as e:
        print(f"[错误] {e}")
        return 2

    print(f"[完成] 已渲染 {len(pages)} 页 -> {os.path.dirname(pages[0])}")
    if contact:
        print(f"       总览图：{contact}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
