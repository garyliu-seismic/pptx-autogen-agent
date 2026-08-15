"""生成示例素材：PNG 图片与 PDF 报告（txt/csv 已作为静态文件存在）。"""
import os

from PIL import Image, ImageDraw


def main() -> None:
    base = os.path.normpath(os.path.join(os.path.dirname(__file__), "..",
                                         "examples", "sample_data"))
    os.makedirs(base, exist_ok=True)

    # 1) 生成示例产品图（png）
    for name, color, label in [
        ("产品图_数据看板.png", "#1F4E79", "Data Dashboard"),
        ("产品图_智能报告.png", "#0F6C4F", "Smart Report"),
    ]:
        img = Image.new("RGB", (960, 600), color)
        d = ImageDraw.Draw(img)
        d.rectangle([60, 60, 900, 540], outline="#FFFFFF", width=6)
        # 画一个简易“柱状图”装饰
        bars = [(160, 300, 280, 460), (320, 240, 440, 460),
                (480, 180, 600, 460), (640, 320, 760, 460)]
        for x0, y0, x1, y1 in bars:
            d.rectangle([x0, y0, x1, y1], fill="#FFFFFF", outline="#FFFFFF", width=2)
        d.line([160, 480, 760, 480], fill="#DDDDDD", width=3)
        d.text((330, 100), label, fill="#FFFFFF")
        img.save(os.path.join(base, name))
        print(f"[生成] {name}")

    # 2) 生成示例 PDF（PyMuPDF 绘制一页英文报告，避免中文字体问题）
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        doc = fitz.open()
        page = doc.new_page(width=595, height=842)  # A4
        page.insert_text((72, 80), "Quarterly Business Report", fontsize=22)
        lines = [
            "This PDF is a sample hint file for the PPTX auto-generation agent.",
            "It summarizes key business metrics of the last four quarters.",
            "",
            "Highlights",
            "1. Revenue grew steadily from Q1 to Q4.",
            "2. Profit margin improved after Q2 optimization.",
            "3. Customer acquisition cost dropped by 15%.",
            "",
            "Key Risks",
            "1. Supply chain volatility in Q3.",
            "2. Increased competition in the domestic market.",
        ]
        y = 130
        for ln in lines:
            page.insert_text((72, y), ln, fontsize=12)
            y += 22
        doc.save(os.path.join(base, "季度业务报告.pdf"))
        print("[生成] 季度业务报告.pdf")
    except ImportError:
        print("[跳过] 未安装 PyMuPDF，跳过 PDF 示例。")

    print("\n示例素材已就绪：examples/sample_data/")


if __name__ == "__main__":
    main()
