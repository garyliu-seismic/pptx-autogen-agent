"""PPTX 自动生成 Agent —— 命令行入口。

用法示例：
    python main.py --input examples/sample_data -o output/demo.pptx
    python main.py --input input --title "季度汇报" --subtitle "2026 Q3" --author "张三" --theme business
    python main.py --input input --llm      # 使用 LLM 规划大纲（需 OPENAI_API_KEY）
"""
from __future__ import annotations

import argparse
import glob
import os

import config
from src.analyzer import apply_suggestion, build_deck
from src.llm_advisor import suggest
from src.parsers import SUPPORTED_EXT, parse_file
from src.renderer import render


def collect_files(input_dir: str) -> list[str]:
    files: list[str] = []
    for ext in SUPPORTED_EXT:
        files.extend(glob.glob(os.path.join(input_dir, "*" + ext)))
        files.extend(glob.glob(os.path.join(input_dir, "**", "*" + ext),
                               recursive=True))
    # 排除解析 PDF 时的临时图片目录
    files = [f for f in files if ".ppt_agent_images" not in f]
    return sorted(set(files))


def main() -> int:
    ap = argparse.ArgumentParser(
        description="根据 txt/csv/pdf/png 素材自动生成可编辑的 PPTX")
    ap.add_argument("--input", default=config.DEFAULT_INPUT_DIR,
                    help="素材目录（默认 input）")
    ap.add_argument("-o", "--output", default=config.DEFAULT_OUTPUT_PATH,
                    help="输出 pptx 路径（默认 output/deck.pptx）")
    ap.add_argument("--title", default="", help="封面标题（留空自动推断）")
    ap.add_argument("--subtitle", default="", help="封面副标题")
    ap.add_argument("--author", default="", help="作者署名")
    ap.add_argument("--theme", default="", choices=list(config.THEMES) + [""],
                    help="强制指定主题（默认根据内容自动选择）")
    ap.add_argument("--llm", action="store_true",
                    help="调用 LLM 规划大纲（需设置 OPENAI_API_KEY）")
    args = ap.parse_args()

    if not os.path.isdir(args.input):
        print(f"[错误] 输入目录不存在：{args.input}")
        return 1

    files = collect_files(args.input)
    if not files:
        print(f"[警告] 目录中没有支持的素材文件：{', '.join(sorted(SUPPORTED_EXT))}")
        return 1

    blocks = []
    for f in files:
        print(f"[解析] {f}")
        blocks.extend(parse_file(f))
    if not blocks:
        print("[错误] 未能从素材中提取到任何内容。")
        return 1

    print(f"[分析] 共提取 {len(blocks)} 个内容块，正在规划版式与主题…")
    plan = build_deck(blocks, title=args.title, subtitle=args.subtitle,
                      author=args.author, theme_name=args.theme)

    if args.llm:
        print("[LLM] 正在请求 AI 规划大纲…")
        suggestion = suggest(blocks)
        if suggestion:
            plan = apply_suggestion(plan, suggestion)
            print(f"[LLM] 已采用 AI 建议，主题「{suggestion.get('theme')}」")
        else:
            print("[LLM] 未获得建议（无 Key 或调用失败），继续使用启发式规划。")

    out = render(plan, args.output)
    print(f"\n[完成] 已生成：{os.path.abspath(out)}")
    print(f"       共 {len(plan.slides)} 页，主题「{plan.theme_name}」，16:9")
    for i, s in enumerate(plan.slides, 1):
        print(f"       {i:02d}. [{s.kind:8}] {s.title}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
