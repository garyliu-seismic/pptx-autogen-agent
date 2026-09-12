"""PPTX 自动生成 Agent —— 命令行入口。

用法示例：
    python main.py --input examples/sample_data -o output/demo.pptx
    python main.py --input input --title "季度汇报" --subtitle "2026 Q3" --author "张三" --theme business
    python main.py --input input --llm      # 使用 LLM 规划大纲（需 OPENAI_API_KEY）
"""
from __future__ import annotations

import argparse
import glob
import json
import os

import config
from src.analyzer import apply_suggestion, build_deck
from src.constraints import load_constraints
from src.llm_advisor import suggest
from src.llm_repair import repair_plan_with_llm, repair_render_with_llm
from src.parsers import SUPPORTED_EXT, parse_file
from src.planio import load_plan, save_plan
from src.quality import LayoutConstraintError, render_with_quality
from src.validate_plan import repair_deterministic, validate_plan


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
    ap.add_argument("--planner", action="store_true",
                    help="用工具调用式规划器构建大纲（需同时 --llm，兼容 function calling）")
    ap.add_argument("--preview", action="store_true",
                    help="生成后用 LibreOffice 渲染每页 PNG 预览（需安装 LibreOffice）")
    ap.add_argument("--no-plan-qa", action="store_true",
                    help="跳过渲染前规划校验与修复（默认开启）")
    ap.add_argument("--no-readback", action="store_true",
                    help="跳过渲染后读回内容 QA（默认开启）")
    ap.add_argument("--plan-only", nargs="?", const=config.DEFAULT_PLAN_PATH,
                    metavar="PLAN_JSON", default=None,
                    help="只输出计划 JSON 供预览/编辑（默认 %(const)s），不渲染")
    ap.add_argument("--plan", metavar="PLAN_JSON", default=None,
                    help="载入已确认/编辑过的计划 JSON 再渲染（跳过 LLM 规划）")
    ap.add_argument("--constraints", metavar="CONSTRAINTS_JSON", default=None,
                    help="约束档 JSON（主题锁定/数据来源/版式硬拒/结构要求），默认不启用")
    args = ap.parse_args()

    constraints = load_constraints(args.constraints)

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
    theme_name = (constraints.locked_theme if constraints else "") or args.theme
    plan = build_deck(blocks, title=args.title, subtitle=args.subtitle,
                      author=args.author, theme_name=theme_name)

    if args.llm and not args.plan:
        if args.planner:
            print("[LLM] 正在用工具调用式规划器构建大纲…")
            from src.planner import plan_with_llm
            planned = plan_with_llm(blocks, constraints=constraints)
            if planned is not None:
                plan = planned
                print(f"[LLM] 规划器完成：共 {len(plan.slides)} 页，主题"
                      f"「{plan.theme_name}」")
            else:
                print("[LLM] 规划器失败（无 Key 或调用错误），回退启发式 + 大纲建议。")
                suggestion = suggest(blocks)
                if suggestion:
                    plan = apply_suggestion(plan, suggestion)
        else:
            print("[LLM] 正在请求 AI 规划大纲…")
            suggestion = suggest(blocks)
            if suggestion:
                plan = apply_suggestion(plan, suggestion)
                print(f"[LLM] 已采用 AI 建议，主题「{suggestion.get('theme')}」")
            else:
                print("[LLM] 未获得建议（无 Key 或调用失败），继续使用启发式规划。")

    # P2-E：--plan 载入已确认/编辑过的计划（跳过上面的 LLM 规划）
    if args.plan:
        try:
            plan = load_plan(args.plan)
            print(f"[计划] 已载入 {args.plan}：{len(plan.slides)} 页")
        except Exception as e:
            print(f"[错误] 无法载入计划文件 {args.plan}：{e}")
            return 1

    # Phase B：渲染前规划校验 + 修复（无 Key 时仅做确定性修复）
    known_sources = {b.source for b in blocks if b.source}
    if not args.no_plan_qa:
        pqa_issues = validate_plan(plan, constraints=constraints, known_sources=known_sources)
        if pqa_issues:
            print(f"[规划质检] 发现 {len(pqa_issues)} 项问题")
            if args.llm:
                repaired = repair_plan_with_llm(plan, pqa_issues)
                if repaired is not None:
                    plan = repaired
                    print("[规划质检] 已用 LLM 修复问题页")
                    pqa_issues = validate_plan(plan, constraints=constraints,
                                               known_sources=known_sources)
            plan, nfixed = repair_deterministic(plan, constraints=constraints,
                                                known_sources=known_sources)
            if nfixed:
                print(f"[规划质检] 确定性修复 {nfixed} 项")
        else:
            print("[规划质检] 通过")

    # P2-E：--plan-only 输出计划 JSON 供预览/编辑后退出
    if args.plan_only:
        save_plan(plan, args.plan_only)
        print(f"\n[计划] 已写出：{os.path.abspath(args.plan_only)}")
        print(f"       共 {len(plan.slides)} 页，主题「{plan.theme_name}」")
        for i, s in enumerate(plan.slides, 1):
            print(f"       {i:02d}. [{s.kind:8}] {s.title}")
        print("       编辑该 JSON 后用 --plan 载入渲染。")
        return 0

    # 逐图旋转：素材目录下的 rotate.json（{文件名: 度数}）
    rotate_map: dict = {}
    rot_file = os.path.join(args.input, "rotate.json")
    if os.path.isfile(rot_file):
        try:
            with open(rot_file, encoding="utf-8") as f:
                rotate_map = json.load(f)
            print(f"[旋转] 已载入 {len(rotate_map)} 条旋转设置（{rot_file}）")
        except Exception as e:
            print(f"[警告] 无法解析 {rot_file}：{e}")

    repairer = repair_render_with_llm if args.llm else None
    hard_fail = bool(constraints and constraints.hard_fail_on_layout)
    try:
        out, issues, tries = render_with_quality(plan, args.output, repairer=repairer,
                                                 rotate_map=rotate_map, hard_fail=hard_fail)
    except LayoutConstraintError as e:
        print(f"[错误] 版式约束未通过（hard_fail_on_layout）：{e}")
        return 1

    # Phase D：读回内容 QA + 空壳页自愈（默认开启，--no-readback 关闭）
    if not args.no_readback:
        from src.readback import auto_fix, compare_plan_render, summary as rb_summary
        rb_issues = compare_plan_render(plan, out)
        if rb_issues:
            print(f"[读回QA] {rb_summary(rb_issues)}")
            new_plan, ndrop = auto_fix(plan, rb_issues)
            if ndrop:
                print(f"[读回QA] 剔除 {ndrop} 个空壳/残缺页，重新渲染…")
                plan = new_plan
                try:
                    out, issues, tries = render_with_quality(
                        plan, args.output, repairer=repairer,
                        rotate_map=rotate_map, hard_fail=hard_fail)
                except LayoutConstraintError as e:
                    print(f"[错误] 版式约束未通过（hard_fail_on_layout）：{e}")
                    return 1
        else:
            print("[读回QA] 通过")

    print(f"\n[完成] 已生成：{os.path.abspath(out)}（质检尝试 {tries} 次）")
    print(f"       共 {len(plan.slides)} 页，主题「{plan.theme_name}」，16:9")
    for i, s in enumerate(plan.slides, 1):
        print(f"       {i:02d}. [{s.kind:8}] {s.title}")
    if issues:
        from src.quality import summary
        print(f"       [提示] {summary(issues)}")

    if args.preview:
        try:
            from tools.render_preview import render_pipeline, print_install_hint
            pages, contact = render_pipeline(out)
            print(f"[预览] 已渲染 {len(pages)} 页 -> {os.path.dirname(pages[0])}")
            if contact:
                print(f"[预览] 总览图：{contact}")
        except RuntimeError:
            print_install_hint()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
