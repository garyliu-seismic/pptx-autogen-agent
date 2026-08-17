"""工具调用式规划器 —— Phase A + Phase F（借鉴 ppt master 的 function calling）。

把「页型」抽象成计划级工具，LLM 通过结构化 tool_calls 逐页构建 DeckPlan，
而不是一次性吐一个脆弱的大纲 JSON。

  - PLAN_TOOLS：工具 JSON Schema（含必填/枚举约束）
  - execute_tool：确定性执行器，就地更新 DeckPlan；参数非法返回 ERROR 供 LLM 修正
  - build_system_prompt / build_user_prompt：上下文构建器 + few-shot（Phase F）
  - plan_with_llm：多轮循环，直到 done() 或达到轮数上限

兼容性：原生 function calling 优先，端点不支持时回退文本 JSON（chat_json_tools）。
失败一律返回 None，调用方降级到启发式 build_deck + suggest。
"""
from __future__ import annotations

import config
from .llm_advisor import _serialize, chat_json_tools
from .models import ChartData, DeckPlan, SlideData, TableData

MAX_BULLETS = 6
MAX_TURNS = 4
MAX_CHARS_INTOOL = 6000

# ---------------------------------------------------------------------------
# 计划级工具定义（OpenAI function calling 格式）
# ---------------------------------------------------------------------------
PLAN_TOOLS: list[dict] = [
    {"type": "function", "function": {"name": "set_title",
        "description": "设置封面标题与副标题（第 1 页）。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"}, "subtitle": {"type": "string"}},
            "required": ["title"]}}},
    {"type": "function", "function": {"name": "set_theme",
        "description": "设置主题配色。可选：tech/business/finance/health/education/nature/elegant/modern",
        "parameters": {"type": "object", "properties": {
            "theme": {"type": "string", "enum": list(config.THEMES)}},
            "required": ["theme"]}}},
    {"type": "function", "function": {"name": "add_section",
        "description": "添加章节分隔页。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"}}, "required": ["title"]}}},
    {"type": "function", "function": {"name": "add_bullets",
        "description": f"添加要点页（bullets 最多 {MAX_BULLETS} 条，每条尽量≤60字）。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "bullets": {"type": "array", "items": {"type": "string"}},
            "takeaway": {"type": "string"}},
            "required": ["title", "bullets"]}}},
    {"type": "function", "function": {"name": "add_cards",
        "description": "添加卡片页（items 为 2~6 条短句，每条≤26字）。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "items": {"type": "array", "items": {"type": "string"}}},
            "required": ["title", "items"]}}},
    {"type": "function", "function": {"name": "add_timeline",
        "description": "添加时间轴页（milestones 为里程碑，最多 6 个）。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "milestones": {"type": "array", "items": {"type": "string"}}},
            "required": ["title", "milestones"]}}},
    {"type": "function", "function": {"name": "add_table",
        "description": "添加数据表页（数据必须来自素材，禁止编造）。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "headers": {"type": "array", "items": {"type": "string"}},
            "rows": {"type": "array", "items": {"type": "array", "items": {"type": "string"}}}},
            "required": ["title", "headers", "rows"]}}},
    {"type": "function", "function": {"name": "add_chart",
        "description": "添加原生图表页。chart_type: column/bar/line/pie。数据必须来自素材。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "chart_type": {"type": "string", "enum": ["column", "bar", "line", "pie"]},
            "categories": {"type": "array", "items": {"type": "string"}},
            "series": {"type": "array", "items": {"type": "object", "properties": {
                "name": {"type": "string"},
                "values": {"type": "array", "items": {"type": "number"}}},
                "required": ["name", "values"]}}},
            "required": ["title", "chart_type", "categories", "series"]}}},
    {"type": "function", "function": {"name": "add_summary",
        "description": "添加总结页（points 最多 6 条）。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "points": {"type": "array", "items": {"type": "string"}}},
            "required": ["points"]}}},
    {"type": "function", "function": {"name": "add_thanks",
        "description": "添加结尾致谢页。",
        "parameters": {"type": "object", "properties": {}}}},
    {"type": "function", "function": {"name": "done",
        "description": "结束规划。", "parameters": {"type": "object", "properties": {}}}},
]

_ALLOWED_CHART = ("column", "bar", "line", "pie")


def execute_tool(plan: DeckPlan, name: str, args: dict) -> str:
    """执行一个计划级工具调用，就地修改 plan。返回描述；ERROR 开头的表示参数非法。"""
    try:
        if name == "set_title":
            plan.title = str(args.get("title") or plan.title or "演示文稿")
            if args.get("subtitle"):
                plan.subtitle = str(args["subtitle"])
            return f"OK 封面标题={plan.title}"
        if name == "set_theme":
            t = args.get("theme")
            if t in config.THEMES:
                plan.theme_name = t
                return f"OK 主题={t}"
            return f"ERROR 主题 {t} 不在 {list(config.THEMES)}"

        if name == "add_section":
            plan.slides.append(SlideData(
                title=str(args.get("title") or "章节"), kind="section"))
            return "OK 章节页"
        if name == "add_bullets":
            bullets = [str(x) for x in (args.get("bullets") or [])]
            if not bullets:
                return "ERROR bullets 不能为空"
            if len(bullets) > MAX_BULLETS:
                return f"ERROR bullets 最多 {MAX_BULLETS} 条，收到 {len(bullets)}"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "要点"), kind="bullets",
                bullets=bullets, takeaway=str(args.get("takeaway") or "")))
            return f"OK 要点页 {len(bullets)} 条"
        if name == "add_cards":
            items = [str(x) for x in (args.get("items") or [])]
            if not (2 <= len(items) <= 6):
                return f"ERROR cards 需 2~6 条，收到 {len(items)}"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "要点"), kind="cards", bullets=items))
            return f"OK 卡片页 {len(items)} 张"
        if name == "add_timeline":
            ms = [str(x) for x in (args.get("milestones") or [])]
            if not ms:
                return "ERROR milestones 不能为空"
            if len(ms) > 6:
                return f"ERROR timeline 最多 6 个，收到 {len(ms)}"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "里程碑"), kind="timeline", bullets=ms))
            return f"OK 时间轴 {len(ms)} 节点"
        if name == "add_table":
            headers = [str(x) for x in (args.get("headers") or [])]
            rows = [[str(c) for c in r] for r in (args.get("rows") or [])]
            if not headers:
                return "ERROR headers 不能为空"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "数据表"), kind="table",
                table=TableData(headers=headers, rows=rows)))
            return f"OK 表格 {len(rows)} 行"
        if name == "add_chart":
            ct = args.get("chart_type")
            if ct not in _ALLOWED_CHART:
                return f"ERROR chart_type {ct} 非法，应为 {_ALLOWED_CHART}"
            cats = [str(c) for c in (args.get("categories") or [])]
            ser = args.get("series") or []
            if not cats or not ser:
                return "ERROR 图表数据为空"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "数据图表"), kind="chart",
                chart=ChartData(chart_type=ct, categories=cats, series=ser,
                                title=str(args.get("title") or ""))))
            return f"OK 图表 {ct} {len(cats)} 类 {len(ser)} 系列"
        if name == "add_summary":
            pts = [str(x) for x in (args.get("points") or [])]
            if not pts:
                return "ERROR points 不能为空"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "总结与展望"),
                kind="summary", bullets=pts[:MAX_BULLETS]))
            return f"OK 总结 {len(pts)} 条"
        if name == "add_thanks":
            plan.slides.append(SlideData(
                title=str(args.get("text") or "谢谢 / Thank You"), kind="thanks"))
            return "OK 致谢页"
        if name == "done":
            return "DONE"
        return f"ERROR 未知工具 {name}"
    except Exception as e:  # noqa: BLE001
        return f"ERROR {e}"


# ---------------------------------------------------------------------------
# 上下文构建器 + few-shot（Phase F）
# ---------------------------------------------------------------------------
def build_system_prompt() -> str:
    themes = "、".join(f"{k}({v.name})" for k, v in config.THEMES.items())
    return (
        "你是演示文稿规划 Agent。根据用户提供的素材，用工具逐页构建一份专业 PPT。\n"
        "规则：\n"
        "1. 只能通过工具创建页：先 set_title / set_theme，再逐页添加内容，"
        "最后 add_thanks 与 done。\n"
        f"2. 每页要点最多 {MAX_BULLETS} 条、每条尽量≤60字；内容过多时拆成多页。\n"
        "3. 画布 16:9（13.33×7.5 英寸）。不要编造素材里没有的数据；"
        "表格/图表数据必须来自素材。\n"
        "4. 结构建议：封面 → (目录/章节) → 要点/卡片/时间轴/图表 → 总结 → 致谢；"
        "页数 6~20 页为宜。\n"
        f"5. 可用主题：{themes}。\n"
        "6. 只调用工具，不要输出工具以外的文字。\n"
        "示例（few-shot）：\n"
        '1) set_title({"title":"季度业务汇报","subtitle":"2026 Q3"})  '
        'set_theme({"theme":"business"})\n'
        '2) add_section({"title":"经营概览"})  '
        'add_bullets({"title":"核心指标","bullets":["营收同比增长 23%","毛利率 41%"],'
        '"takeaway":"整体稳健增长"})\n'
        '3) add_cards({"title":"三大战略","items":["市场扩张","产品创新","客户成功"]})  '
        'add_summary({"points":["持续增长","控制风险"]})\n'
    )


def build_user_prompt(blocks, plan: DeckPlan) -> str:
    material = _serialize(blocks, max_chars=MAX_CHARS_INTOOL)
    parts = ["素材内容：\n" + material]
    if plan.slides:
        done = [f"{i}. {s.kind}: {s.title}"
                for i, s in enumerate(plan.slides, 1)]
        parts.append("\n当前已完成页：\n" + "\n".join(done))
        parts.append("\n请继续添加后续页；若已完整请调用 done。")
    else:
        parts.append("\n请开始规划（先 set_title / set_theme，然后逐页添加）。")
    return "\n".join(parts)


def finalize(plan: DeckPlan) -> bool:
    """补全封面/致谢；保证至少 2 页。返回是否有效。"""
    if not plan.slides:
        return False
    if plan.slides[0].kind != "title":
        plan.slides.insert(0, SlideData(title=plan.title or "演示文稿", kind="title"))
    if not any(s.kind == "thanks" for s in plan.slides):
        plan.slides.append(SlideData(title="谢谢 / Thank You", kind="thanks"))
    return True


def plan_with_llm(blocks, max_turns: int = MAX_TURNS) -> DeckPlan | None:
    """工具调用式规划主循环；失败/无 Key 返回 None。"""
    plan = DeckPlan()
    system = build_system_prompt()
    done = False
    for _ in range(max_turns):
        user = build_user_prompt(blocks, plan)
        calls = chat_json_tools(system, user, PLAN_TOOLS)
        if not calls:
            break
        progressed = False
        for c in calls:
            result = execute_tool(plan, c.get("tool") or c.get("name"),
                                  c.get("args") or {})
            if result == "DONE":
                done = True
            elif result.startswith("OK"):
                progressed = True
        if done:
            break
        if not progressed:
            break  # 本轮无有效进展，避免死循环
    if not finalize(plan):
        return None
    return plan
