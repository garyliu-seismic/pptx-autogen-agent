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

import json
import re
import urllib.parse
import urllib.request

import config
from .llm_advisor import _serialize, chat_json_tools, chat_tools
from .models import ChartData, DeckPlan, MetricData, SlideData, TableData

MAX_BULLETS = 5       # 每页要点上限（比原来 6 更严格，避免密集）
MAX_TURNS = 8
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
        "description": (
            "设置主题配色。可选："
            "tech(科技蓝)/business(商务金)/finance(金融墨)/"
            "health(健康绳)/education(教育橙)/nature(自然青)/"
            "elegant(简约灰)/modern(现代靖)/ai(AI科技)/"
            "dark_pro(深色专业)/corporate(企业蓝灰)/"
            "minimal(极简白)/warm(温暖橙红)/"
            "medical(医疗白蓝)/luxury(黑金奢华)/fresh(清新薄荷)/"
            "academic(学术蓝灰)/government(政务红)"
        ),
        "parameters": {"type": "object", "properties": {
            "theme": {"type": "string", "enum": list(config.THEMES)}},
            "required": ["theme"]}}},
    {"type": "function", "function": {"name": "add_section",
        "description": "添加章节分隔页。每隔 4~6 页内容页应有一个章节页。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"}}, "required": ["title"]}}},
    {"type": "function", "function": {"name": "add_bullets",
        "description": (
            f"添加要点页（bullets 最多 {MAX_BULLETS} 条，每条尽量≤40字）。"
            "适用场景：不属于其他任何专用页型的杂项要点。"
        ),
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "bullets": {"type": "array", "items": {"type": "string"},
                        "maxItems": MAX_BULLETS},
            "takeaway": {"type": "string",
                         "description": "一句话结论（可省）"},
            "notes": {"type": "string",
                      "description": "演讲者备注：本页口头讲解要点，2-4句话（可省）"}},
            "required": ["title", "bullets"]}}},
    {"type": "function", "function": {"name": "add_cards",
        "description": (
            "添加卡片页（items 为 2~6 条并列短句，每条≤26字）。"
            "适用场景：并列的技能点/特性/产品内容等。"
        ),
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "items": {"type": "array", "items": {"type": "string"},
                      "minItems": 2, "maxItems": 6},
            "notes": {"type": "string",
                      "description": "演讲者备注（可省）"}},
            "required": ["title", "items"]}}},
    {"type": "function", "function": {"name": "add_timeline",
        "description": (
            "添加时间轴页（milestones 为 2~8 个里程碑/阶段）。"
            "适用场景：含『第X步』/『阶段X』/『Phase X』/『Step X』 等时序特征的内容。"
        ),
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "milestones": {"type": "array", "items": {"type": "string"},
                          "minItems": 2, "maxItems": 8},
            "notes": {"type": "string",
                      "description": "演讲者备注（可省）"}},
            "required": ["title", "milestones"]}}},
    {"type": "function", "function": {"name": "add_kpi",
        "description": (
            "添加 KPI 指标看板页（3~6 个大数字指标卡）。"
            "适用场景：含有明确数字值（手机个数、增长率、营收额等）的核心指标页。"
        ),
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "metrics": {
                "type": "array",
                "minItems": 1, "maxItems": 6,
                "items": {"type": "object", "properties": {
                    "label": {"type": "string", "description": "指标名，如 营收增长"},
                    "value": {"type": "string", "description": "数字字符串，如 23% / 1.2亿"},
                    "trend": {"type": "string", "enum": ["up", "down", "flat"],
                              "description": "趋势：up上升/down下降/flat持平"},
                    "note": {"type": "string", "description": "补充说明，如 同比上季度（可省）"},
                }, "required": ["label", "value"]}},
            "notes": {"type": "string",
                      "description": "演讲者备注（可省）"}},
            "required": ["title", "metrics"]}}},
    {"type": "function", "function": {"name": "add_quote",
        "description": (
            "添加金句/引语页（大字引语——作者署名）。"
            "适用场景：章节之间的视觉休息，或需要冲击力的核心观点。"
        ),
        "parameters": {"type": "object", "properties": {
            "quote": {"type": "string", "description": "引语正文（建议≤60字）"},
            "author": {"type": "string", "description": "作者/小标题（可省）"},
            "title": {"type": "string", "description": "页标题（可省）"},
            "notes": {"type": "string", "description": "演讲者备注（可省）"}},
            "required": ["quote"]}}},
    {"type": "function", "function": {"name": "add_process",
        "description": (
            "添加流程图页（横向 Step1→Step2→...2~6 步）。"
            "适用场景：A导致B/工作流程/方法论 等有时序层进关系的内容。"
        ),
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "steps": {"type": "array",
                      "items": {"type": "string"},
                      "minItems": 2, "maxItems": 6,
                      "description": "每个步骤的说明（建议≤20字/步）"},
            "notes": {"type": "string", "description": "演讲者备注（可省）"}},
            "required": ["title", "steps"]}}},
    {"type": "function", "function": {"name": "add_comparison",
        "description": (
            "添加左右对比页（Before/After 或 A vs B）。"
            "适用场景：两种方案/状态的对比。"
        ),
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "label_left":  {"type": "string", "description": "左栏标签，如 现状/Before"},
            "label_right": {"type": "string", "description": "右栏标签，如 目标/After"},
            "bullets_left":  {"type": "array", "items": {"type": "string"},
                             "description": "左栏要点"},
            "bullets_right": {"type": "array", "items": {"type": "string"},
                              "description": "右栏要点"},
            "notes": {"type": "string", "description": "演讲者备注（可省）"}},
            "required": ["title", "bullets_left", "bullets_right"]}}},
    {"type": "function", "function": {"name": "add_table",
        "description": "添加数据表页（数据必须来自素材，禁止编造）。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "headers": {"type": "array", "items": {"type": "string"}},
            "rows": {"type": "array", "items": {"type": "array", "items": {"type": "string"}}},
            "source_ref": {"type": "string",
                           "description": "数据来源（素材文件名或片段说明）"}},
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
                "required": ["name", "values"]}},
            "source_ref": {"type": "string",
                           "description": "数据来源（素材文件名或片段说明）"}},
            "required": ["title", "chart_type", "categories", "series"]}}},
    {"type": "function", "function": {"name": "add_summary",
        "description": "添加总结页（points 最多 5 条）。",
        "parameters": {"type": "object", "properties": {
            "title": {"type": "string"},
            "points": {"type": "array", "items": {"type": "string"},
                       "maxItems": 5},
            "notes": {"type": "string",
                      "description": "演讲者备注（可省）"}},
            "required": ["points"]}}},
    {"type": "function", "function": {"name": "add_thanks",
        "description": "添加结尾致谢页。",
        "parameters": {"type": "object", "properties": {
            "notes": {"type": "string",
                      "description": "演讲者备注（可省）"}}}}},
    {"type": "function", "function": {"name": "web_search",
        "description": "联网搜索资料补充内容。当素材信息不足时使用。",
        "parameters": {"type": "object", "properties": {
            "query": {"type": "string"}}, "required": ["query"]}}},
    {"type": "function", "function": {"name": "done",
        "description": "结束规划。", "parameters": {"type": "object", "properties": {}}}},
]

_ALLOWED_CHART = ("column", "bar", "line", "pie")


def execute_tool(plan: DeckPlan, name: str, args: dict,
                 locked_theme: str | None = None) -> str:
    """执行一个计划级工具调用，就地修改 plan。返回描述；ERROR 开头的表示参数非法。

    locked_theme 非空时代表主题已被约束档锁定：set_theme 会被拒绝（防御性检查——
    正常情况下调用方应已把 set_theme 从工具列表里移除，这里是兜底）。
    """
    try:
        if name == "set_title":
            plan.title = str(args.get("title") or plan.title or "演示文稿")
            if args.get("subtitle"):
                plan.subtitle = str(args["subtitle"])
            return f"OK 封面标题={plan.title}"
        if name == "set_theme":
            if locked_theme:
                return f"ERROR 主题已锁定为 {locked_theme}，不可更改"
            t = args.get("theme")
            if t in config.THEMES:
                plan.theme_name = t
                return f"OK 主题={t}"
            return f"ERROR 主题 {t} 不在 {list(config.THEMES)}"

        if name == "add_section":
            plan.slides.append(SlideData(
                title=str(args.get("title") or "章节"), kind="section"))
            return "OK 章节页"
        def _notes(a: dict) -> str:
            return str(a.get("notes") or "")

        if name == "add_bullets":
            bullets = [str(x) for x in (args.get("bullets") or [])]
            if not bullets:
                return "ERROR bullets 不能为空"
            if len(bullets) > MAX_BULLETS:
                return f"ERROR bullets 最多 {MAX_BULLETS} 条，收到 {len(bullets)}"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "要点"), kind="bullets",
                bullets=bullets, takeaway=str(args.get("takeaway") or ""),
                notes=_notes(args)))
            return f"OK 要点页 {len(bullets)} 条"
        if name == "add_cards":
            items = [str(x) for x in (args.get("items") or [])]
            if not (2 <= len(items) <= 6):
                return f"ERROR cards 需 2~6 条，收到 {len(items)}"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "要点"), kind="cards",
                bullets=items, notes=_notes(args)))
            return f"OK 卡片页 {len(items)} 张"
        if name == "add_timeline":
            ms = [str(x) for x in (args.get("milestones") or [])]
            if not ms:
                return "ERROR milestones 不能为空"
            if len(ms) > 8:
                return f"ERROR timeline 最多 8 个，收到 {len(ms)}"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "里程碑"), kind="timeline",
                bullets=ms, notes=_notes(args)))
            return f"OK 时间轴 {len(ms)} 节点"
        if name == "add_kpi":
            raw = args.get("metrics") or []
            if not raw:
                return "ERROR metrics 不能为空"
            metrics = []
            for m in raw:
                if not isinstance(m, dict):
                    continue
                metrics.append(MetricData(
                    label=str(m.get("label") or ""),
                    value=str(m.get("value") or ""),
                    trend=str(m.get("trend") or "flat"),
                    note=str(m.get("note") or ""),
                ))
            if not metrics:
                return "ERROR metrics 解析失败"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "KPI 指标"),
                kind="kpi", metrics=metrics, notes=_notes(args)))
            return f"OK KPI 页 {len(metrics)} 个指标"
        if name == "add_quote":
            q = str(args.get("quote") or "").strip()
            if not q:
                return "ERROR quote 不能为空"
            plan.slides.append(SlideData(
                title=str(args.get("title") or ""),
                kind="quote", takeaway=q,
                author=str(args.get("author") or ""),
                notes=_notes(args)))
            return "OK 金句页"
        if name == "add_process":
            steps = [str(x) for x in (args.get("steps") or [])]
            if not (2 <= len(steps) <= 6):
                return f"ERROR steps 需 2~6 个，收到 {len(steps)}"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "流程"),
                kind="process", bullets=steps, notes=_notes(args)))
            return f"OK 流程页 {len(steps)} 步"
        if name == "add_comparison":
            bl = [str(x) for x in (args.get("bullets_left") or [])]
            br = [str(x) for x in (args.get("bullets_right") or [])]
            if not bl and not br:
                return "ERROR 左右栏内容不能均为空"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "对比分析"),
                kind="comparison",
                label_left=str(args.get("label_left") or "现状 / Before"),
                label_right=str(args.get("label_right") or "目标 / After"),
                bullets=bl, bullets_right=br,
                notes=_notes(args)))
            return f"OK 对比页（左{len(bl)}条/右{len(br)}条）"
        if name == "add_table":
            headers = [str(x) for x in (args.get("headers") or [])]
            rows = [[str(c) for c in r] for r in (args.get("rows") or [])]
            if not headers:
                return "ERROR headers 不能为空"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "数据表"), kind="table",
                table=TableData(headers=headers, rows=rows,
                                source_ref=str(args.get("source_ref") or ""))))
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
                                title=str(args.get("title") or ""),
                                source_ref=str(args.get("source_ref") or ""))))
            return f"OK 图表 {ct} {len(cats)} 类 {len(ser)} 系列"
        if name == "add_summary":
            pts = [str(x) for x in (args.get("points") or [])]
            if not pts:
                return "ERROR points 不能为空"
            plan.slides.append(SlideData(
                title=str(args.get("title") or "总结与展望"),
                kind="summary", bullets=pts[:MAX_BULLETS],
                notes=_notes(args)))
            return f"OK 总结 {len(pts)} 条"
        if name == "add_thanks":
            plan.slides.append(SlideData(
                title=str(args.get("text") or "谢谢 / Thank You"), kind="thanks",
                notes=_notes(args)))
            return "OK 致谢页"
        if name == "web_search":
            q = str(args.get("query") or "").strip()
            if not q:
                return "ERROR query 不能为空"
            return _web_search(q)
        if name == "done":
            return "DONE"
        return f"ERROR 未知工具 {name}"
    except Exception as e:  # noqa: BLE001
        return f"ERROR {e}"


# ---------------------------------------------------------------------------
# 上下文构建器 + few-shot（Phase F）
# ---------------------------------------------------------------------------
def build_system_prompt() -> str:
    themes_str = "、".join(
        f"{k}({v.name})" for k, v in config.THEMES.items()
    )
    return (
        "你是演示文稿规划 Agent，深谙信息设计原则。"
        "根据用户提供的素材，用工具逐页构建一份专业 PPT。\n\n"

        "═══ 核心规则 ═══\n"
        "1. 只能通过工具创建页；先 set_title + set_theme，再逐页添加内容，"
        "最后 add_thanks + done。不要输出工具以外的文字。\n"
        f"2. 每页要点最多 {MAX_BULLETS} 条、每条≤40字；内容超量时宁可拆成多页。\n"
        "3. 画布 16:9（13.33×7.5 英寸）。不要编造素材里没有的数据；"
        "表格/图表/KPI 数字必须来自素材。\n"
        "4. 素材信息不足时可调用 web_search 补充背景事实（结果会回传）。\n\n"

        "═══ 版式选择决策树（每次建页前必须判断）═══\n"
        "• 含明确数字指标（百分比/金额/用户数等）→ add_kpi\n"
        "• 含时序关键词（第X步/阶段X/Phase/Step）→ add_timeline\n"
        "• 两种方案/状态对比（Before&After/A vs B）→ add_comparison\n"
        "• 2~4 个并列短概念（≤22字/条）→ add_cards\n"
        "• 有先后依赖的步骤链（A→B→C）→ add_process\n"
        "• 章节之间需要视觉节奏/核心金句 → add_quote\n"
        "• 其他 ≤5 条要点 → add_bullets\n\n"

        "═══ 视觉节奏控制（全局）═══\n"
        "• 连续 3 页不得是同一 kind；每 4~6 页插 add_section。\n"
        "• KPI/chart 页后面接 bullets 解读会更好。\n"
        "• 总结前可加 quote 页提升收尾感。\n"
        "• 全文页数 6~18 页为宜（素材丰富时上限 20 页）。\n\n"

        f"═══ 可用主题 ═══\n{themes_str}\n"
        "主题选择原则：深色科技/AI→ai或tech；商务财务→business或finance；"
        "企业报告→corporate；高端演讲→dark_pro；极简→minimal；"
        "医疗健康→medical；高端品牌/奖项→luxury；年轻互联网→fresh；"
        "学术研究→academic；政府党政→government\n\n"

        "═══ Few-shot 示例 ═══\n"
        "场景1 季度业务汇报:\n"
        '  set_title({"title":"2026 Q3 业务汇报","subtitle":"季度经营分析"})\n'
        '  set_theme({"theme":"business"})\n'
        '  add_kpi({"title":"核心经营指标","metrics":[\n'
        '    {"label":"营收","value":"3.2亿","trend":"up","note":"同比+23%"},\n'
        '    {"label":"毛利率","value":"41%","trend":"up"},\n'
        '    {"label":"活跃用户","value":"180万","trend":"flat"}]})\n'
        '  add_section({"title":"增长驱动因素"})\n'
        '  add_comparison({"title":"战略转型对比","label_left":"原模式","label_right":"新模式",\n'
        '    "bullets_left":["线下渠道为主","人工处理订单"],\n'
        '    "bullets_right":["全渠道数字化","AI 自动化处理"]})\n'
        '  add_summary({"points":["营收稳健增长","数字化转型提速","下季度重点：用户留存"]})\n'
        '  add_thanks()\n\n'

        "场景2 AI 产品发布:\n"
        '  set_title({"title":"产品发布：NextGen AI 助手"})\n'
        '  set_theme({"theme":"ai"})\n'
        '  add_quote({"quote":"让 AI 真正理解你的意图","author":"产品理念"})\n'
        '  add_cards({"title":"三大核心能力","items":["多模态理解","实时推理","个性化记忆"]})\n'
        '  add_timeline({"title":"发布路线图","milestones":[\n'
        '    "Q1 内测","Q2 公测","Q3 正式发布","Q4 企业版"]})\n'
        '  add_process({"title":"用户使用流程",\n'
        '    "steps":["输入需求","AI 理解意图","生成方案","用户确认","自动执行"]})\n'
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


def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text or "").strip()


def _baike_search(query: str, timeout: int = 15) -> str | None:
    """百度百科 lemma 卡片 API -> 摘要；失败/无结果返回 None。"""
    try:
        url = ("https://baike.baidu.com/api/openapi/BaikeLemmaCardApi"
               "?scope=103&format=json&appid=379020&bk_key=" +
               urllib.parse.quote(query) + "&bk_length=600")
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            data = json.loads(r.read().decode("utf-8"))
        title = str(data.get("title") or query)
        text = _strip_html(str(data.get("abstract") or ""))
        if not text:
            text = _strip_html(str(data.get("desc") or ""))
        if not text:
            return None
        return f"[来源: 百度百科] {title}：{text[:400]}"
    except Exception:
        return None


def _wikipedia_zh_search(query: str, timeout: int = 15) -> str | None:
    """中文 Wikipedia 搜索（search API -> 取首个条目摘要）；失败/无结果返回 None。"""
    try:
        search_url = (
            "https://zh.wikipedia.org/w/api.php?action=query&list=search"
            "&srsearch=" + urllib.parse.quote(query) +
            "&format=json&utf8=1&srlimit=3")
        with urllib.request.urlopen(search_url, timeout=timeout) as r:
            data = json.loads(r.read().decode("utf-8"))
        results = (data.get("query") or {}).get("search") or []
        if not results:
            return None
        try:
            page = urllib.parse.quote(results[0]["title"])
            with urllib.request.urlopen(
                    f"https://zh.wikipedia.org/api/rest_v1/page/summary/{page}",
                    timeout=timeout) as r2:
                d2 = json.loads(r2.read().decode("utf-8"))
            if d2.get("extract"):
                return (f"[来源: Wikipedia 中文] {d2.get('title')}："
                        f"{d2['extract'][:400]}")
        except Exception:
            pass
        lines = []
        for item in results[:3]:
            snippet = _strip_html(str(item.get("snippet") or ""))
            lines.append(f"- {item.get('title', '')}：{snippet}")
        return "[来源: Wikipedia 中文 搜索结果]\n" + "\n".join(lines)
    except Exception:
        return None


def _web_search(query: str, timeout: int = 15) -> str:
    """联网搜索补充内容：百度百科优先，维基百科备选；都失败返回提示文本。"""
    for fn in (_baike_search, _wikipedia_zh_search):
        try:
            r = fn(query, timeout=timeout)
        except Exception:
            r = None
        if r:
            return r
    return ("Web 搜索无结果或网络不通（已试 百度百科 / 维基百科），"
            "请基于已有素材继续。")


def _parse_text_tool_calls(content: str) -> list[dict]:
    """文本 JSON 兜底：解析 {'tool_calls':[...]} 或 [...] 形式的工具列表。"""
    content = (content or "").strip()
    if not content:
        return []
    if content.startswith("```"):
        content = content.split("```", 2)[1]
        if content.startswith("json"):
            content = content[4:]
    try:
        data = json.loads(content)
    except Exception:
        return []
    arr = data.get("tool_calls") if isinstance(data, dict) else data
    out: list[dict] = []
    for c in arr if isinstance(arr, list) else []:
        if not isinstance(c, dict):
            continue
        name = c.get("tool") or c.get("name") or ""
        args = c.get("args") or c.get("arguments") or {}
        if name:
            out.append({"tool": name,
                        "args": args if isinstance(args, dict) else {}})
    return out


def plan_with_llm(blocks, max_turns: int = MAX_TURNS, constraints=None) -> DeckPlan | None:
    """工具调用式规划主循环（带历史的多轮 function calling）；失败/无 Key 返回 None。

    维护完整消息历史（system/user/assistant含tool_calls/tool结果），
    模型才能记住已完成的步骤并继续添加页；端点不支持 function calling 时
    回退文本 JSON 一次执行。

    constraints 非空且设置了 locked_theme 时：预先锁定 plan.theme_name，
    并把 set_theme 从工具列表里剔除，防止 LLM 中途改主题。
    """
    locked_theme = getattr(constraints, "locked_theme", None) if constraints else None
    plan = DeckPlan()
    if locked_theme:
        plan.theme_name = locked_theme
    tools = [t for t in PLAN_TOOLS if not locked_theme
             or t["function"]["name"] != "set_theme"]
    system = build_system_prompt()
    messages: list[dict] = [
        {"role": "system", "content": system},
        {"role": "user", "content": build_user_prompt(blocks, plan)},
    ]
    done = False
    last_count = 0
    stall = 0
    for _ in range(max_turns):
        msg = chat_tools(messages, tools)
        if not msg:
            break
        tool_calls = msg.get("tool_calls")
        if not tool_calls:
            # 端点不支持 function calling -> 文本 JSON 兜底，一次执行完
            for c in _parse_text_tool_calls(msg.get("content")):
                result = execute_tool(plan, c.get("tool"), c.get("args") or {},
                                      locked_theme=locked_theme)
                if result == "DONE":
                    done = True
            break
        # 回填 assistant 消息（含 tool_calls），保持多轮上下文
        messages.append({"role": "assistant",
                         "content": msg.get("content") or "",
                         "tool_calls": tool_calls})
        grew = False
        for tc in tool_calls:
            fn = tc.get("function", {}) or {}
            name = fn.get("name", "")
            try:
                args = json.loads(fn.get("arguments") or "{}")
            except Exception:
                args = {}
            result = execute_tool(plan, name,
                                  args if isinstance(args, dict) else {},
                                  locked_theme=locked_theme)
            if result == "DONE":
                done = True
            elif result.startswith("OK") or not result.startswith("ERROR"):
                grew = True   # 建页或 web_search 等有效调用都算进展
            # 工具结果回填（含 tool_call_id），供模型下一轮继续
            messages.append({"role": "tool",
                             "tool_call_id": tc.get("id") or "",
                             "content": result})
        if done:
            break
        # 防卡死：连续两轮无进展（全是 ERROR）则停止
        if len(plan.slides) > last_count or grew:
            last_count = len(plan.slides)
            stall = 0
        else:
            stall += 1
            if stall >= 2:
                break
    if not finalize(plan):
        return None
    return plan
