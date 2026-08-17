"""内容分析与版面规划：把解析出的 Blocks 组织成 DeckPlan（封面/目录/正文/总结/结尾）。

版式决策规则（启发式，可被 LLM 建议覆盖）：
- 首个一级标题  -> 封面标题（未指定时）
- 多个一级标题  -> 自动生成“目录”页
- 二级及以下标题 -> 章节分隔页 + 内容页
- 要点/段落    -> 要点页（每页最多 6 条，自动分页）
- 表格+数值    -> 原生图表页（PowerPoint 可编辑）
- 表格（纯文本）-> 表格页
- 图片        -> 图片页 / 画廊页
- 含“总结/结论”小节 -> 结尾加“总结与展望”页
"""
from __future__ import annotations

import os
import re

import config
from .models import Block, DeckPlan, SlideData
from .theme_engine import pick_theme

BULLETS_PER_SLIDE = 6
SUMMARY_KEYWORDS = ["总结", "结论", "展望", "下一步", "行动项",
                    "summary", "conclusion", "next steps", "action items"]

# 时间轴特征：第1周 / 第 2 阶段 / Step 3 / Phase 2 ...
_TIMELINE_RE = re.compile(
    r"第\s*\d+\s*(周|天|月|步|阶段|期|季度)|(?:step|phase|stage)\s*\d+",
    re.IGNORECASE)


def _refine_kind(sd: SlideData) -> SlideData:
    """根据内容特征自动升级版式：bullets -> 卡片 / 时间轴。"""
    if sd.kind != "bullets":
        return sd
    items = [b for b in sd.bullets if b.strip()]
    if not items:
        return sd
    # 分阶段/里程碑 -> 时间轴
    if len(items) <= 6 and any(_TIMELINE_RE.search(it) for it in items):
        sd.kind = "timeline"
        return sd
    # 2~4 条短要点 -> 卡片
    if 2 <= len(items) <= 4 and all(len(it) <= 26 for it in items):
        sd.kind = "cards"
        return sd
    return sd


def _split_long(text: str, max_len: int = 80) -> list[str]:
    """把超长文本按句子/逗号切分成多条短要点，防止排版溢出、提升可读性。"""
    text = text.strip()
    if len(text) <= max_len:
        return [text]
    sentences = [s for s in re.split(r"(?<=[。！？.!?])", text) if s.strip()]
    pieces: list[str] = []
    buf = ""
    for s in sentences:
        if not buf:
            buf = s
        elif len(buf) + len(s) <= max_len:
            buf += s
        else:
            pieces.append(buf)
            buf = s
        while len(buf) > max_len:
            cut = max(buf.rfind("，"), buf.rfind(","))
            if cut <= 0:
                cut = max_len
            pieces.append(buf[:cut + 1])
            buf = buf[cut + 1:]
    if buf:
        pieces.append(buf)
    return [p.strip() for p in pieces if p.strip()]


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------
def _theme_key(theme) -> str:
    """由 Theme 对象反查 config.THEMES 的 key（如 business），找不到用默认。"""
    for k, v in config.THEMES.items():
        if v is theme or v.name == theme.name:
            return k
    return config.DEFAULT_THEME


def build_deck(blocks: list[Block], title: str = "", subtitle: str = "",
               author: str = "", theme_name: str = "") -> DeckPlan:
    theme = pick_theme(blocks, theme_name or None)
    # 存主题 key（如 business）而非显示名（如 商务金），供 config.THEMES 索引
    theme_key = theme_name if theme_name in config.THEMES else _theme_key(theme)
    if not title:
        title = _guess_title(blocks)

    # 提取“图解说明”小节（标题含“图解/Figure”，后跟要点），供图片配对成图文页
    blocks, figure_specs = _extract_figures(blocks)

    slides: list[SlideData] = [SlideData(title=title, kind="title",
                                         notes=f"主题配色：{theme.name}")]

    # 小节分隔页仅在数量适中(2~6)时启用，避免页数臃肿
    sec_heads = [b for b in blocks if b.kind == "heading" and b.level >= 2]
    use_dividers = 2 <= len(sec_heads) <= 6

    content_slides = _plan_content(blocks, use_dividers, figure_specs)

    # 目录页
    if len(content_slides) >= 3:
        agenda_items = [s.title for s in content_slides
                        if s.kind != "section" and s.title][:8]
        if len(agenda_items) >= 2:
            slides.append(SlideData(title="目录 / Agenda", kind="agenda",
                                    bullets=agenda_items))

    slides.extend(content_slides)

    # 总结页
    if _has_summary(blocks):
        slides.append(SlideData(title="总结与展望", kind="summary",
                                bullets=_collect_summary(blocks)))

    slides.append(SlideData(title="谢谢 / Thank You", kind="thanks"))

    return DeckPlan(title=title, subtitle=subtitle, author=author,
                    theme_name=theme_key, slides=slides)


# ---------------------------------------------------------------------------
# 正文规划
# ---------------------------------------------------------------------------
def _norm_figure(s: str) -> str:
    """规范化图片/标题名用于配对：去掉“图解/图_/图”等前缀与空格标点，转小写。"""
    t = s.replace("图解", "").replace("图_", "").replace("图-", "").replace("图", "")
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", t.lower())


def _extract_figures(blocks: list[Block]) -> tuple[list[Block], list[tuple]]:
    """提取“图解说明”小节（标题含 图解/Figure，后跟要点），从正文中移除并返回说明。

    返回 (剩余 blocks, [(key, 标题, [要点...], 一句话结论)])；key 用于与图片文件名配对。
    """
    figs: list[tuple] = []
    out: list[Block] = []
    i, n = 0, len(blocks)
    while i < n:
        b = blocks[i]
        if (b.kind == "heading"
                and ("图解" in b.text or "figure" in b.text.lower())):
            points: list[str] = []
            takeaway = ""
            j = i + 1
            while j < n and blocks[j].kind != "heading":
                if blocks[j].kind == "bullets":
                    points.extend(blocks[j].bullets)
                elif blocks[j].text:
                    t = blocks[j].text.strip()
                    if t:
                        takeaway = t if not takeaway else takeaway + " " + t
                j += 1
            if points:
                title = re.sub(r"^图解\s*", "", b.text).strip()
                figs.append((_norm_figure(b.text), title, points, takeaway))
                i = j          # 跳过说明内容，避免重复成普通页
                continue
        out.append(b)
        i += 1
    return out, figs


def _plan_content(blocks: list[Block], use_dividers: bool,
                  figure_specs: list[tuple] | None = None) -> list[SlideData]:
    """把 blocks 组织成正文幻灯片；图片优先配对成图文页(figure)，其余进画廊(gallery)。"""
    figure_specs = figure_specs or []
    fig_map = {key: (title, points, takeaway)
               for key, title, points, takeaway in figure_specs}
    slides: list[SlideData] = []
    cur: SlideData | None = None
    gallery: SlideData | None = None

    def flush() -> None:
        nonlocal cur, gallery
        if cur is not None:
            # 空的“要点页”不生成（内容可能已被提取成图文页/图表等）
            if not (cur.kind == "bullets" and not cur.bullets):
                slides.append(_refine_kind(cur))
            cur = None
        if gallery is not None and gallery.images:
            slides.append(gallery)
            gallery = None

    for b in blocks:
        if b.kind == "heading":
            flush()
            if b.level >= 2 and use_dividers:
                # 章节分隔页 + 后续内容沿用该标题
                slides.append(SlideData(title=b.text, kind="section",
                                        notes="章节分隔页"))
                cur = SlideData(title=b.text, kind="bullets", level=b.level)
            else:
                cur = SlideData(title=b.text, kind="bullets", level=b.level)

        elif b.kind in ("bullets", "paragraph"):
            items = b.bullets if b.kind == "bullets" else [b.text]
            if cur is None:
                cur = SlideData(title="要点 Highlights", kind="bullets")
            for it in items:
                it = (it or "").strip()
                if not it:
                    continue
                for piece in _split_long(it):
                    if len(cur.bullets) >= BULLETS_PER_SLIDE:
                        flush()
                        cur = SlideData(title="要点 Highlights", kind="bullets")
                    cur.bullets.append(piece)

        elif b.kind == "table":
            flush()
            if b.chart is not None:
                slides.append(SlideData(
                    title=b.text or "数据一览", kind="chart",
                    table=b.table, chart=b.chart,
                    notes="原生图表 + 数据表，可在 PowerPoint 中直接编辑"))
            else:
                slides.append(SlideData(title=b.text or "数据表", kind="table",
                                        table=b.table))

        elif b.kind == "chart":
            flush()
            slides.append(SlideData(
                title=b.text or "数据图表", kind="chart",
                table=b.table, chart=b.chart,
                notes="原生图表，可在 PowerPoint 中直接编辑"))

        elif b.kind == "image":
            key = _norm_figure(os.path.splitext(os.path.basename(b.image_path))[0])
            if key in fig_map:
                flush()
                title, points, takeaway = fig_map[key]
                slides.append(SlideData(title=title, kind="figure",
                                        images=[b.image_path], bullets=points,
                                        takeaway=takeaway,
                                        notes="图文讲解页"))
            else:
                if gallery is None:
                    gallery = SlideData(title="图解与图片 / Diagrams & Images",
                                        kind="gallery", images=[])
                gallery.images.append(b.image_path)

    flush()
    return slides


# ---------------------------------------------------------------------------
# 标题 / 总结推断
# ---------------------------------------------------------------------------
def _guess_title(blocks: list[Block]) -> str:
    for b in blocks:
        if b.kind == "heading":
            return b.text
    if blocks:
        return os.path.splitext(os.path.basename(blocks[0].source))[0]
    return "演示文稿"


def _has_summary(blocks: list[Block]) -> bool:
    return any(b.kind == "heading"
               and any(k.lower() in b.text.lower() for k in SUMMARY_KEYWORDS)
               for b in blocks)


def _collect_summary(blocks: list[Block]) -> list[str]:
    out: list[str] = []
    for i, b in enumerate(blocks):
        if b.kind == "heading" and any(k.lower() in b.text.lower()
                                       for k in SUMMARY_KEYWORDS):
            for nxt in blocks[i + 1:]:
                if nxt.kind == "heading":
                    break
                if nxt.kind == "bullets":
                    out.extend(nxt.bullets)
                elif nxt.text:
                    out.append(nxt.text)
    if not out:
        out = ["内容已梳理完毕", "关键要点已提炼", "建议下一步行动已列出"]
    return out[:6]


# ---------------------------------------------------------------------------
# LLM 建议合并（可选）
# ---------------------------------------------------------------------------
def apply_suggestion(plan: DeckPlan, suggestion: dict) -> DeckPlan:
    """把 LLM 返回的 {title, subtitle, theme, outline} 合并进规划。失败则忽略。"""
    try:
        if suggestion.get("title"):
            plan.title = suggestion["title"]
        if suggestion.get("subtitle"):
            plan.subtitle = suggestion["subtitle"]
        if suggestion.get("theme") in config.THEMES:
            plan.theme_name = suggestion["theme"]

        outline = suggestion.get("outline") or []
        if not outline:
            return plan

        new_slides: list[SlideData] = [plan.slides[0]]  # 保留封面
        for item in outline:
            title = (item.get("title") or "").strip() or "要点"
            kind = item.get("kind") or "bullets"
            bullets = [str(x) for x in (item.get("bullets") or [])]
            if kind == "section":
                new_slides.append(SlideData(title=title, kind="section"))
            elif kind in ("table", "chart", "gallery"):
                # 数据/图片页仍以解析结果为准
                continue
            else:
                new_slides.append(SlideData(title=title, kind="bullets", bullets=bullets))

        # 追加解析出的数据/图片页
        for s in plan.slides:
            if s.kind in ("table", "chart", "gallery"):
                new_slides.append(s)

        # 结尾
        if not any(s.kind == "summary" for s in new_slides):
            summary = suggestion.get("summary") or []
            if summary:
                new_slides.append(SlideData(title="总结与展望", kind="summary",
                                            bullets=[str(x) for x in summary]))
        new_slides.append(SlideData(title="谢谢 / Thank You", kind="thanks"))
        plan.slides = new_slides
    except Exception:
        pass
    return plan
