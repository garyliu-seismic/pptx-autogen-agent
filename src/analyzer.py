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

import config
from .models import Block, DeckPlan, SlideData
from .theme_engine import pick_theme

BULLETS_PER_SLIDE = 6
SUMMARY_KEYWORDS = ["总结", "结论", "展望", "下一步", "行动项",
                    "summary", "conclusion", "next steps", "action items"]


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------
def build_deck(blocks: list[Block], title: str = "", subtitle: str = "",
               author: str = "", theme_name: str = "") -> DeckPlan:
    theme = pick_theme(blocks, theme_name or None)
    if not title:
        title = _guess_title(blocks)

    slides: list[SlideData] = [SlideData(title=title, kind="title",
                                         notes=f"主题配色：{theme.name}")]

    # 小节分隔页仅在数量适中(2~6)时启用，避免页数臃肿
    sec_heads = [b for b in blocks if b.kind == "heading" and b.level >= 2]
    use_dividers = 2 <= len(sec_heads) <= 6

    content_slides = _plan_content(blocks, use_dividers)

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
                    theme_name=theme.name, slides=slides)


# ---------------------------------------------------------------------------
# 正文规划
# ---------------------------------------------------------------------------
def _plan_content(blocks: list[Block], use_dividers: bool) -> list[SlideData]:
    slides: list[SlideData] = []
    cur: SlideData | None = None

    def flush() -> None:
        nonlocal cur
        if cur is not None:
            slides.append(cur)
            cur = None

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
                if len(cur.bullets) >= BULLETS_PER_SLIDE:
                    flush()
                    cur = SlideData(title="要点 Highlights", kind="bullets")
                cur.bullets.append(it)

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
            if slides and slides[-1].kind == "gallery" and len(slides[-1].images) < 4:
                slides[-1].images.append(b.image_path)
            else:
                flush()
                slides.append(SlideData(title="图片 / Images", kind="gallery",
                                        images=[b.image_path]))

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
