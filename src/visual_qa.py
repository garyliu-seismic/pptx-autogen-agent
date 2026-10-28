"""视觉 QA（opt-in，--visual-qa）：缩略图 -> 多模态 LLM 判定 pass/fail -> 理由回灌修复器。

移植自 Seismic app-livedoc-service AIController 的 Visual QA 阶段：
  - 只看「严重版式缺陷」（文字溢出/互相重叠/被裁切），不评价内容；
  - 严格 JSON 判定 {passed, issues, suggestions}（json_schema strict）；
  - 失败时把 issues/suggestions 回灌给修复器（quality.render_with_quality 的 repairer），
    并有轮数上限。LLM/渲染任何一步失败都视为「通过」（不阻塞出稿）。
"""
from __future__ import annotations

from . import llm_advisor

VISUAL_QA_SYSTEM = (
    "You are a slide layout inspector. Examine the rendered slide image for layout "
    "problems ONLY — do NOT evaluate content quality, wording, or relevance.\n\n"
    "Check only for severe layout defects such as:\n"
    "- Text overflowing its bounding box\n"
    "- Text elements overlapping each other\n"
    "- Text clipped or cut off at the edge of a shape\n\n"
    "Set passed=false ONLY when one of the above severe defects is clearly visible. "
    "Ignore everything else.\n"
    "When passed=false, suggestions must instruct how to shorten the text to fix the "
    "overflow (e.g. use fewer words, shorter bullets).\n\n"
    "Return ONLY a JSON object matching the provided schema."
)

# strict：全部字段 required + additionalProperties=false
VERDICT_SCHEMA = {
    "type": "object",
    "properties": {
        "passed": {"type": "boolean"},
        "issues": {"type": "array", "items": {"type": "string"}},
        "suggestions": {"type": "string"},
    },
    "required": ["passed", "issues", "suggestions"],
    "additionalProperties": False,
}


def parse_verdict(data) -> dict | None:
    """校验并规整判定；非法结构返回 None（调用方按「通过」处理）。"""
    if not isinstance(data, dict) or not isinstance(data.get("passed"), bool):
        return None
    issues = data.get("issues")
    return {"passed": data["passed"],
            "issues": [str(i) for i in issues] if isinstance(issues, list) else [],
            "suggestions": str(data.get("suggestions") or "")}


def inspect_slide(image_path: str, slide_title: str = "") -> dict | None:
    data = llm_advisor.chat_vision_json(
        VISUAL_QA_SYSTEM, f"Slide title: {slide_title}", image_path,
        schema=VERDICT_SCHEMA, schema_name="visual_qa_verdict")
    return parse_verdict(data)


def default_thumbnailer(pptx_path: str) -> list[str]:
    """pptx -> 每页 PNG（LibreOffice + PyMuPDF，复用 tools.render_preview）；失败返回 []。"""
    try:
        from tools.render_preview import render_pipeline
        pages, _ = render_pipeline(pptx_path)
        return list(pages)
    except Exception:
        return []


def make_visual_checker(thumbnailer=None, inspector=None):
    """返回 checker(pptx_path, plan) -> list[quality.Issue]（kind="visual"）。

    thumbnailer / inspector 可注入，便于测试与替换后端。
    每个失败页一条 Issue，detail 含 LLM 给出的问题与缩短建议。
    """
    from .quality import Issue
    thumbnailer = thumbnailer or default_thumbnailer
    inspector = inspector or inspect_slide

    def check(pptx_path: str, plan) -> list:
        issues = []
        for idx, img in enumerate(thumbnailer(pptx_path), 1):
            title = plan.slides[idx - 1].title if idx <= len(plan.slides) else ""
            v = inspector(img, title)
            if v and not v["passed"]:
                issues.append(Issue(idx, "visual", 1.0,
                                    "; ".join(v["issues"]) + " | 建议: " + v["suggestions"]))
        return issues

    return check
