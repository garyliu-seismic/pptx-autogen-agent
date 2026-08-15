"""主题引擎：根据素材内容的关键词自动挑选配色主题，也可 --theme 强制指定。"""
from __future__ import annotations

import config
from .models import Block

# 每个主题对应的一组关键词（命中越多得分越高）
_THEME_KEYWORDS: dict[str, list[str]] = {
    "tech": [
        "ai", "人工智能", "软件", "代码", "数据", "开发", "系统", "算法", "云", "程序",
        "tech", "software", "cloud", "api", "server", "database", "机器学习", "app",
        "前端", "后端", "运维", "技术", "产品发布",
    ],
    "business": [
        "销售", "营收", "利润", "市场", "商业", "战略", "增长", "季度", "销售额", "财报",
        "sales", "revenue", "profit", "market", "business", "q1", "q2", "q3", "q4",
        "kpi", "目标", "业务",
    ],
    "finance": [
        "投资", "金融", "银行", "理财", "证券", "基金", "风险", "回报", "资产",
        "finance", "invest", "bank", "stock", "portfolio", "收益",
    ],
    "health": [
        "健康", "医疗", "患者", "疾病", "营养", "医院", "临床", "药物", "护理",
        "health", "medical", "patient", "care", "wellness",
    ],
    "education": [
        "教育", "课程", "学习", "培训", "学生", "教学", "知识", "考试", "课堂", "讲师",
        "education", "training", "course", "lecture", "学习计划",
    ],
    "nature": [
        "环境", "生态", "绿色", "能源", "气候", "环保", "碳", "可持续",
        "nature", "energy", "climate", "sustainability",
    ],
    "elegant": [
        "报告", "方案", "总结", "规划", "纪要", "汇报", "review", "plan", "proposal",
        "年度", "季度汇报", "周会",
    ],
}


def pick_theme(blocks: list[Block], forced: str | None = None) -> config.Theme:
    """根据内容关键词打分选主题；forced 非空时直接使用。"""
    if forced and forced in config.THEMES:
        return config.THEMES[forced]

    # 汇总全部文字
    parts: list[str] = []
    for b in blocks:
        if b.text:
            parts.append(b.text)
        parts.extend(b.bullets)
        if b.table:
            parts.extend(b.table.headers)
            parts.extend(" ".join(r) for r in b.table.rows)
        if b.chart:
            parts.append(" ".join(b.chart.categories))
    text = " ".join(parts).lower()

    scores: dict[str, int] = {}
    for theme, kws in _THEME_KEYWORDS.items():
        s = sum(text.count(k.lower()) for k in kws)
        if s:
            scores[theme] = s

    if not scores:
        return config.THEMES[config.DEFAULT_THEME]
    best = max(scores, key=scores.get)
    return config.THEMES.get(best, config.THEMES[config.DEFAULT_THEME])
