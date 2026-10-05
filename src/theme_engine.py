"""主题引擎：根据素材内容的关键词自动挑选配色主题，也可 --theme 强制指定。"""
from __future__ import annotations

import config
from .models import Block

# 每个主题对应的一组关键词（命中越多得分越高）
# 顺序：高识别度关键词在前，通用词在后
_THEME_KEYWORDS: dict[str, list[str]] = {
    "ai": [
        "大模型", "llm", "gpt", "claude", "gemini", "deepseek", "人工智能", "神经网络",
        "transformer", "embedding", "推理", "fine-tune", "rag", "向量数据库",
        "ai", "机器学习", "深度学习", "自然语言", "computer vision", "多模态",
    ],
    "tech": [
        "软件", "代码", "开发", "系统", "算法", "云", "程序", "api", "server",
        "database", "前端", "后端", "运维", "技术", "产品发布", "架构", "微服务",
        "tech", "software", "cloud", "数字化", "平台",
    ],
    "business": [
        "销售", "营收", "利润", "市场", "商业", "战略", "增长", "季度", "销售额", "财报",
        "sales", "revenue", "profit", "market", "business", "q1", "q2", "q3", "q4",
        "kpi", "业务", "客户", "渠道", "转化率", "roi",
    ],
    "finance": [
        "投资", "金融", "银行", "理财", "证券", "基金", "风险", "回报", "资产",
        "finance", "invest", "bank", "stock", "portfolio", "收益", "融资", "估值",
        "ipo", "资本", "杠杆", "对冲",
    ],
    "health": [
        "健康", "医疗", "患者", "疾病", "营养", "医院", "临床", "药物", "护理",
        "health", "medical", "patient", "care", "wellness", "公共卫生", "生命科学",
    ],
    "education": [
        "教育", "课程", "学习", "培训", "学生", "教学", "知识", "考试", "课堂", "讲师",
        "education", "training", "course", "lecture", "学习计划", "技能", "认证",
    ],
    "nature": [
        "环境", "生态", "绿色", "能源", "气候", "环保", "碳", "可持续", "新能源",
        "nature", "energy", "climate", "sustainability", "光伏", "风电", "esg",
    ],
    "corporate": [
        "企业", "集团", "公司治理", "合规", "审计", "董事会", "股东", "年报",
        "corporate", "governance", "annual report", "上市公司", "政府", "机构",
    ],
    "dark_pro": [
        "高端", "旗舰", "发布会", "峰会", "keynote", "演讲", "论坛",
    ],
    "warm": [
        "零售", "消费", "品牌", "活动", "节日", "营销", "推广", "电商",
        "retail", "brand", "campaign", "promotion",
    ],
    "elegant": [
        "报告", "方案", "总结", "规划", "纪要", "汇报", "review", "plan", "proposal",
        "年度", "季度汇报", "周会", "工作总结",
    ],
    "minimal": [
        "极简", "设计", "创意", "portfolio", "作品集", "白皮书", "whitepaper",
    ],
    "modern": [
        "产品", "互联网", "app", "用户", "增长", "运营", "社区", "平台",
        "startup", "创业", "独角兽",
    ],
}

# 标题/一级标题的关键词权重倍率（比正文重要）
_HEADING_WEIGHT = 3


def pick_theme(blocks: list[Block], forced: str | None = None) -> config.Theme:
    """根据内容关键词打分选主题；forced 非空时直接使用。"""
    if forced and forced in config.THEMES:
        return config.THEMES[forced]

    # 分别收集标题文字和正文文字，标题给更高权重
    heading_text = " ".join(
        b.text for b in blocks if b.kind == "heading" and b.level <= 2
    ).lower()
    body_parts: list[str] = []
    for b in blocks:
        if b.kind not in ("heading",) and b.text:
            body_parts.append(b.text)
        body_parts.extend(b.bullets)
        if b.table:
            body_parts.extend(b.table.headers)
            body_parts.extend(" ".join(str(c) for c in r) for r in b.table.rows)
        if b.chart:
            body_parts.append(" ".join(b.chart.categories))
    body_text = " ".join(body_parts).lower()

    scores: dict[str, float] = {}
    for theme, kws in _THEME_KEYWORDS.items():
        s = 0.0
        for k in kws:
            k_low = k.lower()
            s += heading_text.count(k_low) * _HEADING_WEIGHT
            s += body_text.count(k_low)
        if s > 0:
            scores[theme] = s

    if not scores:
        return config.THEMES[config.DEFAULT_THEME]
    best = max(scores, key=scores.__getitem__)
    return config.THEMES.get(best, config.THEMES[config.DEFAULT_THEME])
