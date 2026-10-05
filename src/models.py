"""统一数据模型：把各种来源的素材归一化成 Block / DeckPlan，供分析器与渲染器使用。"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


# ---------------------------------------------------------------------------
# 数据单元
# ---------------------------------------------------------------------------
@dataclass
class TableData:
    """二维表格数据（可编辑表格）。"""
    headers: list[str] = field(default_factory=list)
    rows: list[list] = field(default_factory=list)
    source_ref: str = ""    # 数据来源标注（素材文件名/livedoc 模板 id 等），空表示未标注

    def to_chart(self) -> Optional["ChartData"]:
        """尝试把数值型表格转成图表数据；若无可转列返回 None。"""
        try:
            headers = [str(h).strip() for h in self.headers]
            rows = [[str(c).strip() for c in r] for r in self.rows if r]
            if not headers or not rows:
                return None
            categories = [r[0] for r in rows if r]
            numeric_cols: list[tuple[str, list[float]]] = []
            for j in range(1, len(headers)):
                vals: list[float] = []
                for r in rows:
                    try:
                        cleaned = (str(r[j]) if j < len(r) else "")
                        cleaned = (cleaned.replace(",", "").replace("%", "")
                                   .replace("￥", "").replace("$", "").strip())
                        vals.append(float(cleaned))
                    except (ValueError, IndexError):
                        vals.append(0.0)
                numeric_cols.append((headers[j], vals))
            if not numeric_cols:
                return None
            # 列数少 -> 柱状图，列数多 -> 折线图
            ctype = "column" if len(numeric_cols) <= 3 else "line"
            series = [{"name": name, "values": vals} for name, vals in numeric_cols]
            return ChartData(chart_type=ctype, categories=categories, series=series,
                             source_ref=self.source_ref)
        except Exception:
            return None


@dataclass
class ChartData:
    """原生图表数据（PowerPoint 原生图表，可编辑）。"""
    chart_type: str = "column"                 # column / bar / line / pie
    categories: list[str] = field(default_factory=list)
    series: list[dict] = field(default_factory=list)  # [{name, values}]
    title: str = ""
    source_ref: str = ""    # 数据来源标注（素材文件名/livedoc 模板 id 等），空表示未标注


@dataclass
class Block:
    """解析出的最小内容单元。"""
    source: str = ""                           # 来源文件
    kind: str = "paragraph"                    # heading/paragraph/bullets/table/chart/image
    text: str = ""
    bullets: list[str] = field(default_factory=list)
    table: Optional[TableData] = None
    chart: Optional[ChartData] = None
    image_path: str = ""
    level: int = 0                             # 标题层级（markdown # 数量）


# ---------------------------------------------------------------------------
# 版面规划结果
# ---------------------------------------------------------------------------
@dataclass
class MetricData:
    """KPI 指标卡数据（用于 kpi 页型）。"""
    label: str = ""                # 指标名称，如 "营收增长"
    value: str = ""                # 数值字符串，如 "23%" / "1.2亿"
    trend: str = "flat"            # 趋势：up / down / flat
    note: str = ""                 # 补充说明，如 "同比上季度"
    color: str = ""                # 指定颜色（空=自动按 trend 着色）


@dataclass
class SlideData:
    """单页幻灯片的内容描述。"""
    title: str = ""
    kind: str = "bullets"
    # kind 合法值：
    #   结构页: title / section / agenda / thanks
    #   内容页: bullets / cards / timeline / table / chart
    #           gallery / image / figure / photo_timeline / summary
    #   新增页型: kpi / quote / process / comparison
    bullets: list[str] = field(default_factory=list)
    table: Optional[TableData] = None
    chart: Optional[ChartData] = None
    images: list[str] = field(default_factory=list)
    notes: str = ""
    takeaway: str = ""     # 一句话结论（图文页右下角 / quote 页引语内容）
    level: int = 0
    source: str = ""       # 数据来源标注（渲染为页脚小字，如"各模型官方报告"）
    # ── 新增页型专用字段 ──────────────────────────────────────────────────
    metrics: list[MetricData] = field(default_factory=list)  # kpi 页专用
    author: str = ""       # quote 页署名 / comparison 页左栏标签
    label_left: str = ""   # comparison 页左侧标签（如 "现状" / "Before"）
    label_right: str = ""  # comparison 页右侧标签（如 "目标" / "After"）
    bullets_right: list[str] = field(default_factory=list)  # comparison 页右栏要点


@dataclass
class DeckPlan:
    """整个演示文稿的规划。"""
    title: str = ""
    subtitle: str = ""
    author: str = ""
    theme_name: str = "modern"
    slides: list[SlideData] = field(default_factory=list)
