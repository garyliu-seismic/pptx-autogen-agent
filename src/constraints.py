"""生成约束档（constraints profile）—— 让调用方对 auto gen 结果加更多硬控制。

默认不启用（``--constraints`` 不传时 main.py 各处都拿到 ``None``，行为与之前完全一致）。
四个维度：

  1. 数据真实性 —— ``require_data_source``：图表/表格页必须标注真实来源，否则剔除。
  2. 版式/几何 —— ``hard_fail_on_layout``：质检未消除的严重问题不再「将就」，直接报错。
  3. 品牌/主题 —— ``locked_theme``：锁定主题，自动检测与 LLM 都不能改。
  4. 结构 —— ``required_kinds`` / ``max_kind_counts`` / ``min_slides`` / ``max_slides``。
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field


@dataclass
class Constraints:
    locked_theme: str | None = None
    require_data_source: bool = False
    hard_fail_on_layout: bool = False
    required_kinds: list[str] = field(default_factory=list)
    max_kind_counts: dict[str, int] = field(default_factory=dict)
    min_slides: int | None = None
    max_slides: int | None = None


def load_constraints(path: str | None) -> Constraints | None:
    """从 JSON 文件载入约束档；path 为空返回 None（即不启用任何约束）。"""
    if not path:
        return None
    with open(path, "r", encoding="utf-8") as f:
        d = json.load(f)
    return Constraints(
        locked_theme=d.get("locked_theme") or None,
        require_data_source=bool(d.get("require_data_source", False)),
        hard_fail_on_layout=bool(d.get("hard_fail_on_layout", False)),
        required_kinds=[str(k) for k in (d.get("required_kinds") or [])],
        max_kind_counts={str(k): int(v) for k, v in (d.get("max_kind_counts") or {}).items()},
        min_slides=d.get("min_slides"),
        max_slides=d.get("max_slides"),
    )
