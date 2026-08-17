# 改进思路：借鉴 ppt-addin-claude-code（ppt master）重构 pptx-autogen-agent

> 日期：2026-08-17
> 参考仓库：`C:\Users\garyl\ppt-addin-claude-code`（Office.js 加载体，AI 实时操控 PowerPoint，分支 `personal/gary/AddTaskPane`）
> 目标仓库：`C:\Users\garyl\pptx-autogen-agent`（离线 Python CLI，python-pptx 批量生成）

## 实施进度

| 阶段 | 状态 | 说明 |
| --- | --- | --- |
| **B 规划前校验** | ✅ 2026-08-17 已实现 | `src/validate_plan.py`：`validate_plan` / `repair_deterministic`；`src/llm_repair.py`：`repair_plan_with_llm` |
| **C 渲染后智能修复闭环** | ✅ 2026-08-17 已实现 | `src/quality.py::render_with_quality` 新增 `repairer` 慢路径；`src/llm_repair.py::repair_render_with_llm` |
| A 工具调用式 planner | ⬜ 未开始 | 见 §3-A |
| D 读回内容 QA | ⬜ 未开始 | 见 §3-D |
| E dry-run 计划确认 | ⬜ 未开始 | 见 §3-E |
| F 上下文构建器 + few-shot | ⬜ 未开始 | 见 §3-F |
| G web 搜索 | ⬜ 未开始 | 见 §3-G |

> 附带修复：`analyzer.build_deck` 原把主题**显示名**（如「商务金」）存入 `DeckPlan.theme_name`，
> 而 `renderer` 用**主题 key**（如 `business`）索引 `config.THEMES` → 自动选中的主题一直被静默丢弃回退到 `modern`。
> 已改为存主题 key（`_theme_key`），并顺带修复 `main.py` 的 `IndentationError`（空 `if issues:` 块）。

---

## 1. 两个项目定位对比

| 维度 | ppt master（ppt-addin-claude-code） | 我们的 pptx-autogen-agent |
| --- | --- | --- |
| 运行形态 | 浏览器 Taskpane，跑在 PowerPoint 内 | 离线 CLI，一次性生成 .pptx |
| 操控方式 | Office.js 实时 API（25+ 工具） | python-pptx 离线渲染 |
| AI 角色 | **工具调用式 Agent**（dry-run → 确认 → 批量执行） | 一次性 JSON 大纲建议（`llm_advisor.suggest`） |
| 校验 | **执行前校验 + 执行后 QA + 修复闭环**（`validateService.ts`） | 渲染后**确定性参数盲调**（`quality.py`） |
| 内容来源 | 素材 + `web_search` 工具 | 仅本地素材（txt/csv/pdf/png） |
| 用户确认 | dry-run 计划预览后一键 Apply | 无（直接出整份 PPT） |
| 格式保真 | Office.js 原生对象，格式由 PowerPoint 保真 | python-pptx 原生对象，渲染器自管格式 |

**核心差距**：master 是「**Agent 心智 + 校验修复闭环**」，我们目前是「**流水线 + 盲调**」。master 的每一条工程决策，都能以「离线批处理」的形态落回我们的流水线，且不破坏「原生可编辑对象」的既有优势。

---

## 2. ppt master 值得借鉴的设计点（按价值排序）

### 2.1 工具调用式规划（function calling / JSON Schema）
master 把能力拆成 25+ 个带 JSON Schema 的工具（`add_shape` / `add_card` / `add_table` / `add_chart` / `apply_theme` …），LLM 通过结构化 tool_calls 规划，天然可校验、可重试。
- 我们现状：`llm_advisor.py` 让 LLM 一次性吐一个 outline JSON，解析脆弱、无法约束字段、失败即整段回退。
- **借鉴**：把「页型」抽象成工具，LLM 逐页规划。

### 2.2 校验 + 修复闭环（最大的单点收益）
master 的 `validateService.ts`：
- 执行前：bounds 越界 / 必填字段 / 颜色格式 / 同批 overlap 预测
- 失败 → 错误注入对话 → LLM 自动修正 → 重新调用（repair loop）
- 执行后：`readBackQA` 读回 shape 数量核对
- 我们现状：`quality.py` 检测到问题后只做「缩小字号 / 加留白 / 压间距」的**盲调**，最多 4 次；解决不了就放弃，从不把具体问题告诉 LLM。

### 2.3 复合原子工具（`add_card`）
背景+标题+副标题一次调用生成，AI 想漏文字都不行。
- 我们已有对应概念（`cards` / `timeline` 页型），但那是渲染器内部的事；应把「一个内容块 → 一个页型」作为 LLM 规划的基本单元，保证不产出空壳页。

### 2.4 上下文构建器（context builder）
master 把当前页形状位置 / 画布几何 / 布局规则喂给 LLM，AI 才可能给出坐标正确的调用。
- 我们现状：`_serialize(blocks)` 只给素材原文，LLM 对「每页 ≤6 条 / 每条 ≤40 字 / 13.333×7.5 英寸 / 8 种主题色板」一无所知 → 规划经常不可渲染。
- **借鉴**：prompt 注入画布几何、页型清单、约束规则 + few-shot 示例。

### 2.5 Dry-run → 确认 → 执行
master 先给计划、用户确认再批量执行，避免大改造成本。
- 离线场景等价物：`--plan-only` 输出 `outline.json` 供人工预览，`--plan plan.json` 载入确认过的计划再渲染。

### 2.6 Web 搜索（内容增强）
master 的 `web_search` 支持 CoinGecko / Wikipedia（search API→summary）/ DuckDuckGo，素材不足时用事实补内容。
- 离线场景：可选 `--web`，规划前用 Wikipedia 中文搜索 API 补背景事实。

### 2.7 空页 / 残缺页清理
master 执行后自动删除「只含 0-1 个 shape 的新建 slide」，保护已有 slide。
- 离线场景等价物：读回 QA 时检测「内容元素数 < 预期」的页，删除或重排。

### 2.8 提示词精简 + 内容优先
master 把 40+ 条冲突规则改成 10 条聚焦规则，内容规则放第 1 条；temperature 0.1→0.5、max_tokens 2000→8192。
- 我们现状：`llm_advisor.py` 的 system prompt 只有 3 行，规则不够；对长素材无 token 余量。

---

## 3. 落地映射（分阶段，均可独立交付、不破坏现有启发式）

> 原则：**无 `OPENAI_API_KEY` 时全部优雅降级到现有启发式流水线**（现在是、改完还是）。

### Phase A — 规划层升级：`llm_advisor` → 工具调用式 `planner`
- 新增 `src/tools.py`：定义**页型工具** + JSON Schema + 必填/约束校验：
  - `add_title_slide(title, subtitle, author)`
  - `add_section_slide(title, subtitle)`
  - `add_bullets_slide(title, bullets≤6, takeaway?)`
  - `add_cards_slide(title, cards[{heading, body}])`
  - `add_table_slide(title, headers, rows)`
  - `add_chart_slide(title, chart_type, categories, series)`
  - `add_timeline_slide(title, milestones[{phase, text}])`
  - `add_gallery_slide(title, images)`
  - `add_summary_slide(title, points)`
  - `set_theme(theme_name)`
- 新增 `src/planner.py`：多轮 function-calling 循环，逐步构建 `DeckPlan`（替代 `suggest()` 单次 JSON）；记录每页原始素材引用以便 QA。
- `llm_advisor.suggest` 保留作降级/兜底。

### Phase B — 规划前校验：`validate_plan.py` ✅ 已实现
- 在渲染前校验 `DeckPlan`：
  - 每页 bullets ≤ 6、每条 ≤ 40 字（超长自动分页/截断标记）
  - chart 数据非空、数值列合法、image 文件存在
  - 颜色/主题名在预设集合内
  - 页数上限（如 ≤ 40）
- 不合格项 → 结构化错误回灌 planner → LLM 修复（对应 master repair loop 的**离线版本**）。

> 实现：`src/validate_plan.py`（校验 + 确定性修复）+ `src/llm_repair.py::repair_plan_with_llm`（LLM 修复，无 Key 时返回 None 降级）。

### Phase C — 渲染后智能修复闭环：`quality.py` 升级 ✅ 已实现
- 保留现有确定性参数调节（快路径，font/hpad/gap）
- **新增慢路径**：`max_tries` 用完后，把 `analyze()` 的具体 `Issue`（页号、shape、`overflow_x`/`overlap`/`text_overflow` 数值、涉及文本）序列化给 LLM → 让 LLM 只对问题页重新规划（拆分过密页 / 删减要点 / 换页型）→ 合并回 plan → 重新渲染，最多再加 N 轮。
- 对应 master 的「错误注入 → LLM 自动修正」。

> 实现：`src/quality.py::render_with_quality(..., repairer=...)` 慢路径 + `src/llm_repair.py::repair_render_with_llm`。
> 由 `main.py --llm` 自动启用；无 Key 时 `repairer` 不传，行为与旧版完全一致。

### Phase D — 读回内容 QA：`readback_qa.py`
- 渲染完成后重新打开 pptx，核对每页内容元素数 ≥ 计划预期（`readBackQA` 的离线版）。
- 「空壳页 / 残缺页」：删除或按原素材重排。

### Phase E — Dry-run / 计划确认
- `main.py` 新增 `--plan-only`（输出 `outline.json`）与 `--plan plan.json`（载入已确认计划）。
- 对应 master「dry-run → 用户确认 → 执行」。

### Phase F — 上下文构建器 + few-shot
- `planner` 的 system prompt 注入：画布 13.333×7.5in、页型清单、每页约束、主题色板、2~3 个 few-shot（封面 / 要点页 / 三卡片页的 JSON 示例）。
- 显著提升 LLM 产出「可渲染计划」的首轮命中率，减少 Phase B/C 修复次数。

### Phase G — Web 搜索（可选，`--web`）
- 新增 `web_search` 工具（Wikipedia 中文 search API → summary，兼容 master 实现）；素材过薄时让 LLM 检索补事实，输出 `[来源]` 标注。

---

## 4. 优先级与工作量预估

| 优先级 | 阶段 | 说明 | 预计规模 |
| --- | --- | --- | --- |
| P0 | **C 智能修复闭环** | 单点收益最大：把「盲调」变「LLM 修复」 | 中（改 `quality.py` + 新增 1 模块） |
| P0 | **B 规划前校验** | 成本最低、可独立交付 | 小（新增 1 模块） |
| P1 | **A 工具调用式 planner** | 架构核心，影响 B/C/F | 大（新增 2 模块，改 `llm_advisor`） |
| P1 | **F 上下文构建器 + few-shot** | 提升 A 的准确率，与 A 联调 | 中 |
| P2 | **D 读回内容 QA** | 需配合 C | 小 |
| P2 | **E dry-run 计划确认** | 独立 CLI 开关 | 小 |
| P3 | **G web 搜索** | 依赖 A 的工具框架 | 中 |

**建议实施顺序**：B → C → A → F → D → E → G（先补校验与修复闭环，再动规划层，避免一次重构过大）。

---

## 5. 风险与对策

| 风险 | 对策 |
| --- | --- |
| DeepSeek/Ollama 的 function calling 兼容性不一 | 兼容 OpenAI 风格 tool_calls；不支持时回退 `suggest()` 单次 JSON |
| 多轮循环增加延迟与 token 成本 | 修复循环限定轮数（如 2 轮）；仅对问题页重规划，不复述全 deck |
| LLM 规划不可渲染（越界/超长） | Phase B 校验 + few-shot 约束 + Phase C 兜底 |
| 无 Key / 网络失败 | 全部阶段优雅降级到现有启发式流水线 |
| 与现有 `render_with_quality` 兼容 | 保持 `quality.py` 接口不变，慢路径作为新参数/新入口加入 |

---

## 6. 关键复用参考（master 侧）

- `src/services/validateService.ts` → `validateBounds` / `validateRequired` / `validateColors` / `validateBatch` / `readBackQA`（Phase B/C/D 的规则可直接移植语义）
- `src/services/aiService.ts` → `TOOLS` 定义 / dry-run 循环 / 错误注入修复 / 对话裁剪（Phase A/C 的循环骨架）
- `src/services/shapeService.ts` → `addCard` 复合原子工具（Phase A 的页型原子化思路）
- `src/services/slideBuilderService.ts` → `buildProfessionalSlide` 专业页模板（后续可做「专业样式模板库」）

> 本项目侧现状：`src/quality.py::render_with_quality`（盲调循环）、`src/llm_advisor.py::suggest`（单次 JSON）、`src/analyzer.py::build_deck`（启发式规划）、`src/renderer.py::render`（工具执行器雏形）。
