# PPTX 自动生成 Agent

根据指定目录里的提示素材（`txt / csv / pdf / png`），**自主提炼核心内容、决定主题与版式**，自动生成 **PowerPoint 可直接打开、可直接编辑** 的 `.pptx`。

## 功能特性

- **多格式输入**：`txt`（含 markdown 标题/要点）、`csv`（自动识别数值列）、`pdf`（文字 + 表格 + 内嵌图片）、`png/jpg`（图片，可选 OCR）
- **内容提炼**：归一化所有素材 → 自动分页、提炼要点、识别总结/结论小节
- **自主决策**
  - 主题：按内容关键词自动匹配配色（科技蓝 / 商务金 / 健康绿 / 教育橙 / 自然青 / 金融墨 / 简约灰 / 现代靛），也可 `--theme` 强制指定
  - 版式：封面 → 目录 → 章节页 → 要点页 / 表格页 / 原生图表页 / 图片画廊 → 总结页 → 感谢页
- **可编辑输出**：文本、表格、图表全部是 PowerPoint **原生对象**（图表双击即可改数据），16:9 版式
- **可选 LLM 增强**：设置 `OPENAI_API_KEY` 后可用 `--llm` 让大模型规划大纲（兼容 OpenAI / DeepSeek / 本地 Ollama 等兼容端点）

## 安装

```bash
pip install -r requirements.txt
```

## 使用

```bash
# 1) 最简用法：把素材放到 input 目录，自动生成
python main.py --input input -o output/deck.pptx

# 2) 指定标题/副标题/作者/主题
python main.py --input examples/sample_data -o output/demo.pptx \
    --title "季度经营汇报" --subtitle "2026 Q3" --author "张三" --theme business

# 3) 使用 LLM 规划大纲（需先设置 OPENAI_API_KEY）
set OPENAI_API_KEY=sk-xxx
python main.py --input input -o output/deck.pptx --llm

# 4) 工具调用式规划器（更强，需同时 --llm；兼容 DeepSeek/Ollama）
set OPENAI_API_KEY=sk-xxx
python main.py --input input -o output/deck.pptx --llm --planner
```

> 也可以切换到其它兼容端点，例如 DeepSeek：
> ```bash
> set OPENAI_API_KEY=sk-deepseek-xxx
> set OPENAI_BASE_URL=https://api.deepseek.com/v1
> set OPENAI_MODEL=deepseek-chat
> python main.py --input input -o output/deck.pptx --llm
> ```

### LLMG 模式（可选，Seismic LLM Gateway）

同时设置以下环境变量（或写入 `.env`，已被 gitignore）即启用；不设置则行为不变：

| 变量 | 说明 |
|---|---|
| `LLMG_BASE_URL` | 网关基址，请求发往 `{base}/v1/chat/completions`。QA 形式：`https://gateway-qa-az-westus-wombat.seismic-dev.com/llmg`（来自笔记，未验证） |
| `LLMG_TOKEN` | 委托令牌（scope `aiml_llm`，约 30 分钟过期，需手动获取）；也可用 `LLMG_TOKEN_FILE` 指向存放令牌的文件。令牌不会写入日志 |
| `LLMG_FEATURE_ID` | 每个请求携带的 `featureId`，默认 `demo` |
| `LLMG_MODEL` | LLMG 模型 id，如 `gpt-5-mini-2025-08-07`、`gpt-5.4-2026-03-05` |

- 请求体使用 LLMG 形状：`featureId`、`maxTokens`、多模态 `contents[]`+`imageUrl`、`responseFormat.jsonSchema`；响应经归一化后交给原有逻辑。
- 429/503 最多重试 2 次（优先 `Retry-After`，否则 1s/2s 退避）；非 2xx 会在 stderr 打一行 `[llmg] HTTP <状态> ...`，随后仍回退到启发式。
- 可手动运行 `python tools/llmg_probe.py` 探测 QA（文本/tools/图片/json schema/无 featureId，仅打印状态与短片段）。
- 合规提醒：提示词与图片会经过 LLMG（SEB topics，留存策略未知），请先用非敏感样例。

## 质量闭环（P0）

- **渲染前规划质检（默认开启）**：校验页数/要点条数/要点长度/空壳页/图表数据/图片存在性/主题名，
  自动做确定性修复（拆分超长/超量要点、删空壳页、清理失效图片）。
  `--no-plan-qa` 可关闭。
- **渲染后智能修复（`--llm` 时启用）**：确定性参数调节仍无法消除溢出/重叠时，
  把具体排版问题回灌 LLM，让它只对问题页重新规划（拆分/精简/换页型）后重新渲染，最多 2 轮。
  无 Key 时自动降级为原有确定性调节，行为不变。
- **读回内容 QA（默认开启）**：渲染后逐页核对「计划 vs 实际」内容（空壳页/缺表格/缺图表/缺图片），
  自动剔除空壳/残缺页并重新渲染；`--no-readback` 关闭。
- **Dry-run 计划确认（P2-E）**：`--plan-only` 输出 `output/plan.json` 供预览/编辑后退出，
  `--plan plan.json` 载入确认过的计划再渲染（跳过 LLM 规划）。
- **主题选择修复**：`build_deck` 现按主题 key（如 `business`）存储，自动选中的配色不再被静默丢弃。

## 工具调用式规划器（P1）

- `--llm --planner`：LLM 通过 **11 个页型工具**（`set_title`/`set_theme`/`add_section`/
  `add_bullets`/`add_cards`/`add_timeline`/`add_table`/`add_chart`/`add_summary`/
  `add_thanks`/`done`）逐页构建大纲，替代一次性 JSON 大纲。
- 系统提示注入画布几何、每页约束、主题色板与 few-shot 示例（Phase F 上下文构建器）。
- 原生 function calling 优先，端点不支持时回退文本 JSON；无 Key/失败自动回退启发式 + `--llm` 大纲。
- 规划器含 `web_search` 工具：素材信息不足时 LLM 可联网补充背景事实（结果带来源标注）。
  后端**百度百科优先、维基百科备选**（国内网络可用）；都失败自动降级。

## 示例

```bash
python scripts/make_samples.py      # 生成示例素材（txt/csv 已内置，png/pdf 由脚本生成）
python main.py --input examples/sample_data -o output/demo.pptx
```

打开 `output/demo.pptx` 即可看到自动生成的完整演示文稿。

## 测试

```bash
pip install -r requirements-dev.txt
python -m pytest        # 34 个用例，覆盖校验/修复/规划/渲染/序列化/读回
```

> 测试全部为本地逻辑 + mock（不发真实网络/LLM 请求），无需 `OPENAI_API_KEY`。

## 目录结构

```
pptx-autogen-agent/
├── main.py                 # 命令行入口
├── config.py               # 默认路径 + 主题预设
├── requirements.txt
├── src/
│   ├── parsers.py          # txt/csv/pdf/png -> 内容块
│   ├── analyzer.py         # 内容分析 + 版面规划
│   ├── theme_engine.py     # 关键词 -> 主题
│   ├── renderer.py         # python-pptx 渲染
│   └── llm_advisor.py      # 可选 LLM 大纲
├── scripts/make_samples.py # 生成示例素材
├── examples/sample_data/   # 示例素材
└── output/                 # 生成结果
```

## 支持的素材与识别规则

| 素材 | 识别内容 |
| ---- | -------- |
| `txt` | `#` 标题、`-/*/•` 要点、`1.` 编号、普通段落、轻量标题启发式 |
| `csv` | 表格；含数值列时自动生成原生图表 |
| `pdf` | 文字、表格、内嵌图片（PyMuPDF 抽取） |
| `png/jpg` | 图片直接入页；安装 `pytesseract` 后可 OCR 提取文字 |

## 常见问题

- **打不开？** 生成的是标准 OOXML `.pptx`，PowerPoint 2007+ / WPS 均可直接打开编辑。
- **想换主题？** `python main.py --help` 查看 `--theme` 可选值。
- **想用本地模型？** 设置 `OPENAI_BASE_URL` 指向本地 OpenAI 兼容服务（如 Ollama）即可。
- **中文字体**：默认使用「微软雅黑」，可在 `config.py` 或渲染器中调整。
