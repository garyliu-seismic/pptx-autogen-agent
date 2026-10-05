"""可选 LLM 大纲规划：调用 OpenAI 兼容 API 提炼核心并给出大纲。

仅当环境变量 OPENAI_API_KEY 存在时才生效；失败时自动回退到启发式分析。
可通过 OPENAI_BASE_URL / OPENAI_MODEL 切换到任意兼容端点（如 DeepSeek、本地 Ollama）。
"""
from __future__ import annotations

import json
import os
import urllib.request

from .models import Block


def _load_dotenv(path: str | None = None) -> None:
    """从项目根目录 .env 读取环境变量（不覆盖已存在的环境变量）。

    支持 `KEY=value` 与带引号的值；忽略空行与 # 注释。
    这样无需改系统环境变量，只需在 .env 里写 OPENAI_API_KEY 等即可。
    """
    if path is None:
        path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            ".env")
    if not os.path.isfile(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, _, v = line.partition("=")
                k = k.strip()
                v = v.strip().strip('"').strip("'")
                if k and k not in os.environ:
                    os.environ[k] = v
    except Exception:
        pass


_load_dotenv()  # 模块导入即加载，保证 available()/chat_* 能看到 .env 配置


def available() -> bool:
    return bool(os.environ.get("OPENAI_API_KEY"))


def _temperature() -> float | None:
    """读取 OPENAI_TEMPERATURE；'none'/'omit' 表示不传 temperature 字段（gpt-5 系列要求）。"""
    raw = os.environ.get("OPENAI_TEMPERATURE", "").strip().lower()
    if raw in ("none", "omit", "default", ""):
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def _build_payload(model: str, messages: list[dict],
                   temperature: float | None = None,
                   tools: list[dict] | None = None) -> dict:
    """构造请求 payload；gpt-5 系列不传 temperature，用 max_completion_tokens 替代 max_tokens。"""
    payload: dict = {"model": model, "messages": messages}
    # gpt-5 / o1 / o3 / o4 系列拒绝 temperature 字段
    model_lower = model.lower()
    is_gpt5_family = any(model_lower.startswith(p) for p in
                         ("gpt-5", "o1", "o3", "o4"))
    env_temp = _temperature()
    if not is_gpt5_family:
        t = env_temp if env_temp is not None else temperature
        if t is not None:
            payload["temperature"] = t
    # max_completion_tokens (gpt-5) vs max_tokens (others)
    mct = os.environ.get("OPENAI_MAX_COMPLETION_TOKENS", "").strip()
    if mct:
        key_name = "max_completion_tokens" if is_gpt5_family else "max_tokens"
        try:
            payload[key_name] = int(mct)
        except ValueError:
            pass
    if tools:
        payload["tools"] = tools
    return payload


def chat_json(system: str, user: str, *, timeout: int = 90,
              temperature: float = 0.4) -> dict | None:
    """调用 OpenAI 兼容端点并要求返回 JSON 对象；失败/超时/无 Key 返回 None。

    兼容 OpenAI / DeepSeek / Azure AI Foundry / 本地 Ollama 等端点。
    gpt-5 系列自动省略 temperature 字段。
    """
    if not available():
        return None
    try:
        base = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
        model = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
        key = os.environ["OPENAI_API_KEY"]
        payload = _build_payload(model,
                                 [{"role": "system", "content": system},
                                  {"role": "user", "content": user}],
                                 temperature=temperature)
        req = urllib.request.Request(
            base.rstrip("/") + "/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {key}"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        content = data["choices"][0]["message"]["content"]
        content = content.strip()
        # 兼容 LLM 把 JSON 包在 ```json ... ``` 里的情况
        if content.startswith("```"):
            content = content.split("```", 2)[1]
            if content.startswith("json"):
                content = content[4:]
        return json.loads(content)
    except Exception:
        return None


def chat_json_tools(system: str, user: str, tools: list[dict], *,
                    timeout: int = 120,
                    temperature: float = 0.4) -> list[dict] | None:
    """发送带 tools schema 的一次对话，返回 [{'tool': name, 'args': dict}] 或 None。

    优先解析原生 function calling 的 message.tool_calls；
    端点不支持时回退解析文本内容里的 JSON（{"tool_calls":[...]} 或 [...]）。
    供 planner 工具调用式规划使用。
    """
    if not available():
        return None
    try:
        base = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
        model = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
        key = os.environ["OPENAI_API_KEY"]
        payload = _build_payload(model,
                                 [{"role": "system", "content": system},
                                  {"role": "user", "content": user}],
                                 temperature=temperature,
                                 tools=tools)
        req = urllib.request.Request(
            base.rstrip("/") + "/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {key}"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        msg = data["choices"][0]["message"]

        # 1) 原生 function calling
        if msg.get("tool_calls"):
            out: list[dict] = []
            for tc in msg["tool_calls"]:
                fn = tc.get("function", {})
                name = fn.get("name", "")
                raw = fn.get("arguments") or "{}"
                try:
                    args = json.loads(raw)
                except Exception:
                    args = {}
                if name:
                    out.append({"tool": name,
                                "args": args if isinstance(args, dict) else {}})
            return out

        # 2) 文本 JSON 兜底（兼容不支持 function calling 的端点）
        content = (msg.get("content") or "").strip()
        if not content:
            return None
        if content.startswith("```"):
            content = content.split("```", 2)[1]
            if content.startswith("json"):
                content = content[4:]
        data = json.loads(content)
        arr = data.get("tool_calls") if isinstance(data, dict) else data
        out = []
        for c in arr if isinstance(arr, list) else []:
            if not isinstance(c, dict):
                continue
            name = c.get("tool") or c.get("name") or ""
            args = c.get("args") or c.get("arguments") or {}
            if name:
                out.append({"tool": name,
                            "args": args if isinstance(args, dict) else {}})
        return out or None
    except Exception:
        return None


def chat_tools(messages: list[dict], tools: list[dict], *,
               timeout: int = 120, temperature: float = 0.4) -> dict | None:
    """发送一轮带 tools 的对话；返回 assistant message dict 或 None。

    这是**多轮 function-calling 协议**的基础：messages 需维护完整历史
    （system/user/assistant(含 tool_calls)/tool(结果)），由调用方负责回填。
    返回的 message 含 `tool_calls`（含 id）或 `content`。
    """
    if not available():
        return None
    try:
        base = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
        model = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
        key = os.environ["OPENAI_API_KEY"]
        payload = _build_payload(model, messages, temperature=temperature, tools=tools)
        req = urllib.request.Request(
            base.rstrip("/") + "/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {key}"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        return data["choices"][0]["message"]
    except Exception:
        return None


def _serialize(blocks: list[Block], max_chars: int = 6000) -> str:
    """把素材压缩成给 LLM 的文本。"""
    parts: list[str] = []
    for b in blocks:
        if b.kind == "heading":
            parts.append(f"[标题] {b.text}")
        elif b.kind == "bullets":
            parts.append("[要点] " + "；".join(b.bullets))
        elif b.kind == "paragraph":
            parts.append(b.text)
        elif b.kind == "table" and b.table:
            hdr = ", ".join(b.table.headers)
            rows = " | ".join("; ".join(r) for r in b.table.rows[:12])
            parts.append(f"[表格] {hdr}\n{rows}")
        elif b.kind == "image":
            parts.append("[图片] " + os.path.basename(b.image_path))
    text = "\n".join(parts)
    return text[:max_chars]


def suggest(blocks: list[Block]) -> dict | None:
    """返回 {title, subtitle, theme, outline, summary} 或 None。"""
    if not available():
        return None
    import config as _cfg
    theme_opts = "|".join(_cfg.THEMES.keys())
    system = (
        "你是一个演示文稿设计规划助手。请根据素材提炼核心内容，并规划一份专业 PPT 大纲。"
        "只输出一个 JSON 对象，不要包含 markdown 代码块。JSON 格式："
        '{"title":"封面标题","subtitle":"副标题",'
        f'"theme":"{theme_opts}",'
        '"outline":[{"title":"页标题","kind":"section|bullets|cards|timeline","bullets":["要点1"]}],'
        '"summary":["总结要点1"]}\n'
        "主题选择原则：深色科技/AI→ai或tech；商务财务→business或finance；"
        "企业报告→corporate；高端演讲→dark_pro；极简风格→minimal；"
        "医疗健康→medical；高端品牌/奖项→luxury；年轻互联网→fresh；"
        "学术研究→academic；政府党政→government。"
    )
    return chat_json(system, _serialize(blocks))
