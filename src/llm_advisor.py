"""可选 LLM 大纲规划：调用 OpenAI 兼容 API 提炼核心并给出大纲。

仅当环境变量 OPENAI_API_KEY 存在时才生效；失败时自动回退到启发式分析。
可通过 OPENAI_BASE_URL / OPENAI_MODEL 切换到任意兼容端点（如 DeepSeek、本地 Ollama）。

可选 LLMG 模式（Seismic LLM Gateway）：同时设置 LLMG_BASE_URL 与 LLMG_TOKEN（或 LLMG_TOKEN_FILE）
时启用，请求改走 {LLMG_BASE_URL}/v1/chat/completions（featureId / maxTokens / contents+imageUrl /
responseFormat.jsonSchema），并对 429/503 做有限重试；未设置时行为与之前完全一致。
"""
from __future__ import annotations

import json
import os
import sys
import time
import urllib.error
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


def _llmg_token() -> str:
    """读取 LLMG 令牌：LLMG_TOKEN 优先，其次 LLMG_TOKEN_FILE 指向的文件（不落日志）。"""
    tok = os.environ.get("LLMG_TOKEN", "").strip()
    if tok:
        return tok
    path = os.environ.get("LLMG_TOKEN_FILE", "").strip()
    if path:
        try:
            with open(path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except OSError:
            return ""
    return ""


def llmg_active() -> bool:
    """LLMG 模式：LLMG_BASE_URL 与令牌都存在时启用，否则走原 OpenAI 兼容路径。"""
    return bool(os.environ.get("LLMG_BASE_URL", "").strip() and _llmg_token())


def available() -> bool:
    return llmg_active() or bool(os.environ.get("OPENAI_API_KEY"))


def _model() -> str:
    if llmg_active():
        return os.environ.get("LLMG_MODEL", "").strip() or "gpt-5-mini-2025-08-07"
    return os.environ.get("OPENAI_MODEL", "gpt-4o-mini")


def _endpoint() -> tuple[str, dict]:
    """返回 (url, headers)。LLMG 模式用 {base}/v1/chat/completions + 委托令牌。"""
    if llmg_active():
        url = os.environ["LLMG_BASE_URL"].strip().rstrip("/") + "/v1/chat/completions"
        token = _llmg_token()
    else:
        base = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
        url = base.rstrip("/") + "/chat/completions"
        token = os.environ["OPENAI_API_KEY"]
    return url, {"Content-Type": "application/json",
                 "Authorization": f"Bearer {token}"}


_RETRY_STATUS = (429, 503)
_MAX_RETRIES = 2
_sleep = time.sleep  # 测试可替换，避免真实等待


def _retry_delay(headers, attempt: int) -> float:
    """优先 Retry-After（秒；可能为空/非数字），否则 1s、2s 指数退避；上限 30s。"""
    try:
        ra = (headers.get("Retry-After") or "").strip() if headers else ""
        if ra:
            return min(max(float(ra), 0.0), 30.0)
    except Exception:
        pass
    return float(2 ** attempt)


def _redact(text: str) -> str:
    tok = _llmg_token()
    return text.replace(tok, "***") if tok else text


def _post_chat(payload: dict, timeout: int) -> dict:
    """POST 并返回响应 JSON；失败抛异常（调用方吞掉并回退）。

    非 LLMG 模式与原实现一致（单次 urlopen）。LLMG 模式：429/503 最多重试 2 次；
    非 2xx 往 stderr 打一行（状态码 + 短 body 片段，不含令牌）后抛出；响应经 normalizer。
    """
    url, headers = _endpoint()
    body = json.dumps(payload).encode("utf-8")
    if not llmg_active():
        req = urllib.request.Request(url, data=body, headers=headers)
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    attempt = 0
    while True:
        req = urllib.request.Request(url, data=body, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return _normalize_llmg_response(json.loads(resp.read().decode("utf-8")))
        except urllib.error.HTTPError as e:
            if e.code in _RETRY_STATUS and attempt < _MAX_RETRIES:
                _sleep(_retry_delay(getattr(e, "headers", None), attempt))
                attempt += 1
                continue
            try:
                snippet = e.read().decode("utf-8", "replace")[:200]
            except Exception:
                snippet = ""
            snippet = _redact(snippet).replace("\n", " ")
            print(f"[llmg] HTTP {e.code} (retries={attempt}): {snippet}", file=sys.stderr)
            raise RuntimeError(f"LLMG HTTP {e.code}") from None
        except Exception as e:
            print(f"[llmg] request failed: {type(e).__name__}", file=sys.stderr)
            raise


def _key(obj: dict, name: str):
    """大小写/下划线不敏感取值（LLMG 可能返回 camelCase 或 PascalCase）。"""
    want = name.replace("_", "").lower()
    for k, v in obj.items():
        if k.replace("_", "").lower() == want:
            return v
    return None


def _normalize_llmg_response(raw: dict) -> dict:
    """LLMG 响应 -> OpenAI 形状 {choices:[{message:{content,tool_calls}}]}（仿 LlmgResponseNormalizer）。"""
    choices = _key(raw, "choices")
    choice = choices[0] if isinstance(choices, list) and choices else choices
    msg = _key(choice, "message") if isinstance(choice, dict) else None
    if not isinstance(msg, dict):
        raise ValueError("LLMG response has no choices[0].message")
    content = _key(msg, "content")
    out: dict = {"role": "assistant",
                 "content": content if isinstance(content, str) else None}
    calls = _key(msg, "toolcalls")
    tcs = []
    for c in calls if isinstance(calls, list) else []:
        if not isinstance(c, dict):
            continue
        fn = _key(c, "function")
        fn = fn if isinstance(fn, dict) else {}
        args = _key(fn, "arguments")
        tcs.append({"id": _key(c, "id"), "type": "function",
                    "function": {"name": _key(fn, "name"),
                                 "arguments": args if isinstance(args, str) else "{}"}})
    if tcs:
        out["tool_calls"] = tcs
    return {"choices": [{"index": 0, "message": out,
                         "finish_reason": _key(choice, "finishreason")}]}


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
    if llmg_active():
        # LLMG：featureId 必填；最大 token 字段为 camelCase maxTokens；不传 temperature
        payload["featureId"] = os.environ.get("LLMG_FEATURE_ID", "").strip() or "demo"
        mct = os.environ.get("OPENAI_MAX_COMPLETION_TOKENS", "").strip()
        if mct:
            try:
                payload["maxTokens"] = int(mct)
            except ValueError:
                pass
        if tools:
            payload["tools"] = tools
        return payload
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
        model = _model()
        payload = _build_payload(model,
                                 [{"role": "system", "content": system},
                                  {"role": "user", "content": user}],
                                 temperature=temperature)
        data = _post_chat(payload, timeout)
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


def chat_vision_json(system: str, user: str, image_path: str, *,
                     schema: dict | None = None, schema_name: str = "result",
                     timeout: int = 120) -> dict | None:
    """多模态：一张图 + 文本，返回 JSON 对象；失败/无 Key 返回 None。

    传入 schema 时用 response_format=json_schema(strict) 约束输出；
    端点不支持时调用方会拿到 None 并降级（不阻塞主流程）。
    """
    if not available():
        return None
    try:
        import base64
        with open(image_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("ascii")
        ext = os.path.splitext(image_path)[1].lower().lstrip(".") or "png"
        mime = "jpeg" if ext in ("jpg", "jpeg") else ext
        model = _model()
        data_url = f"data:image/{mime};base64,{b64}"
        if llmg_active():
            # LLMG DTO：多模态走独立的 contents[]（content 为 string），imageUrl={url,detail}
            user_msg = {"role": "user", "contents": [
                {"type": "text", "text": user},
                {"type": "image_url",
                 "imageUrl": {"url": data_url, "detail": "auto"}}]}
        else:
            user_msg = {"role": "user", "content": [
                {"type": "text", "text": user},
                {"type": "image_url", "image_url": {"url": data_url}}]}
        payload = _build_payload(model, [
            {"role": "system", "content": system}, user_msg], temperature=0)
        fmt = ({"type": "json_schema",
                "json_schema": {"name": schema_name, "strict": True, "schema": schema}}
               if schema else {"type": "json_object"})
        if llmg_active():
            # LLMG ResponseFormat={type, jsonSchema}；jsonSchema 原样透传为上游 json_schema 对象
            payload["responseFormat"] = {"type": fmt["type"]}
            if schema:
                payload["responseFormat"]["jsonSchema"] = fmt["json_schema"]
        else:
            payload["response_format"] = fmt
        data = _post_chat(payload, timeout)
        return json.loads(data["choices"][0]["message"]["content"])
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
        model = _model()
        payload = _build_payload(model,
                                 [{"role": "system", "content": system},
                                  {"role": "user", "content": user}],
                                 temperature=temperature,
                                 tools=tools)
        data = _post_chat(payload, timeout)
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
        model = _model()
        payload = _build_payload(model, messages, temperature=temperature, tools=tools)
        data = _post_chat(payload, timeout)
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
