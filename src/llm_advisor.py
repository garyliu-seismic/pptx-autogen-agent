"""可选 LLM 大纲规划：调用 OpenAI 兼容 API 提炼核心并给出大纲。

仅当环境变量 OPENAI_API_KEY 存在时才生效；失败时自动回退到启发式分析。
可通过 OPENAI_BASE_URL / OPENAI_MODEL 切换到任意兼容端点（如 DeepSeek、本地 Ollama）。
"""
from __future__ import annotations

import json
import os
import urllib.request

from .models import Block


def available() -> bool:
    return bool(os.environ.get("OPENAI_API_KEY"))


def chat_json(system: str, user: str, *, timeout: int = 90,
              temperature: float = 0.4) -> dict | None:
    """调用 OpenAI 兼容端点并要求返回 JSON 对象；失败/超时/无 Key 返回 None。

    兼容 OpenAI / DeepSeek / 本地 Ollama 等端点（OPENAI_BASE_URL / OPENAI_MODEL）。
    供 llm_advisor.suggest 与 llm_repair 复用。
    """
    if not available():
        return None
    try:
        base = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
        model = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
        key = os.environ["OPENAI_API_KEY"]
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            "temperature": temperature,
        }
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
    system = (
        "你是一个演示文稿规划助手。请根据用户提供的素材，提炼核心内容并规划一份 PPT 大纲。"
        "只输出一个 JSON 对象，不要输出任何其它文字或 markdown。JSON 格式："
        '{"title":"封面标题","subtitle":"副标题","theme":"tech|business|finance|health|education|nature|elegant|modern",'
        '"outline":[{"title":"页标题","kind":"section|bullets","bullets":["要点1","要点2"]}],'
        '"summary":["总结要点1","总结要点2"]}'
    )
    return chat_json(system, _serialize(blocks))
