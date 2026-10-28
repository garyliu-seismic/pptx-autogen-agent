"""LLMG 探针（手动运行；import 时不发任何请求）。

读取环境变量 LLMG_BASE_URL / LLMG_TOKEN(或 LLMG_TOKEN_FILE) / LLMG_FEATURE_ID / LLMG_MODEL，
向 {base}/v1/chat/completions 发 5 个最小请求，只打印状态码 + 短片段（不打印令牌）：
  1) 纯文本  2) tools  3) 图片（contents + imageUrl）  4) json schema（responseFormat.jsonSchema）
  5) 与 1 相同但不带 featureId（确认 featureId 是否必填）

用法：python tools/llmg_probe.py
注意：请求会经过 LLMG（SEB topics，留存策略未知），探针仅发送无敏感内容的占位文本/1x1 图片。
"""
from __future__ import annotations

import json
import os
import sys
import urllib.error
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 1x1 透明 PNG
_PNG_B64 = ("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5E"
            "rkJggg==")


def _post(url: str, token: str, payload: dict, timeout: int = 60) -> str:
    req = urllib.request.Request(
        url, data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {token}"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            status, body = r.status, r.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        status, body = e.code, e.read().decode("utf-8", "replace")
    except Exception as e:  # 网络错误等
        return f"ERR {type(e).__name__}"
    return f"{status} {body[:200].replace(chr(10), ' ')}".replace(token, "***")


def main() -> int:
    from src import llm_advisor as la
    if not la.llmg_active():
        print("Set LLMG_BASE_URL and LLMG_TOKEN (or LLMG_TOKEN_FILE) first.")
        return 2
    url, hdrs = la._endpoint()
    token = hdrs["Authorization"].split(" ", 1)[1]
    model = la._model()
    feat = os.environ.get("LLMG_FEATURE_ID", "").strip() or "demo"
    base = {"model": model, "featureId": feat, "maxTokens": 256}
    text = {"messages": [{"role": "user", "content": "Reply with the word: pong"}]}
    tool = {"type": "function", "function": {
        "name": "echo", "description": "Echo a word",
        "parameters": {"type": "object", "properties": {"word": {"type": "string"}},
                       "required": ["word"]}}}
    schema = {"name": "r", "strict": True, "schema": {
        "type": "object", "properties": {"ok": {"type": "boolean"}},
        "required": ["ok"], "additionalProperties": False}}
    cases = [
        ("1 text", {**base, **text}),
        ("2 tools", {**base, "messages": [{"role": "user", "content": "Call echo with word=hi"}],
                     "tools": [tool], "tool_choice": "auto"}),
        ("3 image(contents+imageUrl)", {**base, "messages": [{"role": "user", "contents": [
            {"type": "text", "text": "Describe this image in 3 words."},
            {"type": "image_url", "imageUrl": {
                "url": "data:image/png;base64," + _PNG_B64, "detail": "auto"}}]}]}),
        ("4 json_schema", {**base, "messages": [{"role": "user", "content": "Return JSON ok=true"}],
                           "responseFormat": {"type": "json_schema", "jsonSchema": schema}}),
        ("5 no featureId", {k: v for k, v in {**base, **text}.items() if k != "featureId"}),
    ]
    print(f"POST {url}  model={model} featureId={feat}")
    for name, payload in cases:
        print(f"[{name}] {_post(url, token, payload)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
