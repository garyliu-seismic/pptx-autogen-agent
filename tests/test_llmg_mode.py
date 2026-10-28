"""LLMG 模式：payload 形状、默认路径回归、429 重试、令牌不泄露、响应归一化（mock urlopen，无网络）。"""
import io
import json
import os
import urllib.error
import urllib.request

import pytest

import src.llm_advisor as la

_KEYS = ("OPENAI_API_KEY", "OPENAI_BASE_URL", "OPENAI_MODEL", "OPENAI_MAX_COMPLETION_TOKENS",
         "LLMG_BASE_URL", "LLMG_TOKEN", "LLMG_TOKEN_FILE", "LLMG_FEATURE_ID", "LLMG_MODEL")
TOKEN = "tok-SECRET-123"


@pytest.fixture(autouse=True)
def _env(monkeypatch):
    for k in _KEYS:
        monkeypatch.delenv(k, raising=False)
    monkeypatch.setattr(la, "_sleep", lambda s: sleeps.append(s))
    sleeps.clear()


sleeps: list = []


def _llmg(monkeypatch):
    monkeypatch.setenv("LLMG_BASE_URL", "https://gw.example/llmg/")
    monkeypatch.setenv("LLMG_TOKEN", TOKEN)
    monkeypatch.setenv("LLMG_MODEL", "gpt-5.4-2026-03-05")


class _Resp:
    def __init__(self, payload):
        self._b = json.dumps(payload).encode()

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def read(self):
        return self._b


def _http_error(code, body=b'{"key":"busy"}', retry_after=None):
    hdrs = {"Retry-After": retry_after} if retry_after is not None else {}
    return urllib.error.HTTPError("http://x", code, "err", hdrs, io.BytesIO(body))


def _capture(monkeypatch, outcomes):
    """outcomes: 依次返回/抛出；记录每次 Request。"""
    seen = []

    def fake(req, timeout=None):
        seen.append(req)
        o = outcomes[min(len(seen) - 1, len(outcomes) - 1)]
        if isinstance(o, Exception):
            raise o
        return _Resp(o)

    monkeypatch.setattr(urllib.request, "urlopen", fake)
    return seen


OK = {"choices": [{"message": {"content": '{"a": 1}'}}]}


def test_llmg_inactive_by_default_and_needs_both(monkeypatch):
    assert not la.llmg_active()
    monkeypatch.setenv("LLMG_BASE_URL", "https://gw")
    assert not la.llmg_active()
    monkeypatch.setenv("LLMG_TOKEN", TOKEN)
    assert la.llmg_active() and la.available()


def test_token_file(monkeypatch, tmp_path):
    f = tmp_path / "t.txt"
    f.write_text(TOKEN + "\n", encoding="utf-8")
    monkeypatch.setenv("LLMG_BASE_URL", "https://gw")
    monkeypatch.setenv("LLMG_TOKEN_FILE", str(f))
    assert la.llmg_active() and la._llmg_token() == TOKEN


def test_llmg_payload_shape_text(monkeypatch):
    _llmg(monkeypatch)
    monkeypatch.setenv("LLMG_FEATURE_ID", "pptx-agent")
    monkeypatch.setenv("OPENAI_MAX_COMPLETION_TOKENS", "512")
    seen = _capture(monkeypatch, [OK])
    assert la.chat_json("sys", "usr") == {"a": 1}
    req = seen[0]
    assert req.full_url == "https://gw.example/llmg/v1/chat/completions"
    assert req.get_header("Authorization") == f"Bearer {TOKEN}"
    p = json.loads(req.data)
    assert p["model"] == "gpt-5.4-2026-03-05" and p["featureId"] == "pptx-agent"
    assert p["maxTokens"] == 512
    assert "max_completion_tokens" not in p and "max_tokens" not in p and "temperature" not in p


def test_llmg_feature_id_default(monkeypatch):
    _llmg(monkeypatch)
    seen = _capture(monkeypatch, [OK])
    la.chat_json("s", "u")
    assert json.loads(seen[0].data)["featureId"] == "demo"


def test_llmg_vision_contents_and_json_schema(monkeypatch, tmp_path):
    _llmg(monkeypatch)
    img = tmp_path / "a.png"
    img.write_bytes(b"\x89PNG-fake")
    seen = _capture(monkeypatch, [OK])
    schema = {"type": "object", "properties": {"a": {"type": "integer"}}}
    assert la.chat_vision_json("s", "u", str(img), schema=schema, schema_name="r") == {"a": 1}
    p = json.loads(seen[0].data)
    user = p["messages"][1]
    assert "content" not in user
    c = user["contents"]
    assert c[0] == {"type": "text", "text": "u"}
    assert c[1]["type"] == "image_url"
    assert c[1]["imageUrl"]["url"].startswith("data:image/png;base64,")
    assert p["responseFormat"] == {"type": "json_schema",
                                   "jsonSchema": {"name": "r", "strict": True, "schema": schema}}
    assert "response_format" not in p


def test_llmg_tools_passthrough_and_normalizer(monkeypatch):
    _llmg(monkeypatch)
    raw = {"Choices": [{"Message": {"Content": None, "ToolCalls": [
        {"Id": "c1", "Function": {"Name": "add_slide", "Arguments": '{"x": 1}'}}]},
        "FinishReason": "tool_calls"}], "warnings": ["deprecated"]}
    seen = _capture(monkeypatch, [raw])
    tools = [{"type": "function", "function": {"name": "add_slide", "parameters": {}}}]
    calls = la.chat_json_tools("s", "u", tools)
    assert calls == [{"tool": "add_slide", "args": {"x": 1}}]
    assert json.loads(seen[0].data)["tools"] == tools
    msg = la.chat_tools([{"role": "user", "content": "u"}], tools)
    assert msg["tool_calls"][0]["id"] == "c1"
    assert msg["tool_calls"][0]["function"]["arguments"] == '{"x": 1}'


def test_normalizer_text_and_bad_shape():
    out = la._normalize_llmg_response({"choices": [{"message": {"content": "hi"},
                                                    "finishReason": "stop"}]})
    assert out["choices"][0]["message"]["content"] == "hi"
    assert out["choices"][0]["finish_reason"] == "stop"
    with pytest.raises(ValueError):
        la._normalize_llmg_response({"nope": 1})


def test_default_openai_path_unchanged(monkeypatch, tmp_path):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    monkeypatch.setenv("OPENAI_BASE_URL", "https://api.example/v1")
    monkeypatch.setenv("OPENAI_MODEL", "gpt-5-mini")
    monkeypatch.setenv("OPENAI_MAX_COMPLETION_TOKENS", "100")
    img = tmp_path / "a.png"
    img.write_bytes(b"x")
    seen = _capture(monkeypatch, [OK])
    la.chat_json("s", "u")
    la.chat_vision_json("s", "u", str(img), schema={"type": "object"})
    r1, r2 = seen
    assert r1.full_url == "https://api.example/v1/chat/completions"
    assert r1.get_header("Authorization") == "Bearer sk-test"
    p1, p2 = json.loads(r1.data), json.loads(r2.data)
    assert p1["max_completion_tokens"] == 100
    assert not {"featureId", "maxTokens", "responseFormat"} & set(p1)
    assert p2["messages"][1]["content"][1]["image_url"]["url"].startswith("data:image/png")
    assert p2["response_format"]["json_schema"]["strict"] is True
    assert "contents" not in p2["messages"][1] and "featureId" not in p2


def test_openai_path_no_retry_no_log(monkeypatch, capsys):
    monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
    seen = _capture(monkeypatch, [_http_error(429)])
    assert la.chat_json("s", "u") is None
    assert len(seen) == 1 and not sleeps
    assert capsys.readouterr().err == ""


def test_retry_429_then_success(monkeypatch):
    _llmg(monkeypatch)
    seen = _capture(monkeypatch, [_http_error(429, retry_after=""), _http_error(503), OK])
    assert la.chat_json("s", "u") == {"a": 1}
    assert len(seen) == 3 and sleeps == [1.0, 2.0]


def test_retry_after_header_honored(monkeypatch):
    _llmg(monkeypatch)
    _capture(monkeypatch, [_http_error(429, retry_after="7"), OK])
    assert la.chat_json("s", "u") == {"a": 1}
    assert sleeps == [7.0]


def test_give_up_after_two_retries_logs_without_token(monkeypatch, capsys):
    _llmg(monkeypatch)
    body = ('{"key":"busy","echo":"%s"}' % TOKEN).encode()
    seen = _capture(monkeypatch, [_http_error(429, body=body)])
    assert la.chat_json("s", "u") is None
    assert len(seen) == 3 and len(sleeps) == 2
    err = capsys.readouterr().err
    assert "[llmg] HTTP 429" in err and "busy" in err
    assert TOKEN not in err


def test_non_retryable_error_logged_once_and_exception_text_clean(monkeypatch, capsys):
    _llmg(monkeypatch)
    seen = _capture(monkeypatch, [_http_error(401, body=b'{"error":"bad token"}')])
    assert la.chat_json("s", "u") is None
    assert len(seen) == 1 and not sleeps
    assert "HTTP 401" in capsys.readouterr().err
    with pytest.raises(RuntimeError) as ei:
        la._post_chat({"model": "m"}, 5)
    assert TOKEN not in str(ei.value)
