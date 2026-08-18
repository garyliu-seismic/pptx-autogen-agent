"""llm_advisor：.env 加载、chat_json_tools 解析（mock urlopen，不发真实请求）。"""
import json
import os
import urllib.request

import pytest

import src.llm_advisor as la

_KEYS = ("OPENAI_API_KEY", "OPENAI_BASE_URL", "OPENAI_MODEL")


@pytest.fixture(autouse=True)
def _clean_env():
    for k in _KEYS:
        os.environ.pop(k, None)
    yield
    for k in _KEYS:
        os.environ.pop(k, None)


class _FakeResp:
    def __init__(self, payload):
        self._payload = payload

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def read(self):
        return json.dumps(self._payload).encode("utf-8")


def _fake_urlopen(payload):
    return lambda *a, **k: _FakeResp(payload)


def test_dotenv_loads(monkeypatch, tmp_path):
    p = tmp_path / ".env"
    p.write_text(
        'OPENAI_API_KEY="sk-x"\n'
        "OPENAI_BASE_URL=https://api.deepseek.com/v1\n"
        "OPENAI_MODEL=deepseek-chat\n",
        encoding="utf-8")
    la._load_dotenv(str(p))
    assert os.environ["OPENAI_API_KEY"] == "sk-x"
    assert os.environ["OPENAI_BASE_URL"] == "https://api.deepseek.com/v1"
    assert os.environ["OPENAI_MODEL"] == "deepseek-chat"


def test_dotenv_does_not_override_existing(tmp_path):
    os.environ["OPENAI_API_KEY"] = "sk-existing"
    p = tmp_path / ".env"
    p.write_text("OPENAI_API_KEY=sk-fromfile\n", encoding="utf-8")
    la._load_dotenv(str(p))
    assert os.environ["OPENAI_API_KEY"] == "sk-existing"


def test_chat_json_tools_parses_native_tool_calls(monkeypatch):
    os.environ.update(OPENAI_API_KEY="sk-x", OPENAI_BASE_URL="http://x/v1",
                      OPENAI_MODEL="m")
    payload = {"choices": [{"message": {"tool_calls": [
        {"id": "1", "type": "function",
         "function": {"name": "set_title", "arguments": '{"title":"T"}'}}]}}]}
    monkeypatch.setattr(urllib.request, "urlopen", _fake_urlopen(payload))
    calls = la.chat_json_tools("s", "u", [])
    assert calls == [{"tool": "set_title", "args": {"title": "T"}}]


def test_chat_json_tools_parses_text_fallback(monkeypatch):
    os.environ.update(OPENAI_API_KEY="sk-x", OPENAI_BASE_URL="http://x/v1",
                      OPENAI_MODEL="m")
    payload = {"choices": [{"message": {"content": json.dumps(
        {"tool_calls": [{"tool": "add_bullets",
                         "args": {"title": "t", "bullets": ["1"]}}]})}}]}
    monkeypatch.setattr(urllib.request, "urlopen", _fake_urlopen(payload))
    calls = la.chat_json_tools("s", "u", [])
    assert calls == [{"tool": "add_bullets", "args": {"title": "t", "bullets": ["1"]}}]


def test_chat_json_tools_none_without_key(monkeypatch):
    monkeypatch.setattr(urllib.request, "urlopen", lambda *a, **k: 1 / 0)
    assert la.chat_json_tools("s", "u", []) is None   # 无 Key 直接返回 None
