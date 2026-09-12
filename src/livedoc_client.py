"""livedoc-service 客户端：拉取 AI-ready 模板列表 / 原始文件内容（pptx/pdf 等）。

仅用于 QA/UAT 环境的实验性接入。认证走环境变量传入的已获取 Bearer token，
不在这里实现任何 OAuth / client-credentials 流程。
"""
from __future__ import annotations

import os
from dataclasses import dataclass

import requests

API_ROUTE_PREFIX = "api/ai/v1"


class LivedocClientError(RuntimeError):
    """livedoc-service 调用失败（非 2xx 响应或缺少必要配置）。"""


@dataclass
class TemplateContent:
    version_id: str
    content: bytes
    content_type: str
    file_name: str


def _require_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise LivedocClientError(
            f"缺少环境变量 {name}。请先设置 LIVEDOC_SERVICE_BASE_URL 和 "
            "LIVEDOC_SERVICE_TOKEN（QA/UAT 环境的 base URL 和已获取的 Bearer token）。"
        )
    return value


class LivedocClient:
    """对接 app-livedoc-service 的 AIController（/api/ai/v1/...）。"""

    def __init__(self, base_url: str | None = None, token: str | None = None, timeout: float = 30.0):
        self.base_url = (base_url or _require_env("LIVEDOC_SERVICE_BASE_URL")).rstrip("/")
        self.token = token or _require_env("LIVEDOC_SERVICE_TOKEN")
        self.timeout = timeout

    def _headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self.token}"}

    def list_templates(self, template_types: list[str] | None = None, use_v2: bool = False) -> list[dict]:
        """列出 AI-ready 模板（get-templates / get-templates/v2）。

        返回的每个模板字典可能包含 versionId/contentId/name/description 等，
        用 versionId 传给 download_template_content 拉取原始内容。
        """
        path = "get-templates/v2" if use_v2 else "get-templates"
        url = f"{self.base_url}/{API_ROUTE_PREFIX}/{path}"
        body = {"versionIds": [], "templateTypes": template_types, "resolveValues": False}

        response = requests.post(url, json=body, headers=self._headers(), timeout=self.timeout)
        if not response.ok:
            raise LivedocClientError(
                f"get-templates failed: {response.status_code} {response.text[:500]}"
            )

        payload = response.json()
        if not payload.get("success"):
            raise LivedocClientError(f"get-templates returned success=false: {payload.get('error')}")

        return payload.get("data") or []

    def download_template_content(self, version_id: str) -> TemplateContent:
        """下载模板原始文件（GET /templates/{versionId}/content）。"""
        url = f"{self.base_url}/{API_ROUTE_PREFIX}/templates/{version_id}/content"

        response = requests.get(url, headers=self._headers(), timeout=self.timeout)
        if response.status_code == 404:
            raise LivedocClientError(f"No downloadable source content found for versionId: {version_id}")
        if not response.ok:
            raise LivedocClientError(
                f"download content failed: {response.status_code} {response.text[:500]}"
            )

        content_disposition = response.headers.get("Content-Disposition", "")
        file_name = version_id
        if "filename=" in content_disposition:
            file_name = content_disposition.split("filename=", 1)[1].strip('"; ')

        return TemplateContent(
            version_id=version_id,
            content=response.content,
            content_type=response.headers.get("Content-Type", "application/octet-stream"),
            file_name=file_name,
        )
