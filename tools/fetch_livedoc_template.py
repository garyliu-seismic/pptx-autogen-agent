"""livedoc-service 模板拉取手测工具（QA/UAT 实验性接入）。

用法：
    # 列出 AI-ready 模板（拿 versionId）
    python tools/fetch_livedoc_template.py list

    # 下载指定模板的原始内容到 output/ 下
    python tools/fetch_livedoc_template.py download <versionId> [--out output]

调用前需设置环境变量：
    LIVEDOC_SERVICE_BASE_URL   例如 https://<qa-tenant>.seismic.com/livedoc-service
    LIVEDOC_SERVICE_TOKEN      已获取的 Bearer token（不含 "Bearer " 前缀）
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.livedoc_client import LivedocClient, LivedocClientError


def list_templates(use_v2: bool) -> int:
    client = LivedocClient()
    templates = client.list_templates(use_v2=use_v2)
    if not templates:
        print("未找到模板。")
        return 0

    for t in templates:
        version_id = t.get("versionId")
        name = t.get("name")
        ai_type = t.get("aiLiveDocType")
        print(f"{version_id}  [{ai_type}]  {name}")
    return 0


def download_template(version_id: str, out_dir: str) -> int:
    client = LivedocClient()
    result = client.download_template_content(version_id)

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, result.file_name)
    with open(out_path, "wb") as f:
        f.write(result.content)

    print(f"已下载: {out_path}  ({len(result.content)} bytes, {result.content_type})")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="livedoc-service 模板拉取手测工具")
    sub = ap.add_subparsers(dest="command", required=True)

    list_parser = sub.add_parser("list", help="列出 AI-ready 模板")
    list_parser.add_argument("--v2", action="store_true", help="使用 get-templates/v2（Slide Layout）")

    download_parser = sub.add_parser("download", help="下载模板原始内容")
    download_parser.add_argument("version_id", help="模板的 versionId")
    download_parser.add_argument("--out", default="output", help="输出目录，默认 output/")

    args = ap.parse_args()

    try:
        if args.command == "list":
            return list_templates(args.v2)
        if args.command == "download":
            return download_template(args.version_id, args.out)
    except LivedocClientError as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
