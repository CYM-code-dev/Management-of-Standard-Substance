#!/usr/bin/env python3
"""
Apifox LIMS API 文档同步工具

用法:
  python scripts/apifox_sync.py analyze   # 导出 Apifox 数据并分析重复
  python scripts/apifox_sync.py sync      # 用 LIMS_API_openapi.json 同步到 Apifox（覆盖+清理）
"""

import json
import os
import sys
import requests
import io

# 修复 Windows 控制台编码
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
OPENAPI_FILE = os.path.join(PROJECT_DIR, "LIMS_API_openapi.json")

APIFOX_API_BASE = "https://api.apifox.com"
PROJECT_ID = "8238274"
ACCESS_TOKEN = "afxp_d213b4FGYKwP7aSUGNZmEeXHfEzXZF9yOHRR"

HEADERS = {
    "Authorization": f"Bearer {ACCESS_TOKEN}",
    "X-Apifox-Api-Version": "2024-03-28",
    "Content-Type": "application/json",
}


def load_local_openapi():
    """读取本地 LIMS_API_openapi.json"""
    with open(OPENAPI_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def get_local_endpoints(spec):
    """从 OpenAPI spec 中提取所有端点 (method, path)"""
    endpoints = []
    for path, methods in spec.get("paths", {}).items():
        for method in methods:
            if method.lower() in ("get", "post", "put", "delete", "patch"):
                endpoints.append((method.upper(), path))
    return endpoints


def export_apifox():
    """从 Apifox 导出当前项目的 OpenAPI 数据"""
    url = f"{APIFOX_API_BASE}/v1/projects/{PROJECT_ID}/export-openapi"
    params = {"locale": "zh-CN"}
    payload = {
        "version": "3.0",
        "format": "json",
    }
    resp = requests.post(url, headers=HEADERS, params=params, json=payload, timeout=30)
    resp.raise_for_status()
    result = resp.json()
    # 导出 API 直接返回 OpenAPI JSON，也可能包在 {"data": ...} 中
    if "openapi" in result:
        return result
    elif result.get("data") and isinstance(result["data"], dict) and "openapi" in result["data"]:
        return result["data"]
    else:
        print(f"❌ 导出失败: {json.dumps(result, ensure_ascii=False)[:500]}")
        return None


def import_to_apifox(openapi_data):
    """将 OpenAPI 数据导入到 Apifox（覆盖模式，清理多余接口）"""
    url = f"{APIFOX_API_BASE}/v1/projects/{PROJECT_ID}/import-openapi"
    params = {"locale": "zh-CN"}
    payload = {
        "input": {
            "data": json.dumps(openapi_data, ensure_ascii=False)
        },
        "options": {
            "endpointOverwriteBehavior": "OVERWRITE",
            "schemaOverwriteBehavior": "OVERWRITE",
            "updateFolderOfChangedEndpoint": True,
        },
    }
    resp = requests.post(url, headers=HEADERS, params=params, json=payload, timeout=60)
    resp.raise_for_status()
    return resp.json()


def cmd_analyze():
    """分析 Apifox 项目中的接口，与本地 JSON 对比"""
    print("📥 正在从 Apifox 导出数据...")
    apifox_data = export_apifox()
    if not apifox_data:
        return

    # 导出数据直接就是 OpenAPI spec
    apifox_spec = apifox_data

    local_spec = load_local_openapi()

    local_endpoints = get_local_endpoints(local_spec)
    apifox_endpoints = get_local_endpoints(apifox_spec)

    local_set = set(local_endpoints)
    apifox_set = set(apifox_endpoints)

    # 统计 Apifox 中的重复（同 method+path 在不同目录下可能出现）
    from collections import Counter
    apifox_counter = Counter(apifox_endpoints)
    duplicates = {ep: count for ep, count in apifox_counter.items() if count > 1}

    # 差异分析
    only_in_apifox = apifox_set - local_set
    only_in_local = local_set - apifox_set

    print(f"\n{'='*50}")
    print(f"  Apifox 项目接口分析报告")
    print(f"{'='*50}")
    print(f"  本地 JSON 端点数: {len(local_set)}")
    print(f"  Apifox 端点总数:  {len(apifox_endpoints)} (去重后 {len(apifox_set)})")

    if duplicates:
        print(f"\n  ⚠️  重复接口 ({len(duplicates)} 个):")
        for (method, path), count in sorted(duplicates.items()):
            print(f"    - {method} {path} (×{count})")

    if only_in_apifox:
        print(f"\n  📋 仅 Apifox 中有 ({len(only_in_apifox)} 个):")
        for method, path in sorted(only_in_apifox):
            print(f"    - {method} {path}")

    if only_in_local:
        print(f"\n  📄 仅本地 JSON 中有 ({len(only_in_local)} 个):")
        for method, path in sorted(only_in_local):
            print(f"    - {method} {path}")

    common = local_set & apifox_set
    print(f"\n  ✅ 已匹配的接口: {len(common)} 个")

    if not duplicates and not only_in_apifox and not only_in_local:
        print(f"\n  🎉 完美匹配，无重复！")
    else:
        print(f"\n  💡 建议: 运行 'python scripts/apifox_sync.py sync' 进行同步清理")


def cmd_sync():
    """将本地 LIMS_API_openapi.json 同步到 Apifox"""
    spec = load_local_openapi()
    endpoints = get_local_endpoints(spec)
    print(f"📂 本地 JSON 包含 {len(endpoints)} 个端点")
    print(f"📤 正在同步到 Apifox 项目 {PROJECT_ID}...")

    result = import_to_apifox(spec)

    if result.get("code") and result.get("code") != 0:
        print(f"❌ 同步失败: {result}")
        return

    counters = result.get("data", {}).get("counters", {})
    print(f"\n{'='*40}")
    print(f"  同步结果")
    print(f"{'='*40}")
    print(f"  接口创建: {counters.get('endpointCreated', 0)}")
    print(f"  接口更新: {counters.get('endpointUpdated', 0)}")
    print(f"  接口失败: {counters.get('endpointFailed', 0)}")
    print(f"  接口忽略: {counters.get('endpointIgnored', 0)}")

    failed = counters.get("endpointFailed", 0)
    if failed == 0:
        print(f"\n  ✅ 同步完成，无失败项！")
    else:
        print(f"\n  ⚠️  有 {failed} 个接口同步失败，请检查")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    cmd = sys.argv[1].lower()
    if cmd == "analyze":
        cmd_analyze()
    elif cmd == "sync":
        cmd_sync()
    else:
        print(f"未知命令: {cmd}")
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
