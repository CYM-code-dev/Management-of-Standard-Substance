#!/usr/bin/env python3
"""
LIMS API 接口扫描器

扫描代码中对 http://192.168.12.234:60015 的 HTTP 请求，
对比 LIMS_API_openapi.json 中已记录的端点，发现新接口。

用法:
  python scripts/lims_api_scanner.py          # 扫描并输出新端点
  python scripts/lims_api_scanner.py --json    # JSON 格式输出（供 CI 使用）
"""

import json
import os
import re
import sys
import io
from collections import defaultdict

# 修复 Windows 控制台编码
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
OPENAPI_FILE = os.path.join(PROJECT_DIR, "LIMS_API_openapi.json")

# LIMS 服务器特征
LIMS_HOST = "192.168.12.234"
LIMS_PORT = "60015"
LIMS_PATH_PREFIX = "/detectionManager/"

# 正则：匹配 requests/session 的 HTTP 方法调用
# 匹配: session.get(".../detectionManager/..."), requests.post(f"...{base_url}/detectionManager/...")
REQUEST_PATTERN = re.compile(
    r'(?:session|requests|self\.\w*session\w*)\.'
    r'(get|post|put|delete|patch)\s*\('
    r'[^"\']*["\']'
    r'([^"\']*)',
    re.IGNORECASE
)

# 正则：匹配 url 变量赋值中的 detectionManager 路径
URL_ASSIGN_PATTERN = re.compile(
    r'(?:url|endpoint|api_url|req_url)\s*=\s*["\']([^"\']*/detectionManager/[^"\']*)["\']',
    re.IGNORECASE
)


def load_known_endpoints():
    """从 LIMS_API_openapi.json 读取已知端点集合"""
    with open(OPENAPI_FILE, "r", encoding="utf-8") as f:
        spec = json.load(f)

    endpoints = set()
    for path, methods in spec.get("paths", {}).items():
        for method in methods:
            if method.lower() in ("get", "post", "put", "delete", "patch"):
                endpoints.add((method.upper(), path))
    return endpoints


def extract_path(full_url):
    """从完整 URL 或 f-string 中提取 /detectionManager/... 路径部分"""
    # 匹配 /detectionManager/ 开始的路径
    match = re.search(r'(/detectionManager/[^\s"\'`,)}\]]+)', full_url)
    if match:
        path = match.group(1)
        # 移除 query string
        path = path.split("?")[0]
        # 移除尾部斜杠
        path = path.rstrip("/")
        return path
    return None


def scan_py_files():
    """扫描所有 .py 文件，提取对 LIMS 的 HTTP 请求"""
    found = []  # [(method, path, file, line_num)]

    for root, dirs, files in os.walk(PROJECT_DIR):
        # 跳过虚拟环境和缓存
        dirs[:] = [d for d in dirs if d not in (".venv", "__pycache__", "node_modules", ".git")]

        for fname in files:
            if not fname.endswith(".py"):
                continue

            filepath = os.path.join(root, fname)
            rel_path = os.path.relpath(filepath, PROJECT_DIR)

            with open(filepath, "r", encoding="utf-8", errors="replace") as f:
                for line_num, line in enumerate(f, 1):
                    # 跳过注释
                    stripped = line.strip()
                    if stripped.startswith("#"):
                        continue

                    # 方法1: 匹配 session.get/post(..."url"...)
                    for match in REQUEST_PATTERN.finditer(line):
                        method = match.group(1).upper()
                        url_str = match.group(2)
                        path = extract_path(url_str)
                        if path:
                            found.append((method, path, rel_path, line_num))

                    # 方法2: 匹配 url = ".../detectionManager/..."
                    for match in URL_ASSIGN_PATTERN.finditer(line):
                        url_str = match.group(1)
                        path = extract_path(url_str)
                        if path:
                            # 需要找同一区域的 HTTP 方法
                            found.append(("UNKNOWN", path, rel_path, line_num))

    return found


def normalize_path(path):
    """将路径中的动态变量归一化，如 {id} 或 /123 统一为占位符"""
    # f-string 中的 {variable} 已经是占位符
    # 数字 ID 归一化
    path = re.sub(r'/\d+', '/{id}', path)
    return path


def run_scan(output_json=False):
    """执行扫描并返回结果"""
    known = load_known_endpoints()
    found = scan_py_files()

    # 去重并归一化
    seen = {}  # (method, path) -> (file, line)
    new_endpoints = []

    for method, path, rel_file, line_num in found:
        norm_path = normalize_path(path)

        # 尝试匹配已知端点（考虑路径参数差异）
        is_known = False
        for known_method, known_path in known:
            if known_method == method and known_path == norm_path:
                is_known = True
                break
            # 也检查路径模式匹配（如 /record/{id} vs /record/123）
            if method != "UNKNOWN":
                known_pattern = re.sub(r'\{[^}]+\}', '{id}', known_path)
                if known_pattern == norm_path:
                    is_known = True
                    break

        key = (method, norm_path)
        if key not in seen:
            seen[key] = (rel_file, line_num)
            if not is_known:
                new_endpoints.append({
                    "method": method,
                    "path": norm_path,
                    "file": rel_file,
                    "line": line_num,
                })

    if output_json:
        return new_endpoints

    # 人类可读输出
    print(f"{'='*50}")
    print(f"  LIMS 接口扫描报告")
    print(f"{'='*50}")
    print(f"  已知端点: {len(known)} 个")
    print(f"  代码中发现: {len(seen)} 个")
    print(f"  新端点: {len(new_endpoints)} 个")

    if new_endpoints:
        print(f"\n  🔍 发现新端点:")
        for ep in new_endpoints:
            print(f"    - {ep['method']} {ep['path']}")
            print(f"      ({ep['file']}:{ep['line']})")
    else:
        print(f"\n  ✅ 没有发现新端点")

    return new_endpoints


def main():
    output_json = "--json" in sys.argv
    result = run_scan(output_json=output_json)
    if output_json:
        print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
