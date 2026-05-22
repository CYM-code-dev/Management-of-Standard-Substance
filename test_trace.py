"""
独立溯源测试脚本 — 与主程序隔离
用法: python test_trace.py <配置编号或LIMS_ID或solutionCode> [配置编号或LIMS_ID或solutionCode ...]

示例:
  python test_trace.py CK-CG-2025102-1.00-20260508   # solutionCode 精准搜索（推荐）
  python test_trace.py B-3377                     # configureOrder 全量匹配（慢）
  python test_trace.py 23564                      # 数字 ID 直接查
  python test_trace.py D-8559 23320              # 多个目标
"""

import sys
import re
import os
import json
import time
import requests
import hashlib
from io import BytesIO

# Fix Windows console encoding
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if sys.stderr.encoding != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# 禁用 SSL 警告
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from PIL import Image

BASE_URL = "http://192.168.12.234:60015"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE_URL,
    "Referer": f"{BASE_URL}/web/solutionConfigure.html?menuId=544",
}


def md5_1024_times(text: str) -> str:
    current = text.encode('utf-8')
    for _ in range(1024):
        current = hashlib.md5(current).hexdigest().encode('utf-8')
    return current.decode()


def get_session():
    """加载已有 session 或引导登录"""
    session = requests.Session()
    session.verify = False
    session.headers.update(HEADERS)

    # 尝试加载已保存的 session
    for f in sorted(os.listdir('.')):
        if f.startswith('session_') and f.endswith('.json'):
            try:
                with open(f, 'r', encoding='utf-8') as fh:
                    info = json.load(fh)
                login_time = info.get("login_time", "")
                if login_time:
                    elapsed = time.time() - time.mktime(time.strptime(login_time, "%Y-%m-%d %H:%M:%S"))
                    if elapsed > 24 * 3600:
                        continue
                for name, value in info["cookies"].items():
                    session.cookies.set(name, value)
                if "headers" in info:
                    session.headers.update(info["headers"])
                print(f"  已加载 session: {f} (用户: {info['username']})")
                return session
            except Exception:
                continue

    # 没有可用 session，引导登录
    print("未找到有效 session，请登录 LIMS...")
    username = input("  用户名: ").strip()
    password = input("  密码: ").strip()

    # 获取验证码
    temp_session = requests.Session()
    temp_session.verify = False
    temp_session.headers.update(HEADERS)
    ts = str(int(time.time() * 1000))
    url = f"{BASE_URL}/detectionManager/core/security/validatecodes?{ts}&r={ts}"
    resp = temp_session.get(url)
    if resp.status_code != 200:
        print("获取验证码失败")
        sys.exit(1)
    img = Image.open(BytesIO(resp.content))
    img.show()
    captcha = input("  验证码: ").strip()

    pwd_hash = md5_1024_times(password)
    login_data = {"account": username, "password": pwd_hash, "validCode": captcha}
    login_resp = temp_session.post(f"{BASE_URL}/detectionManager/core/security/login", data=login_data)
    if login_resp.status_code != 200:
        print(f"登录请求失败: {login_resp.status_code}")
        sys.exit(1)
    result = login_resp.json()
    if not result.get("success"):
        print(f"登录失败: {result.get('errorCtx', {}).get('errorMsg', '未知错误')}")
        sys.exit(1)

    # 把登录 cookies 转到主 session
    for cookie in temp_session.cookies:
        session.cookies.set(cookie.name, cookie.value)

    print(f"  登录成功: {username}")
    return session


# ==================== 改进版溯源函数 ====================

_type_list_cache = {}  # {sol_type: [items]}
_order_to_id_cache = {}  # {configure_order: lims_id}
_consumable_cache = {}  # {configure_order: record_dict}  A类缓存
_a_solution_code_hints = {}  # {configure_order: solutionCode}  A类搜索提示


def fetch_detail(session, solution_id):
    """通过 LIMS 数字 ID 获取记录（旧接口，saveDetailList 为 null）"""
    url = f"{BASE_URL}/detectionManager/manager/dtSolutionConfigure/detail"
    headers = {"Referer": f"{BASE_URL}/web/solutionConfigure.html?menuId=544"}
    resp = session.get(url, params={"id": solution_id}, headers=headers)
    resp.raise_for_status()
    data = resp.json()
    return data.get('resultData') or {}


def fetch_view(session, solution_id, order_str=None):
    """通过 viewDtSolutionConfigure 获取完整记录（含 detailList）
    type 参数需匹配记录类型，否则 detailList 可能为空。
    如果 order_str 未提供，先尝试 SOLUTION_TYPE_D，再用 SOLUTION_TYPE_E。
    """
    url = f"{BASE_URL}/detectionManager/manager/dtSolutionConfigure/viewDtSolutionConfigure"
    headers = {"Referer": f"{BASE_URL}/web/solutionConfigure.html?menuId=544"}
    view_type = "SOLUTION_TYPE_E"
    if order_str:
        prefix = order_str.split('-')[0].upper()
        type_map = {'D': 'SOLUTION_TYPE_D', 'C': 'SOLUTION_TYPE_C', 'B': 'SOLUTION_TYPE_E', 'E': 'SOLUTION_TYPE_E'}
        view_type = type_map.get(prefix, 'SOLUTION_TYPE_E')
    params = {"id": solution_id, "type": view_type}
    resp = session.get(url, params=params, headers=headers)
    resp.raise_for_status()
    data = resp.json()
    result = data.get('resultData') or {}
    # order_str 未提供时，如果 detailList 为空，尝试 D 类型重试
    if not order_str and view_type != 'SOLUTION_TYPE_D' and not result.get('detailList'):
        params2 = {"id": solution_id, "type": "SOLUTION_TYPE_D"}
        resp2 = session.get(url, params=params2, headers=headers)
        resp2.raise_for_status()
        data2 = resp2.json()
        result2 = data2.get('resultData') or {}
        if result2.get('detailList'):
            result = result2
    return result


def get_session_auth():
    """从 session 文件读取 pid/pname/loginId"""
    for f in sorted(os.listdir('.')):
        if f.startswith('session_') and f.endswith('.json'):
            try:
                with open(f, 'r', encoding='utf-8') as fh:
                    info = json.load(fh)
                return {
                    'pid': info.get('pid', ''),
                    'pname': info.get('real_name', info.get('username', '')),
                    'loginId': info.get('pid', ''),
                }
            except Exception:
                continue
    return {'pid': '', 'pname': '', 'loginId': ''}


def fetch_consumable(session, keyword='', status='normal', auth=None):
    """查询消耗品记录（A 类标准品）"""
    if not auth:
        auth = get_session_auth()
    url = f"{BASE_URL}/detectionManager/manager/consumableBill/pageObj"
    headers = {"Referer": f"{BASE_URL}/web/solutionConfigure.html?menuId=544"}
    params = {
        "_search": "false",
        "nd": str(int(time.time() * 1000)),
        "pageSize": 200,
        "pageNo": 1,
        "sidx": "",
        "sord": "asc",
        "type": "CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE",
        "orgName": "",
        "groupId": "",
        "status": status,
        "keyword": keyword,
        "state": status,
        "pid": auth.get('pid', ''),
        "pname": auth.get('pname', ''),
        "loginId": auth.get('loginId', ''),
    }
    resp = session.get(url, params=params, headers=headers)
    resp.raise_for_status()
    data = resp.json()
    return data.get('resultData', {}).get('voList', [])


def resolve_order_to_id(session, keyword, search_by='solutionCode'):
    """将 solutionCode 或 configureOrder 解析为 (lims_id, configure_order) 或 (None, None)"""
    global _order_to_id_cache, _type_list_cache

    if keyword in _order_to_id_cache:
        lid = _order_to_id_cache[keyword]
        print(f"    [cache hit] {keyword} → id={lid}")
        return lid, keyword

    if search_by == 'solutionCode':
        return _resolve_by_solution_code(session, keyword)
    else:
        lid = _resolve_by_configure_order(session, keyword)
        return (lid, keyword) if lid else (None, None)


def _resolve_by_solution_code(session, solution_code):
    """用 solutionCode 参数精准搜索，返回 (lims_id, configure_order) 或 (None, None)"""
    global _order_to_id_cache
    auth = get_session_auth()
    url = f"{BASE_URL}/detectionManager/manager/dtSolutionConfigure/getSolutionAdata"
    headers = {"Referer": f"{BASE_URL}/web/solutionConfigure.html?menuId=544"}
    for sol_type in ['SOLUTION_TYPE_B', 'SOLUTION_TYPE_C', 'SOLUTION_TYPE_D', 'SOLUTION_TYPE_E']:
        params = {
            "_search": "false",
            "nd": str(int(time.time() * 1000)),
            "pageSize": 30,
            "pageNo": 1,
            "sidx": "",
            "sord": "asc",
            "type": sol_type,
            "solutionCode": solution_code,
            "status": "1",
            "pid": auth.get('pid', ''),
            "pname": auth.get('pname', ''),
            "loginId": auth.get('loginId', ''),
        }
        print(f"    [API] getSolutionAdata solutionCode={solution_code} type={sol_type} ...")
        resp = session.get(url, params=params, headers=headers)
        if resp.status_code == 500:
            continue
        resp.raise_for_status()
        data = resp.json()
        items = data.get('resultData', {}).get('voList', [])
        for item in items:
            sc = str(item.get('solutionCode', '')).strip()
            if sc == solution_code:
                lid = item.get('id')
                co = str(item.get('configureOrder', '')).strip()
                if lid:
                    _order_to_id_cache[solution_code] = lid
                    _order_to_id_cache[co] = lid
                    print(f"    [exact match] {solution_code} → id={lid} ({co})")
                    return lid, co
    print(f"    [not found] solutionCode={solution_code} 未找到精确匹配")
    return None, None


def _resolve_by_configure_order(session, configure_order):
    """用 configureOrder 全量列表匹配（慢，仅作为回退）"""
    global _order_to_id_cache, _type_list_cache

    parts = configure_order.split('-')
    if len(parts) < 2:
        return None
    prefix = parts[0].upper()
    if prefix == 'A':
        return None
    type_map = {'B': 'SOLUTION_TYPE_B', 'C': 'SOLUTION_TYPE_C', 'D': 'SOLUTION_TYPE_D', 'E': 'SOLUTION_TYPE_E'}
    sol_type = type_map.get(prefix)
    if not sol_type:
        return None

    try:
        if sol_type in _type_list_cache:
            items = _type_list_cache[sol_type]
        else:
            auth = get_session_auth()
            url = f"{BASE_URL}/detectionManager/manager/dtSolutionConfigure/getSolutionAdata"
            headers = {"Referer": f"{BASE_URL}/web/solutionConfigure.html?menuId=544"}
            params = {
                "_search": "false", "nd": str(int(time.time() * 1000)),
                "pageSize": 9999, "pageNo": 1,
                "sidx": "", "sord": "asc",
                "type": sol_type, "status": "1",
                "pid": auth.get('pid', ''), "pname": auth.get('pname', ''),
                "loginId": auth.get('loginId', ''),
            }
            print(f"    [API] getSolutionAdata type={sol_type} pageSize=9999 (全量回退) ...")
            resp = session.get(url, params=params, headers=headers)
            if resp.status_code == 500:
                return None
            resp.raise_for_status()
            data = resp.json()
            items = data.get('resultData', {}).get('voList', [])
            _type_list_cache[sol_type] = items
            print(f"    [API ok] 获取到 {len(items)} 条记录")

        # 精确匹配
        for item in items:
            order = str(item.get('configureOrder', '')).strip()
            if order == configure_order:
                lid = item.get('id')
                if lid:
                    _order_to_id_cache[configure_order] = lid
                    return lid
        # 前缀+数字精确匹配
        order_num = parts[1].strip()
        for item in items:
            item_order = str(item.get('configureOrder', '')).strip()
            item_parts = item_order.split('-')
            if (len(item_parts) >= 2
                    and item_parts[0].upper() == prefix
                    and item_parts[1].strip() == order_num):
                lid = item.get('id')
                if lid:
                    _order_to_id_cache[configure_order] = lid
                    return lid

        print(f"    [not found] {configure_order} 在 {sol_type} 列表中未找到")
        # 打印最接近的几条记录帮助调试
        similar = [str(i.get('configureOrder', '')) for i in items
                    if str(i.get('configureOrder', '')).startswith(prefix + '-')]
        if similar:
            print(f"    [hint] 同前缀记录: {similar[:5]}")

    except Exception as e:
        print(f"    [ERROR] resolve {configure_order}: {e}")

    return None


def fetch_record(session, lims_id=None, order_str=None):
    """获取记录：A类用 consumableBill 缓存，其他用 viewDtSolutionConfigure"""
    global _consumable_cache

    # A 类：直接从 consumable 缓存返回
    if order_str and order_str.upper().startswith('A-'):
        if order_str in _consumable_cache:
            return _consumable_cache[order_str], f"consumable({order_str})"
        # 尝试解析后再查缓存
        resolved_id = resolve_order_to_id(session, order_str)
        if order_str in _consumable_cache:
            return _consumable_cache[order_str], f"consumable({order_str})"
        return None, None

    method = None

    # 判断输入格式：solutionCode (CK-CG-/CK-FCM-) 或 configureOrder (B-/C-/D-)
    resolved_id = lims_id
    view_order_str = order_str  # 用于 fetch_view 的 type 检测
    if not resolved_id and order_str:
        if order_str.startswith('CK-'):
            resolved_id, resolved_co = resolve_order_to_id(session, order_str, search_by='solutionCode')
            if resolved_co and resolved_co != order_str:
                view_order_str = resolved_co  # 用 configureOrder 做 type 检测
        elif not order_str.startswith('A-'):
            resolved_id, _ = resolve_order_to_id(session, order_str, search_by='configureOrder')
    # 回退：configureOrder 数字部分
    if not resolved_id and order_str and '-' in order_str:
        num_part = order_str.split('-')[1].strip()
        if num_part.isdigit():
            resolved_id = int(num_part)
            print(f"    [fallback] 使用数字部分作为 ID: {resolved_id}")

    if resolved_id:
        # 优先用 view 接口（返回完整 detailList）
        try:
            rec = fetch_view(session, resolved_id, order_str=view_order_str)
            if rec and rec.get('id'):
                # 验证：configureOrder 输入时检查匹配；solutionCode 输入时检查 solutionCode 匹配
                if order_str:
                    if order_str.startswith('CK-'):
                        rec_sc = str(rec.get('solutionCode', '')).strip()
                        if rec_sc and rec_sc != order_str:
                            print(f"    [mismatch] ID {resolved_id} solutionCode={rec_sc}，非 {order_str}，跳过")
                        else:
                            method = f"view?id={resolved_id}"
                            return rec, method
                    else:
                        rec_order = str(rec.get('configureOrder', '')).strip()
                        if rec_order and rec_order != order_str:
                            print(f"    [mismatch] ID {resolved_id} 返回 {rec_order}，非 {order_str}，跳过")
                        else:
                            method = f"view?id={resolved_id}"
                        return rec, method
                else:
                    method = f"view?id={resolved_id}"
                    return rec, method
        except Exception as e:
            print(f"    [ERROR] fetch_view({resolved_id}): {e}")

        # 回退用 detail 接口
        try:
            rec = fetch_detail(session, resolved_id)
            if rec and rec.get('id'):
                if order_str:
                    if order_str.startswith('CK-'):
                        rec_sc = str(rec.get('solutionCode', '')).strip()
                        if rec_sc and rec_sc != order_str:
                            print(f"    [mismatch] ID {resolved_id} solutionCode={rec_sc}，非 {order_str}，跳过")
                        else:
                            method = f"detail?id={resolved_id} (fallback)"
                            return rec, method
                    else:
                        rec_order = str(rec.get('configureOrder', '')).strip()
                        if rec_order and rec_order != order_str:
                            print(f"    [mismatch] ID {resolved_id} 返回 {rec_order}，非 {order_str}，跳过")
                        else:
                            method = f"detail?id={resolved_id} (fallback)"
                            return rec, method
                else:
                    method = f"detail?id={resolved_id} (fallback)"
                    return rec, method
        except Exception as e:
            print(f"    [ERROR] fetch_detail({resolved_id}): {e}")

    return None, None


def parse_conc(conc_str):
    """解析浓度字段，返回 (数值, 单位)"""
    if not conc_str:
        return None, None
    text = str(conc_str).strip()
    # 带括号: "0.01(mg/L)"
    m = re.match(r'^([\d.]+)\s*\(([^)]+)\)', text)
    if m:
        return float(m.group(1)), m.group(2)
    # 不带括号: "1000μg/mL"
    m = re.match(r'^([\d.]+)\s*(\S+)', text)
    if m:
        return float(m.group(1)), m.group(2)
    return None, None


def print_record(rec, method, depth=0, dump_all=False):
    """打印单条记录的关键信息"""
    indent = "  " * (depth + 1)
    # A 类字段名不同
    is_consumable = rec.get('consumableType') is not None or rec.get('type') == 'CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE'
    if is_consumable:
        order = str(rec.get('num', '')).strip()
        name = str(rec.get('consumeName', '') or rec.get('name', '') or '').strip()
        conc = str(rec.get('concentration', '')).strip()
        person = str(rec.get('configuratorName') or rec.get('receiveUserName') or '').strip()
        date = str(rec.get('receiveDate', '') or rec.get('configureDate', '') or '')[:10]
        ctrl_no = str(rec.get('controlledNo', '') or rec.get('certNo', '') or '').strip()
        conc_unit_name = str(rec.get('concentrationUnitName', '') or '').strip()
        validity = str(rec.get('validityDate', '') or '')[:10]
    else:
        order = str(rec.get('configureOrder', '')).strip()
        name = str(rec.get('solutionName', '')).strip()
        conc = str(rec.get('concentration', '')).strip()
        person = str(rec.get('configuratorName') or rec.get('creatorName') or '').strip()
        date = str(rec.get('configureDate', '') or '')[:10]
        ctrl_no = str(rec.get('controlledNo', '')).strip()
        total_vol = str(rec.get('totalConstantVolume', '')).strip()
        const_vol = str(rec.get('constantVolume', '')).strip()
        medium = str(rec.get('medium', '')).strip()
        validity = str(rec.get('validityDate', '') or '')[:10]

    conc_val, conc_unit = parse_conc(conc)

    print(f"{indent}┌─ {order} (id={rec.get('id')})")
    print(f"{indent}│  获取方式: {method}")
    print(f"{indent}│  名称: {name}")
    print(f"{indent}│  浓度: {conc}")

    if is_consumable:
        print(f"{indent}│  浓度单位(concentrationUnitName): {conc_unit_name}")
        print(f"{indent}│  受控编号: {ctrl_no}")
        print(f"{indent}│  领用人: {person}  领用日期: {date}  有效期: {validity}")
        if dump_all:
            print(f"{indent}│  ── 全部字段 ──")
            for k, v in sorted(rec.items()):
                vstr = str(v)
                if len(vstr) > 120:
                    vstr = vstr[:120] + '...'
                print(f"{indent}│  {k}: {vstr}")
    else:
        original_code = str(rec.get('originalCode', '')).strip()
        sol_code = str(rec.get('solutionCode', '')).strip()
        print(f"{indent}│  编号(solutionCode): {sol_code}")
        print(f"{indent}│  受控编号(controlledNo): {ctrl_no}")
        print(f"{indent}│  父级(originalCode): {original_code}")
        print(f"{indent}│  配置人: {person}  配置日期: {date}  有效期: {validity}")
        print(f"{indent}│  定容体积(totalConstantVolume): {total_vol}  剩余数量(constantVolume): {const_vol}")
        print(f"{indent}│  溶剂介质: {medium}")

        # 打印 detailList
        detail_list = rec.get('detailList')
        if detail_list:
            print(f"{indent}│  detailList ({len(detail_list)} 行):")
            # 判断是 D 类（稀释点格式）还是 B/C 类（标准格式）
            has_dilution = any(dl.get('dilutionOne') is not None and dl.get('groupName') for dl in detail_list)

            if has_dilution:
                # D 类格式：groupName + dilutionOne~Twelve
                # 先收集稀释点数量
                max_pts = 0
                for dl in detail_list:
                    for n in range(1, 13):
                        if dl.get(f'dilution{["One","Two","Three","Four","Five","Six","Seven","Eight","Nine","Ten","Eleven","Twelve"][n-1]}') is not None:
                            max_pts = max(max_pts, n)
                pt_labels = ["One","Two","Three","Four","Five","Six","Seven","Eight","Nine","Ten","Eleven","Twelve"][:max_pts]

                print(f"{indent}│    稀释点数: {max_pts}")
                for dl in detail_list:
                    group = str(dl.get('groupName', '')).strip()
                    unit = str(dl.get('groupUnit', '')).strip()
                    values = []
                    for p in pt_labels:
                        v = dl.get(f'dilution{p}')
                        values.append(str(v) if v is not None else '—')
                    extra = ''
                    if dl.get('originalCode'):
                        extra = f"  [源: {dl.get('originalCode')} 浓度={dl.get('originalConcentration')} {dl.get('originalUnit','')} 领用={dl.get('receivedQuantity')} {dl.get('receivedUint','')}]"
                    print(f"{indent}│    {group}({unit}): {' | '.join(values)}{extra}")
            else:
                # B/C 类标准格式
                for i, dl in enumerate(detail_list):
                    dl_name = str(dl.get('originalName', '')).strip()
                    dl_conc = str(dl.get('originalConcentration', '')).strip()
                    dl_recv = str(dl.get('receivedQuantity', '')).strip()
                    dl_recv_unit = str(dl.get('receivedUint', '')).strip()
                    dl_use = str(dl.get('useUantity', '')).strip()
                    dl_use_unit = str(dl.get('useUnit', '')).strip()
                    dl_vol = str(dl.get('volume', '')).strip()
                    dl_unit = str(dl.get('unit', '')).strip()
                    dl_med = str(dl.get('medium', '')).strip()
                    dl_cfg_conc = str(dl.get('configurationConcentration', '')).strip()
                    dl_cfg_unit = str(dl.get('configurationUnit', '')).strip()
                    dl_no = str(dl.get('originalNo', '')).strip().replace('\n', ' | ')
                    dl_code = str(dl.get('originalCode', '')).strip()
                    print(f"{indent}│    [{i+1}] 源名称: {dl_name}")
                    print(f"{indent}│        源浓度: {dl_conc}")
                    print(f"{indent}│        领用数量: {dl_recv} {dl_recv_unit}  使用数量: {dl_use} {dl_use_unit}")
                    print(f"{indent}│        溶剂: {dl_med}  定容: {dl_vol} {dl_unit}")
                    print(f"{indent}│        配置浓度: {dl_cfg_conc} {dl_cfg_unit}")
                    print(f"{indent}│        编号: {dl_no}")
                    print(f"{indent}│        父级: {dl_code}")
        else:
            print(f"{indent}│  detailList: (空)")

    return {
        'order': order,
        'name': name,
        'conc': conc,
        'conc_val': conc_val,
        'conc_unit': conc_unit,
        'original_code': str(rec.get('originalCode', '')).strip() if not is_consumable else '',
        'person': person,
        'date': date,
        'rec': rec,
    }


def _has_percent_source_conc(rec, parent_order):
    """从 detailList 中查找指定父级的源浓度，判断是否为 % 单位（称量型）"""
    detail_list = rec.get('detailList') or []
    for dl in detail_list:
        dl_code = str(dl.get('originalCode', '')).strip()
        dl_no = str(dl.get('originalNo', '')).strip()
        no_first = dl_no.split('\n')[0].strip().split('|')[0].strip() if dl_no else ''
        if dl_code == parent_order or no_first == parent_order:
            conc = str(dl.get('originalConcentration', '') or '').strip()
            return '%' in conc, conc
    return None, ''


def trace_chain(session, lims_id=None, order_str=None, depth=0, visited=None, dump_all=False):
    """从给定起点向上递归溯源，打印完整链路。
    停止条件：父级为 A 类，或源浓度单位含 %（称量型）。
    """
    if visited is None:
        visited = set()
    if depth > 10:
        print("  " * (depth + 1) + "⚠ 超过最大深度 10，停止")
        return

    key = str(lims_id or order_str)
    if key in visited:
        print("  " * (depth + 1) + "⚠ 循环引用，停止")
        return
    visited.add(key)

    label = f"溯源层级 {depth}"
    if lims_id and order_str:
        print(f"\n{'─' * 60}")
        print(f"  {label}: id={lims_id} / order={order_str}")
    elif lims_id:
        print(f"\n{'─' * 60}")
        print(f"  {label}: id={lims_id}")
    else:
        print(f"\n{'─' * 60}")
        print(f"  {label}: order={order_str}")

    rec, method = fetch_record(session, lims_id=lims_id, order_str=order_str)
    if not rec:
        print("  " * (depth + 1) + "✗ 未获取到记录，链路断裂")
        return

    info = print_record(rec, method, depth, dump_all=dump_all)

    order = info['order']
    prefix = order.split('-')[0].upper() if '-' in order else ''

    # 检查自身是否为称量型 B（receivedUint=g）
    if prefix == 'B' and info['rec'].get('receivedUint', '').strip() == 'g':
        print("  " * (depth + 1) + "✓ 领用单位=g，称量型 B，溯源结束")
        return

    # 继续向上溯源
    parent_orders = [oc.strip() for oc in re.split(r'[,，]', info['original_code']) if oc.strip()]
    if not parent_orders:
        print("  " * (depth + 1) + "✓ 无父级记录（originalCode 为空），溯源结束")
        return

    # 从 detailList 中提取父级 originalId（数字 ID），避免走 resolve
    parent_id_map = {}  # {configure_order: lims_id}
    detail_list = info['rec'].get('detailList') or []
    for dl in detail_list:
        dl_code = str(dl.get('originalCode', '')).strip()
        dl_no = str(dl.get('originalNo', '')).strip()
        no_first = dl_no.split('\n')[0].strip().split('|')[0].strip() if dl_no else ''
        dl_id = dl.get('originalId')
        if dl_id:
            for po in parent_orders:
                if dl_code == po or no_first == po:
                    parent_id_map.setdefault(po, dl_id)

    # 过滤停止条件：父级为 A 类 或 源浓度含 %
    traceable = []
    for po in parent_orders:
        po_prefix = po.split('-')[0].upper() if '-' in po else ''
        if po_prefix == 'A':
            print(f"  " * (depth + 1) + f"✓ 父级 {po} 为 A 类，停止追溯")
            continue
        is_pct, conc = _has_percent_source_conc(info['rec'], po)
        if is_pct:
            print(f"  " * (depth + 1) + f"✓ 父级 {po} 源浓度={conc}（称量型），停止追溯")
            continue
        traceable.append(po)

    if not traceable:
        print("  " * (depth + 1) + "✓ 所有父级均不满足追溯条件，溯源结束")
        return

    print(f"  " * (depth + 1) + f"→ 发现 {len(traceable)} 个可追溯父级: {traceable}")
    for po in traceable:
        pid = parent_id_map.get(po)
        trace_chain(session, lims_id=pid, order_str=po, depth=depth + 1, visited=visited, dump_all=dump_all)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("\n额外选项:")
        print("  --dump    打印完整字段（排查字段名问题）")
        sys.exit(1)

    targets = []
    dump_mode = False
    for arg in sys.argv[1:]:
        if arg == '--dump':
            dump_mode = True
        else:
            targets.append(arg.strip())
    print(f"溯源目标: {targets}\n")

    session = get_session()

    # 验证 session（宽松模式：验证失败只打印警告，不阻断）
    try:
        resp = session.get(f"{BASE_URL}/detectionManager/core/security/getLoginUser")
        if resp.status_code == 200 and resp.json().get("success"):
            print("Session 验证通过\n")
        else:
            print("⚠ Session 验证未通过，但继续尝试（可能仍可用）\n")
    except Exception as e:
        print(f"⚠ Session 验证异常: {e}，继续尝试\n")

    for target in targets:
        target = target.strip()
        if not target:
            continue

        # 判断是数字 ID 还是 configureOrder
        if target.isdigit():
            trace_chain(session, lims_id=int(target), dump_all=dump_mode)
        else:
            trace_chain(session, order_str=target, dump_all=dump_mode)

    print(f"\n{'═' * 60}")
    print("溯源完成")


if __name__ == '__main__':
    main()
