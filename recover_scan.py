# -*- coding: utf-8 -*-
"""只读扫描：定位 customNum 被清空的异常记录。只 GET。"""
import json, time, requests, sys
sys.stdout.reconfigure(encoding='utf-8')

BASE = "http://192.168.12.234:60015"
SESS_FILE = "session_cym.json"
with open(SESS_FILE, encoding="utf-8") as f:
    info = json.load(f)
s = requests.Session()
for c in info["cookies"]:
    kw = {"name": c["name"], "value": c["value"]}
    if c.get("domain"): kw["domain"] = c["domain"]
    if c.get("path"): kw["path"] = c["path"]
    s.cookies.set(**kw)
s.headers.update(info["headers"])
pid = info.get("pid"); pname = info.get("real_name", info.get("username"))

def query(status, keyword="", page=1, size=100):
    params = {
        "_search": "false", "nd": str(int(time.time()*1000)),
        "pageSize": size, "pageNo": page, "sidx": "", "sord": "asc",
        "type": "CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE",
        "receiveUserName": "", "receiveStartDate": "", "receiveEndDate": "",
        "confirmUserName": "", "confirmStartDate": "", "confirmEndDate": "",
        "invoiceNo": "", "groupId": "",
        "status": status, "pid": pid, "pname": pname, "loginId": pid,
    }
    if keyword: params["keyword"] = keyword
    r = s.get(f"{BASE}/detectionManager/manager/consumableBill/pageObj", params=params, timeout=25)
    return r.json().get("resultData") or {}

# 1) 概览 + 抽样确认 cym 可见 CK-FCM
for status in ("normal","history","overdue"):
    rd = query(status)
    vo = rd.get("voList") or []
    total = rd.get("records") or rd.get("totalCount") or len(vo)
    fcm = [v.get("customNum") for v in vo if str(v.get("customNum","")).startswith("CK-FCM")]
    print(f"[{status}] total={total}  sampled_voList={len(vo)}  CK-FCM_in_page={len(fcm)}  e.g. {fcm[:5]}")

# 2) 扫描 normal 全部分页，找 customNum 为空的异常记录
print("\n=== 扫描 normal 状态下 customNum 为空的记录 ===")
anomalies = []
for status in ("normal","history","overdue"):
    page = 1
    while True:
        rd = query(status, page=page, size=200)
        vo = rd.get("voList") or []
        if not vo: break
        for it in vo:
            cn = it.get("customNum")
            if cn is None or str(cn).strip() == "":
                anomalies.append((status, it))
        total = rd.get("records") or rd.get("totalCount") or 0
        if page * 200 >= (total or 0): break
        page += 1
        if page > 20: break  # 安全上限

print(f"customNum 为空的记录数: {len(anomalies)}")
for status, it in anomalies[:20]:
    print(f"  [{status}] id={it.get('id')!r} name={it.get('name')!r} cas={it.get('casNo')!r} "
          f"conc={it.get('concentration')!r} unit={it.get('concentrationUnitName')!r} "
          f"factory={it.get('factoryName')!r} create={it.get('createDatetime')!r} creator={it.get('creatorName')!r}")
