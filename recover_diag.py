# -*- coding: utf-8 -*-
"""只读诊断：查 CK-FCM-2026039 在 LIMS 中的现状。只 GET，不写。"""
import json, os, time, requests

BASE = "http://192.168.12.234:60015"
TARGET = "CK-FCM-2026039"
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

pid = info.get("pid")
pname = info.get("real_name", info.get("username"))

# 1) 先探活
try:
    probe = s.get(f"{BASE}/detectionManager/core/users/info", timeout=15).json()
    print("[probe] alive:", probe.get("success"), "| user:", pname, "| pid:", pid)
except Exception as e:
    print("[probe] FAILED:", e)
    raise SystemExit

# 2) 按 keyword 在三种状态下查
found_any = False
for status in ("normal", "history", "overdue"):
    params = {
        "_search": "false", "nd": str(int(time.time()*1000)),
        "pageSize": 50, "pageNo": 1, "sidx": "", "sord": "asc",
        "type": "CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE",
        "receiveUserName": "", "receiveStartDate": "", "receiveEndDate": "",
        "confirmUserName": "", "confirmStartDate": "", "confirmEndDate": "",
        "invoiceNo": "", "groupId": "",
        "status": status, "keyword": TARGET,
        "pid": pid, "pname": pname, "loginId": pid,
    }
    r = s.get(f"{BASE}/detectionManager/manager/consumableBill/pageObj",
              params=params, timeout=20)
    print(f"\n[status={status}] HTTP {r.status_code}")
    try:
        data = r.json()
    except Exception:
        print("  non-json:", r.text[:200]); continue
    rd = data.get("resultData") or {}
    vo = rd.get("voList") or []
    print(f"  success={data.get('success')} records={rd.get('records') or rd.get('totalCount')} voList_len={len(vo)}")
    for it in vo:
        found_any = True
        print("  --- HIT ---")
        for k in ("id","customNum","name","casNo","standard","concentration",
                  "concentrationUnitName","factoryName","batchNo","validDate",
                  "storageCondition","storagePositionName","usedNum","totalNum",
                  "createDatetime","creatorName"):
            print(f"    {k} = {it.get(k)!r}")

if not found_any:
    print("\n>>> keyword 搜索未命中，符合 customNum 被清空的假设。")
    print(">>> 下一步：按 cas/name 或列出最近修改的记录来定位被清空的 id。")
