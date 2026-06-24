# -*- coding: utf-8 -*-
"""按 CAS / 名称 keyword 检索 LIMS，定位记录 id 与残存字段。只 GET。"""
import json, time, sys, requests
sys.stdout.reconfigure(encoding='utf-8')
BASE = "http://192.168.12.234:60015"
info = json.load(open('session_cym.json', encoding='utf-8'))
s = requests.Session()
for c in info["cookies"]:
    kw = {"name": c["name"], "value": c["value"]}
    if c.get("domain"): kw["domain"] = c["domain"]
    if c.get("path"): kw["path"] = c["path"]
    s.cookies.set(**kw)
s.headers.update(info["headers"])
pid = info.get("pid"); pname = info.get("real_name", info.get("username"))

def query(status, keyword):
    params = {"_search":"false","nd":str(int(time.time()*1000)),"pageSize":50,"pageNo":1,
              "sidx":"","sord":"asc","type":"CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE",
              "receiveUserName":"","receiveStartDate":"","receiveEndDate":"",
              "confirmUserName":"","confirmStartDate":"","confirmEndDate":"",
              "invoiceNo":"","groupId":"","status":status,"keyword":keyword,
              "pid":pid,"pname":pname,"loginId":pid}
    r = s.get(f"{BASE}/detectionManager/manager/consumableBill/pageObj", params=params, timeout=25)
    return (r.json().get("resultData") or {}).get("voList") or []

KEYS = ["2244-16-8", "右旋香芹酮", "CK-FCM-2026039", "香芹酮"]
for kw in KEYS:
    print(f"\n##### keyword={kw!r} #####")
    hit = False
    for status in ("normal","history","overdue"):
        vo = query(status, kw)
        if vo:
            hit = True
            for it in vo:
                print(f"  [{status}] id={it.get('id')!r} customNum={it.get('customNum')!r} "
                      f"name={it.get('name')!r} cas={it.get('casNo')!r} conc={it.get('concentration')!r} "
                      f"unit={it.get('concentrationUnitName')!r} factory={it.get('factoryName')!r} "
                      f"create={it.get('createDatetime')!r} creator={it.get('creatorName')!r}")
    if not hit:
        print("  (无命中)")
