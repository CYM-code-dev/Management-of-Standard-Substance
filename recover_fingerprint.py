# -*- coding: utf-8 -*-
"""精确定位：name/customNum/cas/conc 全空但 concentrationUnitName 非空的记录。只 GET。"""
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
pid = info.get("pid"); pname = info.get('real_name', info.get('username'))

def query(status, page, size=200):
    params = {"_search":"false","nd":str(int(time.time()*1000)),"pageSize":size,"pageNo":page,
              "sidx":"","sord":"asc","type":"CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE",
              "receiveUserName":"","receiveStartDate":"","receiveEndDate":"",
              "confirmUserName":"","confirmStartDate":"","confirmEndDate":"",
              "invoiceNo":"","groupId":"","status":status,
              "pid":pid,"pname":pname,"loginId":pid}
    r = s.get(f"{BASE}/detectionManager/manager/consumableBill/pageObj", params=params, timeout=30)
    return r.json().get("resultData") or {}

def blank(v):
    return v is None or str(v).strip() == ""

matches = []
for status in ("normal","history","overdue"):
    page = 1
    while True:
        rd = query(status, page)
        vo = rd.get("voList") or []
        if not vo: break
        for it in vo:
            unit = it.get("concentrationUnitName")
            if (not blank(unit)) and blank(it.get("customNum")) and blank(it.get("name")) \
               and blank(it.get("casNo")) and blank(it.get("concentration")) and blank(it.get("factoryName")):
                matches.append((status, it))
        total = rd.get("records") or rd.get("totalCount") or 0
        if page * 200 >= (total or 0): break
        page += 1
        if page > 30: break

print(f"指纹命中数: {len(matches)}\n")
for status, it in matches:
    print(f"=== [{status}] id={it.get('id')!r} ===")
    print(json.dumps(it, ensure_ascii=False, indent=2))
    print()
