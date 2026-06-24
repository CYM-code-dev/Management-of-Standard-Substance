# -*- coding: utf-8 -*-
"""只读：探测详情接口的正确调用方式。"""
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
rid = 27198  # CK-FCM-2026068
common = {"pid":pid,"pname":pname,"loginId":pid,"_":str(int(time.time()*1000))}

def show(tag, resp):
    print(f"\n### {tag} ### HTTP {resp.status_code}")
    try:
        j = resp.json()
        print("  success:", j.get("success"), "errorCtx:", j.get("errorCtx"))
        rd = j.get("resultData")
        if isinstance(rd, dict):
            print("  字段数:", len(rd), "| keys:", list(rd.keys())[:20])
    except Exception:
        print("  body:", resp.text[:200])

# 1) POST + _method=GET
show("POST _method=GET (form)", s.post(f"{BASE}/detectionManager/manager/consumableBill/{rid}",
     data={**common, "_method":"GET"}, timeout=20))
# 2) POST 无 _method
show("POST plain (form)", s.post(f"{BASE}/detectionManager/manager/consumableBill/{rid}",
     data=common, timeout=20))
# 3) POST + _method=GET, id 在 query
show("POST _method=GET id-in-query", s.post(f"{BASE}/detectionManager/manager/consumableBill",
     data={"id":rid, **common, "_method":"GET"}, timeout=20))
# 4) GET pageObj 用 id 作为 keyword? 跳过
# 5) view / detail 子路径
for sub in ("detail","view","info","get"):
    show(f"GET /{sub}/{rid}", s.get(f"{BASE}/detectionManager/manager/consumableBill/{sub}/{rid}",
         params=common, timeout=15))
