# -*- coding: utf-8 -*-
"""只读：测 GET /consumableBill/{id} 详情接口返回的字段集是否完整。"""
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

# 取一条可见的 CK-FCM 记录的 id
params = {"_search":"false","nd":str(int(time.time()*1000)),"pageSize":50,"pageNo":1,
          "sidx":"","sord":"asc","type":"CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE",
          "receiveUserName":"","receiveStartDate":"","receiveEndDate":"",
          "confirmUserName":"","confirmStartDate":"","confirmEndDate":"",
          "invoiceNo":"","groupId":"","status":"normal",
          "pid":pid,"pname":pname,"loginId":pid}
vo = (s.get(f"{BASE}/detectionManager/manager/consumableBill/pageObj", params=params, timeout=25)
       .json().get("resultData") or {}).get("voList") or []
target = next((v for v in vo if str(v.get("customNum","")).startswith("CK-FCM")), None)
if not target:
    print("no CK-FCM in first page"); raise SystemExit
rid = target["id"]
print(f"测试 id={rid}  customNum={target.get('customNum')}\n")

# GET 详情
d = s.get(f"{BASE}/detectionManager/manager/consumableBill/{rid}",
          params={"pid":pid,"pname":pname,"loginId":pid,"_":str(int(time.time()*1000))}, timeout=25)
print("HTTP", d.status_code)
data = d.json()
print("success:", data.get("success"), "errorCtx:", data.get("errorCtx"))
rd = data.get("resultData")
print("\n=== 详情返回字段数:", len(rd) if isinstance(rd, dict) else "N/A", "===")
if isinstance(rd, dict):
    for k, v in rd.items():
        print(f"  {k} = {v!r}")
