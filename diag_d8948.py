# -*- coding: utf-8 -*-
"""只读：对比 D-8948 在 列表(getSolutionAdata) 与 详情(viewDtSolutionConfigure) 的 configureDate/solutionCode。"""
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
H = {"Referer": f"{BASE}/web/solutionConfigure.html?menuId=544"}
ORDER = "D-8948"

# 1) 列表
r = s.get(f"{BASE}/detectionManager/manager/dtSolutionConfigure/getSolutionAdata",
          params={"type":"SOLUTION_TYPE_D","pageSize":9999,"pageNo":1,"status":"1",
                  "pid":pid,"pname":pname,"loginId":pid}, headers=H, timeout=30)
rd = r.json().get("resultData") or {}
vo = rd.get("voList") or []
item = next((v for v in vo if str(v.get("configureOrder","")).strip()==ORDER), None)
print("=== 列表 getSolutionAdata ===")
if not item:
    print("  未找到 D-8948, 共", len(vo));
    # 退化：打印前几个 configureOrder
    print("  样例:", [(v.get("configureOrder"), v.get("id")) for v in vo[:5]])
else:
    print("  id:", item.get("id"))
    print("  solutionCode:", repr(item.get("solutionCode")))
    print("  configureDate raw:", repr(item.get("configureDate")))
    print("  configureDate[:10]:", repr(str(item.get("configureDate"))[:10]))
    # 2) 详情
    rv = s.get(f"{BASE}/detectionManager/manager/dtSolutionConfigure/viewDtSolutionConfigure",
               params={"id":item["id"],"type":"SOLUTION_TYPE_D"}, headers=H, timeout=30)
    print("\n=== 详情 viewDtSolutionConfigure ===  HTTP", rv.status_code)
    det = (rv.json().get("resultData") or {})
    print("  solutionCode:", repr(det.get("solutionCode")))
    print("  configureDate raw:", repr(det.get("configureDate")))
    print("  configureDate[:10]:", repr(str(det.get("configureDate"))[:10]))
