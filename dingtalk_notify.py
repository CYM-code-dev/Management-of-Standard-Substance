# -*- coding: utf-8 -*-
"""钉钉群消息自动通知：标液即将过期提醒。

每个工作日 15:00（trigger_hour 可配置）：
  1. 获取一个存活的 LIMS 会话（内存→磁盘 session 文件→OCR 自动登录）；
  2. 查询全部 储备液(B)/应用液(C)/工作液(D)，筛选「即将过期」且配制人在名单内、
     总有效期>阈值、未废弃的记录；
  3. 加签后以 markdown 推送到钉钉群。

设计要点：
- 不复制 login_html 的登录/查询逻辑，复用其 RemoteSystem 与 getSolutionAdata 调用范式；
- 通过 sys.modules['__main__'] 取得「正在运行的」login_html（生产里它是 __main__），
  避免 `import login_html` 造成二次执行/第二份 Flask 实例；
- 验证码是算术题 A+B（0~9，仅加法）：ddddocr 读出后取前两位数字字符求和作为答案，
  天然屏蔽尾部 '=?' 的误读（实测 6/6 准确）。

无前端；由 login_html.__main__ 在 WERKZEUG_RUN_MAIN 守卫下调用 start() 拉起后台线程。
也可命令行手动触发：python dingtalk_notify.py --run-now / --test-send
"""
import base64
import datetime
import glob
import hashlib
import hmac
import json
import sys
import threading
import time
from urllib.parse import quote_plus

import requests

import lims_auto_login

CONFIG_FILE = "config.json"
STATE_FILE = "dingtalk_state.json"
_DINGTALK_LOG = "[DingTalk]"

# 调度状态（_last_run_date / _last_eod_alert_date 持久化到 STATE_FILE，重启不丢）
_flag = False
_thread = None
_last_run_date = None     # 当天已成功发送（或确认无到期项）→ 当天不再重试/补发
_last_eod_alert_date = None  # 当天已发过"截止查询失败"最终警告 → 每天最多1条
_last_attempt_dt = None   # 上次尝试时刻 → 控制 retry_interval_minutes 重试间隔
_last_excel_notify_date = None  # Excel月度提醒：当天已成功发送（或确认无到期项）→ 当天不再重试
_last_excel_attempt_dt = None   # Excel月度提醒：上次尝试时刻 → 控制重试间隔
_last_device_remind_date = None  # 设备使用率提醒：当天已发送 → 当天不再重试
_last_device_check_date = None   # 设备使用率检查：当天已处理 → 当天不再重试
_last_device_attempt_dt = None   # 设备使用率：上次尝试时刻 → 控制重试间隔
_last_inventory_run = None       # 入库提醒：已完成时段 "YYYY-MM-DD HH:MM"（每天两个时段各自判重）
_last_inventory_attempt_dt = None  # 入库提醒：上次尝试时刻 → 控制重试间隔

# 入库提醒时段（工作日各查一次 LIMS 未入库样品并提醒）
_INVENTORY_SLOTS = ("09:00", "14:30")


# ==================== 配置 / 工作日 ====================
def _load_cfg():
    try:
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f).get("dingtalk") or {}
    except Exception as e:
        print(f"{_DINGTALK_LOG} 读取 config.json 失败: {e}")
        return {}


def _departments():
    """读 config.json 根的 departments（_load_cfg 只返回 dingtalk 段，取不到）。"""
    try:
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return (json.load(f).get("departments") or {})
    except Exception:
        return {}


def _load_state():
    """读取持久化状态（last_run_date / last_eod_alert_date / last_excel_notify_date /
    last_device_remind_date / last_device_check_date / last_inventory_run），
    使重启后判重/补发正确。"""
    global _last_run_date, _last_eod_alert_date, _last_excel_notify_date
    global _last_device_remind_date, _last_device_check_date
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            st = json.load(f)
    except Exception:
        return
    for key, target in (
        ("last_run_date", "_last_run_date"),
        ("last_excel_notify_date", "_last_excel_notify_date"),
        ("last_device_remind_date", "_last_device_remind_date"),
        ("last_device_check_date", "_last_device_check_date"),
        ("last_eod_alert_date", "_last_eod_alert_date"),
    ):
        v = st.get(key)
        if not v:
            continue
        try:
            d = datetime.datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
        except Exception:
            continue
        globals()[target] = d
    # last_inventory_run 是 "日期 时段" 复合串（如 "2026-09-10 09:00"），原样恢复
    global _last_inventory_run
    _last_inventory_run = st.get("last_inventory_run") or None


def _save_state():
    """持久化当天状态。"""
    try:
        st = {
            "last_run_date": _last_run_date.isoformat() if _last_run_date else None,
            "last_eod_alert_date": _last_eod_alert_date.isoformat() if _last_eod_alert_date else None,
            "last_excel_notify_date": _last_excel_notify_date.isoformat() if _last_excel_notify_date else None,
            "last_device_remind_date": _last_device_remind_date.isoformat() if _last_device_remind_date else None,
            "last_device_check_date": _last_device_check_date.isoformat() if _last_device_check_date else None,
            "last_inventory_run": _last_inventory_run,
        }
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(st, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"{_DINGTALK_LOG} 写状态文件失败: {e}")


def _date_set(cfg, key):
    out = set()
    for s in cfg.get(key, []) or []:
        try:
            out.add(datetime.datetime.strptime(str(s).strip()[:10], "%Y-%m-%d").date())
        except Exception:
            continue
    return out


def _lh_holidays(year):
    """从正在运行的 login_html 取该年度节假日映射（{iso_date: bool}，true=放假/false=补班）；无则 None。"""
    try:
        lh = _get_lh()
        fn = getattr(lh, "get_holidays", None)
        if fn:
            mapping, _src = fn(year)
            return mapping
    except Exception as e:
        print(f"{_DINGTALK_LOG} 读取节假日缓存失败 year={year}: {e}")
    return None


def is_workday(d, cfg=None):
    """d: datetime.date。
    优先级：config.json 手动覆盖(extra_workdays/holidays) > 共享节假日缓存(法定假日/调休) > 周一~周五。"""
    cfg = cfg if cfg is not None else _load_cfg()
    ds_h = _date_set(cfg, "holidays")
    ds_w = _date_set(cfg, "extra_workdays")
    if d in ds_w:
        return True
    if d in ds_h:
        return False
    mapping = _lh_holidays(d.year)
    if mapping:
        key = d.strftime("%Y-%m-%d")
        if key in mapping:
            return not mapping[key]   # true=放假→非工作日；false=补班→工作日
    return d.weekday() < 5


def next_workday(d, cfg=None):
    """从 d 的次日开始，首个工作日。"""
    cfg = cfg if cfg is not None else _load_cfg()
    cur = d + datetime.timedelta(days=1)
    # 上限 40 天，防配置异常死循环
    for _ in range(40):
        if is_workday(cur, cfg):
            return cur
        cur += datetime.timedelta(days=1)
    return cur


def _parse_expiry_date(val):
    """把 Excel 有效期单元格原值解析成 datetime.date；失败返回 None。
    支持 datetime/date 对象与多种日期字符串格式。"""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = str(val).strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d",
                "%d-%m-%Y", "%m/%d/%Y", "%Y年%m月%d日"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _excel_target_date(year, month, day, cfg=None):
    """月度提醒的目标发送日：从 day 号往前回溯到首个工作日
    （day 是工作日即用 day；否则回到上一个工作日）。25 号每月都有，无月末越界。"""
    cfg = cfg if cfg is not None else _load_cfg()
    d = datetime.date(year, month, day)
    for _ in range(day):
        if is_workday(d, cfg):
            return d
        d -= datetime.timedelta(days=1)
    return d


def _nth_workday_after_26(year, month, n, cfg=None):
    """从 month 的 27 号起严格向后数的第 n 个工作日（n>=1）。
    27 号本身若是工作日即第 1 个；否则顺延。跨月/跨年由 timedelta 自然推进，
    is_workday 按各日期所在年取节假日缓存。"""
    cfg = cfg if cfg is not None else _load_cfg()
    cur = datetime.date(year, month, 27)
    found = 0
    for _ in range(40):
        if is_workday(cur, cfg):
            found += 1
            if found == n:
                return cur
        cur += datetime.timedelta(days=1)
    return cur


def _device_period(today):
    """本期 = 上月27号 ~ 本月26号（与 device_usage.csv 的 period_start/period_end 对齐，返回 ISO 串）。"""
    end = datetime.date(today.year, today.month, 26)
    first = datetime.date(today.year, today.month, 1)
    start = (first - datetime.timedelta(days=1)).replace(day=27)  # 上月末日→replace 为上月27号
    return start.isoformat(), end.isoformat()


# ==================== LIMS 会话获取 ====================
def _get_lh():
    """取得正在运行的 login_html 模块。
    生产环境 login_html 以 __main__ 运行；此处直接取 __main__，避免二次 import 执行。"""
    m = sys.modules.get("__main__")
    if m and hasattr(m, "RemoteSystem") and hasattr(m, "user_systems"):
        return m
    import login_html  # 独立测试时（login_html 非 __main__）正常导入一次
    return login_html


def _acquire_system(cfg):
    """按 内存会话→磁盘 session 文件→OCR 登录 的顺序取一个可用 RemoteSystem；无则 None。"""
    lh = _get_lh()
    RemoteSystem = lh.RemoteSystem

    # 1) 内存中存活会话
    for sys_ in list(lh.user_systems.values()):
        if getattr(sys_, "current_user", None):
            try:
                if sys_.verify_session():
                    return sys_
            except Exception:
                continue

    # 2) 磁盘 session 文件（任意账号）
    for sess_file in glob.glob("session_*.json"):
        username = sess_file[len("session_"):-len(".json")]
        sys_ = RemoteSystem("dt_" + username)
        sys_.current_user = username
        try:
            if sys_.load_session() and sys_.verify_session():
                print(f"{_DINGTALK_LOG} 复用磁盘会话: {username}")
                return sys_
        except Exception:
            continue

    # 3) OCR 自动登录（专用账号）—— 复用 lims_auto_login，返回 LoginResult
    #    （字段与 RemoteSystem 鸭子兼容：session/base_url/current_pid/current_real_name）
    acc = cfg.get("ocr_account") or {}
    if acc.get("username") and acc.get("password"):
        try:
            res = lims_auto_login.auto_login(acc["username"], acc["password"])
            if res:
                print(f"{_DINGTALK_LOG} OCR 登录成功（{res.current_real_name}）")
                return res
        except Exception as e:
            print(f"{_DINGTALK_LOG} OCR 登录异常: {e}")
    return None


# ==================== LIMS 查询 / 过滤 ====================
def _fetch_all(system, sol_type):
    """分页取某类型的全部有效记录（status=1）。复用 getSolutionAdata 范式。"""
    base = system.base_url
    url = f"{base}/detectionManager/manager/dtSolutionConfigure/getSolutionAdata"
    headers = {"Referer": f"{base}/web/solutionConfigure.html?menuId=544"}
    items, page = [], 1
    while page <= 50:
        params = {
            "_search": "false", "nd": str(int(time.time() * 1000)),
            "pageSize": 9999, "pageNo": page, "sidx": "", "sord": "asc",
            "solutionName": "", "solutionCode": "", "customType": "", "configStatus": "",
            "controlledNo": "", "storageLocation": "",
            "configureStartDate": "", "configureEndDate": "", "configureUserName": "",
            "receiveUserName": "", "auditStatus": "", "status": "1", "type": sol_type,
            "pid": system.current_pid or "", "pname": system.current_real_name or "",
            "loginId": system.current_pid or "",
        }
        resp = system.session.get(url, params=params, headers=headers, timeout=20)
        rd = (resp.json() or {}).get("resultData") or {}
        page_items = rd.get("voList", []) or []
        items.extend(page_items)
        if not rd.get("hasNext", False):
            break
        page += 1
    return items


def _fetch_pending_inventory(system):
    """查「轻工-释放项目统计」近30天样品，返回受理超4小时仍未入库的记录（按受理时间升序）。

    pageCustomQueryStatistics：query 为 JSON 串（受理起止日期），acceptTime 形如
    "2026-09-08 08"（小时精度），inventoryTime 为 null 表示未入库。"""
    base = system.base_url
    url = f"{base}/detectionManager/manager/statisticalAnalysis/pageCustomQueryStatistics"
    headers = {"Referer": f"{base}/web/"}
    today = datetime.date.today()
    query = json.dumps({
        "acceptStartTime": (today - datetime.timedelta(days=29)).isoformat(),
        "acceptEndTime": today.isoformat(),
    }, ensure_ascii=False)
    now = datetime.datetime.now()
    matches, page = [], 1
    while page <= 50:
        params = {
            "_search": "false", "nd": str(int(time.time() * 1000)),
            "pageSize": 9999, "pageNo": page, "sidx": "", "sord": "asc",
            "query": query, "name": "自定义查询", "deputyName": "轻工-释放项目统计",
            "pid": system.current_pid or "", "pname": system.current_real_name or "",
            "loginId": system.current_pid or "",
        }
        resp = system.session.get(url, params=params, headers=headers, timeout=20)
        rd = (resp.json() or {}).get("resultData") or {}
        for it in rd.get("voList", []) or []:
            if it.get("inventoryTime"):
                continue
            try:
                at = datetime.datetime.strptime(str(it.get("acceptTime") or "").strip(),
                                                "%Y-%m-%d %H")
            except Exception:
                continue
            if now >= at + datetime.timedelta(hours=4):
                matches.append(it)
        if not rd.get("hasNext", False):
            break
        page += 1
    matches.sort(key=lambda x: str(x.get("acceptTime") or ""))
    return matches


def _fetch_accept_users(system, detection_nos):
    """按检测单号（detectionNo 去掉末3位样品序号，如 TN26090315001→TN26090315）查
    detection/pageObj，返回 {单号: acceptUserName(客服)}。查不到的不入字典。"""
    base = system.base_url
    url = f"{base}/detectionManager/manager/detection/pageObj"
    headers = {"Referer": f"{base}/web/"}
    out = {}
    for no in detection_nos:
        params = {
            "_search": "false", "nd": str(int(time.time() * 1000)),
            "pageSize": 1, "pageNo": 1, "sidx": "", "sord": "asc",
            "detectionType": "PRODUCT", "orgId": "", "acceptUserId": "",
            "auditStatus": "", "keyword": no,
            "pid": system.current_pid or "", "pname": system.current_real_name or "",
            "loginId": system.current_pid or "",
        }
        try:
            resp = system.session.get(url, params=params, headers=headers, timeout=20)
            vo = ((resp.json() or {}).get("resultData") or {}).get("voList") or []
            if vo and vo[0].get("acceptUserName"):
                out[no] = vo[0]["acceptUserName"]
        except Exception as e:
            print(f"{_DINGTALK_LOG} 查客服失败 no={no}: {e}")
    return out


def _type_label(order):
    pfx = str(order or "").split("-")[0].upper()
    return {"B": "储备液", "C": "应用液", "D": "工作液"}.get(pfx, "标液")


def _collect_expiring(system, cfg):
    """返回即将过期、需通知的记录列表（不按配制人筛选）。
    按 notify_groups 的配制人名单分组路由由 run_once 完成。"""
    today = datetime.date.today()
    nw = next_workday(today, cfg)
    min_days = int(cfg.get("min_total_validity_days", 3))

    seen, matches = set(), []
    for sol_type in ("SOLUTION_TYPE_E", "SOLUTION_TYPE_D"):  # E=B+C, D=工作液
        for it in _fetch_all(system, sol_type):
            lid = it.get("id")
            if lid and lid in seen:
                continue
            if lid:
                seen.add(lid)

            vstr = str(it.get("validityDate") or "").strip()[:10]
            cstr = str(it.get("configureDate") or "").strip()[:10]
            if not vstr or not cstr:
                continue
            try:
                vdate = datetime.datetime.strptime(vstr, "%Y-%m-%d").date()
                cdate = datetime.datetime.strptime(cstr, "%Y-%m-%d").date()
            except Exception:
                continue

            # 即将过期：今天 < 有效期 <= 下一个工作日（严格大于今天避免重复推送；
            # 周五时 next_workday=下周一，自然覆盖本周末到期项）
            if not (today < vdate <= nw):
                continue
            # 整瓶有效期 > 阈值 天才通知（短效液不刷屏）
            if (vdate - cdate).days <= min_days:
                continue
            # 已废弃的跳过
            if it.get("disposeUserName"):
                continue
            matches.append(it)

    # 按到期日升序
    matches.sort(key=lambda x: str(x.get("validityDate") or ""))
    return matches


def _person_of(it):
    return str(it.get("configuratorName") or it.get("creatorName") or "").strip()


def _route_by_group(items, cfg):
    """按 notify_groups 把记录按配制人路由到不同群：返回 [(group, [items]), ...]，
    仅保留有命中的组，顺序同 notify_groups。不在任何组名单内的记录被丢弃。
    无 notify_groups 时回退单组（顶层 webhook + 顶层 configurators），兼容旧配置。"""
    groups = cfg.get("notify_groups") or []
    if not groups:
        top = {
            "webhook": cfg.get("webhook"),
            "secret": cfg.get("secret"),
            "configurators": cfg.get("configurators") or [],
        }
        names = set(top["configurators"])
        routed = [it for it in items if _person_of(it) in names]
        return [(top, routed)] if routed else []
    out = []
    for g in groups:
        names = set(g.get("configurators") or [])
        routed = [it for it in items if _person_of(it) in names]
        if routed:
            out.append((g, routed))
    return out


def _collect_excel_expiring(path, today, min_days, max_days):
    """读指定 Excel 路径，返回「今天+min_days ≤ 有效期 ≤ 今天+max_days」的条目（按有效期升序）。
    每条 {group, labNo, name, cas, expiry(datetime.date)}。读/解析异常上抛由调用方处理。"""
    rows = _get_lh().read_excel_expiry_rows(path)
    lo = today + datetime.timedelta(days=min_days)
    hi = today + datetime.timedelta(days=max_days)
    out = []
    for r in rows:
        # 使用情况为「用完」的不通知
        if "用完" in (r.get("usage") or ""):
            continue
        exp = _parse_expiry_date(r.get("expiry_raw"))
        if exp is None or not (lo <= exp <= hi):
            continue
        out.append({
            "group": r.get("group") or "",
            "labNo": r.get("labNo") or "",
            "name": r.get("name") or "",
            "cas": r.get("cas") or "",
            "expiry": exp,
        })
    out.sort(key=lambda x: x["expiry"])
    return out


# ==================== 钉钉发送（加签） ====================
def _sign_url(webhook, secret):
    ts = str(round(time.time() * 1000))
    string_to_sign = f"{ts}\n{secret}"
    hmac_code = hmac.new(secret.encode("utf-8"), string_to_sign.encode("utf-8"),
                         digestmod=hashlib.sha256).digest()
    sign = quote_plus(base64.b64encode(hmac_code))
    return f"{webhook}&timestamp={ts}&sign={sign}"


def send_markdown(title, md_text, cfg=None, webhook=None, secret=None):
    """加签后发送 markdown；返回 (ok, resp_json)。
    webhook/secret 未传则取 cfg.webhook/cfg.secret（LIMS 标液提醒群）。
    Excel 月度提醒走另一个群时由调用方传入 excel_webhook/excel_secret。"""
    cfg = cfg if cfg is not None else _load_cfg()
    webhook = webhook or cfg.get("webhook")
    secret = secret or cfg.get("secret")
    if not webhook or not secret:
        print(f"{_DINGTALK_LOG} 缺少 webhook/secret 配置")
        return False, {"errcode": -1, "errmsg": "missing webhook/secret"}
    url = _sign_url(webhook, secret)
    payload = {"msgtype": "markdown", "markdown": {"title": title, "text": md_text}}
    resp = requests.post(url, json=payload, timeout=10)
    try:
        data = resp.json()
    except Exception:
        data = {"errcode": resp.status_code, "errmsg": resp.text[:200]}
    return data.get("errcode") == 0, data


def _build_message(items, cfg):
    # 排序：第一关键字=配制人（按 configurators 名单顺序，名单外排最后），
    #       第二关键字=标液类型（工作液 D → 应用液 C → 储备液 B），
    #       末位再按到期日升序，保证组内稳定。
    configurators = cfg.get("configurators") or []
    person_order = {name: i for i, name in enumerate(configurators)}
    type_order = {"D": 0, "C": 1, "B": 2}
    _pidx_fallback = len(person_order)

    def _sort_key(it):
        order = str(it.get("configureOrder") or "")
        pfx = order.split("-")[0].upper()
        person = str(it.get("configuratorName") or it.get("creatorName") or "").strip()
        return (person_order.get(person, _pidx_fallback), person,
                type_order.get(pfx, 99), str(it.get("validityDate") or ""))

    items = sorted(items, key=_sort_key)

    # 日期范围取自实际到期日（最早~最晚），单日则只显示一个日期
    vdates = sorted(str(it.get("validityDate") or "")[:10] for it in items)
    md = [d[5:] for d in vdates if len(d) >= 10]  # -> "MM-DD"
    if not md:
        rng = "?"
    elif md[0] == md[-1]:
        rng = md[0]
    else:
        rng = f"{md[0]} ~ {md[-1]}"
    n = len(items)
    title = f"⏰ LIMS标准溶液即将过期提醒（{n}瓶）"
    lines = [
        f"以下 **{n}** 瓶标液将于 **{rng}** 到期，请尽快处理：",
        "",
    ]
    for i, it in enumerate(items, 1):
        name = it.get("solutionName") or "(未命名)"
        order = str(it.get("configureOrder") or "")
        code = str(it.get("solutionCode") or "")
        person = str(it.get("configuratorName") or it.get("creatorName") or "").strip() or "-"
        vd = str(it.get("validityDate") or "")[:10]      # "YYYY-MM-DD"
        vd_md = vd[5:] if len(vd) >= 10 else ""           # -> "MM-DD"
        lines.append(f"{i}. {_type_label(order)}·**{name}**（{order} · {code}）｜ {person} ｜ {vd_md}")
        lines.append("")
    return title, "\n".join(lines)


def _build_excel_message(rows, min_days, max_days):
    """标准品 Excel 即将过期月度提醒的 markdown。rows 来自 _collect_excel_expiring。"""
    dates = sorted(r["expiry"] for r in rows)
    rng = dates[0].isoformat() if dates[0] == dates[-1] else f"{dates[0].isoformat()} ~ {dates[-1].isoformat()}"
    n = len(rows)
    title = f"📅 标准品即将过期月度提醒（{n}个）"
    lines = [
        f"以下 **{n}** 个标准品将于 **{min_days}~{max_days}天内**（{rng}）到期，请及时确认：",
        "",
    ]
    for i, r in enumerate(rows, 1):
        group = r["group"] or "未分组"
        lab = r["labNo"] or "-"
        name = r["name"] or "(未命名)"
        cas = r["cas"] or "-"
        lines.append(f"{i}. 【{group}】{lab} · **{name}** ｜ CAS号: {cas} ｜ 有效期: {r['expiry'].isoformat()}")
        lines.append("")
    return title, "\n".join(lines)


# ==================== 主流程 ====================
def _eod_warning(cutoff_h, cfg):
    """到截止点仍当天未成功发送 → 发一次最终警告（每天最多1条）。"""
    global _last_eod_alert_date
    today = datetime.date.today()
    if _last_eod_alert_date == today or not cfg.get("alert_on_session_failure", True):
        return
    send_markdown(
        "标液提醒-今日查询失败",
        f"⚠️ **今日标液查询失败，请人工核查**\n\n"
        f"已自动重试至 {cutoff_h}:00 截止，仍未成功推送今日即将过期提醒。",
        cfg,
    )
    _last_eod_alert_date = today
    _save_state()


def run_once():
    """执行一次完整的：取会话→采集→按配制人分组发送到各自钉钉群。
    采集成功（或确认无到期项）→ 标记当天完成（持久化）；
    采集环节失败 → 不标记，由调度器按 retry_interval_minutes 重试。"""
    global _last_run_date
    cfg = _load_cfg()
    if not (cfg.get("notify_groups") or cfg.get("webhook")):
        print(f"{_DINGTALK_LOG} dingtalk 配置缺失，跳过")
        return
    today = datetime.date.today()

    # 取会话 + 采集：任一异常都视为失败（告警一次 + 不标记完成 → 等待重试）
    try:
        system = _acquire_system(cfg)
        if system is None:
            raise RuntimeError("无可用 LIMS 登录会话（内存/磁盘/OCR 均失败）")
        all_items = _collect_expiring(system, cfg)
    except Exception as e:
        print(f"{_DINGTALK_LOG} 执行失败: {e}")
        return

    if not all_items:
        print(f"{_DINGTALK_LOG} 今日无即将过期标液，不发送")
        _last_run_date = today
        _save_state()
        return

    # 按配制人分组，每组发到各自群
    groups = _route_by_group(all_items, cfg)
    if not groups:
        print(f"{_DINGTALK_LOG} 采集到 {len(all_items)} 条，但无配制人落在任一通知组名单内，不发送")
        _last_run_date = today
        _save_state()
        return

    for g, items in groups:
        # _build_message 用 configurators 决定组内排序，故把本组名单注入子 cfg
        sub_cfg = dict(cfg)
        sub_cfg["configurators"] = g.get("configurators") or []
        title, md = _build_message(items, sub_cfg)
        ok, data = send_markdown(title, md, cfg,
                                 webhook=g.get("webhook"), secret=g.get("secret"))
        names = ",".join(g.get("configurators") or [])
        print(f"{_DINGTALK_LOG} 组[{names}] 发送 {len(items)} 条，{'成功' if ok else '失败'}: {data}")

    # 采集成功即标记当天完成：避免重试导致已成功群组重复发送刷屏；
    # 单组发送失败当天接受丢失（见日志），不再重试。
    _last_run_date = today
    _save_state()


def run_excel_notify(min_days=None, max_days=None):
    """执行一次 Excel 月度提醒：遍历 departments，每个启用且配置了 excelPath 的部门
    读其 Excel→筛「今天+min_days~今天+max_days」到期→发到该部门提醒群。
    循环结束即标记当天完成（持久化）：单个部门失败不触发整体重试，避免已发部门重复发。"""
    global _last_excel_notify_date
    cfg = _load_cfg()
    ee = cfg.get("excel_expiry") or {}
    if min_days is None:
        min_days = int(ee.get("min_days", 30))
    if max_days is None:
        max_days = int(ee.get("max_days", 60))
    today = datetime.date.today()

    depts = _departments()
    if not depts:
        print(f"{_DINGTALK_LOG} departments 未配置，跳过 Excel 月度提醒")
        _last_excel_notify_date = today
        _save_state()
        return

    for name, d in depts.items():
        rem = d.get("excelRemind") or {}
        if not rem.get("enabled", True):
            continue
        path = (d.get("excelPath") or "").strip()
        if not path:
            continue
        mn = int(rem.get("min_days", min_days))
        mx = int(rem.get("max_days", max_days))
        webhook = rem.get("webhook") or cfg.get("excel_webhook") or cfg.get("webhook")
        secret = rem.get("secret") or cfg.get("excel_secret") or cfg.get("secret")
        if not webhook or not secret:
            print(f"{_DINGTALK_LOG} Excel 提醒[{name}] 缺少 webhook/secret，跳过")
            continue
        try:
            rows = _collect_excel_expiring(path, today, mn, mx)
        except Exception as e:
            print(f"{_DINGTALK_LOG} Excel 提醒[{name}] 采集失败: {e}")
            continue
        if not rows:
            continue
        title, md = _build_excel_message(rows, mn, mx)
        ok, data = send_markdown(title, md, cfg, webhook=webhook, secret=secret)
        print(f"{_DINGTALK_LOG} Excel 提醒[{name}] 发送 {len(rows)} 条，{'成功' if ok else '失败'}: {data}")

    _last_excel_notify_date = today
    _save_state()


# ==================== 设备使用率 月度提醒 ====================
DEVICE_REMIND_TEXT = '【月度提醒】请于今日下班前完成"设备使用率"数据的提交，谢谢配合'


def _required_device_ids(cfg):
    return [str(x).strip() for x in ((cfg.get("device_usage") or {}).get("required_ids") or [])
            if str(x).strip()]


def run_device_remind():
    """26号后第1个工作日：发固定提交提醒到默认群；成功即标记当天完成。"""
    global _last_device_remind_date
    cfg = _load_cfg()
    if not cfg.get("webhook"):
        print(f"{_DINGTALK_LOG} dingtalk webhook 缺失，跳过设备使用率提醒")
        return
    ok, data = send_markdown("设备使用率-月度提交提醒", DEVICE_REMIND_TEXT, cfg)
    print(f"{_DINGTALK_LOG} 设备使用率提醒发送，{'成功' if ok else '失败'}: {data}")
    if ok:
        _last_device_remind_date = datetime.date.today()
        _save_state()


def run_device_check():
    """26号后第2个工作日：读 device_usage.csv，对照本期(上月27~本月26)找缺失编号并通知。
    全部已提交 → 标记完成不发；发送失败 → 不标记，窗口内自动重试。"""
    global _last_device_check_date
    cfg = _load_cfg()
    required = _required_device_ids(cfg)
    if not required:
        print(f"{_DINGTALK_LOG} device_usage.required_ids 未配置，跳过检查")
        return
    today = datetime.date.today()
    p_start, p_end = _device_period(today)
    import device_usage
    rows = device_usage._load_device_usage()
    present = {r.get("device_id") for r in rows
               if r.get("period_start") == p_start and r.get("period_end") == p_end}
    missing = [i for i in required if i not in present]
    if not missing:
        print(f"{_DINGTALK_LOG} 设备使用率本期({p_start}~{p_end})全部已提交，不发送")
        _last_device_check_date = today
        _save_state()
        return
    md = ("以下编号缺少本期使用率统计信息，请及时补充：\n\n"
          + "\n\n".join(f"- {i}" for i in missing))
    ok, data = send_markdown("设备使用率-本期缺失提醒", md, cfg)
    print(f"{_DINGTALK_LOG} 设备使用率缺失通知({len(missing)}个)发送，{'成功' if ok else '失败'}: {data}")
    if ok:
        _last_device_check_date = today
        _save_state()


def run_inventory_remind():
    """样品入库提醒：查近30天「轻工-释放项目统计」中受理超4小时未入库的样品并通知。
    无未入库项 → 标记本时段完成不发；发送失败 → 不标记，窗口内自动重试。"""
    global _last_inventory_run
    cfg = _load_cfg()
    iv = cfg.get("inventory_remind") or {}
    if not (iv.get("webhook") and iv.get("secret")):
        print(f"{_DINGTALK_LOG} inventory_remind webhook/secret 未配置，跳过入库提醒")
        return
    system = _acquire_system(cfg)
    if not system:
        print(f"{_DINGTALK_LOG} 无可用 LIMS 会话，跳过入库提醒（不标记，下轮重试）")
        return
    try:
        items = _fetch_pending_inventory(system)
    except Exception as e:
        print(f"{_DINGTALK_LOG} 入库提醒查询失败（不标记，下轮重试）: {e}")
        return
    now = datetime.datetime.now()
    slot = max((s for s in _INVENTORY_SLOTS if now.strftime("%H:%M") >= s), default=None)
    if not items:
        print(f"{_DINGTALK_LOG} 近30天无受理超4小时未入库样品，不发送")
        if slot:
            _last_inventory_run = f"{now.date()} {slot}"
            _save_state()
        return
    # 补查每条对应检测单的客服（acceptUserName）
    nos = sorted({str(it.get("detectionNo") or "")[:-3] for it in items} - {""})
    users = _fetch_accept_users(system, nos)
    lines = []
    for it in items:
        segs = [f"受理 {it.get('acceptTime') or ''}",
                f"项目 {it.get('projectName') or ''}"]
        u = users.get(str(it.get("detectionNo") or "")[:-3])
        if u:
            segs.append(f"客服 {u}")
        lines.append(f"- {it.get('detectionNo') or ''} {it.get('sampleName') or ''}"
                     f"（{'，'.join(segs)}）")
    md = (f"### 释放样品入库提醒\n以下 {len(items)} 个样品受理已超过4小时仍未入库，请及时入库：\n\n"
          + "\n\n".join(lines))
    ok, data = send_markdown("释放样品入库提醒", md, cfg,
                             webhook=iv.get("webhook"), secret=iv.get("secret"))
    print(f"{_DINGTALK_LOG} 入库提醒({len(items)}个)发送，{'成功' if ok else '失败'}: {data}")
    if ok and slot:
        _last_inventory_run = f"{now.date()} {slot}"
        _save_state()


def _worker():
    global _last_attempt_dt, _last_excel_attempt_dt, _last_device_attempt_dt
    global _last_inventory_attempt_dt
    print(f"{_DINGTALK_LOG} 调度线程已启动")
    while _flag:
        try:
            cfg = _load_cfg()
            now = datetime.datetime.now()
            today = now.date()
            # 活动窗口 [trigger_hour, retry_until_hour]：准点起、截止点止，之后当天不再尝试
            start_h = int(cfg.get("trigger_hour", 15))
            cutoff_h = int(cfg.get("retry_until_hour", 21))
            start = now.replace(hour=start_h, minute=0, second=0, microsecond=0)
            cutoff = now.replace(hour=cutoff_h, minute=0, second=0, microsecond=0)
            retry_min = int(cfg.get("retry_interval_minutes", 5))
            # 触发：工作日、在活动窗口内、当天未完成、距上次尝试≥重试间隔。
            # 既准点发，也能在错过/失败/重启后补发与重试，到截止点当天停止。
            due = (is_workday(today, cfg)
                   and start <= now <= cutoff
                   and _last_run_date != today
                   and (_last_attempt_dt is None
                        or (now - _last_attempt_dt).total_seconds() >= retry_min * 60))
            if due:
                _last_attempt_dt = now
                run_once()
            elif (is_workday(today, cfg) and now > cutoff
                  and _last_run_date != today):
                # 过了截止点仍当天未发送 → 发一次最终警告（每天1条）
                _eod_warning(cutoff_h, cfg)

            # ===== Excel 月度提醒（与 LIMS 提醒独立）：目标日(默认25号，休息日则前一工作日)
            # 的 [hour, cutoff_hour] 窗口内、当天未完成、距上次尝试≥重试间隔 → 触发。
            # 不复用 elif：两条支线可同日各自触发。
            ee = cfg.get("excel_expiry") or {}
            if ee.get("enabled", True):
                e_day = int(ee.get("day", 25))
                e_hour = int(ee.get("hour", 9))
                e_cutoff = int(ee.get("cutoff_hour", 11))
                e_min = int(ee.get("min_days", 30))
                e_max = int(ee.get("max_days", 60))
                target = _excel_target_date(today.year, today.month, e_day, cfg)
                e_start = now.replace(hour=e_hour, minute=0, second=0, microsecond=0)
                e_end = now.replace(hour=e_cutoff, minute=0, second=0, microsecond=0)
                e_due = (today == target
                         and e_start <= now <= e_end
                         and _last_excel_notify_date != today
                         and (_last_excel_attempt_dt is None
                              or (now - _last_excel_attempt_dt).total_seconds() >= retry_min * 60))
                if e_due:
                    _last_excel_attempt_dt = now
                    run_excel_notify(e_min, e_max)

            # ===== 设备使用率月度提醒（默认群，与 LIMS/Excel 支线独立）：
            #   26号后第1个工作日 [9,12) 发提交提醒；
            #   第2个工作日 [17,21) 读 device_usage.csv 查本期缺失并通知。
            du = cfg.get("device_usage") or {}
            if du.get("enabled", True):
                t1 = _nth_workday_after_26(today.year, today.month, 1, cfg)
                t2 = _nth_workday_after_26(today.year, today.month, 2, cfg)
                dev_retry_ok = (_last_device_attempt_dt is None
                                or (now - _last_device_attempt_dt).total_seconds() >= retry_min * 60)
                r_start = now.replace(hour=9, minute=0, second=0, microsecond=0)
                r_end = now.replace(hour=12, minute=0, second=0, microsecond=0)
                c_start = now.replace(hour=17, minute=0, second=0, microsecond=0)
                c_end = now.replace(hour=21, minute=0, second=0, microsecond=0)
                if (today == t1 and r_start <= now <= r_end
                        and _last_device_remind_date != today and dev_retry_ok):
                    _last_device_attempt_dt = now
                    run_device_remind()
                elif (today == t2 and c_start <= now <= c_end
                      and _last_device_check_date != today and dev_retry_ok):
                    _last_device_attempt_dt = now
                    run_device_check()

            # ===== 样品入库提醒（专用机器人，与上述支线独立）：
            #   工作日 9:00 / 14:30 各查一次近30天受理超4小时未入库样品并提醒。
            #   取"最新已到点时段"：错过 9:00 可补发，9:00 完成后 14:30 自然再触发。
            iv = cfg.get("inventory_remind") or {}
            if iv.get("enabled", True):
                slot = max((s for s in _INVENTORY_SLOTS
                            if now.strftime("%H:%M") >= s), default=None)
                iv_retry_ok = (_last_inventory_attempt_dt is None
                               or (now - _last_inventory_attempt_dt).total_seconds()
                               >= retry_min * 60)
                if (slot and is_workday(today, cfg)
                        and _last_inventory_run != f"{today} {slot}"
                        and iv_retry_ok):
                    _last_inventory_attempt_dt = now
                    run_inventory_remind()
        except Exception as e:
            print(f"{_DINGTALK_LOG} 调度异常: {e}")
        time.sleep(60)


def start():
    """启动后台调度线程（幂等）。应由 login_html.__main__ 在服务子进程调用。"""
    global _flag, _thread
    if _flag:
        return
    _load_state()  # 恢复"今天是否已发"，使重启后判重/补发正确
    _flag = True
    _thread = threading.Thread(target=_worker, daemon=True)
    _thread.start()


# ==================== 手动 / 测试入口 ====================
if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="钉钉标液提醒：手动触发/测试")
    ap.add_argument("--run-now", action="store_true", help="立即执行一次采集+发送")
    ap.add_argument("--run-excel-now", action="store_true", help="立即执行一次 Excel 月度提醒采集+发送")
    ap.add_argument("--run-device-remind", action="store_true", help="立即发送一次设备使用率提交提醒")
    ap.add_argument("--run-device-check", action="store_true", help="立即查 device_usage.csv 本期缺失并通知")
    ap.add_argument("--run-inventory-now", action="store_true", help="立即执行一次样品入库提醒查询+发送")
    ap.add_argument("--test-send", action="store_true", help="发送一条测试消息验证加签通道")
    ap.add_argument("--test-ocr", action="store_true", help="强制走 OCR 登录并验证查询（不发钉钉、不写状态）")
    ap.add_argument("--check-workday", action="store_true", help="打印今天/下一工作日/Excel月度提醒目标日")
    args = ap.parse_args()

    if args.check_workday:
        cfg = _load_cfg()
        t = datetime.date.today()
        ee = cfg.get("excel_expiry") or {}
        e_day = int(ee.get("day", 25))
        print(f"今天 {t} is_workday={is_workday(t, cfg)} 下一工作日={next_workday(t, cfg)}")
        print(f"Excel月度提醒目标日(day={e_day})={_excel_target_date(t.year, t.month, e_day, cfg)}")
        print(f"设备使用率 26号后第1个工作日={_nth_workday_after_26(t.year, t.month, 1, cfg)} "
              f"第2个工作日={_nth_workday_after_26(t.year, t.month, 2, cfg)} "
              f"本期={_device_period(t)}")
    elif args.test_send:
        ok, data = send_markdown("测试-标液提醒通道", "✅ 钉钉标液提醒通道测试成功。")
        print(f"test-send -> ok={ok} data={data}")
    elif args.test_ocr:
        # 强制走 OCR 登录（跳过内存/磁盘会话），验证 ocr_account 能否登录并查询。
        # 不发钉钉、不写状态，纯诊断。
        cfg = _load_cfg()
        acc = cfg.get("ocr_account") or {}
        if not (acc.get("username") and acc.get("password")):
            print(f"{_DINGTALK_LOG} ocr_account 未配置 username/password，无法测试")
            sys.exit(1)
        t0 = time.time()
        try:
            system = lims_auto_login.auto_login(acc["username"], acc["password"])
        except Exception as e:
            print(f"{_DINGTALK_LOG} OCR 登录异常（检查 ddddocr 是否装好 / LIMS 是否可达）: {e}")
            sys.exit(1)
        print(f"{_DINGTALK_LOG} OCR 登录: {'成功' if system else '失败'}（耗时 {time.time()-t0:.1f}s）")
        if not system:
            print(f"{_DINGTALK_LOG} OCR 登录失败：检查 ddddocr 识别 / ocr_account 账号密码是否正确")
            sys.exit(1)
        try:
            items = _fetch_all(system, "SOLUTION_TYPE_D")
            print(f"{_DINGTALK_LOG} 查询验证: SOLUTION_TYPE_D 返回 {len(items)} 条")
            matches = _collect_expiring(system, cfg)
            print(f"{_DINGTALK_LOG} 即将过期且需通知(全部): {len(matches)} 条")
            for g, g_items in _route_by_group(matches, cfg):
                names = ",".join(g.get("configurators") or [])
                print(f"{_DINGTALK_LOG}   组[{names}] -> {len(g_items)} 条")
            print(f"{_DINGTALK_LOG} OK OCR 登录 + 查询 全链路正常")
        except Exception as e:
            print(f"{_DINGTALK_LOG} 登录成功但查询失败: {e}")
            sys.exit(1)
    elif args.run_now:
        run_once()
    elif args.run_excel_now:
        run_excel_notify()
    elif args.run_device_remind:
        run_device_remind()
    elif args.run_device_check:
        run_device_check()
    elif args.run_inventory_now:
        run_inventory_remind()
    else:
        ap.print_help()
