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
import re
import sys
import threading
import time
from urllib.parse import quote_plus

import requests

CONFIG_FILE = "config.json"
STATE_FILE = "dingtalk_state.json"
_DINGTALK_LOG = "[DingTalk]"

# 调度状态（_last_run_date / _last_eod_alert_date 持久化到 STATE_FILE，重启不丢）
_flag = False
_thread = None
_last_run_date = None     # 当天已成功发送（或确认无到期项）→ 当天不再重试/补发
_last_eod_alert_date = None  # 当天已发过"截止查询失败"最终警告 → 每天最多1条
_last_attempt_dt = None   # 上次尝试时刻 → 控制 retry_interval_minutes 重试间隔


# ==================== 配置 / 工作日 ====================
def _load_cfg():
    try:
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f).get("dingtalk") or {}
    except Exception as e:
        print(f"{_DINGTALK_LOG} 读取 config.json 失败: {e}")
        return {}


def _load_state():
    """读取持久化状态（last_run_date / last_eod_alert_date），使重启后判重/补发正确。"""
    global _last_run_date, _last_eod_alert_date
    try:
        with open(STATE_FILE, encoding="utf-8") as f:
            st = json.load(f)
    except Exception:
        return
    for key in ("last_run_date", "last_eod_alert_date"):
        v = st.get(key)
        if not v:
            continue
        try:
            d = datetime.datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
        except Exception:
            continue
        if key == "last_run_date":
            _last_run_date = d
        else:
            _last_eod_alert_date = d


def _save_state():
    """持久化当天状态。"""
    try:
        st = {
            "last_run_date": _last_run_date.isoformat() if _last_run_date else None,
            "last_eod_alert_date": _last_eod_alert_date.isoformat() if _last_eod_alert_date else None,
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


# ==================== LIMS 会话获取 ====================
def _get_lh():
    """取得正在运行的 login_html 模块。
    生产环境 login_html 以 __main__ 运行；此处直接取 __main__，避免二次 import 执行。"""
    m = sys.modules.get("__main__")
    if m and hasattr(m, "RemoteSystem") and hasattr(m, "user_systems"):
        return m
    import login_html  # 独立测试时（login_html 非 __main__）正常导入一次
    return login_html


def _ocr_solve(ocr, img_bytes):
    """算术验证码 A+B：取 ddddocr 文本前两位数字字符求和（屏蔽尾部 '=?' 误读）。"""
    text = ocr.classification(img_bytes)
    digits = [int(c) for c in re.findall(r"\d", text)]
    if len(digits) < 2:
        return None, text
    return digits[0] + digits[1], text


def _ocr_login(system, username, password, max_retry=3):
    """用 ddddocr 识别验证码并登录 system；成功返回 True。"""
    import ddddocr
    from io import BytesIO
    ocr = ddddocr.DdddOcr(show_ad=False)
    for _ in range(max_retry):
        img = system.get_captcha_image()
        if img is None:
            time.sleep(1)
            continue
        buf = BytesIO()
        img.save(buf, format="JPEG")
        total, text = _ocr_solve(ocr, buf.getvalue())
        if total is None:
            time.sleep(1)
            continue
        ok, msg = system.login(username, password, str(total))
        if ok:
            print(f"{_DINGTALK_LOG} OCR 登录成功（识别={text!r} 答案={total}）")
            return True
        # 验证码错会重试；账号错则继续重试无意义但无害
    return False


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

    # 3) OCR 自动登录（专用账号）
    acc = cfg.get("ocr_account") or {}
    if acc.get("username") and acc.get("password"):
        sys_ = RemoteSystem("dt_ocr")
        try:
            if _ocr_login(sys_, acc["username"], acc["password"]):
                return sys_
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


def _type_label(order):
    pfx = str(order or "").split("-")[0].upper()
    return {"B": "储备液", "C": "应用液", "D": "工作液"}.get(pfx, "标液")


def _collect_expiring(system, cfg):
    """返回即将过期、需通知的记录列表。"""
    today = datetime.date.today()
    nw = next_workday(today, cfg)
    configurators = set(cfg.get("configurators") or [])
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
            # 配制人在名单内
            person = str(it.get("configuratorName") or it.get("creatorName") or "").strip()
            if person not in configurators:
                continue
            # 已废弃的跳过
            if it.get("disposeUserName"):
                continue
            matches.append(it)

    # 按到期日升序
    matches.sort(key=lambda x: str(x.get("validityDate") or ""))
    return matches


# ==================== 钉钉发送（加签） ====================
def _sign_url(webhook, secret):
    ts = str(round(time.time() * 1000))
    string_to_sign = f"{ts}\n{secret}"
    hmac_code = hmac.new(secret.encode("utf-8"), string_to_sign.encode("utf-8"),
                         digestmod=hashlib.sha256).digest()
    sign = quote_plus(base64.b64encode(hmac_code))
    return f"{webhook}&timestamp={ts}&sign={sign}"


def send_markdown(title, md_text, cfg=None):
    """加签后发送 markdown；返回 (ok, resp_json)。"""
    cfg = cfg if cfg is not None else _load_cfg()
    webhook = cfg.get("webhook")
    secret = cfg.get("secret")
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
    """执行一次完整的：取会话→采集→发送。
    成功发送 或 确认无到期项 → 标记当天完成（持久化）；
    任一环节失败 → 不标记，由调度器按 retry_interval_minutes 重试。"""
    global _last_run_date
    cfg = _load_cfg()
    if not cfg.get("webhook"):
        print(f"{_DINGTALK_LOG} dingtalk 配置缺失，跳过")
        return
    today = datetime.date.today()

    # 取会话 + 采集：任一异常都视为失败（告警一次 + 不标记完成 → 等待重试）
    try:
        system = _acquire_system(cfg)
        if system is None:
            raise RuntimeError("无可用 LIMS 登录会话（内存/磁盘/OCR 均失败）")
        items = _collect_expiring(system, cfg)
    except Exception as e:
        print(f"{_DINGTALK_LOG} 执行失败: {e}")
        return

    if not items:
        print(f"{_DINGTALK_LOG} 今日无即将过期标液，不发送")
        _last_run_date = today
        _save_state()
        return

    title, md = _build_message(items, cfg)
    ok, data = send_markdown(title, md, cfg)
    print(f"{_DINGTALK_LOG} 发送 {len(items)} 条，{'成功' if ok else '失败'}: {data}")
    if ok:
        _last_run_date = today
        _save_state()
    # 发送失败：不标记完成 → 自动重试


def _worker():
    global _last_attempt_dt
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
    ap.add_argument("--test-send", action="store_true", help="发送一条测试消息验证加签通道")
    ap.add_argument("--test-ocr", action="store_true", help="强制走 OCR 登录并验证查询（不发钉钉、不写状态）")
    ap.add_argument("--check-workday", action="store_true", help="打印今天/下一工作日")
    args = ap.parse_args()

    if args.check_workday:
        cfg = _load_cfg()
        t = datetime.date.today()
        print(f"今天 {t} is_workday={is_workday(t, cfg)} 下一工作日={next_workday(t, cfg)}")
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
        lh = _get_lh()
        system = lh.RemoteSystem("dt_ocr")
        t0 = time.time()
        try:
            ok = _ocr_login(system, acc["username"], acc["password"])
        except Exception as e:
            print(f"{_DINGTALK_LOG} OCR 登录异常（检查 ddddocr 是否装好 / LIMS 是否可达）: {e}")
            sys.exit(1)
        print(f"{_DINGTALK_LOG} OCR 登录: {'成功' if ok else '失败'}（耗时 {time.time()-t0:.1f}s）")
        if not ok:
            print(f"{_DINGTALK_LOG} OCR 登录失败：检查 ddddocr 识别 / ocr_account 账号密码是否正确")
            sys.exit(1)
        try:
            items = _fetch_all(system, "SOLUTION_TYPE_D")
            print(f"{_DINGTALK_LOG} 查询验证: SOLUTION_TYPE_D 返回 {len(items)} 条")
            matches = _collect_expiring(system, cfg)
            print(f"{_DINGTALK_LOG} 即将过期且需通知: {len(matches)} 条")
            print(f"{_DINGTALK_LOG} OK OCR 登录 + 查询 全链路正常")
        except Exception as e:
            print(f"{_DINGTALK_LOG} 登录成功但查询失败: {e}")
            sys.exit(1)
    elif args.run_now:
        run_once()
    else:
        ap.print_help()
