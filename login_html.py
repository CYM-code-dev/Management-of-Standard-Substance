# login_html.py - 完整整合版（修复真实姓名显示 + 更新 Excel 使用情况，动态表头检测）
import sys
import os
import json
import math
import time
import threading
import datetime
import base64
import hashlib
import re
import uuid
from io import BytesIO
from urllib.parse import urlencode

import requests
import urllib3
from flask import Flask, session, jsonify, request, redirect, render_template_string, render_template, send_from_directory
from flask_cors import CORS
from PIL import Image
import openpyxl
from openpyxl.styles import Font, Alignment
import xlrd
import xlwt
from xlutils.copy import copy as xl_copy
from xlutils.styles import Styles

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)
app.secret_key = 'a-fixed-secret-key-for-login-system-2026'
CORS(app, supports_credentials=True)

# 仪器使用率上报（独立模块，开放不认证）
from device_usage import bp as device_usage_bp
app.register_blueprint(device_usage_bp)

# ==================== 配置文件路径 ====================
CONFIG_FILE = 'config.json'
MANUAL_UPLOAD_DIR = 'manual_uploads'

# 导出名称预设迁移锁：防止并发首次访问时多用户同时触发旧格式迁移互相覆盖
_presets_migration_lock = threading.Lock()


# ==================== 数值修约规则 ====================
def _sig_figs(val, n):
    if not isinstance(val, (int, float)) or val == 0 or not math.isfinite(val):
        return 0.0
    magnitude = math.floor(math.log10(abs(val)))
    decimals = max(0, n - 1 - int(magnitude))
    factor = 10 ** decimals
    return round(val * factor) / factor


def _vol_decimals(val):
    n = 4 if val >= 10 else 3
    if val <= 0 or not math.isfinite(val):
        return n - 1
    magnitude = math.floor(math.log10(abs(val)))
    return max(0, n - 1 - int(magnitude))


def _round_vol(val):
    return _sig_figs(val, 4 if val >= 10 else 3)


def _fmt_vol(val):
    return f"{val:.{_vol_decimals(val)}f}"


def _round_qty(val, unit):
    if unit == 'mL':
        return round(val, 3)
    return round(val, 4)


def _fmt_qty(val, unit):
    if unit == 'mL':
        return f"{val:.3f}"
    return f"{val:.4f}"


def _round_conc(val, conc_unit):
    if conc_unit == '%':
        return val
    return round(val, 3) if val < 0.10 else round(val, 2)


def _fmt_conc(val, conc_unit):
    if conc_unit == '%':
        # % 保留全精度自然值（不补零），与前端 solCalcRowConc % 分支一致
        s = f"{val:.10f}" if isinstance(val, float) else str(val)
        if '.' in s:
            int_part, dec_part = s.split('.', 1)
            dec_part = dec_part.rstrip('0')
            return int_part + '.' + dec_part if dec_part else int_part
        return s
    decimals = 3 if val < 0.10 else 2
    return f"{val:.{decimals}f}"


def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # 自动迁移旧格式（扁平 → {defaults, users}）
        if 'defaults' not in data and ('excelPath' in data or 'certPath' in data):
            data = {
                'defaults': {'excelPath': data.get('excelPath', ''), 'certPath': data.get('certPath', '')},
                'users': {}
            }
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        return data
    return {'defaults': {'excelPath': '', 'certPath': ''}, 'users': {}}


def save_config(excel_path, cert_path):
    """保存全局默认路径"""
    config = load_config()
    config['defaults'] = {'excelPath': excel_path, 'certPath': cert_path}
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def get_user_paths(display_name):
    """返回 (excelPath, certPath)，优先用户配置，回退默认值"""
    config = load_config()
    user = config.get('users', {}).get(display_name, {})
    defaults = config.get('defaults', {})
    return (
        user.get('excelPath') or defaults.get('excelPath', ''),
        user.get('certPath') or defaults.get('certPath', '')
    )


def set_user_paths(username, excel_path, cert_path):
    """设置用户专属路径，两个路径都为空时删除该用户条目"""
    config = load_config()
    if 'users' not in config:
        config['users'] = {}
    if excel_path or cert_path:
        config['users'][username] = {'excelPath': excel_path, 'certPath': cert_path}
    else:
        config['users'].pop(username, None)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def _norm_items(arr):
    """规整预设项目 items：每项 {key, exportName}，丢弃无 key 的项"""
    out = []
    if isinstance(arr, list):
        for it in arr:
            if isinstance(it, dict):
                k = it.get('key')
                k = k.strip() if isinstance(k, str) else k
                if k:
                    out.append({'key': k, 'exportName': it.get('exportName', '')})
    return out


def _migrate_presets_to_global(config):
    """把旧的按 display_name 分散的预设（顶层 verifyExportPresets）合并到全局共享池
    verifyExportPresetsGlobal（每个项目带 owner=原 display_name，项目名冲突后者覆盖），
    同时把各用户的 lastProject 记入 verifyExportPresetsLastProject。旧键改名为
    _verifyExportPresetsLegacyBackup（代码级软备份 + 幂等守卫，不直接删除）。返回 config。
    幂等：旧键不存在时直接返回。"""
    legacy = config.get('verifyExportPresets')
    if not isinstance(legacy, dict) or not legacy:
        return config
    glob = config.setdefault('verifyExportPresetsGlobal', {})
    last = config.setdefault('verifyExportPresetsLastProject', {})
    for dn, payload in legacy.items():
        if not isinstance(payload, dict):
            continue
        if payload.get('lastProject'):
            last.setdefault(dn, payload['lastProject'])
        projects = payload.get('projects') or {}
        for name, proj in projects.items():
            if isinstance(proj, list):          # 兼容旧纯数组形式
                items, ignore = proj, []
            elif isinstance(proj, dict):
                items, ignore = proj.get('items', []), proj.get('ignore', [])
            else:
                items, ignore = [], []
            glob[name] = {'owner': dn, 'items': _norm_items(items), 'ignore': ignore or []}
    config['_verifyExportPresetsLegacyBackup'] = legacy
    config.pop('verifyExportPresets', None)
    return config


def get_user_presets(display_name):
    """返回全局共享的导出名称预设：{lastProject(个人), projects(全局,带owner), currentUser}。
    全员可见；编辑权限由前端按 owner 只读控制、后端按 owner 合并兜底。"""
    with _presets_migration_lock:
        config = load_config()
        if config.get('verifyExportPresets'):                # 旧键仍在 → 一次性迁移
            config = _migrate_presets_to_global(config)
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
        glob = config.get('verifyExportPresetsGlobal', {}) or {}
        last = config.get('verifyExportPresetsLastProject', {}) or {}
        projects = {n: {'owner': p.get('owner', ''), 'items': p.get('items', []), 'ignore': p.get('ignore', [])}
                    for n, p in glob.items() if isinstance(p, dict)}
        return {"lastProject": last.get(display_name, ''), "projects": projects, "currentUser": display_name}


def set_user_presets(display_name, presets):
    """合并保存当前用户的预设变更。以服务端全局池为权威：
    他人项目(owner≠dn)钉死保留；我的项目(owner==dn)以提交为准(增/改/删)；
    新建(服务端无此名)仅 owner==dn 或 owner 缺失才接受；伪造他人 owner 的提交忽略。
    管理员例外：可改他人项目内容(items/ignore)，owner 保持原作者；不删除/改名他人项目
    （删除他人项目走专用接口 /export_presets/delete）。他人项目仍一律在步骤1保留，管理员
    的修改在步骤2覆盖，故管理员全量保存不会误删其快照外的他人项目。lastProject 单独按 dn 存。"""
    with _presets_migration_lock:
        config = load_config()
        if config.get('verifyExportPresets'):
            config = _migrate_presets_to_global(config)
        glob = config.get('verifyExportPresetsGlobal', {}) or {}
        last = config.setdefault('verifyExportPresetsLastProject', {})

        submitted = presets.get('projects') if isinstance(presets, dict) else None
        if not isinstance(submitted, dict):
            submitted = {}
        if isinstance(presets, dict) and 'lastProject' in presets:
            last[display_name] = presets.get('lastProject') or ''

        is_admin = (display_name or '').strip() == _ADMIN_DISPLAY_NAME
        new_global = {}
        # 1) 他人项目：原样保留服务端版本（含管理员——兜底防误删；管理员对其的修改在步骤2覆盖）
        for name, sd in glob.items():
            if isinstance(sd, dict) and sd.get('owner') != display_name:
                new_global[name] = {'owner': sd.get('owner', ''), 'items': sd.get('items', []), 'ignore': sd.get('ignore', [])}
        # 2) 提交的项目：我的项目(owner==dn)以提交为准；管理员可改他人项目(items/ignore，owner保持原作者)
        for name, sub in submitted.items():
            if not isinstance(sub, dict):
                continue
            sd = glob.get(name)
            if sd is None:
                # 新建：仅 owner==dn 或 owner 缺失才接受；伪造他人 owner 忽略
                if sub.get('owner', '') == display_name or sub.get('owner', '') == '':
                    new_global[name] = {'owner': display_name, 'items': _norm_items(sub.get('items', [])), 'ignore': sub.get('ignore', []) or []}
            elif isinstance(sd, dict) and (sd.get('owner') == display_name or is_admin):
                # 我的项目以提交为准；管理员改他人项目以提交为准但 owner 保持原作者
                owner = display_name if sd.get('owner') == display_name else sd.get('owner', '')
                new_global[name] = {'owner': owner, 'items': _norm_items(sub.get('items', [])), 'ignore': sub.get('ignore', []) or []}
            # else: 服务端已有且 owner≠dn 且非管理员 → 他人项目，已在步骤1保留，忽略提交
        # 3) 删除：我的项目(owner==dn)在提交中不再出现 → 不写入 new_global（天然删除）。
        #    他人项目即便管理员提交里没有，也已由步骤1保留（删除他人项目须走专用接口）。

        config['verifyExportPresetsGlobal'] = new_global
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)


# ==================== 工作液配置参数预设（按用户私有） ====================
def _norm_ws_preset(p):
    """规整单个工作液预设：标量字段去空白，整数字段校验，takeVolumes 去空串。"""
    if not isinstance(p, dict):
        return None
    def _s(v):
        return str(v).strip() if v is not None else ''
    take = p.get('takeVolumes')
    if not isinstance(take, list):
        take = []
    take = [_s(x) for x in take if _s(x) != '']
    vd = p.get('validityDays')
    dc = p.get('dilutionCount')
    return {
        'validityDays': vd if isinstance(vd, int) else '',
        'storage': _s(p.get('storage')),
        'dilutionCount': dc if isinstance(dc, int) else '',
        'customType': _s(p.get('customType')),
        'medium': _s(p.get('medium')),
        'volume': _s(p.get('volume')),
        'takeVolumes': take,
    }


def get_ws_presets(display_name):
    """返回当前用户私有的工作液配置预设：{lastPreset, presets, currentUser}。"""
    config = load_config()
    allp = config.get('workingSolutionPresets', {}) or {}
    user = allp.get(display_name, {}) or {}
    raw = user.get('presets', {}) or {}
    presets = {n: np for n, p in raw.items() if (np := _norm_ws_preset(p))}
    return {"lastPreset": user.get('lastPreset', '') or '', "presets": presets, "currentUser": display_name}


def set_ws_presets(display_name, data):
    """全量保存当前用户的工作液预设（私有，无跨用户合并）。"""
    config = load_config()
    allp = config.setdefault('workingSolutionPresets', {})
    user = allp.setdefault(display_name, {})
    user['lastPreset'] = (data.get('lastPreset') or '') if isinstance(data, dict) else ''
    presets = data.get('presets') if isinstance(data, dict) else None
    clean = {}
    if isinstance(presets, dict):
        for name, p in presets.items():
            name = (name or '').strip()
            if not name:
                continue
            np = _norm_ws_preset(p)
            if np:
                clean[name] = np
    user['presets'] = clean
    config['workingSolutionPresets'] = allp
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


# ==================== 远程系统连接类 ====================
def md5_1024_times(text: str) -> str:
    current = text.encode('utf-8')
    for _ in range(1024):
        current = hashlib.md5(current).hexdigest().encode('utf-8')
    return current.decode()


class RemoteSystem:
    def __init__(self, sess_id, user_agent=None):
        self.sess_id = sess_id
        self.session = requests.Session()
        self.base_url = "http://192.168.12.234:60015"
        self.session.verify = False
        self.current_user = None
        self.current_pid = None
        self.current_real_name = None
        self.captcha_session = None
        self.keep_alive_flag = False
        self.keep_alive_thread = None
        self.keep_alive_interval = 3600
        self.user_agent = user_agent
        self.update_headers()

    def set_user_agent(self, ua_string):
        self.user_agent = ua_string
        self.update_headers()

    def update_headers(self):
        default_ua = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36"
        self.headers = {
            "User-Agent": self.user_agent or default_ua,
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": self.base_url,
            "Referer": f"{self.base_url}/web/login.html",
            "Connection": "keep-alive",
            "Host": "192.168.12.234:60015"
        }
        self.session.headers.update(self.headers)

    def get_captcha_image(self):
        try:
            temp_session = requests.Session()
            temp_session.verify = False
            temp_session.headers.update(self.headers)
            ts = str(int(time.time() * 1000))
            url = f"{self.base_url}/detectionManager/core/security/validatecodes?{ts}&r={ts}"
            resp = temp_session.get(url)
            if resp.status_code == 200:
                self.captcha_session = temp_session
                return Image.open(BytesIO(resp.content))
            return None
        except Exception as e:
            print(f"获取验证码失败: {e}")
            return None

    def login(self, username, plain_password, captcha):
        password_hash = md5_1024_times(plain_password)
        login_data = {"account": username, "password": password_hash, "validCode": captcha}
        try:
            if self.captcha_session:
                resp = self.captcha_session.post(f"{self.base_url}/detectionManager/core/security/login", data=login_data)
                if resp.status_code == 200:
                    for cookie in self.captcha_session.cookies:
                        self.session.cookies.set(cookie.name, cookie.value)
            else:
                resp = self.session.post(f"{self.base_url}/detectionManager/core/security/login", data=login_data)
            if resp.status_code != 200:
                return False, f"请求失败，状态码: {resp.status_code}"
            result = resp.json()
            if result.get("success"):
                self.current_user = username
                real_name = self._fetch_user_info()
                if not real_name:
                    real_name = result.get("resultData", {}).get("nickName", "")
                if not real_name:
                    return False, "登录失败：无法获取用户信息"
                self.current_real_name = real_name
                self._save_session()
                self.start_keep_alive()
                return True, real_name
            else:
                error_msg = result.get("errorCtx", {}).get("errorMsg", "登录失败")
                return False, "验证码错误" if "验证码" in error_msg else error_msg
        except Exception as e:
            return False, f"登录异常: {str(e)}"

    def _fetch_user_info(self):
        try:
            resp = self.session.get(f"{self.base_url}/detectionManager/core/security/getLoginUser")
            if resp.status_code == 200:
                data = resp.json()
                if data.get("success"):
                    user_info = data.get("resultData", {}).get("userInfo", {})
                    pid = user_info.get("id")
                    real_name = user_info.get("realName")
                    if pid:
                        self.current_pid = str(pid)
                    if real_name:
                        return real_name
            resp2 = self.session.get(f"{self.base_url}/detectionManager/core/users/info")
            if resp2.status_code == 200:
                data2 = resp2.json()
                if data2.get("success"):
                    user_info = data2.get("resultData", {}).get("userInfo", {})
                    pid = user_info.get("id")
                    real_name = user_info.get("realName")
                    if pid:
                        self.current_pid = str(pid)
                    if real_name:
                        return real_name
        except Exception as e:
            print(f"获取用户信息异常: {e}")
        return None

    def _save_session(self):
        sess_file = f"session_{self.current_user}.json"
        cookies_list = [{"name": c.name, "value": c.value, "domain": c.domain, "path": c.path} for c in self.session.cookies]
        info = {
            "username": self.current_user, "pid": self.current_pid, "real_name": self.current_real_name,
            "login_time": time.strftime("%Y-%m-%d %H:%M:%S"),
            "cookies": cookies_list, "headers": dict(self.session.headers)
        }
        with open(sess_file, "w", encoding="utf-8") as f:
            json.dump(info, f, indent=2)

    def load_session(self):
        if not self.current_user: return False
        sess_file = f"session_{self.current_user}.json"
        if not os.path.exists(sess_file): return False
        try:
            with open(sess_file, 'r', encoding='utf-8') as f:
                info = json.load(f)
            if time.time() - time.mktime(time.strptime(info["login_time"], "%Y-%m-%d %H:%M:%S")) > 24*3600:
                os.remove(sess_file)
                return False
            self.session.cookies.clear()
            for c in info["cookies"]:
                kw = {"name": c["name"], "value": c["value"]}
                if c.get("domain"): kw["domain"] = c["domain"]
                if c.get("path"): kw["path"] = c["path"]
                self.session.cookies.set(**kw)
            self.session.headers.update(info["headers"])
            self.current_user = info["username"]
            self.current_pid = info.get("pid")
            self.current_real_name = info.get("real_name", info["username"])
            return True
        except:
            return False

    def verify_session(self):
        if not self.current_user: return False
        # /getLoginUser 在部分 LIMS 上不可用（HTTP 500「资源不存在」），改以 /users/info 为主探针；
        # 两端点任一返回 success 即视为存活，与 _fetch_user_info 的探测保持一致
        for ep in ("/detectionManager/core/users/info", "/detectionManager/core/security/getLoginUser"):
            try:
                resp = self.session.get(f"{self.base_url}{ep}")
                if resp.status_code != 200: continue
                data = resp.json()
                if data.get('success'):
                    err_ctx = data.get('errorCtx') or {}
                    if str(err_ctx.get('errorCode')) == '401': continue
                    return True
            except:
                continue
        return False

    def logout(self):
        self.stop_keep_alive()
        if self.current_user:
            sess_file = f"session_{self.current_user}.json"
            if os.path.exists(sess_file): os.remove(sess_file)
        self.current_user = self.current_pid = self.current_real_name = None
        self.session.cookies.clear()

    def start_keep_alive(self):
        if not self.should_keep_alive() or self.keep_alive_flag: return
        self.keep_alive_flag = True
        self.keep_alive_thread = threading.Thread(target=self._keep_alive_worker, daemon=True)
        self.keep_alive_thread.start()

    def stop_keep_alive(self):
        self.keep_alive_flag = False
        if self.keep_alive_thread: self.keep_alive_thread.join(timeout=1)

    def should_keep_alive(self):
        return datetime.datetime.now().hour < 20

    def _keep_alive_worker(self):
        while self.keep_alive_flag and self.should_keep_alive() and self.current_user:
            try:
                if not self.verify_session():
                    print(f"[KeepAlive] LIMS会话已失效，停止保活")
                    self.keep_alive_flag = False
                    break
            except Exception as e:
                print(f"[KeepAlive] 会话检查异常（将继续重试）: {e}")
            for _ in range(self.keep_alive_interval):
                if not self.keep_alive_flag: break
                time.sleep(1)


user_systems = {}
def get_system():
    sess_id = session.get('sess_id')
    if not sess_id:
        sess_id = os.urandom(16).hex()
        session['sess_id'] = sess_id
    if sess_id not in user_systems:
        user_systems[sess_id] = RemoteSystem(sess_id)
    return user_systems[sess_id]


# ==================== 工作液溯源服务 API（供 mup-web 溯源填充调用）====================
# OCR 服务账号登录 LIMS，模块级缓存 12h；不依赖 Flask 用户会话，可跨服务调用。
_TRACE_SESS = {"res": None, "ts": 0}
_TRACE_SESS_TTL = 43200  # 12h


def _trace_service_session():
    """返回 (system, err)：system 为 lims_auto_login 登录后的 LoginResult
   （鸭类型兼容 RemoteSystem：.session/.base_url/.current_pid/.current_real_name）。"""
    now = time.time()
    if _TRACE_SESS["res"] and now - _TRACE_SESS["ts"] < _TRACE_SESS_TTL:
        return _TRACE_SESS["res"], None
    from lims_auto_login import auto_login
    cfg = (load_config().get("dingtalk") or {}).get("ocr_account") or {}
    user = str(cfg.get("username") or "").strip()
    pwd = str(cfg.get("password") or "").strip()
    if not (user and pwd):
        return None, "未配置溯源服务账号(config.json: dingtalk.ocr_account)"
    try:
        res = auto_login(user, pwd)
    except Exception as e:
        return None, f"LIMS 登录异常: {e}"
    if not res:
        return None, "LIMS 登录失败(验证码/账密), 请稍后重试"
    _TRACE_SESS["res"], _TRACE_SESS["ts"] = res, now
    return res, None


# ==================== 会话过期统一处理 ====================
# 远程 LIMS 会话存活缓存（秒）：避免高频 /api/status 轮询打爆 LIMS
_SESSION_CHECK_TTL = 20
_session_check_cache = {}            # {username: (timestamp, alive)}
_session_check_lock = threading.Lock()


def remote_session_alive(system, username):
    """带短缓存的远程会话存活探测。False 时立即失效缓存，便于重登后重探。"""
    now = time.time()
    with _session_check_lock:
        hit = _session_check_cache.get(username)
    if hit and (now - hit[0]) < _SESSION_CHECK_TTL:
        return hit[1]
    # 远程系统可能尚未载入（如服务重启后内存丢失）：按账号从磁盘恢复，再校验，避免误判过期
    if username and (not system.current_user or system.current_user != username):
        system.current_user = username
        system.load_session()
    alive = bool(system.current_user) and system.verify_session()
    with _session_check_lock:
        _session_check_cache[username] = (now, alive)
    if not alive:
        with _session_check_lock:
            _session_check_cache.pop(username, None)
    return alive


def _expired_response(msg="远程会话已失效，请重新登录"):
    """统一过期信号：HTTP 401 + 哨兵 code + 中文提示。"""
    return jsonify({"success": False, "code": "SESSION_EXPIRED", "message": msg}), 401


def _resp_looks_expired(resp):
    """识别 LIMS 过期标记：HTML 登录页 / errorCtx.errorCode==401 / 未登录文案。"""
    try:
        ct = resp.headers.get('Content-Type', '')
        if 'html' in ct.lower():
            return True
        body = resp.text.lstrip()
        if body.startswith('<!') or body.startswith('<html'):
            return True
        data = resp.json()
        if isinstance(data, dict):
            err = data.get('errorCtx') or {}
            if isinstance(err, dict) and err.get('errorCode') in ('401', 401):
                return True
            msg = (data.get('errorDesc') or '') + (data.get('message') or '')
            if '未登录' in msg or 'login' in msg.lower():
                return True
    except Exception:
        pass
    return False


# ==================== 路由 ====================
@app.route('/')
def index():
    if not session.get('logged_in'): return redirect('/login')
    return redirect('/solution-config')

@app.route('/login')
def login_page():
    error = request.args.get('error', '')
    username = request.args.get('username', '')
    system = get_system()
    img = system.get_captcha_image()
    captcha_url = ''
    if img:
        buffered = BytesIO()
        img.save(buffered, format="PNG")
        captcha_url = f"data:image/png;base64,{base64.b64encode(buffered.getvalue()).decode()}"
    return render_template('login.html', error=error, error_json=json.dumps(error), username=username,
                           captcha_url=captcha_url)

@app.route('/organic-std')
def organic_std_page():
    return render_template('OrganicStd.html')

@app.route('/solution-config')
def solution_config_page():
    return render_template('SolutionConfig.html')

@app.route('/api/captcha')
def get_captcha():
    system = get_system()
    img = system.get_captcha_image()
    if img:
        buffered = BytesIO()
        img.save(buffered, format="PNG")
        return jsonify({"success": True, "image": f"data:image/png;base64,{base64.b64encode(buffered.getvalue()).decode()}"})
    return jsonify({"success": False, "message": "获取验证码失败"})

@app.route('/api/login', methods=['POST'])
def login():
    if request.is_json:
        data = request.json
        username = data.get('username'); plain_password = data.get('password'); captcha = data.get('captcha'); client_ua = data.get('client_ua'); is_json = True
    else:
        username = request.form.get('username'); plain_password = request.form.get('password'); captcha = request.form.get('captcha'); client_ua = request.form.get('client_ua'); is_json = False

    if not is_json and session.get('logged_in'):
        next_page = request.form.get('next', '/')
        return redirect(next_page)

    if not username or not plain_password or not captcha:
        if is_json: return jsonify({"success": False, "message": "账号、密码和验证码不能为空"})
        else: return redirect(f'/login?error=请填写完整信息&username={username}')

    system = get_system()
    if client_ua: system.set_user_agent(client_ua)

    success, display_name = system.login(username, plain_password, captcha)

    if success:
        session['logged_in'] = True
        session['username'] = username
        session['display_name'] = display_name
        session['pid'] = system.current_pid
        if is_json: return jsonify({"success": True, "message": "登录成功", "display_name": display_name})
        else: return redirect(request.form.get('next', '/'))
    else:
        if is_json: return jsonify({"success": False, "message": display_name})
        else: return redirect(f'/login?error={display_name}&username={username}')
@app.route('/api/status')
def status():
    if session.get('logged_in'):
        system = get_system()
        username = session.get('username', '')
        if not remote_session_alive(system, username):
            session.pop('logged_in', None)
            system.logout()
            return jsonify({"logged_in": False, "code": "SESSION_EXPIRED"})
        return jsonify({"logged_in": True, "username": username, "display_name": session.get('display_name'), "pid": session.get('pid')})
    system = get_system()
    if system.current_user and system.verify_session():
        session['logged_in'] = True
        session['username'] = system.current_user
        session['display_name'] = system.current_real_name or system.current_user
        session['pid'] = system.current_pid
        return jsonify({"logged_in": True, "username": system.current_user, "display_name": session['display_name'], "pid": session.get('pid')})
    return jsonify({"logged_in": False})


# ==================== 节假日/调休（按年本地缓存 + 缺失才外部查询）====================
# 数据源：timor.tech/api/holiday/year（需带 User-Agent，否则 403）
# 返回 { 'YYYY-MM-DD': true(放假) | false(周末补班) }，仅含“异常日期”，未收录按周末判断
HOLIDAYS_CACHE_DIR = "holidays_cache"
_HOLIDAYS_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"


def _holiday_cache_path(year):
    return os.path.join(HOLIDAYS_CACHE_DIR, f"{year}.json")


def _load_holiday_cache(year):
    """读本地年度缓存，返回 {date: bool} 或 None。"""
    try:
        with open(_holiday_cache_path(year), encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return None


def _save_holiday_cache(year, mapping):
    try:
        os.makedirs(HOLIDAYS_CACHE_DIR, exist_ok=True)
        with open(_holiday_cache_path(year), "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, sort_keys=True)
    except Exception as e:
        print(f"[Holidays] 写缓存失败 year={year}: {e}")


def _fetch_holidays_from_timor(year):
    """请求 timor.tech 年度接口，返回 {date: bool}（true=放假, false=补班）或 None。"""
    try:
        r = requests.get(f"https://timor.tech/api/holiday/year/{year}-01-01",
                         timeout=15, headers={"User-Agent": _HOLIDAYS_UA})
        if r.status_code != 200:
            print(f"[Holidays] timor.tech HTTP {r.status_code} year={year}")
            return None
        j = r.json()
        if j.get("code") != 0:
            print(f"[Holidays] timor.tech code={j.get('code')} year={year}")
            return None
        mapping = {}
        for info in (j.get("holiday") or {}).values():
            ds = info.get("date")
            if ds:
                mapping[ds] = bool(info.get("holiday"))
        return mapping or None
    except Exception as e:
        print(f"[Holidays] 请求异常 year={year}: {e}")
        return None


def get_holidays(year):
    """按年获取节假日映射：本地缓存优先，缺失则请求 timor.tech 并写回本地。
    返回 (mapping, source)：mapping={date:bool} 或 None；source ∈ cache/fresh/none。"""
    mapping = _load_holiday_cache(year)
    if mapping is not None:
        return mapping, "cache"
    mapping = _fetch_holidays_from_timor(year)
    if mapping is None:
        return None, "none"
    _save_holiday_cache(year, mapping)
    return mapping, "fresh"


@app.route('/api/holidays/<int:year>', methods=['GET'])
def api_holidays(year):
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    if year < 2000 or year > 2100:
        return jsonify({"success": False, "message": "年份无效"}), 400
    mapping, source = get_holidays(year)
    return jsonify({"success": True, "year": year, "holidays": mapping or {}, "source": source})


@app.route('/holidays/<int:year>.json')
def holidays_file(year):
    """开放（无需登录）：直接返回 holidays_cache/<year>.json 原文件，
    供仪器使用率统计等内部 LAN 工具读取。缓存缺失时自动拉取并落盘。"""
    if year < 2000 or year > 2100:
        return "", 400
    get_holidays(year)
    return send_from_directory(os.path.abspath(HOLIDAYS_CACHE_DIR), f"{year}.json")


@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    get_system().logout()
    return jsonify({"success": True})

@app.route('/api/query')
def query():
    if not session.get('logged_in'): return jsonify({"success": False, "message": "未登录"}), 401
    pid = session.get('pid')
    username = session.get('username')
    if not pid:
        system = get_system()
        if system.current_pid: pid = session['pid'] = system.current_pid
        else: return _expired_response("无法获取用户PID，请重新登录")
    system = get_system()
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pageNo = int(request.args.get('pageNo', 1))
    pageSize = int(request.args.get('pageSize', 30))
    keyword = request.args.get('keyword', '').strip()
    casNo = request.args.get('casNo', '').strip()
    org_name = request.args.get('org_name', '').strip()
    params = {
        "_search": "false", "nd": str(int(time.time()*1000)), "pageSize": pageSize, "pageNo": pageNo, "sidx": "", "sord": "asc",
        "type": "CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE",
        "receiveUserName": "", "receiveStartDate": "", "receiveEndDate": "",
        "confirmUserName": "", "confirmStartDate": "", "confirmEndDate": "",
        "invoiceNo": "", "groupId": "",
        "status": request.args.get('status', 'normal'),
        "pid": pid, "pname": username, "loginId": pid
    }
    if keyword:
        params["keyword"] = keyword
    if casNo:
        params["casNo"] = casNo
    if org_name:
        params["orgName"] = org_name
    try:
        resp = system.session.get(f"{system.base_url}/detectionManager/manager/consumableBill/pageObj", params=params)
        if resp.status_code != 200:
            if resp.status_code in (401,403): session.pop('logged_in', None); system.logout(); return _expired_response()
            return jsonify({"success": False, "message": f"请求失败，状态码: {resp.status_code}"})
        result = resp.json()
        rd = result.get("resultData") or {}
        if result.get("success"):
            vo_list = rd.get("voList", [])
            records = rd.get("records") or rd.get("totalCount") or rd.get("totalRecords") or 0
            page = rd.get("page", 1)
            total = rd.get("total")
            if total is None or total <= 0:
                total_page = rd.get("totalPage") or rd.get("totalPages")
                if total_page:
                    total = total_page
                elif records > 0:
                    total = math.ceil(records / pageSize)
                else:
                    total = 1
            return jsonify({"success": True, "data": vo_list, "records": records, "page": page, "total": total})
        error_msg = (result.get("errorCtx") or {}).get("errorMsg", "查询失败")
        if "未登录" in error_msg or "login" in error_msg.lower(): session.pop('logged_in', None); system.logout(); return _expired_response()
        return jsonify({"success": False, "message": error_msg})
    except Exception as e:
        return jsonify({"success": False, "message": f"查询异常: {str(e)}"})

@app.route('/api/config', methods=['GET', 'POST'])
def handle_config():
    if request.method == 'GET':
        excel_path, cert_path = get_user_paths(session.get('display_name', ''))
        return jsonify({"excelPath": excel_path, "certPath": cert_path})
    data = request.get_json()
    excel_path = data.get('excelPath', '').strip()
    cert_path = data.get('certPath', '').strip()
    save_config(excel_path, cert_path)
    excel_ok = not excel_path or os.path.exists(os.path.dirname(excel_path))
    cert_ok = not cert_path or os.path.exists(cert_path)
    if excel_ok and cert_ok: return jsonify({"success": True})
    msg = []
    if not excel_ok: msg.append("Excel 路径不可访问")
    if not cert_ok: msg.append("证书路径不可访问")
    return jsonify({"success": False, "message": "；".join(msg)}), 400


# ==================== 管理员：用户路径管理 ====================
@app.route('/api/admin/paths', methods=['GET'])
def admin_get_paths():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    if (session.get('display_name') or '').strip() != _ADMIN_DISPLAY_NAME:
        return jsonify({"success": False, "message": "无权限"}), 403
    config = load_config()
    return jsonify({"defaults": config.get('defaults', {}), "users": config.get('users', {})})


@app.route('/api/admin/paths', methods=['POST'])
def admin_set_paths():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    if (session.get('display_name') or '').strip() != _ADMIN_DISPLAY_NAME:
        return jsonify({"success": False, "message": "无权限"}), 403
    data = request.get_json()
    username = data.get('username', '').strip()
    excel_path = data.get('excelPath', '').strip()
    cert_path = data.get('certPath', '').strip()
    if not username:
        return jsonify({"success": False, "message": "用户名不能为空"}), 400
    set_user_paths(username, excel_path, cert_path)
    return jsonify({"success": True})


@app.route('/api/admin/paths/<username>', methods=['DELETE'])
def admin_delete_paths(username):
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    if (session.get('display_name') or '').strip() != _ADMIN_DISPLAY_NAME:
        return jsonify({"success": False, "message": "无权限"}), 403
    config = load_config()
    config.get('users', {}).pop(username, None)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
    return jsonify({"success": True})


# ==================== 标液期间核查：导出名称预设 ====================
@app.route('/api/lims/export_presets', methods=['GET'])
def lims_get_export_presets():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    dn = (session.get('display_name') or '').strip()
    if not dn:
        return jsonify({"success": False, "message": "无显示名"}), 400
    return jsonify({"success": True, "presets": get_user_presets(dn)})


@app.route('/api/lims/export_presets', methods=['POST'])
def lims_set_export_presets():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    dn = (session.get('display_name') or '').strip()
    if not dn:
        return jsonify({"success": False, "message": "无显示名"}), 400
    data = request.get_json() or {}
    presets = data.get('presets')
    if not isinstance(presets, dict) or 'projects' not in presets:
        return jsonify({"success": False, "message": "presets 格式错误"}), 400
    set_user_presets(dn, presets)
    return jsonify({"success": True})


@app.route('/api/lims/export_presets/delete', methods=['POST'])
def lims_delete_export_preset_project():
    """删除指定预设项目：仅 owner==自己 可删；管理员可删任意项目。不碰其他项目。
    供前端「删除项目」即时生效用——避免全量保存顺带提交窗口内其他未保存编辑。"""
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    dn = (session.get('display_name') or '').strip()
    if not dn:
        return jsonify({"success": False, "message": "无显示名"}), 400
    data = request.get_json() or {}
    name = (data.get('project') or '').strip()
    if not name:
        return jsonify({"success": False, "message": "未指定项目名"}), 400
    with _presets_migration_lock:
        config = load_config()
        if config.get('verifyExportPresets'):
            config = _migrate_presets_to_global(config)
        glob = config.get('verifyExportPresetsGlobal', {}) or {}
        proj = glob.get(name)
        if isinstance(proj, dict) and (proj.get('owner') == dn or dn == _ADMIN_DISPLAY_NAME):
            del glob[name]
        elif proj is None:
            pass  # 项目不存在，视为已删（幂等）
        else:
            return jsonify({"success": False, "message": "无权删除该项目（他人项目只读）"}), 403
        # 清理所有用户 lastProject 中指向该项目的值（管理员删他人项目时，原作者等人的偏好也要清）
        lp = config.get('verifyExportPresetsLastProject', {})
        if isinstance(lp, dict):
            for k, v in list(lp.items()):
                if v == name:
                    lp[k] = ''
        config['verifyExportPresetsGlobal'] = glob
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    return jsonify({"success": True})


# ==================== 工作液配置参数预设（按用户私有） ====================
@app.route('/api/lims/ws_presets', methods=['GET'])
def lims_get_ws_presets():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    dn = (session.get('display_name') or '').strip()
    if not dn:
        return jsonify({"success": False, "message": "无显示名"}), 400
    return jsonify({"success": True, "presets": get_ws_presets(dn)})


@app.route('/api/lims/ws_presets', methods=['POST'])
def lims_set_ws_presets():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    dn = (session.get('display_name') or '').strip()
    if not dn:
        return jsonify({"success": False, "message": "无显示名"}), 400
    # 前端发送 {presets: {lastPreset, presets:{...}, currentUser}}，取内层对象
    payload = request.get_json() or {}
    set_ws_presets(dn, payload.get('presets') if isinstance(payload.get('presets'), dict) else {})
    return jsonify({"success": True})


# ==================== 手动导入 Excel ====================
@app.route('/api/upload_excel', methods=['POST'])
def upload_excel():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    file = request.files.get('file')
    if not file:
        return jsonify({"success": False, "message": "未选择文件"}), 400
    os.makedirs(MANUAL_UPLOAD_DIR, exist_ok=True)
    username = session.get('username', 'anonymous')
    ext = os.path.splitext(file.filename)[1] or '.xlsx'
    save_path = os.path.join(MANUAL_UPLOAD_DIR, f'{username}{ext}')
    file.save(save_path)
    session['manual_excel_path'] = save_path
    return jsonify({"success": True, "path": save_path})


@app.route('/api/clear_manual_excel', methods=['POST'])
def clear_manual_excel():
    session.pop('manual_excel_path', None)
    return jsonify({"success": True})


@app.route('/manual_excel/<path:filename>')
def serve_manual_excel(filename):
    return send_from_directory(MANUAL_UPLOAD_DIR, filename)


@app.route('/organic_excel/<path:filename>')
def serve_excel(filename):
    excel_path, _ = get_user_paths(session.get('display_name', ''))
    excel_dir = os.path.dirname(excel_path)
    if not excel_dir:
        return jsonify({"success": False, "message": "未配置 Excel 路径"}), 400
    return send_from_directory(excel_dir, filename)

@app.route('/certificates/<path:filename>')
def serve_cert(filename):
    _, cert_dir = get_user_paths(session.get('display_name', ''))
    if not cert_dir:
        return jsonify({"success": False, "message": "未配置证书路径"}), 400
    return send_from_directory(cert_dir, filename)


# ==================== 智能插入写入 Excel ====================
_KNOWN_SHEET_NAMES = ['有机标准物质', 'FCM标准品']

def detect_sheet_name(wb):
    """从 workbook 中自动检测目标 sheet 名"""
    for name in _KNOWN_SHEET_NAMES:
        if name in wb.sheetnames:
            return name
    # 都没有则返回第一个 sheet
    return wb.sheetnames[0] if wb.sheetnames else None

def get_prefix_type(original_id):
    if not original_id: return 'number'
    first = original_id[0].upper()
    if first == 'A': return 'A'     # FCM 室温
    if first == 'B': return 'B'     # FCM 4℃
    if first == 'C': return 'C'     # FCM -18℃
    if first == 'D': return 'D'     # 有机 D
    if first == 'E': return 'E'     # 有机 E
    return 'number'

def extract_number(original_id):
    if not original_id: return 0
    m = re.search(r'\d+', original_id)
    return int(m.group()) if m else 0

def is_row_empty(row_values):
    return all(v is None or str(v).strip() == '' for v in row_values)

def parse_existing_records_xlsx(ws):
    records = []
    for row in ws.iter_rows(min_row=2):
        row_vals = [c.value for c in row]
        if is_row_empty(row_vals): continue
        original_id = str(row_vals[0]) if row_vals[0] is not None else ''
        lab_no = str(row_vals[2]) if len(row_vals) > 2 and row_vals[2] is not None else ''
        records.append({'row_index': row[0].row, 'original_id': original_id, 'lab_no': lab_no})
    return records

def parse_existing_records_xls(sheet):
    records = []
    for row_idx in range(1, sheet.nrows):
        row_vals = [sheet.cell_value(row_idx, c) for c in range(sheet.ncols)]
        if is_row_empty(row_vals): continue
        original_id = str(row_vals[0]) if row_vals[0] else ''
        lab_no = str(row_vals[2]) if len(row_vals) > 2 and row_vals[2] else ''
        records.append({'row_index': row_idx, 'original_id': original_id, 'lab_no': lab_no})
    return records


# 标准品 Excel「即将过期」提醒所需字段别名（与前端 OrganicStd.html 的 colMap 保持一致，
# 按表头名定位列号，不依赖固定列顺序）
_EXPIRY_FIELD_ALIASES = {
    'group': ['组别'],
    'labNo': ['实验室编号', '内部编号'],
    'name': ['标品名称', '名称'],
    'cas': ['CAS号', 'CAS'],
    'expiry': ['有效期', '效期', '有效期至'],
    'usage': ['使用情况', '使用'],
}


def _map_headers_to_idx(headers):
    """按 _EXPIRY_FIELD_ALIASES 把表头映射成 {field: 列号}。"""
    norm = {}
    for col, h in enumerate(headers):
        if h is None:
            continue
        norm[str(h).strip()] = col
    idx = {}
    for field, names in _EXPIRY_FIELD_ALIASES.items():
        for nm in names:
            if nm in norm:
                idx[field] = norm[nm]
                break
    return idx


def _find_header_row(peeked):
    """在 peeked（前若干行）中找出表头行：匹配已知列别名最多(≥3)的那行。
    返回 (在 peeked 中的下标, {field: 列号})；找不到返回 (None, {})。
    用于跳过标题行（如首行「有机标品存放记录表」，真正表头在第 2 行）。"""
    best_pos, best_idx, best_score = None, {}, 0
    for pos, row in enumerate(peeked):
        idx = _map_headers_to_idx(row)
        if len(idx) > best_score:
            best_score, best_pos, best_idx = len(idx), pos, idx
    if best_score >= 3:
        return best_pos, best_idx
    return None, {}


def read_excel_expiry_rows(path):
    """读取标准品 Excel，按【表头名】映射返回每行提醒所需字段（不依赖固定列号/行号）。

    供 dingtalk_notify「标准品即将过期」月度提醒使用：只读提醒需要的列，纯读取不改文件。
    自动跳过标题行：在前 12 行里找出真正的表头行（匹配已知列别名最多者）。
    返回 [{group, labNo, name, cas, expiry_raw, row_index}, ...]
    expiry_raw 为单元格原值（datetime / str / int / None），日期解析交给调用方。
    文件不存在/打不开时抛异常，由调用方捕获后当天继续重试。
    """
    def _str(v):
        return str(v).strip() if v is not None else ''

    def _build(vals, idx, exp_col, row_no, exp_raw):
        return {
            'group': _str(vals[idx['group']]) if 'group' in idx and idx['group'] < len(vals) else '',
            'labNo': _str(vals[idx['labNo']]) if 'labNo' in idx and idx['labNo'] < len(vals) else '',
            'name': _str(vals[idx['name']]) if 'name' in idx and idx['name'] < len(vals) else '',
            'cas': _str(vals[idx['cas']]) if 'cas' in idx and idx['cas'] < len(vals) else '',
            'usage': _str(vals[idx['usage']]) if 'usage' in idx and idx['usage'] < len(vals) else '',
            'expiry_raw': exp_raw,
            'row_index': row_no,
        }

    rows = []
    low = path.lower()
    if low.endswith('.xls') and not low.endswith('.xlsx'):
        book = xlrd.open_workbook(path)
        sheet = None
        for nm in _KNOWN_SHEET_NAMES:
            try:
                sheet = book.sheet_by_name(nm); break
            except Exception:
                continue
        if sheet is None:
            sheet = book.sheet_by_index(0)
        if sheet.nrows < 2:
            return rows
        peeked = [[sheet.cell_value(r, c) for c in range(sheet.ncols)]
                  for r in range(min(12, sheet.nrows))]
        hdr_pos, idx = _find_header_row(peeked)
        if hdr_pos is None:
            return rows
        exp_col = idx.get('expiry')
        for r in range(hdr_pos + 1, sheet.nrows):  # 0-based，表头行之后
            vals = [sheet.cell_value(r, c) for c in range(sheet.ncols)]
            if is_row_empty(vals):
                continue
            exp_raw = None
            if exp_col is not None and exp_col < len(vals):
                # xlrd 日期单元格 ctype=3，需按 datemode 转回 datetime；否则原样
                if sheet.cell_type(r, exp_col) == xlrd.XL_CELL_DATE:
                    try:
                        exp_raw = xlrd.xldate.xldate_as_datetime(vals[exp_col], book.datemode)
                    except Exception:
                        exp_raw = vals[exp_col]
                else:
                    exp_raw = vals[exp_col]
            rows.append(_build(vals, idx, exp_col, r + 1, exp_raw))
    else:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb[detect_sheet_name(wb)]
            row_iter = ws.iter_rows(min_row=1, values_only=True)
            peeked = []
            for _ in range(12):
                try:
                    peeked.append(list(next(row_iter)))
                except StopIteration:
                    break
            hdr_pos, idx = _find_header_row(peeked)
            if hdr_pos is None:
                return rows
            exp_col = idx.get('expiry')
            # 数据行 = peeked 中表头之后 + 剩余迭代；row_index 为 Excel 行号(1-based)
            data = [(vals, hdr_pos + 2 + offset)
                    for offset, vals in enumerate(peeked[hdr_pos + 1:])]
            next_row_no = len(peeked) + 1
            for vals in row_iter:
                data.append((list(vals), next_row_no))
                next_row_no += 1
            for vals, row_no in data:
                if is_row_empty(vals):
                    continue
                exp_raw = vals[exp_col] if exp_col is not None and exp_col < len(vals) else None
                rows.append(_build(vals, idx, exp_col, row_no, exp_raw))
        finally:
            wb.close()
    return rows


def find_insert_position(records, new_id):
    new_type = get_prefix_type(new_id)
    new_num = extract_number(new_id)
    groups = ['number', 'A', 'B', 'C', 'D', 'E']
    typed = {g: [] for g in groups}
    for r in records:
        typed[get_prefix_type(r['original_id'])].append(r)
    base = 2
    for g in groups:
        group = typed[g]
        group.sort(key=lambda x: extract_number(x['original_id']))
        if new_type == g:
            idx = 0
            for r in group:
                if extract_number(r['original_id']) < new_num: idx += 1
                else: break
            return base + idx
        base += len(group)
    return base

def check_duplicate_labno(records, new_lab_no):
    if not new_lab_no or new_lab_no.strip() == '/' or new_lab_no.strip() == '':
        return False
    return any(r['lab_no'] == new_lab_no for r in records)


def copy_cell_style(src, dst):
    if src.has_style:
        dst.font = src.font.copy()
        dst.border = src.border.copy()
        dst.fill = src.fill.copy()
        dst.number_format = src.number_format
        dst.protection = src.protection.copy()
        dst.alignment = src.alignment.copy()


def _excel_save_error_detail(e):
    """把 Excel 写入异常翻译成用户可读提示。PermissionError 通常是文件正被 Excel 打开被锁。"""
    if isinstance(e, PermissionError):
        return "该 Excel 可能正被别人打开，请关闭后重试"
    return str(e)


@app.route('/api/add_to_excel', methods=['POST'])
def add_to_excel():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    try:
        data = request.get_json()
        record = data.get('record')
        if not record: return jsonify({"success": False, "message": "无数据"}), 400
        excel_path, _ = get_user_paths(session.get('display_name', ''))
        excel_path = excel_path.strip()
        if not excel_path: return jsonify({"success": False, "message": "未配置 Excel 路径"}), 400
        ext = os.path.splitext(excel_path)[1].lower()
        if ext not in ['.xls', '.xlsx']: return jsonify({"success": False, "message": "仅支持 .xls 或 .xlsx"}), 400

        headers = ['编号', '组别', '实验室编号', '标品名称', 'CAS号', '规格/浓度', '生产商',
                   '有效期', '存放地点', '入库日期', '使用情况', '备注']
        new_lab_no = record.get('labNo', '')
        new_original_id = record.get('originalId', '')
        new_row_data = [new_original_id, record.get('group', ''), new_lab_no, record.get('name', ''),
                        record.get('cas', ''), record.get('spec', ''), record.get('manufacturer', ''),
                        record.get('expiry', ''), record.get('location', ''), record.get('storageDate', ''),
                        record.get('usage', ''), record.get('remarks', '')]

        if ext == '.xlsx':
            wb = openpyxl.load_workbook(excel_path)
            sheet_name = detect_sheet_name(wb) or '有机标准物质'
            if sheet_name not in wb.sheetnames:
                ws = wb.create_sheet(sheet_name)
                for col, h in enumerate(headers, 1):
                    c = ws.cell(row=1, column=col, value=h)
                    c.font = Font(bold=True)
                    c.alignment = Alignment(horizontal='center')
            else:
                ws = wb[sheet_name]
            records = parse_existing_records_xlsx(ws)
            if check_duplicate_labno(records, new_lab_no):
                wb.close()
                return jsonify({"success": False, "message": f"实验室编号 {new_lab_no} 已存在"}), 400
            insert_row = find_insert_position(records, new_original_id)
            ws.insert_rows(insert_row)
            for col, val in enumerate(new_row_data, 1):
                ws.cell(row=insert_row, column=col, value=val)

            source_row = insert_row - 1
            while source_row >= 1:
                row_empty = all(ws.cell(row=source_row, column=c).value is None for c in range(1, len(headers)+1))
                if not row_empty: break
                source_row -= 1
            if source_row < 1:
                source_row = insert_row + 1
                max_row = ws.max_row
                while source_row <= max_row:
                    row_empty = all(ws.cell(row=source_row, column=c).value is None for c in range(1, len(headers)+1))
                    if not row_empty: break
                    source_row += 1
            if 1 <= source_row <= ws.max_row:
                for col in range(1, len(new_row_data) + 1):
                    src_cell = ws.cell(row=source_row, column=col)
                    dst_cell = ws.cell(row=insert_row, column=col)
                    copy_cell_style(src_cell, dst_cell)

            wb.save(excel_path)
            wb.close()

        else:
            rb = xlrd.open_workbook(excel_path, formatting_info=True)
            wb = xl_copy(rb)
            styles = Styles(rb)
            # 自动检测 sheet 名
            sheet_name = None
            for name in _KNOWN_SHEET_NAMES:
                if name in rb.sheet_names():
                    sheet_name = name
                    break
            if not sheet_name and rb.sheet_names():
                sheet_name = rb.sheet_names()[0]
            if sheet_name in rb.sheet_names():
                sheet_rb = rb.sheet_by_name(sheet_name)
                ws = wb.get_sheet(sheet_name)
                records = parse_existing_records_xls(sheet_rb)
                if check_duplicate_labno(records, new_lab_no):
                    return jsonify({"success": False, "message": f"实验室编号 {new_lab_no} 已存在"}), 400
                insert_row_idx = find_insert_position(records, new_original_id) - 1
                for r in range(sheet_rb.nrows - 1, insert_row_idx - 1, -1):
                    for c in range(sheet_rb.ncols):
                        val = sheet_rb.cell_value(r, c)
                        try:
                            style = styles[r][c]
                        except:
                            style = xlwt.easyxf()
                        ws.write(r + 1, c, val, style)
                for c, val in enumerate(new_row_data):
                    try:
                        style = styles[insert_row_idx - 1][c] if insert_row_idx > 0 else styles[0][c]
                    except:
                        style = xlwt.easyxf()
                    ws.write(insert_row_idx, c, val, style)
            else:
                ws = wb.add_sheet(sheet_name)
                for c, h in enumerate(headers): ws.write(0, c, h)
                for c, val in enumerate(new_row_data): ws.write(1, c, val)
            wb.save(excel_path)

        return jsonify({"success": True, "message": "数据已同步到 Excel"})
    except Exception as e:
        return jsonify({"success": False, "message": f"写入 Excel 失败: {_excel_save_error_detail(e)}"}), 500


@app.route('/api/update_lims_unit', methods=['POST'])
def update_lims_unit():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    # 已禁用：原实现对 /consumableBill/{recordId} 做全量替换 PUT（仅带 concentrationUnitName），
    # 会把记录其余字段清空（故障实例：CK-FCM-2026039）。LIMS 详情接口 GET 返回 405、无法按 id
    # 安全 merge；在向 LIMS 管理员确认正确的更新契约之前，停止一切回写，避免再损坏记录。
    return jsonify({"success": False, "message": "LIMS 浓度单位回写已禁用（全量替换会清空记录，待确认安全更新方式）"})


@app.route('/api/lims/receive', methods=['POST'])
def lims_receive():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    data = request.get_json()
    consumable_id = str(data.get('consumable_id', ''))
    quantity = data.get('quantity')
    unit = data.get('unit', 'g')
    receive_date = data.get('receive_date') or datetime.datetime.now().strftime("%Y-%m-%d 00:00:00")
    if not consumable_id or quantity is None:
        return jsonify({"success": False, "message": "缺少参数"}), 400

    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    if not pid:
        return jsonify({"success": False, "message": "无法获取用户PID，请重新登录"}), 401
    try:
        qty_val = float(quantity)
        if unit == 'g':
            formatted_qty = f"{qty_val:.4f}"
        else:  # mL
            formatted_qty = f"{qty_val:.3f}"
    except (ValueError, TypeError):
        formatted_qty = str(quantity)
    form_data = {
        "num": formatted_qty,
        "receiveDate": receive_date,
        "purpose": "",
        "receiveType": "CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE",
        "ids": consumable_id,
        "receiveUserName": pname,
        "isLevelTwo": "NO",
        "pid": pid,
        "pname": pname,
        "loginId": pid
    }
    headers = {
        "Referer": f"{system.base_url}/web/consumablesReceiveListMgt.html?menuId=289",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
    }
    url = f"{system.base_url}/detectionManager/manager/consumableReceive/receive"
    try:
        resp = system.session.post(url, data=form_data, headers=headers)
        resp.raise_for_status()
        result = resp.json()
        if not result.get("success"):
            return jsonify({"success": False, "message": result.get('errorDesc') or str(result.get('errorCtx', '领用失败'))})
        time.sleep(1)
        rec_url = f"{system.base_url}/detectionManager/manager/consumableReceive/record/{consumable_id}"
        rec_params = {"_search": "false", "nd": str(int(time.time()*1000)), "pageSize": 9999, "pageNo": 1,
                      "sidx": "", "sord": "asc", "pid": pid, "pname": pname, "loginId": pid}
        rec_resp = system.session.get(rec_url, params=rec_params, headers={"Referer": f"{system.base_url}/web/consumablesReceiveListMgt.html?menuId=289"})
        receive_id = consumable_id
        if rec_resp.status_code == 200:
            rec_data = rec_resp.json()
            records = rec_data.get("resultData", [])
            if records:
                receive_id = str(records[0].get("id", consumable_id))
        return jsonify({"success": True, "receive_id": receive_id})
    except Exception as e:
        return jsonify({"success": False, "message": f"领用异常: {str(e)}"}), 500


@app.route('/api/lims/save_solution', methods=['POST'])
def lims_save_solution():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    p = request.get_json()
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    if not pid:
        return jsonify({"success": False, "message": "无法获取用户PID，请重新登录"}), 401

    try:
        purity_str = str(p.get('purity_str', '99.5'))
        purity_value = float(purity_str.replace('%', '').strip())
        volume_ml = float(p.get('volume_ml', 1))
        use_quantity = float(p.get('use_quantity', 1))
        original_unit = p.get('original_unit', '%')
        received_unit = p.get('received_unit', 'g')
        volume_ml = _round_vol(volume_ml)
        use_quantity = _round_qty(use_quantity, received_unit)
        if original_unit == '%':
            config_conc = purity_value / 100.0 * use_quantity * 1_000_000 / volume_ml
        else:
            config_conc = purity_value * use_quantity / volume_ml
        config_conc = _round_conc(config_conc, original_unit)
    except Exception:
        config_conc = float(p.get('config_conc', 0))
        received_unit = p.get('received_unit', 'g')

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    receive_id = str(p.get('receive_id', ''))
    original_id = str(p.get('original_id', ''))
    original_name = p.get('original_name', '')
    qty_display = _fmt_qty(use_quantity, received_unit)
    vol_display = _fmt_vol(volume_ml)
    conc_display = _fmt_conc(config_conc, original_unit)

    save_detail_item = {
        "id": None, "createDatetime": now_str, "serialVersionUID": None,
        "configureId": None, "originalId": receive_id,
        "originalCode": p.get('original_code', f"L-{receive_id}"),
        "originalName": original_name,
        "originalConcentration": f"{purity_str}({original_unit})",
        "concentrationCount": f"{purity_str}({original_unit})",
        "originalUnit": original_unit,
        "receivedQuantity": qty_display, "receivedUint": received_unit,
        "useUantity": qty_display, "useUnit": received_unit,
        "medium": p.get('medium', ''),
        "volume": vol_display, "unit": "mL",
        "configurationConcentration": conc_display,
        "configurationUnit": "μg/mL", "configurationUncertainty": None,
        "remark": None, "consumableReceive": None, "dataid": original_id,
        "configurationRecordId": None, "creatorName": None, "creatorId": None,
        "modifierName": None, "modifyDatetime": now_str,
        "conversionFactor": None if original_unit == "μg/mL" else "1",
        "dilutionFactor": None,
        "originalNo": p.get('original_no', f"A-{original_id}\n{p.get('batch_no','')}\n{p.get('solution_code','')}"),
        "_X_ID": f"row_{int(time.time())}"
    }

    payload = {
        "id": None, "createDatetime": now_str, "serialVersionUID": None,
        "solutionName": p.get('solution_name', original_name),
        "solutionCode": p.get('solution_code', ''),
        "deviceIds": None, "deviceNames": p.get('device_names'),
        "configureDate": p.get('configure_date', datetime.date.today().strftime('%Y-%m-%d')),
        "validityDate": p.get('validity_date', ''),
        "storageCondition": p.get('storage_condition'),
        "storageLocation": p.get('storage_location', '4-1-华业4-1'),
        "concentration": f"{original_name}:{conc_display}(μg/mL)",
        "concentrationCount": f"{original_name}:{conc_display}(μg/mL)",
        "concentrationUnitName": None, "uncertainty": None, "configureOrder": None,
        "originalCode": f"A-{original_id}",
        "controlledNo": p.get('controlled_no', p.get('batch_no', '')),
        "medium": None, "configuratorId": None,
        "solutionType": p.get('solution_type', 'SOLUTION_TYPE_B'),
        "configuratorName": None, "constantVolume": 0, "totalConstantVolume": 0,
        "usedConstantVolume": 0, "remark": None, "diluteStatus": False,
        "consumableReceive": None, "customType": p.get('customType'), "auditUserName": None,
        "auditTime": None, "diluteConcentrationControl": "[]",
        "pageType": p.get('page_type', 'SOLUTION_TYPE_A'),
        "saveDetailList": json.dumps([save_detail_item], ensure_ascii=False),
        "originalValidityDate": p.get('expiry_date', ''),
        "configurationTemplateId": None, "intermediateTemplateId": None,
        "temperature": p.get('temperature'), "humidity": p.get('humidity'),
        "computingFormula": None, "configurationMethod": None,
        "configurationProcess": None, "nextAuditUserName": None,
        "disposeUserName": None, "disposeTime": None, "disposeWay": None,
        "configurationSolutionTemplateId": None, "configurationRecordId": None,
        "creatorName": None, "creatorId": None, "modifierName": None,
        "modifyDatetime": now_str, "receivedUint": original_unit,
        "pid": pid, "pname": pname, "loginId": pid
    }
    try:
        url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/saveSolutionConfigure"
        headers = {
            "Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544",
            "Content-Type": "application/json;charset=UTF-8"
        }
        resp = system.session.post(url, json=payload, headers=headers)
        if resp.status_code in (401, 403) or _resp_looks_expired(resp):
            return _expired_response()
        if not resp.ok:
            print(f"[saveSolutionConfigure] status={resp.status_code} body={resp.text[:500]}")
            return jsonify({"success": False, "message": f"LIMS 请求失败，状态码: {resp.status_code}"}), 502
        result = resp.json()
        if not result.get("success"):
            return jsonify({"success": False, "message": result.get('errorDesc') or str(result.get('errorCtx', '配置失败'))})
        return jsonify({"success": True})
    except Exception as e:
        print(f"[saveSolutionConfigure] exception: {e}")
        return jsonify({"success": False, "message": f"配置异常: {str(e)}"}), 500


def _extract_direct_source_codes(original_code):
    text = str(original_code or '').strip()
    if not text:
        return []
    seen = []
    seen_set = set()
    for match in re.findall(r'([ABC])-\s*(\d+)', text, flags=re.IGNORECASE):
        code = f"{match[0].upper()}-{match[1]}"
        if code not in seen_set:
            seen.append(code)
            seen_set.add(code)
    return seen


def _normalize_solution_type(value):
    if isinstance(value, dict):
        return str(value.get('key') or '').strip()
    return str(value or '').strip()


def _extract_a_source_codes(original_code):
    text = str(original_code or '').strip()
    if not text:
        return []
    seen = []
    seen_set = set()
    for match in re.findall(r'A-\s*(\d+)', text, flags=re.IGNORECASE):
        code = f"A-{match}"
        if code not in seen_set:
            seen.append(code)
            seen_set.add(code)
    return seen


def _parse_conc_field(conc_str):
    """Parse LIMS concentration field.
    '0.01(mg/L)' → (0.01, 'mg/L')
    'DIDP:447.10(mg/L);DINP:682.32(mg/L)' → [(DIDP,447.10,mg/L), ...]
    """
    if not conc_str:
        return []
    text = str(conc_str).strip()
    if ':' in text and ';' in text:
        parts = text.split(';')
        result = []
        for p in parts:
            p = p.strip()
            if not p:
                continue
            ci = p.rfind(':')
            if ci < 0:
                continue
            name = p[:ci].strip()
            cstr = p[ci+1:].strip()
            m = re.match(r'^([\d.]+)\s*\(([^)]+)\)', cstr)
            if m:
                result.append((name, float(m.group(1)), m.group(2)))
            else:
                m2 = re.match(r'^([\d.]+)\s*(\S+)', cstr)
                if m2:
                    result.append((name, float(m2.group(1)), m2.group(2)))
        return result
    m = re.match(r'^([\d.]+)\s*\(([^)]+)\)', text)
    if m:
        return [(None, float(m.group(1)), m.group(2))]
    m2 = re.match(r'^([\d.]+)\s*(\S+)', text)
    if m2:
        return [(None, float(m2.group(1)), m2.group(2))]
    return []


_order_to_id_cache = {}
_type_list_cache = {}

def _resolve_order_to_lims_id(system, configure_order):
    """Resolve configureOrder (e.g. 'B-3408') to actual LIMS record ID.
    Uses getSolutionAdata API with appropriate type parameter, pageSize=9999,
    exact matching, and per-type list caching."""
    global _order_to_id_cache, _type_list_cache
    if configure_order in _order_to_id_cache:
        return _order_to_id_cache[configure_order]
    parts = configure_order.split('-')
    if len(parts) < 2:
        return None
    prefix = parts[0].upper()
    if prefix == 'A':
        return None  # A-type records not queryable via this API
    type_map = {'B': 'SOLUTION_TYPE_B', 'C': 'SOLUTION_TYPE_C', 'D': 'SOLUTION_TYPE_D', 'E': 'SOLUTION_TYPE_E'}
    sol_type = type_map.get(prefix)
    if not sol_type:
        return None
    try:
        if sol_type in _type_list_cache:
            items = _type_list_cache[sol_type]
        else:
            url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/getSolutionAdata"
            headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
            items = []
            page_no = 1
            while True:
                resp = system.session.get(url, params={
                    "type": sol_type, "pageSize": 9999, "pageNo": page_no,
                    "status": "1",
                    "pid": system.current_pid or '',
                    "pname": system.current_real_name or '',
                    "loginId": system.current_pid or '',
                }, headers=headers)
                resp.raise_for_status()
                data = resp.json()
                page_items = data.get('resultData', {}).get('voList', [])
                items.extend(page_items)
                if not data.get('resultData', {}).get('hasNext', False):
                    break
                page_no += 1
            _type_list_cache[sol_type] = items
        # Exact match
        for item in items:
            order = str(item.get('configureOrder', '')).strip()
            if order == configure_order:
                lid = item.get('id')
                if lid:
                    _order_to_id_cache[configure_order] = lid
                    return lid
        # Prefix + number exact match (not endswith)
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
    except Exception as e:
        pass
    return None


def _resolve_by_solution_code(system, solution_code):
    """Resolve solutionCode (e.g. 'CK-CG-xxx') to (lims_id, configure_order) or (None, None).
    Uses getSolutionAdata with solutionCode parameter for precise search."""
    global _order_to_id_cache
    url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/getSolutionAdata"
    headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
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
            "pid": system.current_pid or '',
            "pname": system.current_real_name or '',
            "loginId": system.current_pid or '',
        }
        try:
            resp = system.session.get(url, params=params, headers=headers)
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
                        return lid, co
        except Exception:
            pass
    return None, None


def _fetch_solution_view(system, solution_id, order_str=None):
    """Fetch record via viewDtSolutionConfigure — returns complete detailList.
    type parameter must match record type (B/C→SOLUTION_TYPE_E, D→SOLUTION_TYPE_D).
    If order_str is None, tries SOLUTION_TYPE_E first, then SOLUTION_TYPE_D as fallback."""
    url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/viewDtSolutionConfigure"
    headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
    view_type = "SOLUTION_TYPE_E"
    if order_str:
        pfx = order_str.split('-')[0].upper()
        type_map = {'D': 'SOLUTION_TYPE_D', 'C': 'SOLUTION_TYPE_C', 'B': 'SOLUTION_TYPE_E', 'E': 'SOLUTION_TYPE_E'}
        view_type = type_map.get(pfx, 'SOLUTION_TYPE_E')
    resp = system.session.get(url, params={"id": solution_id, "type": view_type}, headers=headers)
    resp.raise_for_status()
    data = resp.json()
    result = data.get('resultData') or {}
    if not order_str and view_type != 'SOLUTION_TYPE_D' and not result.get('detailList'):
        resp2 = system.session.get(url, params={"id": solution_id, "type": "SOLUTION_TYPE_D"}, headers=headers)
        resp2.raise_for_status()
        result2 = resp2.json().get('resultData') or {}
        if result2.get('detailList'):
            result = result2
    return result


_ADMIN_DISPLAY_NAME = '崔艳梅'


def _is_configurator_match(sess, record):
    """Check if current session user matches the record's configurator by ID."""
    current_pid = str(sess.get('pid') or '').strip()
    if not current_pid:
        return True
    if not record:
        return True
    rec_pid = str(record.get('configuratorId') or record.get('creatorId') or '').strip()
    if rec_pid:
        return current_pid == rec_pid
    return True


def _check_edit_permission(sess, record):
    """Check if current user has permission to modify a record.
    Returns (allowed, reason) tuple."""
    if (sess.get('display_name') or '').strip() == _ADMIN_DISPLAY_NAME:
        return True, ''
    if record and record.get('auditUserName'):
        return False, '该记录已审核，不允许修改'
    if not _is_configurator_match(sess, record):
        return False, '只能修改自己配置的记录'
    return True, ''


def _check_delete_permission(sess, record):
    """Check if current user has permission to delete a record.
    Configurator can delete own records (including audited). Admin can delete any.
    Returns (allowed, reason) tuple."""
    if (sess.get('display_name') or '').strip() == _ADMIN_DISPLAY_NAME:
        return True, ''
    if not _is_configurator_match(sess, record):
        return False, '只能删除自己配置的记录'
    return True, ''


def _fetch_solution_detail(system, solution_id):
    """Fetch record via detail API (saveDetailList always null). Fallback only."""
    url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/detail"
    headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
    resp = system.session.get(url, params={"id": solution_id}, headers=headers)
    resp.raise_for_status()
    data = resp.json()
    return data.get('resultData') or {}


def _first_number(s):
    """提取字符串前导数字部分，保留原始小数位（避免 str(float) 丢尾零）。
    '10.00(mg/L)' → '10.00'；'' 或无前导数字 → ''
    """
    m = re.match(r'^\s*([\d.]+)', str(s))
    return m.group(1) if m else ''


def _trace_export_chain(system, trace_targets, target_date, target_person):
    """Trace source chain upward (D→C→B→A), find records with same configurator + date.
    trace_targets: list of (lims_id, configure_order) tuples. Use lims_id if available.
    Returns (matched_records, top_ancestor) where matched_records is ordered top→bottom.
    Stops when: parent is A-type, or source concentration unit is % (weighing type).
    Uses viewDtSolutionConfigure API for complete detailList data.
    """
    matched = []
    visiting = set()

    def _fetch_record(lims_id=None, order_str=None):
        resolved = lims_id
        view_order_str = order_str
        if not resolved and order_str:
            if order_str.startswith('CK-'):
                resolved, co = _resolve_by_solution_code(system, order_str)
                if co and co != order_str:
                    view_order_str = co
            else:
                resolved = _resolve_order_to_lims_id(system, order_str)
        # Fallback: use configureOrder number as ID when resolve fails
        if not resolved and order_str and '-' in order_str:
            num_part = order_str.split('-')[1].strip()
            if num_part.isdigit():
                resolved = int(num_part)
        if not resolved:
            return None

        is_solution_code = order_str and order_str.startswith('CK-')

        def _validate(rec):
            if not order_str:
                return True
            if is_solution_code:
                rec_sc = str(rec.get('solutionCode', '')).strip()
                if rec_sc and rec_sc != order_str:
                    return False
            else:
                rec_order = str(rec.get('configureOrder', '')).strip()
                if rec_order and rec_order != order_str:
                    return False
            return True

        try:
            rec = _fetch_solution_view(system, resolved, order_str=view_order_str)
            if rec and rec.get('id') and _validate(rec):
                return rec
        except Exception:
            pass
        try:
            rec = _fetch_solution_detail(system, resolved)
            if rec and rec.get('id') and _validate(rec):
                return rec
        except Exception:
            pass
        return None

    def _has_pct_source_conc(record, parent_order):
        """Check if parent's source concentration in detailList contains %."""
        for dl in (record.get('detailList') or []):
            dl_code = str(dl.get('originalCode', '')).strip()
            dl_no = str(dl.get('originalNo', '')).strip()
            no_first = dl_no.split('\n')[0].strip().split('|')[0].strip() if dl_no else ''
            if dl_code == parent_order or no_first == parent_order:
                conc = str(dl.get('originalConcentration', '') or '').strip()
                return '%' in conc
        return False

    def _trace(lims_id=None, order_str=None, depth=0):
        if depth > 10:
            return None
        key = str(lims_id or order_str)
        if key in visiting:
            return None
        visiting.add(key)

        record = _fetch_record(lims_id, order_str)
        if not record:
            visiting.discard(key)
            return None

        rec_order = str(record.get('configureOrder') or '').strip() or order_str
        rec_person = str(record.get('configuratorName') or record.get('creatorName') or '').strip()
        rec_date = str(record.get('configureDate') or '').strip()[:10]

        # Check same configurator + date (only when filter is specified)
        if target_person and target_date:
            if rec_person != target_person or not rec_date.startswith(target_date):
                visiting.discard(key)
                return record  # stopping point (different person/date)

        prefix = rec_order.split('-')[0].upper() if '-' in rec_order else ''
        original_code = str(record.get('originalCode') or '').strip()
        parent_orders = [oc.strip() for oc in re.split(r'[,，]', original_code) if oc.strip()]

        # Use totalConstantVolume for actual dilution volume (constantVolume = remaining qty)
        total_vol = record.get('totalConstantVolume')
        constant_volume = float(total_vol if total_vol is not None else record.get('constantVolume') or 0)
        rec_conc = str(record.get('concentration') or '').strip()
        rec_parsed = _parse_conc_field(rec_conc)
        rec_conc_val = rec_parsed[0][1] if rec_parsed else 0
        rec_conc_unit = rec_parsed[0][2] if rec_parsed else ''

        # Get detailList for this record
        detail_list = record.get('detailList') or []
        is_d_type = prefix == 'D'
        # 体积优先取 detailList 原始串（保留 10.00 精度）；record 级 totalConstantVolume 是 float，str() 会丢尾零
        constant_volume_raw = ''
        if detail_list:
            constant_volume_raw = str(detail_list[0].get('volume') or '').strip()
        if not constant_volume_raw:
            constant_volume_raw = str(total_vol if total_vol is not None else record.get('constantVolume') or '')

        # Extract parent info from detailList (if available) or from parent record
        parent_name = ''
        parent_conc_val = 0
        parent_conc_unit = ''
        parent_conc_raw = ''
        received_qty = ''
        received_unit = str(record.get('receivedUint') or 'mL').strip()

        if detail_list and not is_d_type:
            # B/C type: use first detailList row for parent info
            dl0 = detail_list[0]
            parent_name = str(dl0.get('originalName', '')).strip()
            p_conc = str(dl0.get('originalConcentration', '') or '').strip()
            parent_conc_raw = p_conc
            pc = _parse_conc_field(p_conc)
            if pc:
                parent_conc_val = pc[0][1]
                parent_conc_unit = pc[0][2]
            received_qty = str(dl0.get('receivedQuantity', '')).strip()
            received_unit = str(dl0.get('receivedUint', '')).strip()
        else:
            # D-type or no detailList: fetch parent record
            if parent_orders:
                po = parent_orders[0]
                try:
                    resolved_id = _resolve_order_to_lims_id(system, po)
                    if resolved_id:
                        parent_rec = _fetch_solution_view(system, resolved_id, order_str=po)
                    else:
                        parent_rec = None
                    if parent_rec:
                        parent_name = str(parent_rec.get('solutionName') or '').strip()
                        p_conc = str(parent_rec.get('concentration') or '').strip()
                        parent_conc_raw = p_conc
                        pc = _parse_conc_field(p_conc)
                        if pc:
                            parent_conc_val = pc[0][1]
                            parent_conc_unit = pc[0][2]
                except Exception:
                    pass

            # Calculate received quantity if not from detailList
            if not received_qty and parent_conc_val > 0 and constant_volume > 0:
                calc_qty = rec_conc_val * constant_volume / parent_conc_val
                received_qty = f"{calc_qty:.4f}" if received_unit == 'g' else f"{calc_qty:.2f}"

        # Stopping conditions
        is_weighing_top = False
        if prefix == 'B' and received_unit == 'g':
            is_weighing_top = True

        matched.append({
            'level': prefix,
            'configure_order': rec_order,
            'solution_name': str(record.get('solutionName') or '').strip(),
            'solution_code': str(record.get('solutionCode') or '').strip(),
            'concentration': rec_conc,
            'constant_volume': _first_number(constant_volume_raw) or str(constant_volume),
            'medium': str(record.get('medium') or '').strip(),
            'configure_date': rec_date,
            'validity_date': str(record.get('validityDate') or '').strip(),
            'controlled_no': str(record.get('controlledNo') or '').strip(),
            'storage_condition': str(record.get('storageCondition') or '').strip(),
            'received_quantity': received_qty,
            'received_unit': received_unit,
            'parent_name': parent_name,
            'parent_concentration': _first_number(parent_conc_raw) or str(parent_conc_val),
            'parent_conc_unit': parent_conc_unit,
            'original_code': original_code,
            'concentration_count': str(record.get('concentrationCount') or '').strip(),
            '_is_weighing_top': is_weighing_top,
            '_detail_list': detail_list,
            'weighing_equipment': record.get('weighingEquipment'),
            'weighing_equipment_json': str(record.get('weighingEquipmentJson') or ''),
            'device_names': str(record.get('deviceNames') or '').strip(),
            'diluteStatus': record.get('diluteStatus'),   # 逐级稀释开关(D 记录持久化字段); 供下游 mup 溯源填充判定 work_serial_dilute
        })

        if is_weighing_top:
            visiting.discard(key)
            return record

        # Extract parent originalId from detailList (avoids resolve when getSolutionAdata returns 500)
        # Only use originalId when detailList.originalCode directly matches parent_order.
        # When match is through originalNo (e.g., originalCode=A-5728 but originalNo contains B-3387),
        # the originalId refers to the A-type record, NOT the B-type — wrong mapping.
        parent_id_map = {}
        for dl in detail_list:
            dl_code = str(dl.get('originalCode', '')).strip()
            dl_id = dl.get('originalId')
            if dl_id and dl_code:
                parent_id_map.setdefault(dl_code, dl_id)

        # Continue tracing: filter out A-type parents and %-concentration parents
        if original_code:
            for po in parent_orders:
                po_prefix = po.split('-')[0].upper() if '-' in po else ''
                if po_prefix == 'A':
                    continue
                if _has_pct_source_conc(record, po):
                    continue
                pid = parent_id_map.get(po)
                _trace(lims_id=pid, order_str=po, depth=depth + 1)

        visiting.discard(key)
        return None

    for lims_id, order in trace_targets:
        _trace(lims_id=lims_id, order_str=order)

    matched.reverse()

    # Build lookup
    matched_by_order = {}
    for rec in matched:
        order = rec.get('configure_order', '')
        if order:
            matched_by_order[order] = rec

    # Post-process: expand multi-source records using detailList
    for rec in matched:
        if rec.get('_is_weighing_top'):
            continue
        src_code = rec.get('original_code', '')
        sources = [s.strip() for s in re.split(r'[,，]', src_code) if s.strip()]
        if len(sources) > 1:
            rec['source_details'] = []
            rec_conc = str(rec.get('concentration', '')).strip()
            rec_parsed = _parse_conc_field(rec_conc)
            rec_conc_val = rec_parsed[0][1] if rec_parsed else 0
            vol = float(rec.get('constant_volume', 0))
            for src_order in sources:
                src_name = ''
                src_conc_val = 0
                src_conc_unit = ''
                src_qty = ''

                src_matched = matched_by_order.get(src_order)
                if src_matched:
                    src_name = src_matched.get('solution_name', '')
                    sp = _parse_conc_field(src_matched.get('concentration', ''))
                    if sp:
                        src_conc_val = sp[0][1]
                        src_conc_unit = sp[0][2]
                else:
                    sp2 = src_order.split('-')
                    if len(sp2) >= 2:
                        try:
                            src_id = int(sp2[1].strip())
                            src_rec = _fetch_solution_detail(system, src_id)
                            if src_rec:
                                src_name = str(src_rec.get('solutionName') or '').strip()
                                sp3 = _parse_conc_field(str(src_rec.get('concentration') or ''))
                                if sp3:
                                    src_conc_val = sp3[0][1]
                                    src_conc_unit = sp3[0][2]
                        except Exception:
                            pass
                    if not src_name:
                        cc_count = str(rec.get('concentration_count') or '').strip()
                        cc_parts = cc_count.split(';')
                        if len(cc_parts) >= len(sources):
                            idx = sources.index(src_order)
                            if idx < len(cc_parts):
                                cpart = cc_parts[idx].strip()
                                ci = cpart.rfind(':')
                                if ci > 0:
                                    src_name = cpart[:ci].strip()
                                else:
                                    src_name = cpart or src_order

                if not src_name:
                    src_name = src_order
                if src_conc_val > 0 and vol > 0:
                    calc = rec_conc_val * vol / src_conc_val
                    src_qty = '' if calc > vol else f"{calc:.2f}"
                rec['source_details'].append({
                    'name': src_name,
                    'conc': src_conc_val,
                    'conc_unit': src_conc_unit,
                    'qty': src_qty,
                })

    # Determine top ancestor
    top_ancestor = {}
    if matched:
        top_names = []
        top_codes = []
        for rec in matched:
            if rec.get('_is_weighing_top'):
                name = rec.get('solution_name', '')
            else:
                name = rec.get('parent_name', '')
            if name and name not in top_names:
                top_names.append(name)
            sc = rec.get('solution_code', '')
            sc_parts = sc.split('-')
            bc = '-'.join(sc_parts[:-2]) if len(sc_parts) >= 3 else sc
            if bc and bc not in top_codes:
                top_codes.append(bc)

        if len(top_names) > 1 or len(top_codes) > 1:
            top_ancestor = {
                'solution_name': '；\n'.join(top_names) + '；',
                'controlled_no': '；\n'.join(top_codes) + '；',
                'concentration': '见下表',
            }
        else:
            rec0 = matched[0]
            if rec0.get('_is_weighing_top'):
                top_ancestor = {
                    'solution_name': top_names[0] if top_names else rec0.get('solution_name', ''),
                    'controlled_no': top_codes[0] if top_codes else '',
                    'concentration': rec0.get('concentration', ''),
                }
            else:
                pn = rec0.get('parent_name', '')
                pc = rec0.get('parent_concentration', '')
                pu = rec0.get('parent_conc_unit', '')
                top_ancestor = {
                    'solution_name': pn or rec0.get('solution_name', ''),
                    'controlled_no': top_codes[0] if top_codes else '',
                    'concentration': f"{pc}({pu})" if pc and pu else str(pc) if pc else '',
                }

    return matched, top_ancestor


def _resolve_bottom_a_codes(system, item, cache=None, items_by_order=None, visiting=None):
    if cache is None:
        cache = {}
    if visiting is None:
        visiting = set()

    item_id = str(item.get('id') or '').strip()
    configure_order = str(item.get('configureOrder') or '').strip()
    cache_key = item_id or configure_order
    if cache_key and cache_key in cache:
        return cache[cache_key]
    if cache_key and cache_key in visiting:
        return []

    if cache_key:
        visiting.add(cache_key)

    original_code = item.get('originalCode')
    solution_type = _normalize_solution_type(item.get('solutionType'))
    direct_a_codes = _extract_a_source_codes(original_code)
    if direct_a_codes and solution_type != 'SOLUTION_TYPE_C':
        result = direct_a_codes
    else:
        result = []
        nested_codes = []
        text = str(original_code or '')
        pattern = r'([BC])-\s*(\d+)'
        if solution_type == 'SOLUTION_TYPE_C':
            pattern = r'B-\s*(\d+)'
        for match in re.findall(pattern, text, flags=re.IGNORECASE):
            if isinstance(match, tuple):
                nested_code = f"{match[0].upper()}-{match[1]}"
            else:
                nested_code = f"B-{match}"
            if nested_code not in nested_codes:
                nested_codes.append(nested_code)
        if nested_codes:
            seen = set()
            for nested_code in nested_codes:
                nested_detail = (items_by_order or {}).get(nested_code)
                if not nested_detail:
                    try:
                        nested_detail = _fetch_solution_detail(system, nested_code.split('-')[1])
                    except Exception:
                        nested_detail = {}
                nested_result = _resolve_bottom_a_codes(system, nested_detail, cache, items_by_order, visiting)
                for code in nested_result:
                    if code not in seen:
                        seen.add(code)
                        result.append(code)
        elif direct_a_codes:
            result = direct_a_codes

    if cache_key:
        visiting.discard(cache_key)
        if result:
            cache[cache_key] = result
    return result


def _extract_a_source_count_from_original_code(original_code):
    codes = _extract_a_source_codes(original_code)
    return len(codes) or 1


def _enrich_solution_a_source_count(items, system):
    cache = {}
    items_by_order = {}
    for it in items:
        co = str(it.get('configureOrder') or '').strip()
        if co:
            items_by_order[co] = it
    for item in items:
        item['sourceCodes'] = _extract_direct_source_codes(item.get('originalCode'))
        source_codes = _resolve_bottom_a_codes(system, item, cache, items_by_order)
        item['aSourceCodes'] = source_codes
        item['aSourceCount'] = len(source_codes) or 1
    return items


@app.route('/api/lims/list_configured_solutions', methods=['GET'])
def lims_list_configured_solutions():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    if not pid:
        return jsonify({"success": False, "message": "无法获取用户PID，请重新登录"}), 401

    params = {
        "_search": "false",
        "nd": str(int(time.time() * 1000)),
        "pageSize": int(request.args.get('page_size', 30)),
        "pageNo": int(request.args.get('page_no', 1)),
        "sidx": "",
        "sord": "asc",
        "solutionName": request.args.get('name', ''),
        "solutionCode": request.args.get('code', ''),
        "customType": request.args.get('custom_type', ''),
        "configStatus": request.args.get('config_status', '0'),
        "controlledNo": "",
        "storageLocation": "",
        "configureStartDate": request.args.get('date_from', ''),
        "configureEndDate": request.args.get('date_to', ''),
        "configureUserName": request.args.get('operator', ''),
        "receiveUserName": "",
        "auditStatus": request.args.get('audit_status', ''),
        "status": "1",
        "type": "SOLUTION_TYPE_E",
        "pid": pid,
        "pname": pname,
        "loginId": pid,
    }
    try:
        url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/getSolutionAdata"
        headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
        resp = system.session.get(url, params=params, headers=headers)
        resp.raise_for_status()
        result = resp.json()
        rd = result.get('resultData') or {}
        items = rd.get('voList', [])
        return jsonify({
            "success": True,
            "data": items,
            "total": rd.get('totalCount', 0),
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"查询异常: {str(e)}"}), 500


@app.route('/api/lims/list_d_solutions', methods=['GET'])
def lims_list_d_solutions():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    if not pid:
        return jsonify({"success": False, "message": "无法获取用户PID，请重新登录"}), 401

    params = {
        "_search": "false",
        "nd": str(int(time.time() * 1000)),
        "pageSize": int(request.args.get('page_size', 30)),
        "pageNo": int(request.args.get('page_no', 1)),
        "sidx": "",
        "sord": "asc",
        "solutionName": request.args.get('name', ''),
        "solutionCode": request.args.get('code', ''),
        "customType": request.args.get('custom_type', ''),
        "configStatus": request.args.get('config_status', '0'),
        "controlledNo": "",
        "storageLocation": "",
        "configureStartDate": request.args.get('date_from', ''),
        "configureEndDate": request.args.get('date_to', ''),
        "configureUserName": request.args.get('operator', ''),
        "receiveUserName": "",
        "auditStatus": request.args.get('audit_status', ''),
        "status": "1",
        "type": "SOLUTION_TYPE_D",
        "pid": pid,
        "pname": pname,
        "loginId": pid,
    }
    try:
        url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/getSolutionAdata"
        headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
        resp = system.session.get(url, params=params, headers=headers)
        result = resp.json()
        # 检测LIMS返回的认证错误（可能包在200或500中）
        err_ctx = result.get('errorCtx') or {}
        if err_ctx.get('errorCode') == '401' or '未登录' in (err_ctx.get('errorMsg') or ''):
            session.pop('logged_in', None)
            return jsonify({"success": False, "message": "远程会话已失效，请重新登录"}), 401
        rd = result.get('resultData') or {}
        items = rd.get('voList', [])
        return jsonify({
            "success": True,
            "data": items,
            "total": rd.get('totalCount', 0),
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"查询异常: {str(e)}"}), 500


@app.route('/api/lims/quick_query', methods=['GET'])
def lims_quick_query():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    if not pid:
        return jsonify({"success": False, "message": "无法获取用户PID，请重新登录"}), 401
    operator = request.args.get('operator', '').strip()
    date_from = request.args.get('date_from', '')
    date_to = request.args.get('date_to', '')

    base_params = {
        "_search": "false", "nd": str(int(time.time() * 1000)),
        "pageSize": 9999, "pageNo": 1, "sidx": "", "sord": "asc",
        "solutionName": "", "solutionCode": "", "customType": "", "configStatus": request.args.get('config_status', ''),
        "controlledNo": "", "storageLocation": "",
        "configureStartDate": date_from, "configureEndDate": date_to,
        "configureUserName": operator, "receiveUserName": "", "auditStatus": request.args.get('audit_status', ''),
        "status": "1", "pid": pid, "pname": pname, "loginId": pid,
    }
    headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
    url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/getSolutionAdata"

    def _fetch(sol_type):
        p = {**base_params, "type": sol_type}
        try:
            resp = system.session.get(url, params=p, headers=headers, timeout=15)
            resp.raise_for_status()
            rd = resp.json().get('resultData') or {}
            items = rd.get('voList', [])
            for item in items:
                order = str(item.get('configureOrder', '')).upper()
                if order.startswith('D-'):
                    item['_solType'] = 'SOLUTION_TYPE_D'
                elif order.startswith('C-'):
                    item['_solType'] = 'SOLUTION_TYPE_C'
                else:
                    item['_solType'] = 'SOLUTION_TYPE_B'
            return items
        except Exception:
            return []

    from concurrent.futures import ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=2) as pool:
        f_e = pool.submit(_fetch, "SOLUTION_TYPE_E")
        f_d = pool.submit(_fetch, "SOLUTION_TYPE_D")
        items = f_e.result() + f_d.result()

    return jsonify({"success": True, "data": items, "total": len(items)})


@app.route('/api/lims/get_source_info', methods=['GET'])
def lims_get_source_info():
    """Fetch source record info by configureOrder or LIMS id.
    Returns solutionCode and solutionName."""
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()

    configure_order = request.args.get('order', '').strip()
    source_id = request.args.get('id', '').strip()
    if not configure_order and not source_id:
        return jsonify({"success": True, "data": {}})

    try:
        lims_id = None
        if configure_order:
            lims_id = _resolve_order_to_lims_id(system, configure_order)
        if not lims_id and source_id:
            lims_id = source_id
        if not lims_id:
            return jsonify({"success": True, "data": {}})
        result = _fetch_solution_view(system, lims_id, order_str=configure_order or None)
        if not result.get('id') and source_id:
            for try_type in ['SOLUTION_TYPE_C', 'SOLUTION_TYPE_E', 'SOLUTION_TYPE_B']:
                try:
                    url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/viewDtSolutionConfigure"
                    headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
                    resp = system.session.get(url, params={"id": source_id, "type": try_type}, headers=headers)
                    resp.raise_for_status()
                    data = resp.json()
                    r = data.get('resultData') or {}
                    if r.get('id'):
                        result = r
                        break
                except Exception:
                    continue
        if not result.get('id'):
            return jsonify({"success": True, "data": {}})
        return jsonify({
            "success": True,
            "data": {
                "solutionCode": result.get('solutionCode', ''),
                "solutionName": result.get('solutionName', ''),
                "concentration": result.get('concentration', ''),
                "concentrationUnit": result.get('concentrationUnit', ''),
            }
        })
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route('/api/lims/save_solution_b', methods=['POST'])
def lims_save_solution_b():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    payload = request.get_json() or {}
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    if not pid:
        return jsonify({"success": False, "message": "无法获取用户PID，请重新登录"}), 401

    payload['pid'] = pid
    payload['pname'] = pname
    payload['loginId'] = pid

    try:
        url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/saveSolutionConfigure"
        headers = {
            "Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544",
            "Content-Type": "application/json;charset=UTF-8",
        }
        resp = system.session.post(url, json=payload, headers=headers)
        if resp.status_code in (401, 403) or _resp_looks_expired(resp):
            return _expired_response()
        if not resp.ok:
            print(f"[save_solution_b] status={resp.status_code} body={resp.text[:500]}")
            return jsonify({"success": False, "message": f"LIMS 请求失败，状态码: {resp.status_code}"}), 502
        result = resp.json()
        if not result.get('success'):
            return jsonify({
                "success": False,
                "message": result.get('errorDesc') or str(result.get('errorCtx', '配置失败')),
            })
        return jsonify({"success": True})
    except Exception as e:
        print(f"[save_solution_b] exception: {e}")
        return jsonify({"success": False, "message": f"配置异常: {str(e)}"}), 500


@app.route('/api/lims/save_solution_mix', methods=['POST'])
def lims_save_solution_mix():
    """N-1：多个标准品 A 混合定容到 1 瓶。
    复刻 lims_save_solution 的 detail 构造，改为 N 个 detail + 共享 volume/medium，
    pageType=A（标准品作源的合法页类型），顶层 originalCode/controlledNo/concentration 分号拼接。
    与 save_solution 同一上游 saveSolutionConfigure。"""
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    p = request.get_json() or {}
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    if not pid:
        return jsonify({"success": False, "message": "无法获取用户PID，请重新登录"}), 401

    items = p.get('items') or []
    if not items:
        return jsonify({"success": False, "message": "无标准品源"}), 400

    volume_ml = _round_vol(float(p.get('volume_ml', 1)))
    medium = p.get('medium', '')
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    vol_display = _fmt_vol(volume_ml)

    detail_items = []
    conc_parts = []
    original_code_parts = []
    controlled_parts = []
    earliest_expiry = ''

    for idx, it in enumerate(items):
        purity_str = str(it.get('purity_str', '99.5'))
        try:
            purity_value = float(purity_str.replace('%', '').strip())
        except Exception:
            purity_value = 99.5
        original_unit = it.get('original_unit', '%')
        received_unit = it.get('received_unit', 'g')
        try:
            use_quantity = _round_qty(float(it.get('use_quantity', 1)), received_unit)
        except Exception:
            use_quantity = float(it.get('use_quantity', 1))
        receive_id = str(it.get('receive_id', ''))
        original_id = str(it.get('original_id', ''))
        original_name = it.get('original_name', '')
        controlled_no = it.get('controlled_no', '')
        batch_no = it.get('batch_no', '')
        solution_code = it.get('solution_code', '')
        expiry_date = it.get('expiry_date', '')
        original_no = it.get('original_no') or f"A-{original_id}\n{batch_no}\n{solution_code}"

        if original_unit == '%':
            config_conc = purity_value / 100.0 * use_quantity * 1_000_000 / volume_ml
        else:
            config_conc = purity_value * use_quantity / volume_ml
        config_conc = _round_conc(config_conc, original_unit)
        conc_display = _fmt_conc(config_conc, original_unit)
        qty_display = _fmt_qty(use_quantity, received_unit)

        detail_items.append({
            "id": None, "createDatetime": now_str, "serialVersionUID": None,
            "configureId": None, "originalId": receive_id,
            "originalCode": it.get('original_code') or f"L-{receive_id}",
            "originalName": original_name,
            "originalConcentration": f"{purity_str}({original_unit})",
            "concentrationCount": f"{purity_str}({original_unit})",
            "originalUnit": original_unit,
            "receivedQuantity": qty_display, "receivedUint": received_unit,
            "useUantity": qty_display, "useUnit": received_unit,
            "medium": medium,
            "volume": vol_display, "unit": "mL",
            "configurationConcentration": conc_display,
            "configurationUnit": "μg/mL", "configurationUncertainty": None,
            "remark": None, "consumableReceive": None, "dataid": original_id,
            "configurationRecordId": None, "creatorName": None, "creatorId": None,
            "modifierName": None, "modifyDatetime": now_str,
            "conversionFactor": None if original_unit == "μg/mL" else "1",
            "dilutionFactor": None,
            "originalNo": original_no,
            "_X_ID": f"row_{int(time.time()) + idx}",
        })
        conc_parts.append(f"{original_name}:{conc_display}(μg/mL)")
        original_code_parts.append(f"A-{original_id}")
        if controlled_no:
            controlled_parts.append(controlled_no)
        if expiry_date and (not earliest_expiry or expiry_date < earliest_expiry):
            earliest_expiry = expiry_date

    payload = {
        "id": None, "createDatetime": now_str, "serialVersionUID": None,
        "solutionName": p.get('solution_name', ''),
        "solutionCode": p.get('solution_code', ''),
        "deviceIds": None, "deviceNames": p.get('device_names'),
        "configureDate": p.get('configure_date'),
        "validityDate": p.get('validity_date', ''),
        "storageCondition": p.get('storage_condition'),
        "storageLocation": p.get('storage_location', '4-1-华业4-1'),
        "concentration": ";".join(conc_parts),
        "concentrationCount": ";".join(conc_parts),
        "concentrationUnitName": None, "uncertainty": None, "configureOrder": None,
        "originalCode": ";".join(original_code_parts),
        "controlledNo": ";".join(controlled_parts),
        "medium": None, "configuratorId": None,
        "solutionType": p.get('solution_type', 'SOLUTION_TYPE_B'),
        "configuratorName": None, "constantVolume": 0, "totalConstantVolume": 0,
        "usedConstantVolume": 0, "remark": None, "diluteStatus": False,
        "consumableReceive": None, "customType": p.get('customType'), "auditUserName": None,
        "auditTime": None, "diluteConcentrationControl": "[]",
        "pageType": "SOLUTION_TYPE_A",
        "saveDetailList": json.dumps(detail_items, ensure_ascii=False),
        "originalValidityDate": earliest_expiry,
        "configurationTemplateId": None, "intermediateTemplateId": None,
        "temperature": p.get('temperature'), "humidity": p.get('humidity'),
        "computingFormula": None, "configurationMethod": None,
        "configurationProcess": None, "nextAuditUserName": None,
        "disposeUserName": None, "disposeTime": None, "disposeWay": None,
        "configurationSolutionTemplateId": None, "configurationRecordId": None,
        "creatorName": None, "creatorId": None, "modifierName": None,
        "modifyDatetime": now_str, "receivedUint": items[0].get('original_unit', '%'),
        "pid": pid, "pname": pname, "loginId": pid,
    }

    try:
        url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/saveSolutionConfigure"
        headers = {
            "Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544",
            "Content-Type": "application/json;charset=UTF-8",
        }
        resp = system.session.post(url, json=payload, headers=headers)
        if resp.status_code in (401, 403) or _resp_looks_expired(resp):
            return _expired_response()
        if not resp.ok:
            print(f"[save_solution_mix] status={resp.status_code} body={resp.text[:500]}")
            return jsonify({"success": False, "message": f"LIMS 请求失败，状态码: {resp.status_code}"}), 502
        result = resp.json()
        if not result.get('success'):
            return jsonify({
                "success": False,
                "message": result.get('errorDesc') or str(result.get('errorCtx', '配置失败')),
            })
        return jsonify({"success": True})
    except Exception as e:
        print(f"[save_solution_mix] exception: {e}")
        return jsonify({"success": False, "message": f"配置异常: {str(e)}"}), 500


@app.route('/api/lims/solution_code_check', methods=['GET'])
def solution_code_check():
    """代理工作液编号检查请求"""
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401

    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()

    params = request.args.to_dict()
    url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/solutionCodeCheck"

    try:
        resp = system.session.get(url, params=params)
        resp.raise_for_status()
        return jsonify(resp.json()), resp.status_code
    except Exception as e:
        return jsonify({"success": False, "message": f"检查异常: {str(e)}"}), 500


@app.route('/api/lims/save_working_solution', methods=['POST'])
def save_working_solution():
    """代理工作液保存请求"""
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401

    payload = request.get_json() or {}
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()

    # pid/pname/loginId 应该已经在前端添加了，这里不需要再添加
    url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/saveSolutionConfigure"
    headers = {
        "Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544",
        "Content-Type": "application/json;charset=UTF-8",
    }

    try:
        resp = system.session.post(url, json=payload, headers=headers)
        if resp.status_code in (401, 403) or _resp_looks_expired(resp):
            return _expired_response()
        if not resp.ok:
            print(f"[save_working_solution] status={resp.status_code} body={resp.text[:500]}")
            return jsonify({"success": False, "message": f"LIMS 请求失败，状态码: {resp.status_code}"}), 502
        result = resp.json()
        if not result.get('success'):
            return jsonify({
                "success": False,
                "message": result.get('errorDesc') or str(result.get('errorCtx', '配置失败')),
            })
        return jsonify({"success": True})
    except Exception as e:
        print(f"[save_working_solution] exception: {e}")
        return jsonify({"success": False, "message": f"配置异常: {str(e)}"}), 500


@app.route('/api/lims/update_solution', methods=['POST'])
def lims_update_solution():
    """代理工作液/储备液/应用液修改请求 → updateObj1"""
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    payload = request.get_json() or {}
    # 权限检查
    allowed, reason = _check_edit_permission(session, None)
    if not allowed:
        solution_id = payload.get('id')
        if solution_id:
            try:
                system = get_system()
                username = session.get('username', '')
                if system.current_user != username:
                    system.current_user = username
                    system.load_session()
                order_str = payload.get('configureOrder') or payload.get('configure_order')
                record = _fetch_solution_view(system, solution_id, order_str=order_str)
                if record:
                    allowed, reason = _check_edit_permission(session, record)
            except Exception:
                pass
        if not allowed:
            return jsonify({"success": False, "message": reason}), 403
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    if not pid:
        return jsonify({"success": False, "message": "无法获取用户PID，请重新登录"}), 401
    payload['pid'] = pid
    payload['pname'] = pname
    payload['loginId'] = pid
    try:
        url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/updateObj1"
        headers = {
            "Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544",
            "Content-Type": "application/json;charset=UTF-8",
        }
        resp = system.session.post(url, json=payload, headers=headers)
        if resp.status_code in (401, 403) or _resp_looks_expired(resp):
            return _expired_response()
        if not resp.ok:
            print(f"[updateObj1] status={resp.status_code} body={resp.text[:500]}")
            return jsonify({"success": False, "message": f"LIMS 请求失败，状态码: {resp.status_code}"}), 502
        result = resp.json()
        if not result.get("success"):
            err_ctx = result.get('errorCtx') or {}
            err_msg = result.get('errorDesc') or (err_ctx.get('errorMsg') if isinstance(err_ctx, dict) else '') or '修改失败'
            return jsonify({"success": False, "message": err_msg})
        return jsonify({"success": True})
    except Exception as e:
        print(f"[updateObj1] exception: {e}")
        return jsonify({"success": False, "message": f"修改异常: {str(e)}"}), 500


@app.route('/api/lims/audit_solution', methods=['POST'])
def lims_audit_solution():
    """代理溶液配置审核（未审核→已审核）→ auditData"""
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    p = request.get_json() or {}
    solution_id = p.get('id')
    if not solution_id:
        return jsonify({"success": False, "message": "缺少记录ID"})
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    try:
        url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/auditData"
        headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
        resp = system.session.get(url, params={
            "id": solution_id, "pid": pid, "pname": pname, "loginId": pid,
        }, headers=headers)
        if resp.status_code in (401, 403) or _resp_looks_expired(resp):
            return _expired_response()
        if not resp.ok:
            print(f"[auditData] status={resp.status_code} body={resp.text[:500]}")
            return jsonify({"success": False, "message": f"LIMS 请求失败，状态码: {resp.status_code}"}), 502
        result = resp.json()
        if not result.get("success"):
            return jsonify({"success": False, "message": result.get('errorDesc') or str(result.get('errorCtx', '审核失败'))})
        return jsonify({"success": True})
    except Exception as e:
        print(f"[auditData] exception: {e}")
        return jsonify({"success": False, "message": f"审核异常: {str(e)}"}), 500


@app.route('/api/lims/unaudit_solution', methods=['POST'])
def lims_unaudit_solution():
    """代理溶液配置取消审核（已审核→未审核）→ cancelAuditData"""
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    p = request.get_json() or {}
    solution_id = p.get('id')
    if not solution_id:
        return jsonify({"success": False, "message": "缺少记录ID"})
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    try:
        url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/cancelAuditData"
        headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
        resp = system.session.get(url, params={
            "id": solution_id, "pid": pid, "pname": pname, "loginId": pid,
        }, headers=headers)
        if resp.status_code in (401, 403) or _resp_looks_expired(resp):
            return _expired_response()
        if not resp.ok:
            print(f"[cancelAuditData] status={resp.status_code} body={resp.text[:500]}")
            return jsonify({"success": False, "message": f"LIMS 请求失败，状态码: {resp.status_code}"}), 502
        result = resp.json()
        if not result.get("success"):
            return jsonify({"success": False, "message": result.get('errorDesc') or str(result.get('errorCtx', '取消审核失败'))})
        return jsonify({"success": True})
    except Exception as e:
        print(f"[cancelAuditData] exception: {e}")
        return jsonify({"success": False, "message": f"取消审核异常: {str(e)}"}), 500


@app.route('/api/lims/delete_solution', methods=['POST'])
def lims_delete_solution():
    """代理删除溶液配置请求 → delById"""
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    p = request.get_json() or {}
    solution_id = p.get('id')
    solution_type = p.get('type', '')
    if not solution_id:
        return jsonify({"success": False, "message": "缺少记录ID"})
    # 权限检查
    allowed, reason = _check_delete_permission(session, None)
    if not allowed:
        try:
            system = get_system()
            username = session.get('username', '')
            if system.current_user != username:
                system.current_user = username
                system.load_session()
            record = _fetch_solution_view(system, solution_id)
            if record:
                allowed, reason = _check_delete_permission(session, record)
        except Exception:
            pass
        if not allowed:
            return jsonify({"success": False, "message": reason}), 403
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    if not pid:
        return jsonify({"success": False, "message": "无法获取用户PID，请重新登录"}), 401
    try:
        url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/delById"
        form_data = {
            'ids': str(solution_id),
            'type': solution_type,
            'pid': str(pid),
            'pname': pname,
            'loginId': str(pid),
            '_method': 'DELETE',
        }
        resp = system.session.post(url, data=form_data)
        if not resp.ok:
            print(f"[delById] status={resp.status_code} body={resp.text[:500]}")
        ct = resp.headers.get('Content-Type', '')
        if 'html' in ct or resp.text.lstrip().startswith('<!') or resp.text.lstrip().startswith('<html'):
            print(f"[delById] session过期，返回了HTML: {resp.text[:200]}")
            return _expired_response("LIMS 会话已过期，请重新登录后再试")
        result = resp.json()
        if not result.get("success"):
            err_ctx = result.get('errorCtx') or {}
            err_msg = result.get('errorDesc') or (err_ctx.get('errorMsg') if isinstance(err_ctx, dict) else '') or '删除失败'
            return jsonify({"success": False, "message": err_msg})
        return jsonify({"success": True})
    except Exception as e:
        print(f"[delById] exception: {e}")
        return jsonify({"success": False, "message": f"删除异常: {str(e)}"}), 500


@app.route('/api/lims/delete_receive', methods=['POST'])
def lims_delete_receive():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    p = request.get_json() or {}
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid') or system.current_pid or ''
    pname = session.get('display_name') or system.current_real_name or username
    if not pid:
        return jsonify({"success": False, "message": "无法获取用户PID，请重新登录"}), 401
    try:
        # 领用记录删除支持三种入参：ids(领用id) / source_ids(源溶液id，按viewConsumableReceive查领用) / consumable_ids(耗材id)
        consumable_ids = p.get('consumable_ids')
        receive_ids = p.get('ids')
        source_ids = p.get('source_ids')
        receive_id_list = []
        if consumable_ids and not receive_ids:
            if not isinstance(consumable_ids, list):
                consumable_ids = [consumable_ids]
            found_ids = []
            for cid in consumable_ids:
                cid = str(cid).strip()
                if not cid:
                    continue
                url = f"{system.base_url}/detectionManager/manager/consumableReceive/record/{cid}"
                resp = system.session.get(url, params={
                    '_search': 'false', 'pageSize': 9999, 'pageNo': 1,
                    'sidx': '', 'sord': 'asc',
                    'pid': str(pid), 'pname': pname, 'loginId': str(pid),
                })
                if resp.ok:
                    data = resp.json()
                    for rec in (data.get('resultData') or []):
                        rid = rec.get('id')
                        if rid:
                            found_ids.append(str(rid))
            receive_ids = found_ids
        if receive_ids:
            if isinstance(receive_ids, list):
                receive_id_list = [str(i) for i in receive_ids]
            else:
                receive_id_list = [str(i).strip() for i in str(receive_ids).split(',') if str(i).strip()]
        elif source_ids:
            # 按源溶液（被领用的储备液/应用液）查领用：viewConsumableReceive?id=<源溶液id>&type=<源溶液类型>
            # 用于删工作液(D)时补删 LIMS 漏删的、挂在储备液名下的领用记录
            if isinstance(source_ids, list) and source_ids and isinstance(source_ids[0], dict):
                sources = [(str(s.get('id', '')).strip(), s.get('type') or 'SOLUTION_TYPE_B') for s in source_ids if s.get('id')]
            else:
                stype = p.get('source_type') or 'SOLUTION_TYPE_B'
                lst = source_ids if isinstance(source_ids, list) else [source_ids]
                sources = [(str(sid).strip(), stype) for sid in lst if str(sid).strip()]
            for sid, stype in sources:
                vurl = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/viewConsumableReceive"
                vresp = system.session.get(vurl, params={
                    '_search': 'false', 'pageSize': 9999, 'pageNo': 1, 'sidx': '', 'sord': 'asc',
                    'id': sid, 'type': stype,
                    'pid': str(pid), 'pname': pname, 'loginId': str(pid),
                }, headers={"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"})
                if vresp.ok:
                    try:
                        for rec in (vresp.json().get('resultData') or []):
                            rid = rec.get('id')
                            if rid:
                                receive_id_list.append(str(rid))
                    except Exception:
                        pass
        if not receive_id_list:
            return jsonify({"success": True, "message": "无领用记录需要删除"})
        # 删领用走溶液配置模块的 delReceiveId（耗材领用通用接口 consumableReceive 会校验原操作人，配制产生的领用会被拒）
        headers = {
            "Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        }
        failed = []
        for rid in receive_id_list:
            url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/delReceiveId"
            form_data = {
                'ids': rid,
                'pid': str(pid),
                'pname': pname,
                'loginId': str(pid),
                '_method': 'DELETE',
            }
            resp = system.session.post(url, data=form_data, headers=headers)
            if not resp.ok:
                print(f"[deleteReceive] status={resp.status_code} body={resp.text[:500]}")
            ct = resp.headers.get('Content-Type', '')
            if 'html' in ct or resp.text.lstrip().startswith('<!') or resp.text.lstrip().startswith('<html'):
                print(f"[deleteReceive] session过期，返回了HTML: {resp.text[:200]}")
                return _expired_response("LIMS 会话已过期，请重新登录后再试")
            result = resp.json()
            if not result.get("success"):
                err_ctx = result.get('errorCtx') or {}
                err_msg = result.get('errorDesc') or (err_ctx.get('errorMsg') if isinstance(err_ctx, dict) else '') or '删除领用记录失败'
                failed.append(f"{rid}: {err_msg}")
        if failed:
            return jsonify({"success": False, "message": "；".join(failed)})
        return jsonify({"success": True})
    except Exception as e:
        print(f"[deleteReceive] exception: {e}")
        return jsonify({"success": False, "message": f"删除领用异常: {str(e)}"}), 500


@app.route('/api/lims/get_solution_detail', methods=['POST'])
def lims_get_solution_detail():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    p = request.get_json() or {}
    solution_id = p.get('id')
    configure_order = str(p.get('configure_order', '')).strip()
    if not solution_id and not configure_order:
        return jsonify({"success": False, "message": "缺少 id 或 configure_order"})
    try:
        system = get_system()
        username = session.get('username', '')
        if system.current_user != username:
            system.current_user = username
            system.load_session()
        lims_id = solution_id
        if not lims_id and configure_order:
            lims_id = _resolve_order_to_lims_id(system, configure_order)
        if not lims_id:
            return jsonify({"success": False, "message": f"未找到记录: {configure_order}"})
        record = _fetch_solution_view(system, lims_id, order_str=configure_order)
        if not record or not record.get('id'):
            return jsonify({"success": False, "message": "获取详情失败"})
        return jsonify({"success": True, "data": record})
    except Exception as e:
        return jsonify({"success": False, "message": f"获取详情异常: {str(e)}"}), 500


@app.route('/api/lims/trace_export_chain', methods=['POST'])
def lims_trace_export_chain():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401

    p = request.get_json() or {}
    source_ids = p.get('source_ids', [])
    source_orders = p.get('source_orders', [])
    configure_date = str(p.get('configure_date', '')).strip()[:10]
    configurator_name = str(p.get('configurator_name', '')).strip()

    if not configure_date or not configurator_name:
        return jsonify({"success": True, "has_candidates": False, "matched_records": [], "top_ancestor": {}})

    trace_targets = []
    if source_ids:
        trace_targets = [(sid, None) for sid in source_ids]
    elif source_orders:
        trace_targets = [(None, order) for order in source_orders]
    if not trace_targets:
        return jsonify({"success": True, "has_candidates": False, "matched_records": [], "top_ancestor": {}})

    try:
        system = get_system()
        username = session.get('username', '')
        if system.current_user != username:
            system.current_user = username
            system.load_session()

        matched_records, top_ancestor = _trace_export_chain(system, trace_targets, configure_date, configurator_name)

        non_weighing = [r for r in matched_records if not r.get('_is_weighing_top')]
        return jsonify({
            "success": True,
            "has_candidates": len(non_weighing) > 0,
            "matched_records": matched_records,
            "top_ancestor": top_ancestor,
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": True, "has_candidates": False, "matched_records": [], "top_ancestor": {}})


def _parse_purity(raw):
    """'99.5(%)' → (0.995, '%')；'0.982' → (0.982, '')；'' → (None, '')。纯度统一归一为分数。"""
    s = _first_number(raw)
    if not s:
        return None, ''
    try:
        val = float(s)
    except ValueError:
        return None, ''
    is_pct = '%' in str(raw)
    return (val / 100.0, '%') if is_pct else (val, '')


def _frac_from_raw(raw):
    """扩展不确定度串 → 绝对分数：'0.5%' → 0.005；'0.005' → 0.005；'' → None。
    ponytail: 假设 % 为质量分数扩展不确定度；真实台账格式待按 dump 数据核对。"""
    s = _first_number(raw)
    if not s:
        return None
    try:
        n = float(s)
    except ValueError:
        return None
    return n / 100.0 if '%' in str(raw) else n


def _num_from_raw(raw):
    """任意串 → float 或 None。"""
    if raw is None or raw == '':
        return None
    s = _first_number(raw)
    try:
        return float(s) if s else None
    except ValueError:
        return None


def _extract_balance_id(text):
    """从称量设备文本提取天平编号(如 CK-SB295-CG)；无匹配返回 None。"""
    m = re.search(r'CK-SB\d+-[A-Z]+', str(text or ''))
    return m.group(0) if m else None


# 多点工作液 detailList 源行里的逐级稀释目标浓度字段（dilutionOne..Twelve）
_DILUTION_KEYS = ('dilutionOne', 'dilutionTwo', 'dilutionThree', 'dilutionFour', 'dilutionFive',
                  'dilutionSix', 'dilutionSeven', 'dilutionEight', 'dilutionNine', 'dilutionTen',
                  'dilutionEleven', 'dilutionTwelve')


def _extract_crm_from_chain(matched):
    """从溯源链提取 A 级标准物质(CRM)：controlledNo(CK-CG号)、名称、纯度。
    lift 自 _trace_to_a（4565-4631）的 detailList 挖掘逻辑。返回 dict 或 None。"""
    for r in matched:
        if r.get('level') == 'D':
            continue
        if not str(r.get('original_code') or '').strip().startswith('A-'):
            continue
        a_name = str(r.get('parent_name') or '').strip()
        a_controlled_no = ''
        a_concentration = ''  # A 级纯度，如 "99.5(%)"（见 1473/2562 行 originalConcentration 约定）
        for dl in (r.get('_detail_list') or []):
            dl_name = str(dl.get('originalName', '')).strip()
            dl_no = str(dl.get('originalNo', '')).strip()
            lines = [l.strip() for l in dl_no.split('\n') if l.strip()]
            potential_no = lines[-1] if lines else ''
            if dl_name and a_name and dl_name == a_name:
                a_controlled_no = potential_no
                a_concentration = str(dl.get('originalConcentration', '') or '').strip()
                break
            if not a_controlled_no and potential_no:
                a_controlled_no = potential_no
                a_concentration = str(dl.get('originalConcentration', '') or '').strip()
        if a_controlled_no or a_name:
            purity_val, purity_unit = _parse_purity(a_concentration)
            return {
                'controlled_no': a_controlled_no,
                'name': a_name,
                'purity': purity_val,
                'purity_raw': a_concentration,
                'purity_unit': purity_unit,
                'storage_condition': str(r.get('storage_condition') or '').strip(),
            }
    return None


def _fetch_crm_ledger(system, controlled_no):
    """按 controlledNo(CK-CG号) 查标准物质台账 consumableBill/pageObj，返回完整 voList[0] 或 None。
    lift 自 /api/lims/storage_conditions（4667-4691），但返回整条记录以取 uncertainty/purity/batchNo。"""
    cb_params = {
        "_search": "false", "nd": str(int(time.time() * 1000)),
        "pageSize": 30, "pageNo": 1, "sidx": "", "sord": "asc",
        "type": "CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE",
        "receiveUserName": "", "receiveStartDate": "", "receiveEndDate": "",
        "confirmUserName": "", "confirmStartDate": "", "confirmEndDate": "",
        "invoiceNo": "", "groupId": "", "casNo": "", "orgName": "", "state": "normal",
        "keyword": controlled_no,
        "pid": system.current_pid or '', "pname": system.current_real_name or '',
        "loginId": system.current_pid or '',
    }
    for _st in ("normal", "history", "overdue"):
        cb_params["status"] = _st
        try:
            r = system.session.get(f"{system.base_url}/detectionManager/manager/consumableBill/pageObj",
                                   params=cb_params)
            if r.status_code != 200:
                continue
            vol = (r.json().get('resultData') or {}).get('voList') or []
            if vol:
                return vol[0]
        except Exception:
            continue
    return None


@app.route('/api/lims/trace_working_solution', methods=['GET'])
def lims_trace_working_solution():
    """工作液溯源 → 原始溯源数据（稀释链 B/C/D + A 级 CRM 不确定度模块），供 mup-web 溯源填充调用。
    服务账号 OCR 登录 LIMS，不依赖 Flask 用户会话。返回 {success, data:{code,stock_source,chain,crm,...}}。"""
    code = (request.args.get('code') or '').strip()
    if not code:
        return jsonify({"success": False, "message": "请输入工作液编号(如 D-9203)"}), 400

    system, err = _trace_service_session()
    if err:
        return jsonify({"success": False, "message": err}), 502

    try:
        # (None, code) → 由 _trace_export_chain 内部 _fetch_record 解析编号；''/'' 不限日期与配置人，自由追溯
        matched, _top = _trace_export_chain(system, [(None, code)], '', '')
        if not matched:
            return jsonify({"success": True, "data": {
                "code": code, "stock_source": None, "chain": [], "crm": None,
                "multi_source": False, "info_chain": ""}})

        # A 级 CRM 提取 + 台账不确定度查询（标准溶液 urel(C) 模块）
        crm = _extract_crm_from_chain(matched)
        if crm and crm.get('controlled_no'):
            ledger = _fetch_crm_ledger(system, crm['controlled_no'])
            if ledger:
                u_raw = str(ledger.get('uncertainty') or '').strip()
                crm['uncertainty_raw'] = u_raw
                crm['u_value'] = _frac_from_raw(u_raw)
                lp = str(ledger.get('purity') or '').strip()  # 台账 purity 优先于链上推断
                if lp:
                    pv, pu = _parse_purity(lp)
                    if pv is not None:
                        crm['purity'], crm['purity_unit'], crm['purity_raw'] = pv, pu, lp
                crm['k_raw'] = ledger.get('k') or ledger.get('coverageFactor') or ledger.get('包含因子')
                crm['k_value'] = _num_from_raw(crm['k_raw'])
                crm['batchNo'] = str(ledger.get('batchNo') or '').strip()

        # 天平设备编号（称量型 B 记录的配置设备 deviceNames；仅 solid 储备液有）
        balance = None
        for r in matched:
            if r.get('_is_weighing_top'):
                dev = str(r.get('device_names') or '').strip()
                we = r.get('weighing_equipment') or {}
                we_json = r.get('weighing_equipment_json') or ''
                raw_name = str(we.get('name') or '').strip() if isinstance(we, dict) else str(we or '').strip()
                txt = dev or ((raw_name + ' ' + we_json).strip())
                balance = {'id': _extract_balance_id(txt), 'raw': txt[:200]}
                break

        # D 型多点工作液：detailList 源行补移取体积；定容体积按 移取量*父浓度/本行首点浓度 反推（记录级常缺）
        for r in matched:
            if r.get('level') != 'D':
                continue
            if not str(r.get('received_quantity') or '').strip():
                oc = str(r.get('original_code') or '')
                parent_order = next((p.strip() for p in re.split(r'[,，]', oc) if p.strip()), '')
                for dl in (r.get('_detail_list') or []):
                    if parent_order and str(dl.get('originalCode', '')).strip() == parent_order:
                        rq = str(dl.get('receivedQuantity', '') or '').strip()
                        if rq:
                            r['received_quantity'] = rq
                        ru = str(dl.get('receivedUint', '') or '').strip()
                        if ru:
                            r['received_unit'] = ru
                        break
            if str(r.get('constant_volume') or '') in ('', '0', '0.0'):
                _rq = _num_from_raw(r.get('received_quantity'))
                _pc = _num_from_raw(r.get('parent_concentration'))
                _rc = _num_from_raw(r.get('concentration'))
                if _rq and _pc and _rc:
                    r['constant_volume'] = f"{_rq * _pc / _rc:.2f}"

        # 多点工作液稀释系列：detailList 按 groupName 分行存 移取体积/定容体积，源行(originalCode)存目标浓度，
        # 三者均以 dilutionOne..Twelve 为各点值。直接读取，无需反推。
        dilution_series = None
        for r in matched:
            if r.get('level') != 'D':
                continue
            pip_vols, flask_vols, targets = [], [], []
            for _dl in (r.get('_detail_list') or []):
                _gn = str(_dl.get('groupName', '')).strip()
                _vals = []
                for _k in _DILUTION_KEYS:
                    _v = _num_from_raw(_dl.get(_k))
                    if _v is not None:
                        _vals.append(_v)
                if not _vals:
                    continue
                if _gn == '移取体积':
                    pip_vols = _vals
                elif _gn == '定容体积':
                    flask_vols = _vals
                elif _dl.get('originalCode'):
                    targets = _vals
            if len(targets) >= 2:
                stock_conc = _num_from_raw(r.get('parent_concentration'))
                rows, _prev = [], stock_conc
                for _i, _t in enumerate(targets):
                    rows.append({
                        'mother_conc': _prev,
                        'target_conc': _t,
                        'pip_vol': pip_vols[_i] if _i < len(pip_vols) else None,
                        'flask_vol': flask_vols[_i] if _i < len(flask_vols) else None,
                    })
                    _prev = _t   # 逐级稀释：下行母液 = 本行目标
                dilution_series = {'stock_conc': stock_conc, 'rows': rows}
            break

        top = matched[0]
        stock_source = 'solid' if (top.get('_is_weighing_top') or (top.get('received_unit') or '').strip() == 'g') \
            else 'liquid_dilute'

        chain_out = []
        for r in matched:        # 剥除内部大字段 _detail_list
            rr = dict(r)
            rr.pop('_detail_list', None)
            chain_out.append(rr)

        multi_source = any(r.get('source_details') for r in matched)
        info_chain = " → ".join(f"{r.get('level', '')}({r.get('configure_order', '')})"
                                for r in matched if r.get('configure_order'))

        return jsonify({"success": True, "data": {
            "code": code,
            "stock_source": stock_source,
            "chain": chain_out,
            "crm": crm,
            "balance": balance,
            "dilution_series": dilution_series,
            "multi_source": multi_source,
            "info_chain": info_chain,
        }})
    except Exception as e:
        import traceback
        traceback.print_exc()
        _TRACE_SESS["res"] = None  # 疑似会话过期，清缓存重登
        return jsonify({"success": False, "message": f"溯源失败: {e}"}), 500


def _fill_rf10_09_item(doc, item_payload):
    """填充 RF10-09 模板的单条记录"""
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    solution_name = item_payload.get('solution_name', '')
    custom_num = item_payload.get('base_custom_num', '') or item_payload.get('custom_num', '')
    purity_str = str(item_payload.get('purity_str', ''))
    original_unit = item_payload.get('original_unit', '%')
    purity_display = f"{purity_str}({original_unit})"
    device_names = item_payload.get('device_names', '') or ''
    solution_code = item_payload.get('solution_code', '')
    validity_date = item_payload.get('validity_date', '')
    configure_date = item_payload.get('configure_date', datetime.date.today().strftime('%Y-%m-%d'))
    medium = item_payload.get('medium', '')
    storage_cond = item_payload.get('storage_condition', '') or ''
    temp_val = item_payload.get('temperature', '') or ''
    humid_val = item_payload.get('humidity', '') or ''
    received_unit = item_payload.get('received_unit', 'g')

    try:
        purity_value = float(purity_str.replace('%', '').strip())
        vol_f = _round_vol(float(item_payload.get('volume_ml', '')))
        qty_f = _round_qty(float(item_payload.get('use_quantity', '')), received_unit)
        if original_unit == '%':
            config_conc_raw = purity_value / 100.0 * qty_f * 1_000_000 / vol_f
        else:
            config_conc_raw = purity_value * qty_f / vol_f
        config_conc = _round_conc(config_conc_raw, original_unit)
        conc = _fmt_conc(config_conc, original_unit)
        use_qty = _fmt_qty(qty_f, received_unit)
        volume = _fmt_vol(vol_f)
    except Exception:
        conc = str(item_payload.get('config_conc', ''))
        use_qty = str(item_payload.get('use_quantity', ''))
        volume = str(item_payload.get('volume_ml', ''))

    if len(doc.paragraphs) > 1:
        p1 = doc.paragraphs[1]
        p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if temp_val and len(p1.runs) > 4:
            p1.runs[2].text = f" {str(temp_val).center(10)} "
            p1.runs[2].font.underline = True
        if len(p1.runs) > 6:
            p1.runs[5].text = "\t\t"
            p1.runs[6].text = "\t\t"
        if humid_val and len(p1.runs) > 10:
            p1.runs[9].text = str(humid_val).center(10)
            p1.runs[9].font.underline = True
            p1.runs[10].text = "%" + " " * 13
        elif len(p1.runs) > 10:
            p1.runs[10].text = "                        "
        if storage_cond and len(p1.runs) > 12:
            p1.runs[12].text = f" {storage_cond} "
            p1.runs[12].font.underline = True

    table = doc.tables[0]
    if received_unit == 'g':
        table.cell(0, 1).text = solution_name
        table.cell(0, 5).text = custom_num
        table.cell(1, 1).text = purity_display
        table.cell(1, 5).text = device_names
        purity_num = purity_str.split('(')[0].strip() if '(' in purity_str else purity_str
        cell_2_1 = table.cell(2, 1)
        cell_2_1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_2_1.paragraphs[0].text = f"称取{use_qty}g标准品，用{medium}定容至{volume}mL容量瓶中，保存于{storage_cond}"
        cell_3_1 = table.cell(3, 1)
        cell_3_1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_3_1.paragraphs[0].text = f"X=（{use_qty}g*{purity_num}%/{volume}mL）*10^6={conc}μg/mL"
        table.cell(4, 1).text = solution_code
        table.cell(5, 1).text = validity_date
        table.cell(7, 6).text = configure_date
    else:
        if '(' in purity_display:
            purity_val = purity_display.split('(')[0].strip()
            purity_unit = purity_display.split('(')[1].rstrip(')').strip()
        else:
            purity_val = purity_display
            purity_unit = ''
        table.cell(1, 1).text = solution_name
        table.cell(1, 5).text = custom_num
        table.cell(1, 7).text = purity_val
        cell_1_6 = table.cell(1, 6)
        if len(cell_1_6.paragraphs) > 1:
            cell_1_6.paragraphs[1].text = f"({purity_unit})"
        for ci in range(len(table.row_cells(3))):
            cell = table.cell(3, ci)
            p0 = cell.paragraphs[0]
            if "(      )" in p0.text:
                p0.text = p0.text.replace("(      )", f"({purity_unit})")
            elif "(     )" in p0.text:
                p0.text = p0.text.replace("(     )", "(mL)")
            elif "(    )" in p0.text:
                p0.text = p0.text.replace("(    )", "(μg/mL)")
        table.cell(4, 0).text = purity_val
        table.cell(4, 1).text = use_qty
        table.cell(4, 2).text = medium
        table.cell(4, 3).text = volume
        table.cell(4, 4).text = conc
        table.cell(4, 5).text = solution_code
        table.cell(4, 6).text = configure_date
        table.cell(4, 8).text = validity_date


@app.route('/api/lims/export_docx', methods=['POST'])
def lims_export_docx():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    from docx import Document as DocxDocument
    from docx.enum.table import WD_ALIGN_VERTICAL
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import io, copy

    p = request.get_json()

    # 合并导出分支
    items_list = p.get('items_list')
    if items_list:
        template_name = 'RF10-09 标准溶液配制记录(1).docx'
        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'word_templates', template_name)
        if not os.path.exists(template_path):
            return jsonify({"success": False, "message": f"模板文件不存在: {template_name}"}), 404

        first_code = items_list[0].get('solution_code', '') if items_list else ''
        doc = DocxDocument(template_path)
        _fill_rf10_09_item(doc, items_list[0])

        # 保存第一个文档的 sectPr（含页眉页脚定义），稍后移到末尾
        body = doc.element.body
        saved_sectPr = None
        for child in list(body):
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'sectPr':
                saved_sectPr = child
                body.remove(child)
                break

        from docx.oxml.ns import qn as _qn
        for item_payload in items_list[1:]:
            doc.add_page_break()
            tpl = DocxDocument(template_path)
            _fill_rf10_09_item(tpl, item_payload)
            for el in list(tpl.element.body):
                tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
                if tag == 'sectPr':
                    continue
                text = ''.join(el.itertext()).strip()
                if not text and not el.findall('.//' + _qn('w:t')):
                    continue
                doc.element.body.append(copy.deepcopy(el))

        # 将 sectPr 放回 body 末尾，恢复页眉页脚
        if saved_sectPr is not None:
            body.append(saved_sectPr)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        download_name = f"{len(items_list)}条_配制记录_合并.docx"
        from urllib.parse import quote as _urlquote
        from flask import send_file
        resp = send_file(buf, as_attachment=True, download_name=download_name, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        resp.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{_urlquote(download_name)}"
        return resp

    # 单条导出（原逻辑）
    received_unit = p.get('received_unit', 'g')
    template_name = "RF10-09 标准溶液配制记录(1).docx" if received_unit == 'g' else "RF10-10 标准溶液配制记录（稀释）(1).docx"
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'word_templates', template_name)
    if not os.path.exists(template_path):
        return jsonify({"success": False, "message": f"模板文件不存在: {template_name}"}), 404

    solution_name = p.get('solution_name', '')
    custom_num = p.get('base_custom_num', '') or p.get('custom_num', '')
    purity_str = str(p.get('purity_str', ''))
    original_unit = p.get('original_unit', '%')
    purity_display = f"{purity_str}({original_unit})"
    device_names = p.get('device_names', '') or ''
    solution_code = p.get('solution_code', '')
    validity_date = p.get('validity_date', '')
    configure_date = p.get('configure_date', datetime.date.today().strftime('%Y-%m-%d'))
    medium = p.get('medium', '')
    storage_cond = p.get('storage_condition', '') or ''
    temp_val = p.get('temperature', '') or ''
    humid_val = p.get('humidity', '') or ''

    try:
        purity_value = float(purity_str.replace('%', '').strip())
        vol_f = _round_vol(float(p.get('volume_ml', '')))
        qty_f = _round_qty(float(p.get('use_quantity', '')), received_unit)
        if original_unit == '%':
            config_conc_raw = purity_value / 100.0 * qty_f * 1_000_000 / vol_f
        else:
            config_conc_raw = purity_value * qty_f / vol_f
        config_conc = _round_conc(config_conc_raw, original_unit)
        conc = _fmt_conc(config_conc, original_unit)
        use_qty = _fmt_qty(qty_f, received_unit)
        volume = _fmt_vol(vol_f)
    except Exception:
        conc = str(p.get('config_conc', ''))
        use_qty = str(p.get('use_quantity', ''))
        volume = str(p.get('volume_ml', ''))

    doc = DocxDocument(template_path)

    if len(doc.paragraphs) > 1:
        p1 = doc.paragraphs[1]
        p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if temp_val and len(p1.runs) > 4:
            p1.runs[2].text = f" {str(temp_val).center(10)} "
            p1.runs[2].font.underline = True
        if len(p1.runs) > 6:
            p1.runs[5].text = "\t\t"
            p1.runs[6].text = "\t\t"
        if humid_val and len(p1.runs) > 10:
            p1.runs[9].text = str(humid_val).center(10)
            p1.runs[9].font.underline = True
            p1.runs[10].text = "%" + " " * 13
        elif len(p1.runs) > 10:
            p1.runs[10].text = "                        "
        if storage_cond and len(p1.runs) > 12:
            p1.runs[12].text = f" {storage_cond} "
            p1.runs[12].font.underline = True

    table = doc.tables[0]
    if received_unit == 'g':
        table.cell(0, 1).text = solution_name
        table.cell(0, 5).text = custom_num
        table.cell(1, 1).text = purity_display
        table.cell(1, 5).text = device_names
        purity_num = purity_str.split('(')[0].strip() if '(' in purity_str else purity_str
        cell_2_1 = table.cell(2, 1)
        cell_2_1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_2_1.paragraphs[0].text = f"称取{use_qty}g标准品，用{medium}定容至{volume}mL容量瓶中，保存于{storage_cond}"
        cell_3_1 = table.cell(3, 1)
        cell_3_1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell_3_1.paragraphs[0].text = f"X=（{use_qty}g*{purity_num}%/{volume}mL）*10^6={conc}μg/mL"
        table.cell(4, 1).text = solution_code
        table.cell(5, 1).text = validity_date
        table.cell(7, 6).text = configure_date
    else:
        if '(' in purity_display:
            purity_val = purity_display.split('(')[0].strip()
            purity_unit = purity_display.split('(')[1].rstrip(')').strip()
        else:
            purity_val = purity_display
            purity_unit = ''
        table.cell(1, 1).text = solution_name
        table.cell(1, 5).text = custom_num
        table.cell(1, 7).text = purity_val
        cell_1_6 = table.cell(1, 6)
        if len(cell_1_6.paragraphs) > 1:
            cell_1_6.paragraphs[1].text = f"({purity_unit})"
        for ci in range(len(table.row_cells(3))):
            cell = table.cell(3, ci)
            p0 = cell.paragraphs[0]
            if "(      )" in p0.text:
                p0.text = p0.text.replace("(      )", f"({purity_unit})")
            elif "(     )" in p0.text:
                p0.text = p0.text.replace("(     )", "(mL)")
            elif "(    )" in p0.text:
                p0.text = p0.text.replace("(    )", "(μg/mL)")
        table.cell(4, 0).text = purity_val
        table.cell(4, 1).text = use_qty
        table.cell(4, 2).text = medium
        table.cell(4, 3).text = volume
        table.cell(4, 4).text = conc
        table.cell(4, 5).text = solution_code
        table.cell(4, 6).text = configure_date
        table.cell(4, 8).text = validity_date

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    download_name = (f"{solution_code}_配制记录.docx" if solution_code else "配制记录.docx").replace("/", "-").replace("\\", "-")
    from flask import send_file
    return send_file(buf, as_attachment=True, download_name=download_name,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


# ── Word XML helpers (模块级，供 export_bbcd_docx / export_verification_docx 复用) ──
from docx.oxml.ns import qn as _docx_qn
from docx.oxml import OxmlElement as _docx_OxmlElement

def _docx_make_run(text, sz=18, underline=False):
    r_el = _docx_OxmlElement('w:r')
    rPr  = _docx_OxmlElement('w:rPr')
    if underline:
        u_el = _docx_OxmlElement('w:u')
        u_el.set(_docx_qn('w:val'), 'single')
        rPr.append(u_el)
    sz_el = _docx_OxmlElement('w:sz')
    sz_el.set(_docx_qn('w:val'), str(sz))
    szCs_el = _docx_OxmlElement('w:szCs')
    szCs_el.set(_docx_qn('w:val'), str(sz))
    rPr.append(sz_el)
    rPr.append(szCs_el)
    r_el.append(rPr)
    t_el = _docx_OxmlElement('w:t')
    t_el.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t_el.set('{http://www.w3.org/XML/1998/XMLSchema-instance}space', 'preserve')
    r_el.append(t_el)
    return r_el

def _docx_set_tc_labelled_inline(tc, title, values, sz=18):
    """Fill cell with title + values joined by ；on same line, left-aligned."""
    el = _docx_unwrap_tc(tc)
    _docx_clear_tc(el)
    p_el = _docx_OxmlElement('w:p')
    pPr = _docx_OxmlElement('w:pPr')
    p_el.append(pPr)
    # Title + values as single run (no bold)
    full_text = title + ('；'.join(values) if values else '')
    p_el.append(_docx_make_run(full_text, sz))
    el.append(p_el)

def _docx_set_v_align_center(tc):
    """Set vertical alignment of a cell to center."""
    el = _docx_unwrap_tc(tc)
    tcPr = el.find(_docx_qn('w:tcPr'))
    if tcPr is None:
        tcPr = _docx_OxmlElement('w:tcPr')
        el.insert(0, tcPr)
    vAlign = tcPr.find(_docx_qn('w:vAlign'))
    if vAlign is None:
        vAlign = _docx_OxmlElement('w:vAlign')
        tcPr.append(vAlign)
    vAlign.set(_docx_qn('w:val'), 'center')

def _docx_set_v_align_bottom(tc):
    """Set vertical alignment of a cell to bottom."""
    el = _docx_unwrap_tc(tc)
    tcPr = el.find(_docx_qn('w:tcPr'))
    if tcPr is None:
        tcPr = _docx_OxmlElement('w:tcPr')
        el.insert(0, tcPr)
    vAlign = tcPr.find(_docx_qn('w:vAlign'))
    if vAlign is None:
        vAlign = _docx_OxmlElement('w:vAlign')
        tcPr.append(vAlign)
    vAlign.set(_docx_qn('w:val'), 'bottom')

def _docx_set_tc_checkbox(tc, text, sz=21, center=False):
    """Fill cell with checkbox text in 宋体(SimSun). Supports \\n for line breaks."""
    el = _docx_unwrap_tc(tc)
    _docx_clear_tc(el)
    p_el = _docx_OxmlElement('w:p')
    pPr = _docx_OxmlElement('w:pPr')
    jc = _docx_OxmlElement('w:jc')
    jc.set(_docx_qn('w:val'), 'center' if center else 'left')
    pPr.append(jc)
    p_el.append(pPr)
    # Build runs: split text on \n, insert w:br between parts
    def _make_r(t):
        r_el = _docx_OxmlElement('w:r')
        rPr = _docx_OxmlElement('w:rPr')
        rFonts = _docx_OxmlElement('w:rFonts')
        rFonts.set(_docx_qn('w:ascii'), 'SimSun')
        rFonts.set(_docx_qn('w:eastAsia'), 'SimSun')
        rFonts.set(_docx_qn('w:hAnsi'), 'SimSun')
        rPr.append(rFonts)
        sz_el = _docx_OxmlElement('w:sz')
        sz_el.set(_docx_qn('w:val'), str(sz))
        szCs_el = _docx_OxmlElement('w:szCs')
        szCs_el.set(_docx_qn('w:val'), str(sz))
        rPr.append(sz_el)
        rPr.append(szCs_el)
        r_el.append(rPr)
        t_el = _docx_OxmlElement('w:t')
        t_el.text = t
        if t and (t[0] == ' ' or t[-1] == ' '):
            t_el.set('{http://www.w3.org/XML/1998/XMLSchema-instance}space', 'preserve')
        r_el.append(t_el)
        return r_el

    parts = text.split('\n')
    for i, part in enumerate(parts):
        p_el.append(_make_r(part))
        if i < len(parts) - 1:
            br_r = _docx_OxmlElement('w:r')
            br_el = _docx_OxmlElement('w:br')
            br_r.append(br_el)
            p_el.append(br_r)
    el.append(p_el)

def _docx_make_para(text, center=False, sz=18):
    p_el = _docx_OxmlElement('w:p')
    pPr  = _docx_OxmlElement('w:pPr')
    if center:
        jc = _docx_OxmlElement('w:jc')
        jc.set(_docx_qn('w:val'), 'center')
        pPr.append(jc)
    p_el.append(pPr)
    p_el.append(_docx_make_run(text, sz))
    return p_el

def _docx_unwrap_tc(tc):
    """Return the underlying lxml element for a table cell (_Cell -> CT_Tc)."""
    return tc._tc if hasattr(tc, '_tc') else tc

def _docx_clear_tc(tc):
    tc = _docx_unwrap_tc(tc)
    for p in tc.findall(_docx_qn('w:p')):
        tc.remove(p)

def _docx_set_tc_text(tc, text, center=True, sz=18):
    el = _docx_unwrap_tc(tc)
    _docx_clear_tc(el)
    el.append(_docx_make_para(text, center=center, sz=sz))

def _docx_set_tc_two_lines(tc, line1, line2, sz=18):
    el = _docx_unwrap_tc(tc)
    _docx_clear_tc(el)
    el.append(_docx_make_para(line1, center=True, sz=sz))
    el.append(_docx_make_para(line2, center=True, sz=sz))

def _docx_set_tc_multi_lines(tc, lines, sz=18):
    el = _docx_unwrap_tc(tc)
    _docx_clear_tc(el)
    for line in lines:
        el.append(_docx_make_para(line, center=True, sz=sz))

def _docx_set_vmerge(tc, mode):
    el = _docx_unwrap_tc(tc)
    tcPr = el.find(_docx_qn('w:tcPr'))
    if tcPr is None:
        tcPr = _docx_OxmlElement('w:tcPr')
        el.insert(0, tcPr)
    vm = tcPr.find(_docx_qn('w:vMerge'))
    if vm is None:
        vm = _docx_OxmlElement('w:vMerge')
        tcPr.append(vm)
    if mode == 'restart':
        vm.set(_docx_qn('w:val'), 'restart')
    else:
        if _docx_qn('w:val') in vm.attrib:
            del vm.attrib[_docx_qn('w:val')]

def _docx_set_row_cant_split(row):
    """设置表格行不允许跨页分割"""
    tr_el = row._tr if hasattr(row, '_tr') else row
    trPr = tr_el.find(_docx_qn('w:trPr'))
    if trPr is None:
        trPr = _docx_OxmlElement('w:trPr')
        tr_el.insert(0, trPr)
    # 添加 cantSplit 属性
    cantSplit = trPr.find(_docx_qn('w:cantSplit'))
    if cantSplit is None:
        cantSplit = _docx_OxmlElement('w:cantSplit')
        trPr.append(cantSplit)

def _docx_set_row_keep_next(row):
    """设置表格行与下一行保持在同一页"""
    tr_el = row._tr if hasattr(row, '_tr') else row
    # 在表格行的每个单元格的段落中设置 keepNext
    for tc_el in tr_el.findall(_docx_qn('w:tc')):
        for p_el in tc_el.findall(_docx_qn('w:p')):
            pPr = p_el.find(_docx_qn('w:pPr'))
            if pPr is None:
                pPr = _docx_OxmlElement('w:pPr')
                p_el.insert(0, pPr)
            keepNext = pPr.find(_docx_qn('w:keepNext'))
            if keepNext is None:
                keepNext = _docx_OxmlElement('w:keepNext')
                pPr.append(keepNext)

def _docx_set_row_height(row, height_twips):
    """设置表格行高度（单位：twips，1磅=20 twips）"""
    tr_el = row._tr if hasattr(row, '_tr') else row
    trPr = tr_el.find(_docx_qn('w:trPr'))
    if trPr is None:
        trPr = _docx_OxmlElement('w:trPr')
        tr_el.insert(0, trPr)
    # 设置行高
    trHeight = trPr.find(_docx_qn('w:trHeight'))
    if trHeight is None:
        trHeight = _docx_OxmlElement('w:trHeight')
        trPr.append(trHeight)
    trHeight.set(_docx_qn('w:val'), str(height_twips))
    trHeight.set(_docx_qn('w:hRule'), 'exact')  # 精确高度


# ── 核查记录导出 PDF 合并辅助函数 ──

def _detect_encoding(raw_bytes):
    """检测字节数组的编码，按常见中文编码顺序尝试。"""
    for enc in ['utf-8', 'gbk', 'gb2312', 'gb18030', 'big5', 'latin-1']:
        try:
            text = raw_bytes.decode(enc)
            return text, enc
        except (UnicodeDecodeError, LookupError):
            continue
    return raw_bytes.decode('gbk', errors='replace'), 'gbk'


def _register_chinese_font():
    """检测并注册可用的中文字体，返回 ReportLab 字体名称。"""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    font_candidates = [
        ('SimSun', 'simsun.ttc'),
        ('NSimSun', 'simsun.ttc'),
        ('KaiTi', 'simkai.ttf'),
        ('STSong', 'STSONG.TTF'),
        ('SimHei', 'simhei.ttf'),
    ]
    fonts_dir = os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts')
    for name, fname in font_candidates:
        fpath = os.path.join(fonts_dir, fname)
        if os.path.exists(fpath):
            try:
                pdfmetrics.registerFont(TTFont(name, fpath))
                return name
            except Exception:
                continue
    return 'Helvetica'


def _epatemp_txt_to_pdf_bytes(text_content):
    """将 epatemp.txt 文本内容用 ReportLab 生成为 PDF 字节。

    排版参数与 epatemp_to_pdf.py 的 txt_to_pdf() 一致：
    A4、零边距、12pt/14pt 行高、&nbsp; 对齐。
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.colors import white
    from reportlab.platypus import SimpleDocTemplate, Paragraph

    buf = BytesIO()
    font_name = _register_chinese_font()

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name='EpatempEmpty', fontName=font_name, fontSize=12, leading=14,
        spaceBefore=0, spaceAfter=0, textColor=white))
    styles.add(ParagraphStyle(
        name='EpatempLine', fontName=font_name, fontSize=12, leading=14,
        spaceBefore=0, spaceAfter=0))

    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=40, leftMargin=0,
                            topMargin=0, bottomMargin=0)
    story = []

    lines = text_content.split('\n')
    # 去除末尾空白行
    while lines and lines[-1].strip() == '':
        lines.pop()

    for line in lines:
        stripped = line.rstrip()
        if stripped == '':
            story.append(Paragraph('&nbsp;', styles['EpatempEmpty']))
        else:
            story.append(Paragraph(stripped.replace(' ', '&nbsp;'), styles['EpatempLine']))

    doc.build(story)
    buf.seek(0)
    return buf.read()


def _docx_to_pdf_bytes(docx_bytes):
    """用 Word COM 将 docx 字节转为 PDF 字节。"""
    import tempfile
    import sys, os
    # NSSM 服务环境下 pywintypes DLL 可能不在 PATH 中，手动加入
    _pywin32_sysdir = os.path.join(os.path.dirname(sys.executable), '..', 'Lib', 'site-packages', 'pywin32_system32')
    if os.path.isdir(_pywin32_sysdir) and _pywin32_sysdir not in sys.path:
        sys.path.insert(0, os.path.abspath(_pywin32_sysdir))
    import win32com.client
    import pythoncom

    tmp_docx = None
    tmp_pdf = None
    word_app = None
    try:
        pythoncom.CoInitialize()
        # 写入临时 docx 文件
        tmp_docx = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
        tmp_docx.write(docx_bytes)
        tmp_docx.close()

        tmp_pdf = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
        tmp_pdf.close()

        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False
        com_doc = word_app.Documents.Open(os.path.abspath(tmp_docx.name))
        if com_doc is None:
            raise RuntimeError("Word.Documents.Open returned None - service profile missing Desktop folder")
        com_doc.SaveAs(os.path.abspath(tmp_pdf.name), FileFormat=17)  # wdFormatPDF
        com_doc.Close(False)

        with open(tmp_pdf.name, 'rb') as f:
            return f.read()
    finally:
        if word_app:
            try:
                word_app.Quit(False)
            except Exception:
                pass
        for p in [tmp_docx, tmp_pdf]:
            if p:
                try:
                    os.unlink(p.name)
                except OSError:
                    pass
        pythoncom.CoUninitialize()


def _merge_pdfs(pdf_pages):
    """合并多个 PDF 字节，返回合并后的 PDF 字节。

    用 PyPDF2 字节级搬运，保留各页原始字体子集不变。pypdfium2 的
    import_pages 会重新序列化页面、扰动 reportlab 子集字体的 cmap/ToUnicode
    编码，曾导致合并后的定量报告页在 XP 打印时数字 5/6 乱码。
    """
    from PyPDF2 import PdfReader, PdfWriter

    writer = PdfWriter()
    for pdf_bytes in pdf_pages:
        reader = PdfReader(BytesIO(pdf_bytes))
        for page in reader.pages:
            writer.add_page(page)

    buf = BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()

@app.route('/api/lims/export_bbcd_docx', methods=['POST'])
def lims_export_bbcd_docx():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401

    from docx import Document as DocxDocument
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import io, copy

    p = request.get_json()
    solution_name   = p.get('solution_name', '')
    solution_code   = p.get('solution_code', '')
    configure_date  = p.get('configure_date', datetime.date.today().strftime('%Y-%m-%d'))
    validity_date   = p.get('validity_date', '')
    temperature     = str(p.get('temperature', '') or '')
    humidity        = str(p.get('humidity', '') or '')
    storage_cond    = str(p.get('storage_condition', '') or '')
    custom_type     = str(p.get('custom_type', '') or '')
    detail_list     = p.get('detail_list', [])
    is_merged       = p.get('is_merged_export', False)
    top_ancestor    = p.get('top_ancestor', {})
    use_uv          = p.get('use_uv_template', False)

    if use_uv:
        template_name = 'RF10-12 标准溶液配制校准曲线记录.docx'
    else:
        template_name = 'RF10-10 标准溶液配制记录（稀释）(1).docx'
    template_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), 'word_templates', template_name)
    if not os.path.exists(template_path):
        return jsonify({"success": False, "message": f"模板文件不存在: {template_name}"}), 404

    doc = DocxDocument(template_path)

    # local aliases for module-level helpers (backward compat with existing code below)
    _make_run = _docx_make_run
    _make_para = _docx_make_para
    _clear_tc = _docx_clear_tc
    _set_tc_text = _docx_set_tc_text
    _set_tc_two_lines = _docx_set_tc_two_lines
    _set_tc_multi_lines = _docx_set_tc_multi_lines
    _set_vmerge = _docx_set_vmerge

    def _parse_conc_val(raw):
        m = re.match(r'^\s*([\d.]+)', str(raw))
        return m.group(1) if m else str(raw)

    def _extract_conc_from_code(code):
        code = str(code)
        # Handle range-format codes like CK-CG-1HG-[10.00-5.00]-20260518-点1
        m = re.search(r'\[([\d.]+)', code)
        if m:
            return m.group(1)
        parts = code.split('-')
        if len(parts) >= 2:
            candidate = parts[-2]
            if re.match(r'^[\d.]+$', candidate):
                return candidate
        return ''

    if len(doc.paragraphs) > 1:
        runs = doc.paragraphs[1].runs

        def _centered_underline(run, value, width, suffix=''):
            padded = value.center(width)
            run.text = padded + suffix
            run.font.underline = True

        if temperature and len(runs) > 2:
            _centered_underline(runs[2], temperature, 10)
        if humidity and len(runs) > 9:
            _centered_underline(runs[9], humidity, 10, '%')
        if storage_cond and len(runs) > 12:
            _centered_underline(runs[12], storage_cond, 14)

    table = doc.tables[0]

    def _fill_tc1(tc, text):
        el = tc._tc if hasattr(tc, '_tc') else tc
        for p in el.findall(qn('w:p')):
            el.remove(p)
        lines = text.split('\n')
        for line in lines:
            p_el = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'center')
            pPr.append(jc)
            p_el.append(pPr)
            r_el = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            sz_el = OxmlElement('w:sz')
            sz_el.set(qn('w:val'), '18')
            szCs_el = OxmlElement('w:szCs')
            szCs_el.set(qn('w:val'), '18')
            rPr.append(sz_el)
            rPr.append(szCs_el)
            r_el.append(rPr)
            t_el = OxmlElement('w:t')
            t_el.text = line
            if line and (line[0] == ' ' or line[-1] == ' '):
                t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r_el.append(t_el)
            p_el.append(r_el)
            el.append(p_el)

    # Determine top-level source items for header
    if is_merged:
        # Group ancestor items (dilutionIdx < 0) by dilutionIdx level
        dil_groups = {}
        for item in detail_list:
            di = item.get('dilutionIdx', 0)
            if di < 0:
                dil_groups.setdefault(di, []).append(item)
        if dil_groups:
            # "Deepest expanded" algorithm: start from deepest level (most negative),
            # find the first level with > 1 non-skip_body item (expanded source).
            # Use ALL items (including skip_body) for header, but only non-skip for count check.
            sorted_levels = sorted(dil_groups.keys())
            top_items = None
            for level in sorted_levels:
                non_skip = [it for it in dil_groups[level] if not it.get('_skip_body')]
                if len(non_skip) > 1:
                    top_items = dil_groups[level]
                    break
            if top_items is None:
                top_items = dil_groups[max(dil_groups.keys())]
        else:
            top_items = []
    else:
        top_items = [item for item in detail_list if item.get('dilutionIdx', 0) == 0]
    if not top_items and detail_list:
        top_items = [detail_list[0]]

    # Header name
    source_names = []
    for item in top_items:
        nm = (item.get('originalName', '') or '').strip()
        if nm and nm not in source_names:
            source_names.append(nm)
    header_name = '；\n'.join(source_names) + '；' if len(source_names) > 1 else (source_names[0] if source_names else solution_name)

    _fill_tc1(table.rows[1]._tr.findall(qn('w:tc'))[1], header_name)

    # Header code (extract from originalNo: "order\ncode" → last part)
    source_codes = []
    for item in top_items:
        parts = (item.get('originalNo', '') or '').split('\n')
        code = parts[-1].strip() if parts else ''
        if code and code not in source_codes:
            source_codes.append(code)
    header_code = '；\n'.join(source_codes) + '；' if source_codes else ''
    _fill_tc1(table.rows[1]._tr.findall(qn('w:tc'))[3], header_code)

    conc_unit = detail_list[0].get('configurationUnit', '') if detail_list else ''
    qty_unit = detail_list[0].get('receivedUint', 'mL') if detail_list else 'mL'

    r1_tcs = table.rows[1]._tr.findall(qn('w:tc'))
    tc4 = r1_tcs[4]
    tc5 = r1_tcs[5]

    _fill_tc1(tc4, f'浓度\n({conc_unit})' if use_uv else f'浓度({conc_unit})')

    # Header concentration: single source → specific value, multi source → 见下表
    if len(top_items) <= 1:
        cv = _parse_conc_val(top_items[0].get('originalConcentration', '')) if top_items else ''
        _fill_tc1(tc5, cv)
        show_source_names = False
    else:
        _fill_tc1(tc5, '见下表')
        show_source_names = True

    r3_tcs = table.rows[3]._tr.findall(qn('w:tc'))
    if use_uv:
        _set_tc_text(r3_tcs[0], f'母体标液({conc_unit})', center=True)
        _set_tc_text(r3_tcs[1], f'取量({qty_unit})', center=True)
        _set_tc_text(r3_tcs[2], '溶剂', center=True)
        _set_tc_text(r3_tcs[3], '稀释至,ml', center=True)
        _set_tc_text(r3_tcs[4], f'浓度/含量Y({conc_unit})', center=True)
        # TC[5] X1, TC[6] X2 — leave as-is (template has header text)
        _set_tc_text(r3_tcs[7], '编号', center=True)
        _set_tc_text(r3_tcs[8], '有效期', center=True)
    else:
        _set_tc_text(r3_tcs[0], f'母体标液({conc_unit})', center=True)
        _set_tc_text(r3_tcs[1], f'取量({qty_unit})', center=True)
        _set_tc_text(r3_tcs[2], '溶剂', center=True)
        _set_tc_text(r3_tcs[3], '稀释至,ml', center=True)
        _set_tc_text(r3_tcs[4], f'浓度({conc_unit})', center=True)
        _set_tc_text(r3_tcs[5], '编号', center=True)
        _set_tc_text(r3_tcs[6], '配制日期', center=True)
        _set_tc_text(r3_tcs[7], '有效期', center=True)

    # For merged export: separate header items (all) from body items (exclude weighing steps)
    if is_merged:
        body_list = [item for item in detail_list if not item.get('_skip_body')]
    else:
        body_list = detail_list

    # Determine top-level source items for header
    n_items = len(body_list)
    n_tpl    = 12
    remark_row_idx = 16

    # For merged export: override qty_unit from body ancestor rows (exclude weighing steps)
    if is_merged:
        for item in body_list:
            if item.get('dilutionIdx', 0) < 0:
                qty_unit = item.get('receivedUint', 'mL')
                break

    if n_items > n_tpl:
        # Clone source row: for RF10-12 use a data row (row 4), for RF10-10 use row before remark
        clone_src_row_idx = 4 if use_uv else (remark_row_idx - 1)
        for _ in range(n_items - n_tpl):
            src_tr    = table.rows[clone_src_row_idx]._tr
            new_tr    = copy.deepcopy(src_tr)
            remark_tr = table.rows[remark_row_idx]._tr
            table._tbl.insert(list(table._tbl).index(remark_tr), new_tr)

    row_values = []
    for item in body_list:
        item_result_conc = _parse_conc_val(item.get('configurationConcentration', ''))
        if not item_result_conc:
            item_result_conc = _extract_conc_from_code(solution_code)
        item_result_code = item.get('resultCode', '')
        if not item_result_code:
            item_result_code = solution_code
        if not item_result_code:
            original_no = item.get('originalNo', '')
            item_result_code = original_no.split('\n')[-1].strip() if original_no else ''
        row_values.append({
            'name':          item.get('originalName', ''),
            'conc_val':      _parse_conc_val(item.get('originalConcentration', '')),
            'qty':           str(item.get('receivedQuantity', '')),
            'medium':        item.get('medium', ''),
            'volume':        str(item.get('volume', '')),
            'result_conc':   item_result_conc,
            'result_code':   item_result_code,
            'dilutionIdx':   item.get('dilutionIdx', 0),
            'configure_date': item.get('configureDate', configure_date),
            'validity_date':  item.get('validityDate', validity_date),
        })

    is_working = bool(body_list and body_list[0].get('resultCode'))

    # Detect multi-source: count how many rows belong to dilution point 0
    first_dil_rows = sum(1 for rv in row_values if rv['dilutionIdx'] == 0)
    is_multi_source = is_working and first_dil_rows > 1

    # Column mapping depends on template
    if use_uv:
        # RF10-12: TC[0]=母体标液, TC[1]=取量, TC[2]=溶剂, TC[3]=稀释至,
        #          TC[4]=浓度, TC[5]=X1(skip), TC[6]=X2(skip), TC[7]=编号, TC[8]=有效期
        col_medium = 2
        col_volume = 3
        col_conc = 4
        col_code = 7
        col_validity = 8
        n_data_tcs = 9
        uv_validity = '一天'
        follow_code_merge_cols = {col_medium, col_volume, col_validity}
        value_merge_cols = {col_medium, col_volume, col_conc, col_code} if not (is_working and not is_multi_source) else {col_conc, col_code}
    else:
        # RF10-10: TC[0]=母体标液, TC[1]=取量, TC[2]=溶剂, TC[3]=稀释至,
        #          TC[4]=浓度, TC[5]=编号, TC[6]=配制日期, TC[7]=有效期
        col_medium = 2
        col_volume = 3
        col_conc = 4
        col_code = 5
        col_date = 6
        col_validity = 7
        n_data_tcs = 8
        uv_validity = None
        follow_code_merge_cols = {col_medium, col_volume, col_date, col_validity}
        value_merge_cols = {col_medium, col_volume, col_conc, col_code, col_date, col_validity} if not (is_working and not is_multi_source) else {col_conc, col_code}

    # Column values for data rows
    col_vals_map = {
        col_medium: [rv['medium']      for rv in row_values],
        col_volume: [rv['volume']      for rv in row_values],
        col_conc:   [rv['result_conc'] for rv in row_values],
        col_code:   [rv['result_code'] for rv in row_values],
        col_validity: [uv_validity or rv['validity_date'] for rv in row_values],
    }
    if not use_uv:
        col_vals_map[col_date] = [rv['configure_date'] for rv in row_values]

    # All fillable TC indices (excluding X1/X2 in UV mode)
    fill_tc_indices = sorted(set(col_vals_map.keys()))

    for i, item in enumerate(body_list):
        tr  = table.rows[4 + i]._tr
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) < n_data_tcs:
            continue

        # 母体标液
        if is_merged and row_values[i]['dilutionIdx'] < 0:
            if show_source_names:
                _set_tc_two_lines(tcs[0], row_values[i]['name'], row_values[i]['conc_val'])
            else:
                _set_tc_text(tcs[0], row_values[i]['conc_val'])
        elif is_working:
            if is_multi_source and row_values[i]['dilutionIdx'] == 0:
                _set_tc_two_lines(tcs[0], row_values[i]['name'], row_values[i]['conc_val'])
            else:
                _set_tc_text(tcs[0], row_values[i]['conc_val'])
        else:
            if show_source_names:
                _set_tc_two_lines(tcs[0], row_values[i]['name'], row_values[i]['conc_val'])
            else:
                _set_tc_text(tcs[0], row_values[i]['conc_val'])

        # 取量
        _set_tc_text(tcs[1], row_values[i]['qty'])

        # Determine if 编号 would merge with previous row
        same_dilution = i > 0 and row_values[i]['dilutionIdx'] == row_values[i - 1]['dilutionIdx']
        code_same_as_prev = i > 0 and row_values[i]['result_code'] == row_values[i - 1]['result_code']
        code_should_merge = same_dilution and code_same_as_prev

        # Fill mapped columns
        for tc_idx in fill_tc_indices:
            val = col_vals_map[tc_idx][i]
            if tc_idx in follow_code_merge_cols:
                if code_should_merge:
                    _clear_tc(tcs[tc_idx])
                    tcs[tc_idx].append(_make_para('', center=True))
                    _set_vmerge(tcs[tc_idx], 'continue')
                else:
                    _set_tc_text(tcs[tc_idx], val)
                    _set_vmerge(tcs[tc_idx], 'restart')
            elif tc_idx in value_merge_cols:
                if same_dilution and val == col_vals_map[tc_idx][i - 1]:
                    _clear_tc(tcs[tc_idx])
                    tcs[tc_idx].append(_make_para('', center=True))
                    _set_vmerge(tcs[tc_idx], 'continue')
                else:
                    _set_tc_text(tcs[tc_idx], val)
                    _set_vmerge(tcs[tc_idx], 'restart')
            else:
                _set_tc_text(tcs[tc_idx], val)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    from flask import send_file
    from urllib.parse import quote as _urlquote
    dl_name = (f"{solution_code}_配制记录.docx" if solution_code else "配制记录.docx") \
              .replace('/', '-').replace('\\', '-')
    resp = send_file(buf, as_attachment=True, download_name=dl_name,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    resp.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{_urlquote(dl_name)}"
    return resp


# ==================== 修复版：更新 Excel 使用情况（动态检测表头） ====================
@app.route('/api/update_excel_usage', methods=['POST'])
def update_excel_usage():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    
    # 防御性 JSON 解析
    data = request.get_json(silent=True) or {}
    lab_no = data.get('labNo', '').strip()
    new_usage = data.get('usage', '').strip()
    
    # 如果实验室编号为空，直接返回成功（跳过），避免报错
    if not lab_no:
        return jsonify({"success": True, "message": "实验室编号为空，跳过更新"})
    if not new_usage:
        return jsonify({"success": False, "message": "使用情况不能为空"}), 400

    excel_path, _ = get_user_paths(session.get('display_name', ''))
    excel_path = excel_path.strip()
    if not excel_path:
        return jsonify({"success": False, "message": "未配置 Excel 路径"}), 400

    ext = os.path.splitext(excel_path)[1].lower()

    try:
        if ext == '.xlsx':
            wb = openpyxl.load_workbook(excel_path)
            sheet_name = detect_sheet_name(wb)
            if not sheet_name or sheet_name not in wb.sheetnames:
                wb.close()
                return jsonify({"success": True, "message": "工作表不存在，跳过"})
            ws = wb[sheet_name]

            # 动态查找表头行（前 5 行中包含“实验室编号”和“使用情况”的行）
            header_row = None
            lab_col = None
            usage_col = None
            for row_idx in range(1, min(6, ws.max_row + 1)):
                row_cells = list(ws[row_idx])
                for col_idx, cell in enumerate(row_cells, 1):
                    val = str(cell.value) if cell.value is not None else ''
                    if '实验室编号' in val:
                        lab_col = col_idx
                    if '使用情况' in val:
                        usage_col = col_idx
                if lab_col is not None and usage_col is not None:
                    header_row = row_idx
                    break

            if header_row is None:
                wb.close()
                return jsonify({"success": False, "message": "未找到包含‘实验室编号’和‘使用情况’的表头行"}), 400

            # 从表头下一行开始查找目标行
            target_row = None
            for row in ws.iter_rows(min_row=header_row + 1, values_only=False):
                cell_val = row[lab_col - 1].value
                if cell_val and str(cell_val).strip() == lab_no:
                    target_row = row
                    break

            if target_row is None:
                wb.close()
                return jsonify({"success": True, "message": "未找到匹配记录，跳过"})

            current_usage = target_row[usage_col - 1].value
            if current_usage == new_usage:
                wb.close()
                return jsonify({"success": True, "message": "已是目标状态，无需更新"})

            target_row[usage_col - 1].value = new_usage
            wb.save(excel_path)
            wb.close()
            return jsonify({"success": True, "message": "使用情况已更新为开封"})

        elif ext == '.xls':
            rb = xlrd.open_workbook(excel_path, formatting_info=True)
            sheet_name = None
            for name in _KNOWN_SHEET_NAMES:
                if name in rb.sheet_names():
                    sheet_name = name
                    break
            if not sheet_name and rb.sheet_names():
                sheet_name = rb.sheet_names()[0]
            if not sheet_name or sheet_name not in rb.sheet_names():
                return jsonify({"success": True, "message": "工作表不存在，跳过"})
            sheet = rb.sheet_by_name(sheet_name)
            
            # 动态查找表头行（前 5 行）
            header_row_idx = None
            lab_col = None
            usage_col = None
            for row_idx in range(min(5, sheet.nrows)):
                for col_idx in range(sheet.ncols):
                    val = str(sheet.cell_value(row_idx, col_idx)).strip()
                    if '实验室编号' in val:
                        lab_col = col_idx
                    if '使用情况' in val:
                        usage_col = col_idx
                if lab_col is not None and usage_col is not None:
                    header_row_idx = row_idx
                    break

            if header_row_idx is None:
                return jsonify({"success": False, "message": "未找到包含‘实验室编号’和‘使用情况’的表头行"}), 400

            # 查找目标数据行
            target_row = None
            for row_idx in range(header_row_idx + 1, sheet.nrows):
                if str(sheet.cell_value(row_idx, lab_col)).strip() == lab_no:
                    target_row = row_idx
                    break

            if target_row is None:
                return jsonify({"success": True, "message": "未找到匹配记录，跳过"})

            current_usage = sheet.cell_value(target_row, usage_col)
            if current_usage == new_usage:
                return jsonify({"success": True, "message": "已是目标状态，无需更新"})

            wb = xl_copy(rb)
            ws = wb.get_sheet(sheet_name)
            ws.write(target_row, usage_col, new_usage)
            wb.save(excel_path)
            return jsonify({"success": True, "message": "使用情况已更新为开封"})

        else:
            return jsonify({"success": False, "message": "不支持的 Excel 格式"}), 400

    except Exception as e:
        return jsonify({"success": False, "message": f"更新失败: {str(e)}"}), 500


@app.route('/api/update_excel_record', methods=['POST'])
def update_excel_record():
    """编辑标准品：按实验室编号原地更新该行全部字段（不新增）。"""
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    try:
        data = request.get_json(silent=True) or {}
        record = data.get('record') or {}
        lab_no = str(record.get('labNo', '')).strip()
        if not lab_no:
            return jsonify({"success": False, "message": "缺少实验室编号"}), 400

        # 字段 -> 表头（精确匹配定位列，避免“编号”误匹配“实验室编号”）
        field_headers = [
            ('originalId', '编号'), ('group', '组别'), ('labNo', '实验室编号'),
            ('name', '标品名称'), ('cas', 'CAS号'), ('spec', '规格/浓度'),
            ('manufacturer', '生产商'), ('expiry', '有效期'), ('location', '存放地点'),
            ('storageDate', '入库日期'), ('usage', '使用情况'), ('remarks', '备注'),
        ]

        excel_path, _ = get_user_paths(session.get('display_name', ''))
        excel_path = excel_path.strip()
        if not excel_path:
            return jsonify({"success": False, "message": "未配置 Excel 路径"}), 400
        ext = os.path.splitext(excel_path)[1].lower()

        if ext == '.xlsx':
            wb = openpyxl.load_workbook(excel_path)
            sheet_name = detect_sheet_name(wb)
            if not sheet_name or sheet_name not in wb.sheetnames:
                wb.close()
                return jsonify({"success": False, "message": "工作表不存在"}), 400
            ws = wb[sheet_name]

            # 动态查找表头行 + 列映射（前 5 行中含“实验室编号”的行）
            header_row = None
            col_map = {}
            for row_idx in range(1, min(6, ws.max_row + 1)):
                cm = {}
                for col_idx, cell in enumerate(ws[row_idx], 1):
                    val = str(cell.value).strip() if cell.value is not None else ''
                    for field, hdr in field_headers:
                        if val == hdr and field not in cm:
                            cm[field] = col_idx
                if 'labNo' in cm:
                    col_map, header_row = cm, row_idx
                    break
            if header_row is None:
                wb.close()
                return jsonify({"success": False, "message": "未找到含‘实验室编号’的表头行"}), 400

            lab_col = col_map['labNo']
            target_row = None
            for row in ws.iter_rows(min_row=header_row + 1, values_only=False):
                cv = row[lab_col - 1].value
                if cv and str(cv).strip() == lab_no:
                    target_row = row
                    break
            if target_row is None:
                wb.close()
                return jsonify({"success": True, "message": f"Excel 中未找到实验室编号 {lab_no}，跳过"})

            for field, hdr in field_headers:
                if field in col_map:
                    target_row[col_map[field] - 1].value = record.get(field, '')
            wb.save(excel_path)
            wb.close()
            return jsonify({"success": True, "message": "已更新到 Excel"})

        elif ext == '.xls':
            rb = xlrd.open_workbook(excel_path, formatting_info=True)
            sheet_name = None
            for name in _KNOWN_SHEET_NAMES:
                if name in rb.sheet_names():
                    sheet_name = name
                    break
            if not sheet_name and rb.sheet_names():
                sheet_name = rb.sheet_names()[0]
            if not sheet_name or sheet_name not in rb.sheet_names():
                return jsonify({"success": False, "message": "工作表不存在"}), 400
            sheet = rb.sheet_by_name(sheet_name)

            header_row_idx = None
            col_map = {}
            for row_idx in range(min(5, sheet.nrows)):
                cm = {}
                for col_idx in range(sheet.ncols):
                    val = str(sheet.cell_value(row_idx, col_idx)).strip()
                    for field, hdr in field_headers:
                        if val == hdr and field not in cm:
                            cm[field] = col_idx
                if 'labNo' in cm:
                    col_map, header_row_idx = cm, row_idx
                    break
            if header_row_idx is None:
                return jsonify({"success": False, "message": "未找到含‘实验室编号’的表头行"}), 400

            lab_col = col_map['labNo']
            target_row = None
            for row_idx in range(header_row_idx + 1, sheet.nrows):
                if str(sheet.cell_value(row_idx, lab_col)).strip() == lab_no:
                    target_row = row_idx
                    break
            if target_row is None:
                return jsonify({"success": True, "message": f"Excel 中未找到实验室编号 {lab_no}，跳过"})

            wb = xl_copy(rb)
            ws = wb.get_sheet(sheet_name)
            for field, hdr in field_headers:
                if field in col_map:
                    ws.write(target_row, col_map[field], record.get(field, ''))
            wb.save(excel_path)
            return jsonify({"success": True, "message": "已更新到 Excel"})

        else:
            return jsonify({"success": False, "message": "仅支持 .xls 或 .xlsx"}), 400

    except Exception as e:
        return jsonify({"success": False, "message": f"更新 Excel 失败: {_excel_save_error_detail(e)}"}), 500



# ── 标液核查端点 ──

def _extract_d_concentration_points(record):
    """从 D 型工作液记录的 concentration 字段提取各浓度点。

    concentration 格式示例:
      单组分: '10.40mg/L,5.20mg/L,1.04mg/L,0.52mg/L,0.10mg/L,0.052mg/L'
      多组分: '甲苯:10.40(mg/L),5.20(mg/L),...;乙酸丁酯:...'
    返回 [{index, concentration, unit, name, code, conc_str}]，conc_str 保留原始精度字符串。
    """
    conc_str = str(record.get('concentration') or '').strip()
    if not conc_str:
        return []

    points = []
    # 多组分格式: '甲苯:10.40(mg/L),5.20(mg/L);乙酸丁酯:...'
    if ':' in conc_str and ';' in conc_str:
        groups = conc_str.split(';')
        all_names = []
        all_conc_maps = {}
        for gi, group in enumerate(groups):
            group = group.strip()
            if not group:
                continue
            ci = group.rfind(':')
            if ci < 0:
                continue
            name = group[:ci].strip()
            all_names.append(name)
            cstr = group[ci+1:].strip()
            conc_parts = re.findall(r'([\d.]+)\s*\(([^)]+)\)', cstr)
            if not conc_parts:
                conc_parts = re.findall(r'([\d.]+)\s*(mg/L|μg/mL|ng/ml|ppm)', cstr)
            for pi, (cv, cu) in enumerate(conc_parts):
                key = (float(cv), cu)
                if key not in all_conc_maps:
                    all_conc_maps[key] = {'cv': cv, 'cu': cu, 'names': []}
                all_conc_maps[key]['names'].append(name)
        # 按浓度值排序，生成去重点位
        sorted_keys = sorted(all_conc_maps.keys(), key=lambda x: x[0])
        for pi, key in enumerate(sorted_keys):
            m = all_conc_maps[key]
            points.append({
                'index': pi,
                'concentration': key[0],
                'unit': m['cu'],
                'name': ','.join(m['names']),
                'source_name': ','.join(m['names']),
                'conc_str': m['cv'],
            })
    else:
        # 单组分格式: '10.40mg/L,5.20mg/L,...' 或 '10.40(mg/L),5.20(mg/L),...'
        name = str(record.get('solutionName') or '').strip()
        conc_parts = re.findall(r'([\d.]+)\s*\(([^)]+)\)', conc_str)
        if not conc_parts:
            conc_parts = re.findall(r'([\d.]+)\s*(mg/L|μg/mL|ng/ml|ppm)', conc_str)
        for pi, (cv, cu) in enumerate(conc_parts):
            points.append({
                'index': pi,
                'concentration': float(cv),
                'unit': cu,
                'name': name,
                'source_name': name,
                'conc_str': cv,
            })

    return points


def _format_conc_value(value):
    """浓度修约: 浓度<0.10保留至小数点后3位, 浓度≥0.10保留至小数点后2位"""
    if value < 0.10:
        return f"{value:.3f}"
    return f"{value:.2f}"


def _build_concentration_point_code(record, point):
    """构建浓度点编号: CK-CG-{num}-{concentration}-{date}"""
    code = str(record.get('solutionCode', '')).strip()
    date = str(record.get('configureDate', ''))[:10].replace('-', '')
    conc_str = point.get('conc_str') or _format_conc_value(point['concentration'])
    mx_match = re.match(r'.*-\d{8}-(.+)$', code)
    mx = mx_match.group(1) if mx_match else ''
    parts = code.split('-')
    if len(parts) >= 3:
        base = '-'.join(parts[:3])
        return f"{base}-{conc_str}-{date}" + (f"-{mx}" if mx else "")
    return f"{code}-{conc_str}-{date}" + (f"-{mx}" if mx else "")


@app.route('/api/lims/get_verification_info', methods=['POST'])
def lims_get_verification_info():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    p = request.get_json() or {}
    solution_ids = p.get('solution_ids', [])
    configure_orders = p.get('configure_orders', [])

    if len(solution_ids) < 2 and len(configure_orders) < 2:
        return jsonify({"success": False, "message": "请提供两个工作液的ID或编号"})

    try:
        system = get_system()
        username = session.get('username', '')
        if system.current_user != username:
            system.current_user = username
            system.load_session()

        records = []
        for idx in range(2):
            sid = solution_ids[idx] if idx < len(solution_ids) else None
            order = configure_orders[idx] if idx < len(configure_orders) else None
            lims_id = sid
            if not lims_id and order:
                lims_id = _resolve_order_to_lims_id(system, order)
            if not lims_id:
                return jsonify({"success": False, "message": f"未找到记录: {order or sid}"}), 404
            rec = _fetch_solution_view(system, lims_id, order_str=order)
            if not rec or not rec.get('id'):
                return jsonify({"success": False, "message": f"获取详情失败: {order or sid}"}), 404
            records.append(rec)

        # 判断新旧
        d1 = str(records[0].get('configureDate', ''))[:10]
        d2 = str(records[1].get('configureDate', ''))[:10]
        if d1 >= d2:
            new_rec, old_rec = records[0], records[1]
        else:
            new_rec, old_rec = records[1], records[0]

        def _build_solution_info(rec):
            detail_list = rec.get('detailList') or []
            conc_str = str(rec.get('concentration', '')).strip()

            # 解析concentration字段: 物质名 -> [conc0, conc1, ...] (降序)
            substance_concs = {}
            substance_concs_raw = {}  # 保留原始字符串精度 (e.g. "0.60" not 0.6)
            all_unit = ''
            if ':' in conc_str and ';' in conc_str:
                for group in conc_str.split(';'):
                    group = group.strip()
                    if not group: continue
                    ci = group.rfind(':')
                    if ci < 0: continue
                    sname = group[:ci].strip()
                    cstr = group[ci+1:].strip()
                    cparts = re.findall(r'([\d.]+)\s*\(([^)]+)\)', cstr)
                    if not cparts:
                        cparts = re.findall(r'([\d.]+)\s*(mg/L|μg/mL|ug/mL|ng/ml|ppm)', cstr)
                    if cparts and not all_unit: all_unit = cparts[0][1]
                    substance_concs[sname] = [float(cv) for cv, cu in cparts]
                    substance_concs_raw[sname] = [cv for cv, cu in cparts]
            elif conc_str:
                sname = str(rec.get('solutionName', '')).strip()
                cparts = re.findall(r'([\d.]+)\s*\(([^)]+)\)', conc_str)
                if not cparts:
                    cparts = re.findall(r'([\d.]+)\s*(mg/L|μg/mL|ug/mL|ng/ml|ppm)', conc_str)
                if cparts:
                    all_unit = cparts[0][1]
                    substance_concs[sname] = [float(cv) for cv, cu in cparts]
                    substance_concs_raw[sname] = [cv for cv, cu in cparts]
            all_sub_names = list(substance_concs.keys())

            # 从detailList提取resultCode
            code_map = {}
            for dl in detail_list:
                rc = str(dl.get('resultCode', '')).strip()
                if not rc:
                    orig_no = str(dl.get('originalNo', '')).strip()
                    rc = orig_no.split('\n')[-1].strip() if orig_no else ''
                if not rc:
                    rc = str(rec.get('solutionCode', '')).strip()
                name = str(dl.get('originalName', '')).strip()
                if rc not in code_map:
                    code_map[rc] = {'names': [], 'dilutionIdx': dl.get('dilutionIdx', 0)}
                if name:
                    code_map[rc]['names'].append(name)

            points = []
            unique_rcs = list(code_map.keys())
            is_range_format = len(unique_rcs) == 1 and '[' in unique_rcs[0]

            if is_range_format:
                first_concs = substance_concs.get(all_sub_names[0], []) if all_sub_names else []
                for pi, cv in enumerate(first_concs):
                    sub_concs = {sn: substance_concs[sn][pi] for sn in all_sub_names if pi < len(substance_concs[sn])}
                    sub_concs_str = {sn: _format_conc_value(substance_concs[sn][pi]) for sn in all_sub_names if pi < len(substance_concs.get(sn, []))}
                    pt = {'index': pi, 'concentration': cv, 'unit': all_unit, 'conc_str': _format_conc_value(cv),
                          'name': ','.join(all_sub_names), 'sub_concs': sub_concs, 'sub_concs_str': sub_concs_str}
                    pt['code'] = _build_concentration_point_code(rec, pt)
                    points.append(pt)
            else:
                sorted_codes = sorted(code_map.items(), key=lambda x: x[1].get('dilutionIdx', 0))
                for pi, (rc, info) in enumerate(sorted_codes):
                    sub_concs = {sn: substance_concs[sn][pi] for sn in all_sub_names if pi < len(substance_concs.get(sn, []))}
                    sub_concs_str = {sn: _format_conc_value(substance_concs[sn][pi]) for sn in all_sub_names if pi < len(substance_concs.get(sn, []))}
                    conc_from_code = ''
                    parts = rc.split('-')
                    if len(parts) >= 2 and re.match(r'^[\d.]+$', parts[-2]):
                        conc_from_code = parts[-2]
                    elif re.search(r'\[([\d.]+)', rc):
                        conc_from_code = re.search(r'\[([\d.]+)', rc).group(1)
                    conc_val = float(conc_from_code) if conc_from_code else 0
                    points.append({
                        'code': rc, 'concentration': conc_val, 'unit': all_unit,
                        'name': ','.join(info['names']), 'sub_concs': sub_concs,
                        'sub_concs_str': sub_concs_str,
                        'conc_str': conc_from_code or '',
                    })
            point_codes = [pt['code'] for pt in points]

            # 解析直接父级(储备液/应用液)的完整自编号，供"新旧储备液比对"核查方法显示
            parent_codes = []
            for src_order in _extract_direct_source_codes(rec.get('originalCode')):
                try:
                    src_id = _resolve_order_to_lims_id(system, src_order)
                    if not src_id:
                        continue
                    prec = _fetch_solution_view(system, src_id, order_str=src_order)
                    if prec and prec.get('solutionCode'):
                        parent_codes.append(str(prec['solutionCode']).strip())
                except Exception:
                    pass

            return {
                'id': rec.get('id'),
                'configureOrder': rec.get('configureOrder', ''),
                'solutionCode': rec.get('solutionCode', ''),
                'solutionName': rec.get('solutionName', ''),
                'configureDate': str(rec.get('configureDate', ''))[:10],
                'concentrationPoints': points,
                'validityDate': str(rec.get('validityDate', ''))[:10],
                'customType': rec.get('customType', ''),
                'parentSolutionCode': '\n'.join(parent_codes),
            }

        new_info = _build_solution_info(new_rec)
        old_info = _build_solution_info(old_rec)

        # 溯源到A获取标准物质信息
        def _trace_to_a(rec):
            """溯源到A级祖先获取标准物质信息"""
            rec_id = rec.get('id')
            rec_order = rec.get('configureOrder', '')
            if not rec_id:
                return []
            try:
                ancestors = []
                seen = set()
                matched, top = _trace_export_chain(
                    system, [(rec_id, rec_order)],
                    '',  # 不限制日期，溯源需追溯到任意日期的A
                    '',  # 不限制配置人，溯源需跨配置人追溯
                )
                # 从matched中找到父级为A型的B/C级记录，从其detailList提取A级标准物质信息
                for r in matched:
                    if r.get('level') == 'D':
                        continue
                    original_code = str(r.get('original_code') or '').strip()
                    if not original_code.startswith('A-'):
                        continue
                    a_name = r.get('parent_name', '')
                    # 从detailList的originalNo提取A级的controlledNo
                    a_controlled_no = ''
                    a_concentration = ''
                    detail_list = r.get('_detail_list') or []
                    for dl in detail_list:
                        dl_name = str(dl.get('originalName', '')).strip()
                        dl_no = str(dl.get('originalNo', '')).strip()
                        lines = [l.strip() for l in dl_no.split('\n') if l.strip()]
                        potential_no = lines[-1] if lines else ''
                        # 优先精确匹配名称
                        if dl_name and a_name and dl_name == a_name:
                            a_controlled_no = potential_no
                            a_concentration = str(dl.get('originalConcentration', '')).strip()
                            break
                        # 精确匹配失败时，取第一个有CK-CG编号的行作为候选
                        if not a_controlled_no and potential_no:
                            a_controlled_no = potential_no
                            a_concentration = str(dl.get('originalConcentration', '')).strip()
                    if not a_controlled_no:
                        pass
                    a_storage_condition = str(r.get('storage_condition') or '').strip()
                    if a_controlled_no and a_controlled_no not in seen:
                        seen.add(a_controlled_no)
                        ancestors.append({
                            'level': 'A',
                            'configureOrder': a_controlled_no,
                            'solutionName': a_name,
                            'controlledNo': a_controlled_no,
                            'storageCondition': a_storage_condition,
                            'concentration': a_concentration,
                        })
                    elif a_name and a_name not in seen:
                        seen.add(a_name)
                        ancestors.append({
                            'level': 'A',
                            'configureOrder': a_name,
                            'solutionName': a_name,
                            'controlledNo': a_controlled_no,
                            'storageCondition': a_storage_condition,
                            'concentration': a_concentration,
                        })
                return ancestors
            except Exception as e:
                pass
                return []

        new_info['ancestors'] = _trace_to_a(new_rec)
        old_info['ancestors'] = _trace_to_a(old_rec)

        return jsonify({
            "success": True,
            "new_solution": new_info,
            "old_solution": old_info,
            "mode": "new_verifies_old",
        })

    except Exception as e:
        return jsonify({"success": False, "message": f"获取核查信息异常: {str(e)}"}), 500


@app.route('/api/lims/storage_conditions', methods=['POST'])
def lims_storage_conditions():
    """根据controlledNo列表批量查询标准物质保存条件"""
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    data = request.get_json() or {}
    keywords = data.get('keywords', [])
    if not keywords:
        return jsonify({"success": True, "conditions": {}})
    system = get_system()
    username = session.get('username', '')
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    pid = session.get('pid', '')
    conditions = {}
    for kw in keywords:
        kw = str(kw).strip()
        if not kw or kw in conditions:
            continue
        try:
            cb_params = {
                "_search": "false", "nd": str(int(time.time()*1000)),
                "pageSize": 30, "pageNo": 1, "sidx": "", "sord": "asc",
                "type": "CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE",
                "receiveUserName": "", "receiveStartDate": "", "receiveEndDate": "",
                "confirmUserName": "", "confirmStartDate": "", "confirmEndDate": "",
                "invoiceNo": "", "groupId": "",
                "casNo": "", "orgName": "", "state": "normal",
                "keyword": kw,
                "pid": pid, "pname": username, "loginId": pid,
            }
            for _status in ("normal", "history", "overdue"):
                cb_params["status"] = _status
                cb_resp = system.session.get(f"{system.base_url}/detectionManager/manager/consumableBill/pageObj", params=cb_params)
                if cb_resp.status_code != 200:
                    continue
                cb_data = cb_resp.json()
                vo_list = (cb_data.get('resultData') or {}).get('voList') or []
                if vo_list:
                    conditions[kw] = str(vo_list[0].get('storageCondition') or '').strip()
                    break
        except Exception as e:
            pass
    return jsonify({"success": True, "conditions": conditions})


@app.route('/api/lims/parse_verification_pdf', methods=['POST'])
def lims_parse_verification_pdf():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    import tempfile
    from report_parser import parse_pdf_report

    files = request.files.getlist('pdf_files')
    if not files:
        return jsonify({"success": False, "message": "请上传PDF文件"})

    all_compounds = {}  # name -> {value, unit, source_file}

    for f in files:
        if not f.filename or not f.filename.lower().endswith('.pdf'):
            continue
        tmp = None
        try:
            tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
            f.save(tmp.name)
            tmp.close()
            result = parse_pdf_report(tmp.name)
            for name, (val, unit) in result.items():
                all_compounds[name] = {
                    'name': name,
                    'measured_value': val,
                    'unit': unit,
                    'source_file': f.filename,
                }
        except Exception as e:
            pass
        finally:
            if tmp:
                try:
                    os.unlink(tmp.name)
                except OSError:
                    pass

    compounds = list(all_compounds.values())
    return jsonify({"success": True, "compounds": compounds})


@app.route('/api/lims/parse_epatemp_content', methods=['POST'])
def lims_parse_epatemp_content():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    from report_parser import parse_epatemp_txt
    import tempfile

    p = request.get_json() or {}
    d_folders = p.get('d_folders', [])
    if not d_folders:
        return jsonify({"success": False, "message": "请选择.D文件夹"})

    all_compounds = []
    skipped = []
    for folder in d_folders:
        name = folder.get('name', '')
        content_b64 = folder.get('content_b64', '')
        if not content_b64:
            skipped.append(name)
            continue
        tmp = None
        try:
            import base64 as _base64
            raw_bytes = _base64.b64decode(content_b64)
            tmp = tempfile.NamedTemporaryFile(suffix='.txt', delete=False, mode='wb')
            tmp.write(raw_bytes)
            tmp.close()
            parsed = parse_epatemp_txt(tmp.name)
            for compound_name, vals in parsed.items():
                measured_raw = vals[1] if len(vals) > 2 else str(vals[0])
                all_compounds.append({
                    'name': compound_name,
                    'measured_value': vals[0],
                    'measured_raw': measured_raw,
                    'unit': vals[-1],
                    'source_file': name,
                })
        except Exception as e:
            pass
            skipped.append(name)
        finally:
            if tmp:
                try: os.unlink(tmp.name)
                except OSError: pass

    msg = '解析成功，提取到 ' + str(len(all_compounds)) + ' 个化合物'
    if skipped:
        msg += '（以下文件夹无epatemp.txt已跳过: ' + ', '.join(skipped) + '）'
    return jsonify({"success": True, "compounds": all_compounds, "message": msg})


@app.route('/api/lims/export_verification_docx', methods=['POST'])
def lims_export_verification_docx():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    from docx import Document as DocxDocument
    import io

    p = request.get_json() or {}
    template_name = 'RF11-04 标准物质期间核查记录 .docx'
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'word_templates', template_name)
    if not os.path.exists(template_path):
        return jsonify({"success": False, "message": f"模板不存在: {template_name}"}), 404

    try:
        doc = DocxDocument(template_path)
        t0 = doc.tables[0]  # 19行 x 9列
        t1 = doc.tables[1]  # 23行 x 4列

        # ── Table0: 标准物质信息 ──
        table0_data = p.get('table0_data', [])
        verify_date_str = str(p.get('核查时间', ''))[:10]
        for ri, item in enumerate(table0_data):
            row_idx = ri + 2  # 从第3行开始（0/1是表头）
            if row_idx >= len(t0.rows):
                break
            row = t0.rows[row_idx]
            # 序号、标准物质名称、标准物质编号 - 居中对齐+垂直居中
            _docx_set_tc_text(row.cells[0], str(item.get('序号', ri + 1)), center=True, sz=18)
            _docx_set_v_align_center(row.cells[0])
            _docx_set_tc_text(row.cells[1], str(item.get('标准物质名称', '')), center=True, sz=18)
            _docx_set_v_align_center(row.cells[1])
            _docx_set_tc_text(row.cells[2], str(item.get('标准物质编号', '')), center=True, sz=18)
            _docx_set_v_align_center(row.cells[2])
            row.cells[3].text = str(item.get('保存条件', ''))
            # Col4: 是否在有效期 - 根据有效期情况勾选
            item_validity = str(item.get('有效期', '')).strip()[:10]
            if item_validity and verify_date_str:
                try:
                    from datetime import datetime as _dt
                    is_valid = _dt.strptime(item_validity, '%Y-%m-%d') >= _dt.strptime(verify_date_str, '%Y-%m-%d')
                except ValueError:
                    is_valid = True
            else:
                is_valid = True
            _docx_set_tc_checkbox(row.cells[4], '☑是\n□否' if is_valid else '□是\n☑否')
            _docx_set_v_align_center(row.cells[4])
            # Col5: 标志是否齐全 - 默认勾选是
            _docx_set_tc_checkbox(row.cells[5], '☑是\n□否')
            _docx_set_v_align_center(row.cells[5])
            # Col6: 容器是否损伤 - 默认勾选否
            _docx_set_tc_checkbox(row.cells[6], '□是\n☑否')
            _docx_set_v_align_center(row.cells[6])

        # ── Table1: 核查结果 ──
        # R0: 被核查对象编号 - 标题和内容同一行，用；分隔
        target_code = str(p.get('被核查对象编号', ''))
        target_values = [l for l in target_code.replace(', ', '\n').split('\n') if l.strip()]
        _docx_set_tc_labelled_inline(t1.rows[0].cells[0], '被核查对象编号：', target_values, sz=18)

        # R1: 核查对象编号 - 标题和内容同一行，用；分隔
        source_code = str(p.get('核查对象编号', ''))
        source_values = [l for l in source_code.replace(', ', '\n').split('\n') if l.strip()]
        _docx_set_tc_labelled_inline(t1.rows[1].cells[0], '核查对象编号：', source_values, sz=18)

        # R2: 核查方法 + 核查时间
        _docx_set_tc_text(t1.rows[2].cells[0], '核查方法：' + str(p.get('核查方法', '新旧标准曲线比对')), center=False, sz=18)
        _docx_set_tc_text(t1.rows[2].cells[2], '核查时间：' + str(p.get('核查时间', '')), center=False, sz=18)

        # R3: 核查方法描述 - 标题和内容同一行，垂直居中
        method_desc = str(p.get('核查方法描述', ''))
        _docx_set_tc_labelled_inline(t1.rows[3].cells[0], '核查方法描述：', [method_desc] if method_desc else [], sz=18)
        _docx_set_v_align_center(t1.rows[3].cells[0])

        # R5+: 组分数据 (行数不够时自动增加)
        table1_data = p.get('table1_data', [])
        data_start = 5   # R5 开始填数据
        # t1 尾部3行为 结论/备注/签名，数据行不能占用
        data_end = len(t1.rows) - 3  # 留出结论、备注、签名3行
        max_data_rows = data_end - data_start

        if len(table1_data) > max_data_rows:
            # 需要增加行: 在结论行前插入新行
            from copy import deepcopy
            # 以数据区第一行(R5)为模板克隆
            template_tr = t1.rows[data_start]._tr
            conclusion_tr = t1.rows[-3]._tr   # 结论行（倒数第3行）
            tbl_el = template_tr.getparent()
            extra_needed = len(table1_data) - max_data_rows
            for _ in range(extra_needed):
                new_row = deepcopy(template_tr)
                # 清空新行所有单元格
                for tc_el in new_row.findall(_docx_qn('w:tc')):
                    for p_el in tc_el.findall(_docx_qn('w:p')):
                        tc_el.remove(p_el)
                    tc_el.append(_docx_OxmlElement('w:p'))
                tbl_el.insert(list(tbl_el).index(conclusion_tr), new_row)
            # 重新获取t1引用（行数已变化）
            t1 = doc.tables[1]

        for ri, item in enumerate(table1_data):
            row_idx = ri + data_start
            if row_idx >= len(t1.rows) - 3:
                break
            row = t1.rows[row_idx]
            # 组分数据居中对齐
            _docx_set_tc_text(row.cells[0], str(item.get('各组分名称', '')), center=True, sz=18)
            _docx_set_tc_text(row.cells[1], str(item.get('理论值', '')), center=True, sz=18)
            _docx_set_tc_text(row.cells[2], str(item.get('实测值', '')), center=True, sz=18)
            _docx_set_tc_text(row.cells[3], str(item.get('相对偏差', '')), center=True, sz=18)
            # 为每个数据行设置不跨页
            _docx_set_row_cant_split(row)

        # 动态调整行高策略：根据数据行数决定是压缩到一页还是舒适地分成两页
        # 统一应用于：R0-R3信息行 + 数据行 + 结论行 + 备注行（全部同一个行高）
        num_data_rows = len(table1_data)
        if num_data_rows <= 10:
            # 10行以内：尝试压缩行高，争取放入一页（含结论+备注）
            row_height = 320  # 16磅 = 320 twips
        elif num_data_rows <= 14:
            # 11-14行：临界区，使用中等行高
            row_height = 360  # 18磅 = 360 twips
        else:
            # 15行以上：明确分两页，使用舒适行高
            row_height = 400  # 20磅 = 400 twips

        # 收集所有需要统一设置行高的行索引：
        # R0-R3 信息行 + 数据行 + 结论行 + 备注行 + 签名行（全部同一个行高）
        height_row_indices = set()
        for row_idx in range(min(4, len(t1.rows))):   # R0-R3 信息行
            height_row_indices.add(row_idx)
        for ri in range(num_data_rows):                 # 数据行
            row_idx = ri + data_start
            if row_idx < len(t1.rows) - 3:
                height_row_indices.add(row_idx)
        if len(t1.rows) >= 3:
            height_row_indices.add(len(t1.rows) - 3)    # 结论行
            height_row_indices.add(len(t1.rows) - 2)    # 备注行
            # 签名行(-1) 不纳入动态行高，保留模板原有行高

        # 统一应用同一个动态行高
        for row_idx in sorted(height_row_indices):
            _docx_set_row_height(t1.rows[row_idx], row_height)

        # 为最后3个数据行设置 keepNext，强制与结论行保持在同一页
        # 这样确保结论行不会单独成为新页第一行
        last_data_rows = min(3, len(table1_data))  # 至少3行，如果数据少于3行则全部设置
        for i in range(last_data_rows):
            row_idx = data_start + len(table1_data) - last_data_rows + i
            if row_idx < len(t1.rows) - 3:
                _docx_set_row_keep_next(t1.rows[row_idx])

        # 核查结论行(-3)：不跨页，与备注行保持在一起
        _docx_set_row_cant_split(t1.rows[-3])
        _docx_set_row_keep_next(t1.rows[-3])
        # 备注行(-2)：不跨页，与签名行保持在一起
        _docx_set_row_cant_split(t1.rows[-2])
        _docx_set_row_keep_next(t1.rows[-2])
        # 签名行(-1)：不跨页（签名栏已在表格内，全在表格内可靠控制分页）
        _docx_set_row_cant_split(t1.rows[-1])
        # 签名行单元格底端对齐
        for cell in t1.rows[-1].cells:
            _docx_set_v_align_bottom(cell)

        # 核查结论 - 宋体五号
        conclusion = str(p.get('核查结论', '合格'))
        _docx_set_tc_checkbox(t1.rows[-3].cells[0], "核查结论：☑合格；□不合格" if conclusion == '合格' else "核查结论：□合格；☑不合格")
        _docx_set_v_align_center(t1.rows[-3].cells[0])

        # 备注
        t1.rows[-2].cells[0].text = '备注：' + str(p.get('备注', ''))

        new_code = str(p.get('new_solution_code', '')).strip()
        epatemp_contents = p.get('epatemp_contents', [])

        if epatemp_contents:
            # 有 .D 文件夹时，将 docx 页面强制设为 A4 以匹配后续 PDF
            from docx.shared import Cm
            for section in doc.sections:
                section.page_width = Cm(21.0)
                section.page_height = Cm(29.7)

        if not epatemp_contents:
            # 无 .D 文件夹 → 保持原有 docx 导出
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            download_name = (new_code + '期间核查.docx') if new_code else '标准物质期间核查记录.docx'
            from flask import send_file
            from urllib.parse import quote as _urlquote
            resp = send_file(buf, as_attachment=True, download_name=download_name,
                             mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            resp.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{_urlquote(download_name)}"
            return resp
        else:
            # 有 .D 文件夹 → docx 转 PDF + epatemp PDF 合并
            import base64 as _b64

            # 检测 win32com 是否可用
            _has_win32com = False
            try:
                import win32com.client  # noqa: F401
                _has_win32com = True
            except ImportError:
                pass

            if not _has_win32com:
                # 无 win32com：降级为 docx 导出（不合并 epatemp PDF）
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                download_name = (new_code + '期间核查.docx') if new_code else '标准物质期间核查记录.docx'
                from flask import send_file
                from urllib.parse import quote as _urlquote
                resp = send_file(buf, as_attachment=True, download_name=download_name,
                                 mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                resp.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{_urlquote(download_name)}"
                return resp

            # 1. docx → PDF（Word COM）
            docx_buf = io.BytesIO()
            doc.save(docx_buf)
            docx_buf.seek(0)
            docx_pdf_bytes = _docx_to_pdf_bytes(docx_buf.read())

            # 2. 每个 epatemp → PDF（ReportLab）
            epatemp_pdf_pages = [docx_pdf_bytes]
            for item in epatemp_contents:
                content_b64 = item.get('content_b64', '')
                if not content_b64:
                    continue
                raw_bytes = _b64.b64decode(content_b64)
                text, _ = _detect_encoding(raw_bytes)
                epatemp_pdf_pages.append(_epatemp_txt_to_pdf_bytes(text))

            # 3. 合并所有 PDF
            merged_bytes = _merge_pdfs(epatemp_pdf_pages)

            # 4. 返回合并后的 PDF
            download_name = (new_code + '期间核查.pdf') if new_code else '标准物质期间核查记录.pdf'
            from flask import send_file
            from urllib.parse import quote as _urlquote
            resp = send_file(BytesIO(merged_bytes), as_attachment=True, download_name=download_name,
                             mimetype='application/pdf')
            resp.headers['Content-Disposition'] = f"attachment; filename*=UTF-8''{_urlquote(download_name)}"
            return resp

    except Exception as e:
        return jsonify({"success": False, "message": f"导出失败: {str(e)}"}), 500


@app.route('/api/lims/print_label', methods=['POST'])
def lims_print_label():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    from PIL import ImageDraw, ImageFont
    try:
        p = request.get_json() or {}
        solution_code = p.get('solution_code', '')
        solution_name = p.get('solution_name', '')
        concentration = p.get('concentration', '')
        configure_date = p.get('configure_date', '')
        validity_date = p.get('validity_date', '')
        storage_condition = p.get('storage_condition', '')
        medium = p.get('medium', '')
        creator_name = p.get('creator_name', '')

        # 字体（Windows 系统字体）
        system_font_dirs = ['C:/Windows/Fonts', 'C:\\Windows\\Fonts']
        def load_font(size, names):
            for d in system_font_dirs:
                for name in names:
                    fp = os.path.join(d, name)
                    if os.path.exists(fp):
                        try:
                            return ImageFont.truetype(fp, size)
                        except:
                            pass
            return ImageFont.load_default()

        # 标签尺寸 40mm×60mm @300DPI
        # 预览：709×472（横版可读），打印：472×709（竖版旋转）
        PW, PH = 709, 472

        font_title = load_font(32, ['msyh.ttc', 'msyhbd.ttc', 'simhei.ttf'])
        font_value = load_font(32, ['simsun.ttc', 'simsun.ttf'])
        asc_t, _ = font_title.getmetrics()
        asc_v, _ = font_value.getmetrics()
        value_y_offset = asc_t - asc_v

        img_preview = Image.new('RGB', (PW, PH), 'white')
        draw = ImageDraw.Draw(img_preview)

        pad = 20
        max_text_w = PW - pad - 10

        label_items = [
            ('标样编号', solution_code),
            ('标样名称', solution_name),
            ('介质/浓度', f'{medium}/{concentration}' if medium else concentration),
            ('配置人/日期', f'{creator_name}/{configure_date}'),
            ('有效期至', validity_date),
            ('储存地点/温度', storage_condition),
        ]

        # 自动换行（内容部分可能超宽）
        all_lines = []
        for title, value in label_items:
            label_prefix = title + '： '
            prefix_w = draw.textlength(label_prefix, font=font_title)
            value_max_w = max_text_w - prefix_w
            if draw.textlength(value, font=font_value) <= value_max_w:
                all_lines.append((label_prefix, value, 0, False))
            else:
                # 内容超宽，首行标题+部分内容，续行缩进
                cur = ''
                first = True
                for ch in value:
                    if draw.textlength(cur + ch, font=font_value) > (value_max_w if first else max_text_w):
                        if cur:
                            all_lines.append((label_prefix if first else '', cur, 0 if first else prefix_w, not first))
                        cur = ch
                        first = False
                    else:
                        cur += ch
                if cur:
                    all_lines.append((label_prefix if first else '', cur, 0 if first else prefix_w, not first))

        line_gap = 56
        wrap_gap = 36

        # 计算总高度，垂直居中
        total_h = sum(wrap_gap if c else line_gap for _, _, _, c in all_lines)
        total_h -= wrap_gap if all_lines[-1][3] else line_gap
        total_h += 32
        y = max(pad, (PH - total_h) / 2)

        for i, (label_prefix, value, indent, is_cont) in enumerate(all_lines):
            if i > 0:
                y += wrap_gap if is_cont else line_gap
            if label_prefix:
                draw.text((pad, y), label_prefix, fill='black', font=font_title)
                tw = draw.textlength(label_prefix, font=font_title)
            else:
                tw = 0
            draw.text((pad + tw + indent, y + value_y_offset), value, fill='black', font=font_value)

        # 旋转用于打印
        img = img_preview.transpose(Image.Transpose.ROTATE_90)

        # 返回 base64：preview（未旋转可读）+ image（旋转后打印）
        buf_preview = BytesIO()
        img_preview.save(buf_preview, format='PNG')
        preview_b64 = base64.b64encode(buf_preview.getvalue()).decode('utf-8')

        buf_print = BytesIO()
        img.save(buf_print, format='PNG')
        print_b64 = base64.b64encode(buf_print.getvalue()).decode('utf-8')
        return jsonify({"success": True, "preview": f"data:image/png;base64,{preview_b64}", "image": f"data:image/png;base64,{print_b64}"})
    except Exception as e:
        return jsonify({"success": False, "message": f"生成标签失败: {str(e)}"}), 500


def _niimbot_ensure_connected():
    """确保打印机已连接，未连接则自动扫描连接。返回 (成功, 错误信息)。"""
    try:
        resp = requests.get(f"{NIIMBOT_SERVER}/connected", timeout=3)
        if resp.ok and resp.json().get("connected"):
            return True, None
    except requests.exceptions.ConnectionError:
        # 打印服务进程不在 → 自动重启；失败则给出准确信息，避免误导为「扫描打印机失败」
        if not start_niimbot_server():
            return False, "打印服务未启动，自动重启失败，请手动运行 npm start"
    except Exception:
        pass
    # 扫描串口，过滤非打印机设备，逐个尝试连接
    try:
        resp = requests.post(f"{NIIMBOT_SERVER}/scan", json={"transport": "serial"}, timeout=5)
        devices = resp.json().get("devices", [])
        # 跳过已知非打印机设备（CH340 是常见串口转接芯片）
        skip_names = ('ch340', 'cp210', 'ft232', 'pl2303')
        candidates = [d for d in devices if not any(s in d.get('name', '').lower() for s in skip_names)]
        for dev in candidates:
            addr = dev["address"]
            try:
                conn = requests.post(f"{NIIMBOT_SERVER}/connect", json={"transport": "serial", "address": addr}, timeout=5)
                if conn.ok and conn.json().get("message") == "Connected":
                    # 验证打印机是否真正响应
                    try:
                        info = requests.get(f"{NIIMBOT_SERVER}/info", timeout=3)
                        if info.ok and info.json().get("printerInfo"):
                            return True, None
                    except Exception:
                        pass
                    try:
                        requests.post(f"{NIIMBOT_SERVER}/disconnect", timeout=3)
                    except Exception:
                        pass
            except Exception:
                continue
        return False, "未找到打印机，请检查 USB 连接"
    except Exception as e:
        return False, f"扫描打印机失败: {str(e)}"


@app.route('/api/lims/do_print', methods=['POST'])
def do_print():
    try:
        data = request.json or {}
        image_base64 = data.get("image_base64", "")
        quantity = data.get("quantity", 1)

        if not image_base64:
            return jsonify({"success": False, "message": "缺少标签图片"}), 400

        # 去掉 data:image/png;base64, 前缀
        if "," in image_base64:
            image_base64 = image_base64.split(",", 1)[1]

        ok, err = _niimbot_ensure_connected()
        if not ok:
            return jsonify({"success": False, "message": err}), 503

        resp = requests.post(f"{NIIMBOT_SERVER}/print", json={
            "printTask": "B1",
            "printDirection": "top",
            "density": 3,
            "quantity": quantity,
            "imageBase64": image_base64,
            "labelWidth": 472,
            "labelHeight": 709,
            "imageFit": "contain"
        }, timeout=30)

        if resp.ok:
            return jsonify({"success": True})
        else:
            return jsonify({"success": False, "message": resp.json().get("error", "打印失败")}), 500
    except requests.exceptions.ConnectionError:
        return jsonify({"success": False, "message": "打印服务未启动，请先运行 niimblue-cli server -p 5001 --cors"}), 503
    except Exception as e:
        return jsonify({"success": False, "message": f"打印失败: {str(e)}"}), 500


import subprocess

NIIMBOT_SERVER = "http://localhost:5001"


def start_niimbot_server():
    """自动启动 niimblue-cli 打印服务（后台运行）。"""
    try:
        requests.get(f"{NIIMBOT_SERVER}/connected", timeout=2)
        print("  打印服务已在运行")
        return True
    except Exception:
        pass
    try:
        project_dir = os.path.dirname(os.path.abspath(__file__))
        subprocess.Popen(
            f'npm start',
            shell=True,
            cwd=project_dir,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        # 等待服务启动
        for _ in range(10):
            time.sleep(1)
            try:
                requests.get(f"{NIIMBOT_SERVER}/connected", timeout=2)
                print("  打印服务已自动启动")
                return True
            except Exception:
                continue
        print("  警告：打印服务启动超时，请手动运行 npm start")
        return False
    except Exception as e:
        print(f"  警告：打印服务启动失败: {e}，请手动运行 npm start")
        return False


# ==================== 标液核查草稿（仅创建者可见，3 自然日过期） ====================
DRAFT_DIR = 'drafts'
DRAFT_EXPIRE_DAYS = 3


def _draft_user_dir(username):
    """草稿按账号隔离：返回 drafts/<username>/ 目录（已创建）。用户名做安全过滤。"""
    safe = re.sub(r'[^A-Za-z0-9_一-鿿.-]', '_', username or 'anon')
    d = os.path.join(DRAFT_DIR, safe)
    os.makedirs(d, exist_ok=True)
    return d


def _draft_purge_expired(username):
    """删除已过期草稿；逐文件 try/except，损坏文件原样保留（绝不静默销毁数据）。"""
    d = _draft_user_dir(username)
    now = datetime.datetime.now()
    for fn in os.listdir(d):
        if not fn.endswith('.json'):
            continue
        path = os.path.join(d, fn)
        try:
            with open(path, 'r', encoding='utf-8') as f:
                doc = json.load(f)
            exp = doc.get('expires_at', '')
            if exp and datetime.datetime.strptime(exp, '%Y-%m-%d %H:%M:%S') < now:
                os.remove(path)
        except Exception:
            continue


def _draft_write_atomic(path, obj):
    tmp = path + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _draft_iter_all_docs():
    """遍历 drafts/ 下所有用户目录的草稿，yield (user_dir, filename, doc)。供管理员视角用。
    跳过非目录/非 .json/读取失败；不删除任何文件（过期是否显示由调用方按 expires_at 判断）。"""
    if not os.path.isdir(DRAFT_DIR):
        return
    for user_dir in os.listdir(DRAFT_DIR):
        ud = os.path.join(DRAFT_DIR, user_dir)
        if not os.path.isdir(ud):
            continue
        for fn in os.listdir(ud):
            if not fn.endswith('.json'):
                continue
            try:
                with open(os.path.join(ud, fn), 'r', encoding='utf-8') as f:
                    doc = json.load(f)
            except Exception:
                continue
            yield user_dir, fn, doc



# ==================== 标液期间核查：草稿 ====================
@app.route('/api/verification/draft/save', methods=['POST'])
def verification_draft_save():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    username = session.get('username') or 'anon'
    p = request.get_json() or {}
    payload = p.get('payload')
    if not isinstance(payload, dict):
        return jsonify({"success": False, "message": "草稿内容无效"}), 400
    if len(json.dumps(payload, ensure_ascii=False)) > 10_000_000:
        return jsonify({"success": False, "message": "草稿过大，请减少 .D 数据"}), 400

    d = _draft_user_dir(username)
    draft_id = p.get('draft_id') or ''
    path = None
    if re.match(r'^[A-Za-z0-9]{32}$', draft_id):
        cand = os.path.join(d, draft_id + '.json')
        if os.path.exists(cand):
            path = cand  # 复用现有 id（更新）
    if not path:
        draft_id = uuid.uuid4().hex
        path = os.path.join(d, draft_id + '.json')

    now = datetime.datetime.now()
    expires = now + datetime.timedelta(days=DRAFT_EXPIRE_DAYS)
    doc = {
        "draft_id": draft_id,
        "created_by_username": username,
        "created_by_name": session.get('display_name') or username,
        "created_at": now.strftime('%Y-%m-%d %H:%M:%S'),
        "expires_at": expires.strftime('%Y-%m-%d %H:%M:%S'),
        "title": (p.get('title') or draft_id[:8]),
        "payload": payload,
    }
    _draft_write_atomic(path, doc)
    _draft_purge_expired(username)
    return jsonify({"success": True, "draft_id": draft_id, "title": doc["title"], "expires_at": doc["expires_at"]})


@app.route('/api/verification/draft/list', methods=['GET'])
def verification_draft_list():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    username = session.get('username') or 'anon'
    is_admin = (session.get('display_name') or '').strip() == _ADMIN_DISPLAY_NAME
    now = datetime.datetime.now()

    def _expired(doc):
        exp = doc.get('expires_at', '')
        try:
            return bool(exp) and datetime.datetime.strptime(exp, '%Y-%m-%d %H:%M:%S') < now
        except Exception:
            return False

    out = []
    if is_admin:
        # 管理员视角：遍历所有账号草稿，过滤已过期（仅不显示，不删他人文件）
        for user_dir, fn, doc in _draft_iter_all_docs():
            if _expired(doc):
                continue
            payload = doc.get('payload') or {}
            out.append({
                "draft_id": doc.get('draft_id', fn[:-5]),
                "title": doc.get('title', ''),
                "created_at": doc.get('created_at', ''),
                "expires_at": doc.get('expires_at', ''),
                "has_d_data": bool(payload.get('epatempContents')),
                "created_by_name": doc.get('created_by_name', ''),
                "created_by_username": doc.get('created_by_username', user_dir),
                "mine": doc.get('created_by_username', '') == username,
            })
    else:
        _draft_purge_expired(username)
        d = _draft_user_dir(username)
        for fn in os.listdir(d):
            if not fn.endswith('.json'):
                continue
            try:
                with open(os.path.join(d, fn), 'r', encoding='utf-8') as f:
                    doc = json.load(f)
            except Exception:
                continue
            payload = doc.get('payload') or {}
            out.append({
                "draft_id": doc.get('draft_id', fn[:-5]),
                "title": doc.get('title', ''),
                "created_at": doc.get('created_at', ''),
                "expires_at": doc.get('expires_at', ''),
                "has_d_data": bool(payload.get('epatempContents')),
                "created_by_name": doc.get('created_by_name', ''),
                "mine": True,
            })
    out.sort(key=lambda x: x.get('created_at', ''), reverse=True)
    return jsonify({"success": True, "drafts": out})


@app.route('/api/verification/draft/<draft_id>', methods=['GET'])
def verification_draft_load(draft_id):
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    if not re.match(r'^[A-Za-z0-9]{32}$', draft_id):
        return jsonify({"success": False, "message": "草稿不存在或已过期"}), 404
    username = session.get('username') or 'anon'
    is_admin = (session.get('display_name') or '').strip() == _ADMIN_DISPLAY_NAME
    if is_admin:
        # 管理员可加载他人草稿：按 ?username= 定位目录（_draft_user_dir 已过滤防路径穿越）
        target = (request.args.get('username') or '').strip()
        d = _draft_user_dir(target) if target else _draft_user_dir(username)
    else:
        d = _draft_user_dir(username)   # 非管理员忽略 username 参数，强制自己目录
    path = os.path.join(d, draft_id + '.json')
    if not os.path.exists(path):
        return jsonify({"success": False, "message": "草稿不存在或已过期"}), 404
    try:
        with open(path, 'r', encoding='utf-8') as f:
            doc = json.load(f)
    except Exception:
        return jsonify({"success": False, "message": "草稿损坏"}), 500
    try:
        if datetime.datetime.strptime(doc.get('expires_at', ''), '%Y-%m-%d %H:%M:%S') < datetime.datetime.now():
            os.remove(path)
            return jsonify({"success": False, "message": "草稿已过期"}), 404
    except Exception:
        pass
    return jsonify({"success": True, "draft": doc})


@app.route('/api/verification/draft/<draft_id>/delete', methods=['POST', 'DELETE'])
def verification_draft_delete(draft_id):
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    if re.match(r'^[A-Za-z0-9]{32}$', draft_id):
        username = session.get('username') or 'anon'
        path = os.path.join(_draft_user_dir(username), draft_id + '.json')
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception:
            pass
    return jsonify({"success": True})


if __name__ == '__main__':
    if not os.path.exists('templates'):
        os.makedirs('templates')
    print("Flask 服务启动，访问以下地址：")
    print("  登录页面: http://127.0.0.1:5000/login")
    print("  耗材查询主页: http://127.0.0.1:5000/")
    print("  有机标准品管理: http://127.0.0.1:5000/organic-std")
    start_niimbot_server()
    # 钉钉标液过期提醒：仅在 reloader 子进程启动调度线程，避免双实例重复发送
    if os.environ.get('WERKZEUG_RUN_MAIN') == 'true':
        try:
            import dingtalk_notify
            dingtalk_notify.start()
        except Exception as e:
            print(f"[DingTalk] 调度器启动失败: {e}")
    app.run(host='0.0.0.0', port=5000, debug=True)