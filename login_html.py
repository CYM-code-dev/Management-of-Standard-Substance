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

# ==================== 配置文件路径 ====================
CONFIG_FILE = 'config.json'


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
    return round(val, 2) if unit == 'mL' else round(val, 4)


def _fmt_qty(val, unit):
    if unit == 'mL':
        return f"{val:.2f}"
    if val == int(val):
        return str(int(val))
    return f"{val:.4f}".rstrip('0').rstrip('.')


def _round_conc(val, conc_unit):
    if conc_unit == '%':
        return val
    return round(val, 3) if val < 0.10 else round(val, 2)


def _fmt_conc(val, conc_unit):
    if conc_unit == '%':
        return f"{val:.10f}".rstrip('0').rstrip('.') if isinstance(val, float) else str(val)
    decimals = 3 if val < 0.10 else 2
    return f"{val:.{decimals}f}"


def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'excelPath': '', 'certPath': ''}


def save_config(excel_path, cert_path):
    config = {'excelPath': excel_path, 'certPath': cert_path}
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2)


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
        try:
            resp = self.session.get(f"{self.base_url}/detectionManager/core/security/getLoginUser")
            return resp.status_code == 200 and resp.json().get("success")
        except:
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
                self.session.get(f"{self.base_url}/detectionManager/core/security/getLoginUser")
            except: pass
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

# ==================== 路由 ====================
@app.route('/')
def index():
    if not session.get('logged_in'): return redirect('/login')
    return redirect('/organic-std')

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
        return jsonify({"logged_in": True, "username": session.get('username'), "display_name": session.get('display_name')})
    system = get_system()
    if system.current_user and system.verify_session():
        session['logged_in'] = True
        session['username'] = system.current_user
        session['display_name'] = system.current_real_name or system.current_user
        session['pid'] = system.current_pid
        return jsonify({"logged_in": True, "username": system.current_user, "display_name": session['display_name']})
    return jsonify({"logged_in": False})

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    get_system().logout()
    return jsonify({"success": True})

@app.route('/api/query')
def query():
    if not session.get('logged_in'): return jsonify({"success": False, "message": "未登录"})
    pid = session.get('pid')
    username = session.get('username')
    if not pid:
        system = get_system()
        if system.current_pid: pid = session['pid'] = system.current_pid
        else: return jsonify({"success": False, "message": "无法获取用户PID"})
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
            if resp.status_code in (401,403): session.pop('logged_in', None); system.logout(); return jsonify({"success": False, "message": "远程会话已失效"})
            return jsonify({"success": False, "message": f"请求失败，状态码: {resp.status_code}"})
        result = resp.json()
        rd = result.get("resultData") or {}
        if result.get("success"):
            vo_list = rd.get("voList", [])
            records = rd.get("records", 0)
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
        if "未登录" in error_msg or "login" in error_msg.lower(): session.pop('logged_in', None); system.logout(); return jsonify({"success": False, "message": "远程会话已失效"})
        return jsonify({"success": False, "message": error_msg})
    except Exception as e:
        return jsonify({"success": False, "message": f"查询异常: {str(e)}"})

@app.route('/api/config', methods=['GET', 'POST'])
def handle_config():
    if request.method == 'GET': return jsonify(load_config())
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

@app.route('/organic_excel/<path:filename>')
def serve_excel(filename):
    config = load_config()
    excel_dir = os.path.dirname(config.get('excelPath', ''))
    if not excel_dir:
        return jsonify({"success": False, "message": "未配置 Excel 路径"}), 400
    return send_from_directory(excel_dir, filename)

@app.route('/certificates/<path:filename>')
def serve_cert(filename):
    config = load_config()
    cert_dir = config.get('certPath', '')
    if not cert_dir:
        return jsonify({"success": False, "message": "未配置证书路径"}), 400
    return send_from_directory(cert_dir, filename)


# ==================== 智能插入写入 Excel ====================
def get_prefix_type(original_id):
    if not original_id: return 'number'
    first = original_id[0].upper()
    if first == 'D': return 'D'
    if first == 'E': return 'E'
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

def find_insert_position(records, new_id):
    new_type = get_prefix_type(new_id)
    new_num = extract_number(new_id)
    numbers = [r for r in records if get_prefix_type(r['original_id']) == 'number']
    d_records = [r for r in records if get_prefix_type(r['original_id']) == 'D']
    e_records = [r for r in records if get_prefix_type(r['original_id']) == 'E']
    numbers.sort(key=lambda x: extract_number(x['original_id']))
    d_records.sort(key=lambda x: extract_number(x['original_id']))
    e_records.sort(key=lambda x: extract_number(x['original_id']))
    if new_type == 'number':
        idx = 0
        for r in numbers:
            if extract_number(r['original_id']) < new_num: idx += 1
            else: break
        return 2 + idx
    elif new_type == 'D':
        base = 2 + len(numbers)
        idx = 0
        for r in d_records:
            if extract_number(r['original_id']) < new_num: idx += 1
            else: break
        return base + idx
    else:
        base = 2 + len(numbers) + len(d_records)
        idx = 0
        for r in e_records:
            if extract_number(r['original_id']) < new_num: idx += 1
            else: break
        return base + idx

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


@app.route('/api/add_to_excel', methods=['POST'])
def add_to_excel():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    try:
        data = request.get_json()
        record = data.get('record')
        if not record: return jsonify({"success": False, "message": "无数据"}), 400
        config = load_config()
        excel_path = config.get('excelPath', '').strip()
        if not excel_path: return jsonify({"success": False, "message": "未配置 Excel 路径"}), 400
        ext = os.path.splitext(excel_path)[1].lower()
        if ext not in ['.xls', '.xlsx']: return jsonify({"success": False, "message": "仅支持 .xls 或 .xlsx"}), 400

        sheet_name = '有机标准物质'
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
        return jsonify({"success": False, "message": f"写入 Excel 失败: {str(e)}"}), 500


@app.route('/api/update_lims_unit', methods=['POST'])
def update_lims_unit():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    try:
        data = request.get_json()
        record_id = data.get('recordId')
        new_unit = data.get('concentrationUnitName', '').strip()
        lims_item = data.get('limsItem')
        if not record_id or not lims_item:
            return jsonify({"success": False, "message": "缺少记录ID或原始数据"}), 400

        system = get_system()
        if not system.current_user:
            return jsonify({"success": False, "message": "远程会话已失效，请重新登录"})

        form_data = {}
        form_data['concentrationUnitName'] = new_unit
        form_data['pid'] = system.current_pid
        form_data['pname'] = system.current_user
        form_data['loginId'] = system.current_pid
        form_data['_method'] = 'PUT'

        url = f"{system.base_url}/detectionManager/manager/consumableBill/{record_id}"
        for k in list(form_data.keys()):
            if form_data[k] is None:
                form_data[k] = ''
        resp = system.session.post(url, data=form_data)
        if resp.status_code != 200:
            try: detail = resp.text[:2000]
            except: pass
            print(f"[LIMS UPDATE] response: {detail}")
            return jsonify({"success": False, "message": f"LIMS 请求失败，状态码: {resp.status_code}，{detail}"})
        result = resp.json()
        if result.get('success'):
            return jsonify({"success": True, "message": "浓度单位已同步到 LIMS"})
        else:
            return jsonify({"success": False, "message": result.get('errorDesc', '更新失败')})
    except Exception as e:
        return jsonify({"success": False, "message": f"LIMS 同步异常: {str(e)}"}), 500


@app.route('/api/lims/receive', methods=['POST'])
def lims_receive():
    if not session.get('logged_in'):
        return jsonify({"success": False, "message": "未登录"}), 401
    data = request.get_json()
    consumable_id = str(data.get('consumable_id', ''))
    quantity = data.get('quantity')
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
    form_data = {
        "num": str(quantity),
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
        if not resp.ok:
            print(f"[saveSolutionConfigure] status={resp.status_code} body={resp.text[:500]}")
        resp.raise_for_status()
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
            print(f"[ResolveOrder] {configure_order} type={sol_type} fetched {len(items)} records")
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
        print(f"[ResolveOrder] NO match for {configure_order}")
    except Exception as e:
        print(f"[ResolveOrder] ERROR {configure_order}: {e}")
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
                        print(f"[ResolveSC] {solution_code} → id={lid} ({co})")
                        return lid, co
        except Exception as e:
            print(f"[ResolveSC] ERROR solutionCode={solution_code} type={sol_type}: {e}")
    print(f"[ResolveSC] not found: solutionCode={solution_code}")
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


def _fetch_solution_detail(system, solution_id):
    """Fetch record via detail API (saveDetailList always null). Fallback only."""
    url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/detail"
    headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
    resp = system.session.get(url, params={"id": solution_id}, headers=headers)
    resp.raise_for_status()
    data = resp.json()
    return data.get('resultData') or {}


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
                print(f"[TraceFetch] fallback: configureOrder number as ID: {resolved}")
        if not resolved:
            return None

        is_solution_code = order_str and order_str.startswith('CK-')

        def _validate(rec):
            if not order_str:
                return True
            if is_solution_code:
                rec_sc = str(rec.get('solutionCode', '')).strip()
                if rec_sc and rec_sc != order_str:
                    print(f"[TraceFetch] ID {resolved} solutionCode={rec_sc}, expected {order_str}, skipping")
                    return False
            else:
                rec_order = str(rec.get('configureOrder', '')).strip()
                if rec_order and rec_order != order_str:
                    print(f"[TraceFetch] ID {resolved} returned {rec_order}, expected {order_str}, skipping")
                    return False
            return True

        try:
            rec = _fetch_solution_view(system, resolved, order_str=view_order_str)
            if rec and rec.get('id') and _validate(rec):
                return rec
        except Exception as e:
            print(f"[TraceFetch] view error id={resolved}: {e}")
        try:
            rec = _fetch_solution_detail(system, resolved)
            if rec and rec.get('id') and _validate(rec):
                return rec
        except Exception as e:
            print(f"[TraceFetch] detail error id={resolved}: {e}")
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

        # Extract parent info from detailList (if available) or from parent record
        parent_name = ''
        parent_conc_val = 0
        parent_conc_unit = ''
        received_qty = ''
        received_unit = str(record.get('receivedUint') or 'mL').strip()

        if detail_list and not is_d_type:
            # B/C type: use first detailList row for parent info
            dl0 = detail_list[0]
            parent_name = str(dl0.get('originalName', '')).strip()
            p_conc = str(dl0.get('originalConcentration', '') or '').strip()
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

        print(f"[TraceRecord] {rec_order} | level={prefix} | solutionCode={record.get('solutionCode','')} | "
              f"conc='{rec_conc}' | parent='{parent_name}' | parent_conc={parent_conc_val}({parent_conc_unit}) | "
              f"vol={constant_volume} | recv={received_qty}{received_unit} | weighing={is_weighing_top} | "
              f"originalCode='{original_code}' | detailList={len(detail_list)}行")
        if detail_list:
            for di, dl in enumerate(detail_list):
                print(f"  detailList[{di}]: originalCode={dl.get('originalCode','')} originalNo={repr(dl.get('originalNo',''))} originalName={dl.get('originalName','')} "
                      f"originalConcentration={dl.get('originalConcentration','')} receivedQuantity={dl.get('receivedQuantity','')} "
                      f"originalId={dl.get('originalId','')} volume={dl.get('volume','')} medium={dl.get('medium','')}")

        matched.append({
            'level': prefix,
            'configure_order': rec_order,
            'solution_name': str(record.get('solutionName') or '').strip(),
            'solution_code': str(record.get('solutionCode') or '').strip(),
            'concentration': rec_conc,
            'constant_volume': str(constant_volume),
            'medium': str(record.get('medium') or '').strip(),
            'configure_date': rec_date,
            'validity_date': str(record.get('validityDate') or '').strip(),
            'controlled_no': str(record.get('controlledNo') or '').strip(),
            'storage_condition': str(record.get('storageCondition') or '').strip(),
            'received_quantity': received_qty,
            'received_unit': received_unit,
            'parent_name': parent_name,
            'parent_concentration': str(parent_conc_val),
            'parent_conc_unit': parent_conc_unit,
            'original_code': original_code,
            'concentration_count': str(record.get('concentrationCount') or '').strip(),
            '_is_weighing_top': is_weighing_top,
            '_detail_list': detail_list,
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
            print(f"[TraceParents] {rec_order} → parent_orders={parent_orders} | parent_id_map={parent_id_map}")
            for po in parent_orders:
                po_prefix = po.split('-')[0].upper() if '-' in po else ''
                if po_prefix == 'A':
                    print(f"[TraceStop] {rec_order} → 父级 {po} 为 A 类，停止")
                    continue
                if _has_pct_source_conc(record, po):
                    print(f"[TraceStop] {rec_order} → 父级 {po} 源浓度为 %，停止")
                    continue
                pid = parent_id_map.get(po)
                print(f"[TraceContinue] {rec_order} → 追溯父级 {po} (id={pid})")
                _trace(lims_id=pid, order_str=po, depth=depth + 1)

        visiting.discard(key)
        return None

    for lims_id, order in trace_targets:
        _trace(lims_id=lims_id, order_str=order)

    matched.reverse()

    print(f"[TraceSummary] ===== 溯源完成 ===== 共 {len(matched)} 条记录，顺序(top→bottom):")
    for mi, rec in enumerate(matched):
        print(f"  [{mi}] {rec.get('level','')}-{rec.get('configure_order','')} | "
              f"name={rec.get('solution_name','')} | conc={rec.get('concentration','')} | "
              f"originalCode={rec.get('original_code','')} | parent={rec.get('parent_name','')} | "
              f"source_details={len(rec.get('source_details',[]))}条 | "
              f"weighing={rec.get('_is_weighing_top',False)}")

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

    print(f"[TraceSummary] top_ancestor={top_ancestor}")
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
        "solutionName": "", "solutionCode": "", "customType": "", "configStatus": "",
        "controlledNo": "", "storageLocation": "",
        "configureStartDate": date_from, "configureEndDate": date_to,
        "configureUserName": operator, "receiveUserName": "", "auditStatus": "",
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
        print(f"[GetSourceInfo] ERROR order={configure_order} id={source_id}: {e}")
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
        if not resp.ok:
            print(f"[save_solution_b] status={resp.status_code} body={resp.text[:500]}")
        resp.raise_for_status()
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
        if not resp.ok:
            print(f"[save_working_solution] status={resp.status_code} body={resp.text[:500]}")
        resp.raise_for_status()
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
        if not resp.ok:
            print(f"[updateObj1] status={resp.status_code} body={resp.text[:500]}")
        result = resp.json()
        if not result.get("success"):
            err_ctx = result.get('errorCtx') or {}
            err_msg = result.get('errorDesc') or (err_ctx.get('errorMsg') if isinstance(err_ctx, dict) else '') or '修改失败'
            return jsonify({"success": False, "message": err_msg})
        return jsonify({"success": True})
    except Exception as e:
        print(f"[updateObj1] exception: {e}")
        return jsonify({"success": False, "message": f"修改异常: {str(e)}"}), 500


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
        # 如果传了 consumable_ids，先查询对应的领用记录ID
        consumable_ids = p.get('consumable_ids')
        receive_ids = p.get('ids')
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
            if not found_ids:
                return jsonify({"success": True, "message": "无关联领用记录"})
            receive_ids = found_ids
        if not receive_ids:
            return jsonify({"success": True, "message": "无领用记录需要删除"})
        if isinstance(receive_ids, list):
            receive_ids = ','.join(str(i) for i in receive_ids)
        url = f"{system.base_url}/detectionManager/manager/consumableReceive"
        form_data = {
            'ids': str(receive_ids),
            'pid': str(pid),
            'pname': pname,
            'loginId': str(pid),
            '_method': 'DELETE',
        }
        resp = system.session.post(url, data=form_data)
        if not resp.ok:
            print(f"[deleteReceive] status={resp.status_code} body={resp.text[:500]}")
        result = resp.json()
        if not result.get("success"):
            err_ctx = result.get('errorCtx') or {}
            err_msg = result.get('errorDesc') or (err_ctx.get('errorMsg') if isinstance(err_ctx, dict) else '') or '删除领用记录失败'
            return jsonify({"success": False, "message": err_msg})
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

def _docx_clear_tc(tc):
    for p in tc.findall(_docx_qn('w:p')):
        tc.remove(p)

def _docx_set_tc_text(tc, text, center=True, sz=18):
    _docx_clear_tc(tc)
    tc.append(_docx_make_para(text, center=center, sz=sz))

def _docx_set_tc_two_lines(tc, line1, line2, sz=18):
    _docx_clear_tc(tc)
    tc.append(_docx_make_para(line1, center=True, sz=sz))
    tc.append(_docx_make_para(line2, center=True, sz=sz))

def _docx_set_tc_multi_lines(tc, lines, sz=18):
    _docx_clear_tc(tc)
    for line in lines:
        tc.append(_docx_make_para(line, center=True, sz=sz))

def _docx_set_vmerge(tc, mode):
    tcPr = tc.find(_docx_qn('w:tcPr'))
    if tcPr is None:
        tcPr = _docx_OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    vm = tcPr.find(_docx_qn('w:vMerge'))
    if vm is None:
        vm = _docx_OxmlElement('w:vMerge')
        tcPr.append(vm)
    if mode == 'restart':
        vm.set(_docx_qn('w:val'), 'restart')
    else:
        if _docx_qn('w:val') in vm.attrib:
            del vm.attrib[_docx_qn('w:val')]


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
        for p in tc.findall(qn('w:p')):
            tc.remove(p)
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
            tc.append(p_el)

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

    _fill_tc1(tc4, f'浓度({conc_unit})')

    # Header concentration: single source → specific value, multi source → 见下表
    if len(top_items) <= 1:
        cv = _parse_conc_val(top_items[0].get('originalConcentration', '')) if top_items else ''
        _fill_tc1(tc5, cv)
        show_source_names = False
    else:
        _fill_tc1(tc5, '见下表')
        show_source_names = True

    r3_tcs = table.rows[3]._tr.findall(qn('w:tc'))
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
        for _ in range(n_items - n_tpl):
            src_tr    = table.rows[remark_row_idx - 1]._tr
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

    # Column values for data rows (index 2=溶剂, 3=稀释至, 4=浓度, 5=编号, 6=配制日期, 7=有效期)
    col_vals_map = {
        2: [rv['medium']      for rv in row_values],
        3: [rv['volume']      for rv in row_values],
        4: [rv['result_conc'] for rv in row_values],
        5: [rv['result_code'] for rv in row_values],
        6: [rv['configure_date'] for rv in row_values],
        7: [rv['validity_date'] for rv in row_values],
    }

    # Columns that should follow 编号(col5) merge pattern: merge whenever 编号 merges
    follow_code_merge_cols = {2, 3, 6, 7}
    # Columns that merge based on their own value equality
    value_merge_cols = {2, 3, 4, 5, 6, 7} if not (is_working and not is_multi_source) else {4, 5}

    for i, item in enumerate(body_list):
        tr  = table.rows[4 + i]._tr
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) < 8:
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

        # Determine if 编号(col5) would merge with previous row
        same_dilution = i > 0 and row_values[i]['dilutionIdx'] == row_values[i - 1]['dilutionIdx']
        code_same_as_prev = i > 0 and row_values[i]['result_code'] == row_values[i - 1]['result_code']
        code_should_merge = same_dilution and code_same_as_prev

        # Columns 2-7
        for tc_idx in range(2, 8):
            val = col_vals_map[tc_idx][i]
            if tc_idx in follow_code_merge_cols:
                # 溶剂/稀释至/配置日期/有效期 follow 编号 merge pattern
                if code_should_merge:
                    _clear_tc(tcs[tc_idx])
                    tcs[tc_idx].append(_make_para('', center=True))
                    _set_vmerge(tcs[tc_idx], 'continue')
                else:
                    _set_tc_text(tcs[tc_idx], val)
                    _set_vmerge(tcs[tc_idx], 'restart')
            elif tc_idx in value_merge_cols:
                # 浓度/编号 merge based on their own value equality
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
    dl_name = (f"{solution_code}_配制记录.docx" if solution_code else "配制记录.docx") \
              .replace('/', '-').replace('\\', '-')
    return send_file(buf, as_attachment=True, download_name=dl_name,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


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

    config = load_config()
    excel_path = config.get('excelPath', '').strip()
    if not excel_path:
        return jsonify({"success": False, "message": "未配置 Excel 路径"}), 400

    ext = os.path.splitext(excel_path)[1].lower()
    sheet_name = '有机标准物质'

    try:
        if ext == '.xlsx':
            wb = openpyxl.load_workbook(excel_path)
            if sheet_name not in wb.sheetnames:
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
            if sheet_name not in rb.sheet_names():
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


def _build_concentration_point_code(record, point):
    """构建浓度点编号: CK-CG-{num}-{concentration}-{date}"""
    code = str(record.get('solutionCode', '')).strip()
    date = str(record.get('configureDate', ''))[:10].replace('-', '')
    conc_str = point.get('conc_str', f"{point['concentration']:g}")
    parts = code.split('-')
    if len(parts) >= 3:
        base = '-'.join(parts[:3])
        return f"{base}-{conc_str}-{date}"
    return f"{code}-{conc_str}-{date}"


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
            elif conc_str:
                sname = str(rec.get('solutionName', '')).strip()
                cparts = re.findall(r'([\d.]+)\s*\(([^)]+)\)', conc_str)
                if not cparts:
                    cparts = re.findall(r'([\d.]+)\s*(mg/L|μg/mL|ug/mL|ng/ml|ppm)', conc_str)
                if cparts:
                    all_unit = cparts[0][1]
                    substance_concs[sname] = [float(cv) for cv, cu in cparts]
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
                    pt = {'index': pi, 'concentration': cv, 'unit': all_unit, 'conc_str': f"{cv:g}",
                          'name': ','.join(all_sub_names), 'sub_concs': sub_concs}
                    pt['code'] = _build_concentration_point_code(rec, pt)
                    points.append(pt)
            else:
                sorted_codes = sorted(code_map.items(), key=lambda x: x[1].get('dilutionIdx', 0))
                for pi, (rc, info) in enumerate(sorted_codes):
                    sub_concs = {sn: substance_concs[sn][pi] for sn in all_sub_names if pi < len(substance_concs.get(sn, []))}
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
                        'conc_str': conc_from_code or '',
                    })
            point_codes = [pt['code'] for pt in points]
            return {
                'id': rec.get('id'),
                'configureOrder': rec.get('configureOrder', ''),
                'solutionCode': rec.get('solutionCode', ''),
                'solutionName': rec.get('solutionName', ''),
                'configureDate': str(rec.get('configureDate', ''))[:10],
                'concentrationPoints': points,
                'validityDate': str(rec.get('validityDate', ''))[:10],
                'customType': rec.get('customType', ''),
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
                    for dl in (r.get('_detail_list') or []):
                        dl_name = str(dl.get('originalName', '')).strip()
                        if dl_name and dl_name == a_name:
                            dl_no = str(dl.get('originalNo', '')).strip()
                            # originalNo格式如 'A-24092\nCIRS400484-3\nCK-CG-2025214'
                            # 取最后一段（CK-CG-xxx）作为controlledNo
                            lines = [l.strip() for l in dl_no.split('\n') if l.strip()]
                            a_controlled_no = lines[-1] if lines else ''
                            a_concentration = str(dl.get('originalConcentration', '')).strip()
                            break
                    # 通过consumableBill API用controlledNo查询保存条件
                    a_storage_condition = ''
                    if a_controlled_no:
                        try:
                            import time as _time
                            cb_params = {
                                "_search": "false", "nd": str(int(_time.time()*1000)),
                                "pageSize": 30, "pageNo": 1, "sidx": "", "sord": "asc",
                                "type": "CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE", "casNo": "",
                                "orgName": "", "groupId": "", "status": "",
                                "keyword": a_controlled_no,
                                "pid": session.get('user_id', ''), "pname": session.get('username', ''),
                                "loginId": session.get('user_id', ''),
                            }
                            cb_resp = system.session.get(f"{system.base_url}/detectionManager/manager/consumableBill/pageObj", params=cb_params)
                            cb_data = cb_resp.json() if cb_resp.status_code == 200 else {}
                            vo_list = (cb_data.get('resultData') or {}).get('voList') or []
                            if vo_list:
                                a_storage_condition = str(vo_list[0].get('storageCondition') or '').strip()
                        except Exception as e:
                            print(f"[VerificationTrace] 查询consumableBill失败: {e}")
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
                print(f"[VerificationTrace] 溯源失败: {e}")
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
            print(f"[ParsePDF] 解析 {f.filename} 失败: {e}")
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
            for compound_name, (val, unit) in parsed.items():
                all_compounds.append({
                    'name': compound_name,
                    'measured_value': val,
                    'unit': unit,
                    'source_file': name,
                })
        except Exception as e:
            print(f"[ParseEpatemp] 解析 {name} 失败: {e}")
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
        for ri, item in enumerate(table0_data):
            row_idx = ri + 2  # 从第3行开始（0/1是表头）
            if row_idx >= len(t0.rows):
                break
            row = t0.rows[row_idx]
            row.cells[0].text = str(item.get('序号', ri + 1))
            row.cells[1].text = str(item.get('标准物质名称', ''))
            row.cells[2].text = str(item.get('标准物质编号', ''))
            row.cells[3].text = str(item.get('保存条件', ''))

        # ── Table1: 核查结果 ──
        # R0: 被核查对象编号 (合并单元格)
        target_code = str(p.get('被核查对象编号', ''))
        _docx_set_tc_multi_lines(t1.rows[0].cells[0], target_code.split(', '), sz=18)

        # R1: 核查对象编号 (合并单元格)
        source_code = str(p.get('核查对象编号', ''))
        _docx_set_tc_multi_lines(t1.rows[1].cells[0], source_code.split(', '), sz=18)

        # R2: 核查方法 + 核查时间
        t1.rows[2].cells[0].text = '核查方法：' + str(p.get('核查方法', '新旧标准品对比'))
        t1.rows[2].cells[2].text = '核查时间：' + str(p.get('核查时间', ''))

        # R3: 核查方法描述
        method_desc = str(p.get('核查方法描述', ''))
        _docx_set_tc_multi_lines(t1.rows[3].cells[0], [method_desc], sz=18)

        # R5-R20: 组分数据
        table1_data = p.get('table1_data', [])
        for ri, item in enumerate(table1_data):
            row_idx = ri + 5
            if row_idx >= len(t1.rows) - 2:  # 留出结论和备注行
                break
            row = t1.rows[row_idx]
            row.cells[0].text = str(item.get('各组分名称', ''))
            row.cells[1].text = str(item.get('理论值', ''))
            row.cells[2].text = str(item.get('实测值', ''))
            row.cells[3].text = str(item.get('相对偏差', ''))

        # R21: 核查结论
        conclusion = str(p.get('核查结论', '合格'))
        conclusion_text = f"核查结论：{'☑合格；☐不合格' if conclusion == '合格' else '☐合格；☑不合格'}"
        t1.rows[-2].cells[0].text = conclusion_text

        # R22: 备注
        t1.rows[-1].cells[0].text = '备注：' + str(p.get('备注', ''))

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        download_name = '标准物质期间核查记录.docx'
        from flask import send_file
        from urllib.parse import quote as _urlquote
        resp = send_file(buf, as_attachment=True, download_name=download_name,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
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
        wrap_gap = 40

        # 计算总高度，垂直居中
        total_h = sum(wrap_gap if c else line_gap for _, _, _, c in all_lines)
        total_h -= wrap_gap if all_lines[-1][3] else line_gap
        total_h += 32
        y = max(pad, (PH - total_h) / 2)

        for label_prefix, value, indent, is_cont in all_lines:
            if label_prefix:
                draw.text((pad, y), label_prefix, fill='black', font=font_title)
                tw = draw.textlength(label_prefix, font=font_title)
            else:
                tw = 0
            draw.text((pad + tw + indent, y + value_y_offset), value, fill='black', font=font_value)
            y += wrap_gap if is_cont else line_gap

        # 旋转用于打印
        img = img_preview.transpose(Image.Transpose.ROTATE_90)

        # 保存到文件
        output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'print_output')
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        file_path = os.path.join(output_dir, f"{solution_code}.png")
        img.save(file_path, 'PNG')

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


NIIMBOT_SERVER = "http://localhost:5001"


def _niimbot_ensure_connected():
    """确保打印机已连接，未连接则自动扫描连接。返回 (成功, 错误信息)。"""
    try:
        resp = requests.get(f"{NIIMBOT_SERVER}/connected", timeout=3)
        if resp.ok and resp.json().get("connected"):
            return True, None
    except Exception:
        pass
    # 扫描串口
    try:
        resp = requests.post(f"{NIIMBOT_SERVER}/scan", json={"transport": "serial"}, timeout=5)
        devices = resp.json().get("devices", [])
        for dev in devices:
            addr = dev["address"]
            try:
                conn = requests.post(f"{NIIMBOT_SERVER}/connect", json={"transport": "serial", "address": addr}, timeout=5)
                if conn.ok and conn.json().get("message") == "Connected":
                    return True, None
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


if __name__ == '__main__':
    if not os.path.exists('templates'):
        os.makedirs('templates')
    print("Flask 服务启动，访问以下地址：")
    print("  登录页面: http://127.0.0.1:5000/login")
    print("  耗材查询主页: http://127.0.0.1:5000/")
    print("  有机标准品管理: http://127.0.0.1:5000/organic-std")
    app.run(host='0.0.0.0', port=5000, debug=True)