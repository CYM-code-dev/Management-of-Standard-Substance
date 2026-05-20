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
                    real_name = result.get("resultData", {}).get("nickName", username)
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
        cookies_dict = {c.name: c.value for c in self.session.cookies}
        info = {
            "username": self.current_user, "pid": self.current_pid, "real_name": self.current_real_name,
            "login_time": time.strftime("%Y-%m-%d %H:%M:%S"),
            "cookies": cookies_dict, "headers": dict(self.session.headers)
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
            for name, value in info["cookies"].items():
                self.session.cookies.set(name, value)
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
    keyword = request.args.get('keyword', '').strip()
    if not keyword: return jsonify({"success": False, "message": "请输入查询关键词"})
    system = get_system()
    if system.current_user != username:
        system.current_user = username
        system.load_session()
    params = {
        "_search": "false", "nd": str(int(time.time()*1000)), "pageSize": 30, "pageNo": 1, "sidx": "", "sord": "asc",
        "type": "CONSUMABLE_DIR_TYPE_STANDARD_SUBSTANCE", "orgName": request.args.get('org_name', ''), "groupId": "", "status": request.args.get('status', 'normal'),
        "keyword": keyword, "state": request.args.get('status', 'normal'), "pid": pid, "pname": username, "loginId": pid
    }
    try:
        resp = system.session.get(f"{system.base_url}/detectionManager/manager/consumableBill/pageObj", params=params)
        if resp.status_code != 200:
            if resp.status_code in (401,403): session.pop('logged_in', None); system.logout(); return jsonify({"success": False, "message": "远程会话已失效"})
            return jsonify({"success": False, "message": f"请求失败，状态码: {resp.status_code}"})
        result = resp.json()
        if result.get("success"): return jsonify({"success": True, "data": result.get("resultData", {}).get("voList", [])})
        error_msg = result.get("errorCtx", {}).get("errorMsg", "查询失败")
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

    url = f"{system.base_url}/detectionManager/manager/consumableReceive/receive"
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
        "controlledNo": p.get('batch_no', ''),
        "medium": None, "configuratorId": None,
        "solutionType": p.get('solution_type', 'SOLUTION_TYPE_B'),
        "configuratorName": None, "constantVolume": 0, "totalConstantVolume": 0,
        "usedConstantVolume": 0, "remark": None, "diluteStatus": False,
        "consumableReceive": None, "customType": None, "auditUserName": None,
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
        resp.raise_for_status()
        result = resp.json()
        if not result.get("success"):
            return jsonify({"success": False, "message": result.get('errorDesc') or str(result.get('errorCtx', '配置失败'))})
        return jsonify({"success": True})
    except Exception as e:
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


def _fetch_solution_detail(system, solution_id):
    url = f"{system.base_url}/detectionManager/manager/dtSolutionConfigure/detail"
    headers = {"Referer": f"{system.base_url}/web/solutionConfigure.html?menuId=544"}
    resp = system.session.get(url, params={"id": solution_id}, headers=headers)
    resp.raise_for_status()
    data = resp.json()
    return data.get('resultData') or {}


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
        _enrich_solution_a_source_count(items, system)
        return jsonify({
            "success": True,
            "data": items,
            "total": rd.get('totalCount', 0),
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"查询异常: {str(e)}"}), 500


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
        resp.raise_for_status()
        result = resp.json()
        if not result.get('success'):
            return jsonify({
                "success": False,
                "message": result.get('errorDesc') or str(result.get('errorCtx', '配置失败')),
            })
        return jsonify({"success": True})
    except Exception as e:
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
        resp.raise_for_status()
        result = resp.json()
        if not result.get('success'):
            return jsonify({
                "success": False,
                "message": result.get('errorDesc') or str(result.get('errorCtx', '配置失败')),
            })
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "message": f"配置异常: {str(e)}"}), 500


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

    template_name = 'RF10-10 标准溶液配制记录（稀释）(1).docx'
    template_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), 'word_templates', template_name)
    if not os.path.exists(template_path):
        return jsonify({"success": False, "message": f"模板文件不存在: {template_name}"}), 404

    doc = DocxDocument(template_path)

    def _make_run(text, sz=18, underline=False):
        r_el = OxmlElement('w:r')
        rPr  = OxmlElement('w:rPr')
        if underline:
            u_el = OxmlElement('w:u')
            u_el.set(qn('w:val'), 'single')
            rPr.append(u_el)
        sz_el = OxmlElement('w:sz')
        sz_el.set(qn('w:val'), str(sz))
        szCs_el = OxmlElement('w:szCs')
        szCs_el.set(qn('w:val'), str(sz))
        rPr.append(sz_el)
        rPr.append(szCs_el)
        r_el.append(rPr)
        t_el = OxmlElement('w:t')
        t_el.text = text
        if text and (text[0] == ' ' or text[-1] == ' '):
            t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        r_el.append(t_el)
        return r_el

    def _make_para(text, center=False, sz=18):
        p_el = OxmlElement('w:p')
        pPr  = OxmlElement('w:pPr')
        if center:
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'center')
            pPr.append(jc)
        p_el.append(pPr)
        p_el.append(_make_run(text, sz))
        return p_el

    def _clear_tc(tc):
        for p in tc.findall(qn('w:p')):
            tc.remove(p)

    def _set_tc_text(tc, text, center=True, sz=18):
        _clear_tc(tc)
        tc.append(_make_para(text, center=center, sz=sz))

    def _set_tc_two_lines(tc, line1, line2, sz=18):
        _clear_tc(tc)
        tc.append(_make_para(line1, center=True, sz=sz))
        tc.append(_make_para(line2, center=True, sz=sz))

    def _set_tc_multi_lines(tc, lines, sz=18):
        _clear_tc(tc)
        for line in lines:
            tc.append(_make_para(line, center=True, sz=sz))

    def _set_vmerge(tc, mode):
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr')
            tc.insert(0, tcPr)
        vm = tcPr.find(qn('w:vMerge'))
        if vm is None:
            vm = OxmlElement('w:vMerge')
            tcPr.append(vm)
        if mode == 'restart':
            vm.set(qn('w:val'), 'restart')
        else:
            if qn('w:val') in vm.attrib:
                del vm.attrib[qn('w:val')]

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

    # For working solution export, use source names instead of working solution name
    is_working_doc = bool(detail_list and detail_list[0].get('resultCode'))
    if is_working_doc:
        source_names = []
        for item in detail_list:
            nm = (item.get('originalName', '') or '').strip()
            if nm and nm not in source_names:
                source_names.append(nm)
        header_name = '；\n'.join(source_names) + '；' if source_names else solution_name
    else:
        header_name = solution_name

    _fill_tc1(table.rows[1]._tr.findall(qn('w:tc'))[1], header_name)

    source_codes = []
    for item in detail_list:
        parts = (item.get('originalNo', '') or '').split('\n')
        code = parts[1].strip() if len(parts) > 1 else parts[0].strip()
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

    if not p.get('is_multi_source'):
        src_conc = _parse_conc_val(detail_list[0].get('originalConcentration', ''))
        _fill_tc1(tc5, src_conc)
    else:
        _fill_tc1(tc5, '见下表')

    r3_tcs = table.rows[3]._tr.findall(qn('w:tc'))
    _set_tc_text(r3_tcs[0], f'母体标液({conc_unit})', center=True)
    _set_tc_text(r3_tcs[1], f'取量({qty_unit})', center=True)
    _set_tc_text(r3_tcs[2], '溶剂', center=True)
    _set_tc_text(r3_tcs[3], '稀释至,ml', center=True)
    _set_tc_text(r3_tcs[4], f'浓度({conc_unit})', center=True)
    _set_tc_text(r3_tcs[5], '编号', center=True)
    _set_tc_text(r3_tcs[6], '配制日期', center=True)
    _set_tc_text(r3_tcs[7], '有效期', center=True)

    n_items  = len(detail_list)
    n_tpl    = 12
    remark_row_idx = 16

    if n_items > n_tpl:
        for _ in range(n_items - n_tpl):
            src_tr    = table.rows[remark_row_idx - 1]._tr
            new_tr    = copy.deepcopy(src_tr)
            remark_tr = table.rows[remark_row_idx]._tr
            table._tbl.insert(list(table._tbl).index(remark_tr), new_tr)

    row_values = []
    for item in detail_list:
        item_result_conc = _parse_conc_val(item.get('configurationConcentration', ''))
        item_result_code = item.get('resultCode', '')
        if not item_result_conc:
            item_result_conc = _extract_conc_from_code(solution_code)
        if not item_result_code:
            item_result_code = solution_code
        row_values.append({
            'name':          item.get('originalName', ''),
            'conc_val':      _parse_conc_val(item.get('originalConcentration', '')),
            'qty':           str(item.get('receivedQuantity', '')),
            'medium':        item.get('medium', ''),
            'volume':        str(item.get('volume', '')),
            'result_conc':   item_result_conc,
            'result_code':   item_result_code,
            'dilutionIdx':   item.get('dilutionIdx', 0),
        })

    is_working = bool(detail_list and detail_list[0].get('resultCode'))

    # Detect multi-source: count how many rows belong to dilution point 0
    first_dil_rows = sum(1 for rv in row_values if rv['dilutionIdx'] == 0)
    is_multi_source = is_working and first_dil_rows > 1

    # Column values for data rows (index 2=溶剂, 3=稀释至, 4=浓度, 5=编号, 6=配制日期, 7=有效期)
    col_vals_map = {
        2: [rv['medium']      for rv in row_values],
        3: [rv['volume']      for rv in row_values],
        4: [rv['result_conc'] for rv in row_values],
        5: [rv['result_code'] for rv in row_values],
        6: [configure_date    for _ in row_values],
        7: [validity_date     for _ in row_values],
    }

    # Working solution: only merge 编号 within same dilutionIdx; BC/BCD: merge all from 溶剂 onward
    if is_working and is_multi_source:
        merge_cols = {2, 3, 4, 5, 6, 7}  # all data cols merge within same dilutionIdx
    else:
        merge_cols = {4, 5} if is_working else {2, 3, 4, 5, 6, 7}

    for i, item in enumerate(detail_list):
        tr  = table.rows[4 + i]._tr
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) < 8:
            continue

        # 母体标液
        if is_working:
            if is_multi_source and row_values[i]['dilutionIdx'] == 0:
                _set_tc_two_lines(tcs[0], row_values[i]['name'], row_values[i]['conc_val'])
            else:
                _set_tc_text(tcs[0], row_values[i]['conc_val'])
        else:
            _set_tc_two_lines(tcs[0], row_values[i]['name'], row_values[i]['conc_val'])

        # 取量
        _set_tc_text(tcs[1], row_values[i]['qty'])

        # Columns 2-7
        for tc_idx in range(2, 8):
            val = col_vals_map[tc_idx][i]
            same_dilution = i > 0 and row_values[i]['dilutionIdx'] == row_values[i - 1]['dilutionIdx']
            if tc_idx in merge_cols and same_dilution and val == col_vals_map[tc_idx][i - 1]:
                _clear_tc(tcs[tc_idx])
                tcs[tc_idx].append(_make_para('', center=True))
                _set_vmerge(tcs[tc_idx], 'continue')
            else:
                _set_tc_text(tcs[tc_idx], val)
                if tc_idx in merge_cols:
                    _set_vmerge(tcs[tc_idx], 'restart')

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


if __name__ == '__main__':
    if not os.path.exists('templates'):
        os.makedirs('templates')
    print("Flask 服务启动，访问以下地址：")
    print("  登录页面: http://127.0.0.1:5000/login")
    print("  耗材查询主页: http://127.0.0.1:5000/")
    print("  有机标准品管理: http://127.0.0.1:5000/organic-std")
    app.run(host='0.0.0.0', port=5000, debug=True)