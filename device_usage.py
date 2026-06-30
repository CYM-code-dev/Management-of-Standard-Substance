# -*- coding: utf-8 -*-
"""仪器使用率上报（HTA 仪器使用率统计.hta 提交）。

作为独立 Blueprint 供 login_html.py 注册使用；接口开放不认证（内网工具）。
存储：device_usage.csv（utf-8-sig 带 BOM，Excel 直接打开中文不乱码）。
"""
import os
import csv
import threading
import datetime

from flask import Blueprint, request, jsonify

bp = Blueprint('device_usage', __name__)

DEVICE_USAGE_FILE = 'device_usage.csv'
_DEVICE_USAGE_COLS = ['device_id', 'period_start', 'period_end', 'usage_hours',
                      'workdays', 'rate', 'submitted_at']
_device_usage_lock = threading.Lock()


def _load_device_usage():
    """读 device_usage.csv -> list[dict]，数值字段还原为数字。"""
    if not os.path.exists(DEVICE_USAGE_FILE):
        return []
    try:
        with open(DEVICE_USAGE_FILE, 'r', encoding='utf-8-sig', newline='') as f:
            rows = list(csv.DictReader(f))
    except Exception:
        return []
    for r in rows:
        try:
            r['workdays'] = int(float(r.get('workdays') or 0))
        except (ValueError, TypeError):
            r['workdays'] = 0
        try:
            r['rate'] = float(r.get('rate') or 0.0)
        except (ValueError, TypeError):
            r['rate'] = 0.0
    return rows


def _save_device_usage(records):
    """原子写 CSV（utf-8-sig 带 BOM，Excel 直接打开中文不乱码）。"""
    tmp = DEVICE_USAGE_FILE + '.tmp'
    with open(tmp, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(_DEVICE_USAGE_COLS)
        for r in records:
            w.writerow([r.get(c, '') for c in _DEVICE_USAGE_COLS])
    os.replace(tmp, DEVICE_USAGE_FILE)


@bp.route('/api/device/usage/submit', methods=['POST'])
def device_usage_submit():
    p = request.get_json(silent=True) or {}
    device_id = (p.get('device_id') or '').strip()
    period_start = (p.get('period_start') or '').strip()
    period_end = (p.get('period_end') or '').strip()
    if not device_id or not period_start or not period_end:
        return jsonify({"success": False, "message": "缺少 device_id/period_start/period_end"}), 400
    try:
        rate = float(p.get('rate'))
    except (TypeError, ValueError):
        return jsonify({"success": False, "message": "rate 非数值"}), 400
    try:
        usage_sec = float(p.get('usage_sec') or 0)
    except (TypeError, ValueError):
        usage_sec = 0.0
    try:
        maint_sec = float(p.get('maint_sec') or 0)
    except (TypeError, ValueError):
        maint_sec = 0.0
    # usage_hours = 使用时长+维护时长，单位 h，两位小数（A+B 形式）
    usage_hours = "{:.2f}+{:.2f}".format(usage_sec / 3600, maint_sec / 3600)

    # 字段顺序固定：device_id 在 rate 前（HTA 端用正则解析 list 响应依赖此顺序）
    rec = {
        "device_id": device_id,
        "period_start": period_start,
        "period_end": period_end,
        "usage_hours": usage_hours,
        "workdays": int(p.get('workdays') or 0),
        "rate": rate,
        "submitted_at": datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }
    with _device_usage_lock:
        records = _load_device_usage()
        idx = next((i for i, r in enumerate(records)
                    if r.get('device_id') == device_id
                    and r.get('period_start') == period_start
                    and r.get('period_end') == period_end), None)
        if idx is None:
            records.append(rec)
        else:
            records[idx] = rec
        _save_device_usage(records)
    return jsonify({"success": True, "message": "已记录"})


@bp.route('/api/device/usage/list', methods=['GET'])
def device_usage_list():
    period_start = (request.args.get('period_start') or '').strip()
    period_end = (request.args.get('period_end') or '').strip()
    records = _load_device_usage()
    if period_start and period_end:
        records = [r for r in records
                   if r.get('period_start') == period_start and r.get('period_end') == period_end]
    out = [{
        "device_id": r.get('device_id', ''),
        "rate": r.get('rate', 0.0),
        "usage_hours": r.get('usage_hours', ''),
        "workdays": r.get('workdays', 0),
        "submitted_at": r.get('submitted_at', ''),
    } for r in records]
    return jsonify({"success": True, "period_start": period_start,
                    "period_end": period_end, "records": out})


@bp.route('/api/device/usage/delete', methods=['POST'])
def device_usage_delete():
    p = request.get_json(silent=True) or {}
    period_start = (p.get('period_start') or '').strip()
    period_end = (p.get('period_end') or '').strip()
    device_ids = p.get('device_ids') or []
    if not isinstance(device_ids, list) or not device_ids or not period_start or not period_end:
        return jsonify({"success": False, "message": "参数不足"}), 400
    ids = set(str(x) for x in device_ids)
    with _device_usage_lock:
        records = _load_device_usage()
        before = len(records)
        records = [r for r in records if not (r.get('device_id') in ids
                                              and r.get('period_start') == period_start
                                              and r.get('period_end') == period_end)]
        _save_device_usage(records)
    return jsonify({"success": True, "removed": before - len(records)})
