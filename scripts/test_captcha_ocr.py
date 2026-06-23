# -*- coding: utf-8 -*-
"""Phase 0: 验证码 OCR 准确率测试。

LIMS 登录验证码是算术题 `A + B = ?`，答案为 A+B 的和（0~18）。
本脚本验证 ddddocr 能否读出两位数字并正确求和。

两种模式：
  --recognize-only   仅识别（无需账号密码）：抓验证码→ddddocr→提取数字求和，
                     保存图片+识别结果到 manifest，供人工/视觉核对。← 首选用它评估准确率
  （默认）登录模式  需账号密码：识别求和后尝试登录，统计登录成功率（金标准）。

用法示例：
  .venv/Scripts/python scripts/test_captcha_ocr.py --recognize-only -n 12
  .venv/Scripts/python scripts/test_captcha_ocr.py -n 20 -u <账号> -p <密码>
"""
import argparse
import json
import os
import re
import time
from io import BytesIO

import requests
import ddddocr

BASE = "http://192.168.12.234:60015"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36",
    "Accept": "image/webp,image/apng,image/*,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Referer": f"{BASE}/web/login.html",
}
OUT_DIR = os.path.join(os.path.dirname(__file__), "_captcha_samples")


def fetch_captcha(session):
    """复用 login_html.RemoteSystem.get_captcha_image 的请求方式，返回图片字节。"""
    ts = str(int(time.time() * 1000))
    url = f"{BASE}/detectionManager/core/security/validatecodes?{ts}&r={ts}"
    resp = session.get(url, headers=HEADERS, timeout=10)
    resp.raise_for_status()
    return resp.content


def ocr_solve(ocr, img_bytes):
    """算术验证码 A+B=?：ddddocr 能读对两个操作数，但尾部 '=?' 常被误读为
    数字/符号（如 '2+1-9'、'9+89'、'7+7>'）。故只取**前两位数字字符**作为 A、B 求和，
    天然屏蔽尾部杂讯。返回 (原始识别文本, [A,B], A+B 或 None)。"""
    text = ocr.classification(img_bytes)
    digits = [int(d) for d in re.findall(r"\d", text)]  # 逐个数字字符
    if len(digits) < 2:
        return text, digits, None
    ab = digits[:2]
    return text, ab, ab[0] + ab[1]


def run_recognize_only(n):
    os.makedirs(OUT_DIR, exist_ok=True)
    session = requests.Session()
    session.verify = False
    ocr = ddddocr.DdddOcr(show_ad=False)
    manifest = []
    for i in range(n):
        try:
            img = fetch_captcha(session)
        except Exception as e:
            print(f"[{i}] 抓取失败: {e}")
            continue
        text, digits, total = ocr_solve(ocr, img)
        path = os.path.join(OUT_DIR, f"cap_{i:02d}.jpg")
        with open(path, "wb") as f:
            f.write(img)
        manifest.append({"i": i, "raw": text, "digits": digits, "sum": total, "file": path})
        print(f"[{i:02d}] raw={text!r:20} digits={digits} sum={total}")
    with open(os.path.join(OUT_DIR, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print(f"\n保存 {len(manifest)} 张到 {OUT_DIR}（含 manifest.json，可人工核对）")


def run_login(n, username, password):
    # 懒加载：登录模式才 import login_html（避免识别模式触发 Flask 模块副作用）
    from login_html import RemoteSystem  # noqa
    sys_ = RemoteSystem("ocr_test")
    ocr = ddddocr.DdddOcr(show_ad=False)
    ok = 0
    for i in range(n):
        img = sys_.get_captcha_image()
        if img is None:
            print(f"[{i}] 取验证码失败")
            continue
        buf = BytesIO()
        img.save(buf, format="JPEG")
        text, digits, total = ocr_solve(ocr, buf.getvalue())
        success, msg = sys_.login(username, password, str(total))
        print(f"[{i:02d}] raw={text!r:20} sum={total} -> {'OK' if success else 'FAIL: ' + msg}")
        if success:
            ok += 1
    print(f"\n登录成功率: {ok}/{n} = {ok / n * 100:.1f}%")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--recognize-only", action="store_true", help="仅识别，不登录（无需账号密码）")
    ap.add_argument("-n", type=int, default=12, help="测试次数")
    ap.add_argument("-u", dest="username", help="LIMS 账号（登录模式）")
    ap.add_argument("-p", dest="password", help="LIMS 密码（登录模式）")
    args = ap.parse_args()
    if args.recognize_only:
        run_recognize_only(args.n)
    else:
        if not args.username or not args.password:
            ap.error("登录模式需要 -u 账号 -p 密码（或用 --recognize-only 免登录评估）")
        run_login(args.n, args.username, args.password)


if __name__ == "__main__":
    main()
