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
import sys

import ddddocr

# 让脚本能从项目根 import lims_auto_login
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import lims_auto_login

BASE = lims_auto_login.BASE_URL
OUT_DIR = os.path.join(os.path.dirname(__file__), "_captcha_samples")


def run_recognize_only(n):
    os.makedirs(OUT_DIR, exist_ok=True)
    session = lims_auto_login.new_session()
    ocr = ddddocr.DdddOcr(show_ad=False)
    manifest = []
    for i in range(n):
        try:
            img = lims_auto_login.fetch_captcha(session)
            if img is None:
                raise RuntimeError("抓取返回空")
        except Exception as e:
            print(f"[{i}] 抓取失败: {e}")
            continue
        text, digits, total = lims_auto_login.solve_captcha(ocr, img)
        path = os.path.join(OUT_DIR, f"cap_{i:02d}.jpg")
        with open(path, "wb") as f:
            f.write(img)
        manifest.append({"i": i, "raw": text, "digits": digits, "sum": total, "file": path})
        print(f"[{i:02d}] raw={text!r:20} digits={digits} sum={total}")
    with open(os.path.join(OUT_DIR, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print(f"\n保存 {len(manifest)} 张到 {OUT_DIR}（含 manifest.json，可人工核对）")


def run_login(n, username, password):
    # 复用 lims_auto_login 的 session/验证码/登录，不再依赖 login_html（无 Flask 副作用）
    session = lims_auto_login.new_session()
    ocr = ddddocr.DdddOcr(show_ad=False)
    ok = 0
    for i in range(n):
        img = lims_auto_login.fetch_captcha(session)
        if img is None:
            print(f"[{i}] 取验证码失败")
            continue
        text, digits, total = lims_auto_login.solve_captcha(ocr, img)
        if total is None:
            print(f"[{i:02d}] raw={text!r:20} sum=None -> FAIL: 识别不到两位数字")
            continue
        success, msg = lims_auto_login.login(session, username, password, str(total))
        print(f"[{i:02d}] raw={text!r:20} sum={total} -> {'OK' if success else 'FAIL: ' + msg}")
        if success:
            ok += 1
            # 登录成功后 cookie 已绑死该 session；后续轮次需新 session 才能再登录同账号
            session = lims_auto_login.new_session()
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
