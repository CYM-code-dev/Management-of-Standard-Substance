# -*- coding: utf-8 -*-
"""LIMS OCR 自动登录（独立可复用模块，零 Flask 依赖）。

封装「抓验证码 → ddddocr 识别 → MD5 加密 → 登录 → 取用户信息」全流程，
返回一个已登录、cookie 已就绪的 `requests.Session`（包在 `LoginResult` 里）。
任何外部程序 `import lims_auto_login` 即可使用，不触发 login_html 的 Flask 副作用。

验证码是算术题 `A + B = ?`（0~9，仅加法）：ddddocr 能读对两个操作数，但尾部 '=?'
常被误读为数字/符号（如 '2+1-9'、'9+89'），故只取**前两位数字字符**作为 A、B 求和，
天然屏蔽尾部杂讯。

典型用法：
    from lims_auto_login import auto_login
    res = auto_login("cym", "Cui12345")
    if res:
        resp = res.session.get(f"{res.base_url}/detectionManager/core/users/info")

设计说明：
- 仅依赖 requests + Pillow + ddddocr；ddddocr 懒加载（仅 auto_login 内 import）。
- `LoginResult` 字段镜像 login_html.RemoteSystem 的鸭子接口（session/base_url/current_user/
  current_pid/current_real_name），可作为其 OCR 登录路径的等价替代。
- 不含 session 文件持久化（save/load/verify）——那是 RemoteSystem 的另一职责。
"""
import hashlib
import re
import time
from dataclasses import dataclass
from typing import List, Optional, Tuple

import requests

BASE_URL = "http://192.168.12.234:60015"

_DEFAULT_UA = (
    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36"
)


@dataclass
class LoginResult:
    """auto_login 的返回值；字段与 login_html.RemoteSystem 鸭子兼容。"""
    session: requests.Session
    base_url: str
    current_user: str
    current_pid: Optional[str]
    current_real_name: Optional[str]


def md5_1024_times(text: str) -> str:
    """LIMS 密码加密：MD5 迭代 1024 次。"""
    current = text.encode("utf-8")
    for _ in range(1024):
        current = hashlib.md5(current).hexdigest().encode("utf-8")
    return current.decode()


def new_session(base_url: str = BASE_URL) -> requests.Session:
    """建一个带默认 headers、关闭 SSL 校验的 requests.Session。"""
    session = requests.Session()
    session.verify = False
    session.headers.update({
        "User-Agent": _DEFAULT_UA,
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": base_url,
        "Referer": f"{base_url}/web/login.html",
        "Connection": "keep-alive",
        "Host": "192.168.12.234:60015",
    })
    return session


def fetch_captcha(session: requests.Session, base_url: str = BASE_URL,
                  timeout: int = 10) -> Optional[bytes]:
    """抓取登录验证码图片，返回原始字节；失败返回 None。

    注意：LIMS 的验证码与登录请求需共用同一 session（cookie 绑定），故直接在传入的
    session 上请求。
    """
    ts = str(int(time.time() * 1000))
    url = f"{base_url}/detectionManager/core/security/validatecodes?{ts}&r={ts}"
    try:
        resp = session.get(url, timeout=timeout)
        if resp.status_code == 200:
            return resp.content
        return None
    except Exception as e:
        print(f"获取验证码失败: {e}")
        return None


def solve_captcha(ocr, img_bytes: bytes) -> Tuple[str, List[int], Optional[int]]:
    """算术验证码 A+B=?：ddddocr 识别后取**前两位数字字符**求和，屏蔽尾部 '=?' 误读。

    返回 (原始识别文本, 逐位数字列表, A+B 或 None)。
    """
    text = ocr.classification(img_bytes)
    digits = [int(d) for d in re.findall(r"\d", text)]
    if len(digits) < 2:
        return text, digits, None
    ab = digits[:2]
    return text, ab, ab[0] + ab[1]


def login(session: requests.Session, username: str, password: str, captcha: str,
          base_url: str = BASE_URL, timeout: int = 15) -> Tuple[bool, str]:
    """用账号/明文密码/验证码登录；成功时 cookie 已落入 session。

    返回 (是否成功, 消息)。
    """
    login_data = {
        "account": username,
        "password": md5_1024_times(password),
        "validCode": captcha,
    }
    try:
        resp = session.post(
            f"{base_url}/detectionManager/core/security/login",
            data=login_data, timeout=timeout,
        )
        if resp.status_code != 200:
            return False, f"请求失败，状态码: {resp.status_code}"
        result = resp.json()
        if result.get("success"):
            return True, result.get("resultData", {}).get("nickName", "") or "登录成功"
        error_msg = result.get("errorCtx", {}).get("errorMsg", "登录失败")
        return False, "验证码错误" if "验证码" in error_msg else error_msg
    except Exception as e:
        return False, f"登录异常: {e}"


def fetch_user_info(session: requests.Session, base_url: str = BASE_URL,
                    timeout: int = 10) -> Tuple[Optional[str], Optional[str]]:
    """登录后探测用户信息，返回 (pid, real_name)；取不到对应项为 None。

    /getLoginUser 在部分 LIMS 上不可用（HTTP 500「资源不存在」），故以 /users/info 为主探针，
    两端点任一返回 success 即采用。
    """
    for ep in ("/detectionManager/core/users/info",
               "/detectionManager/core/security/getLoginUser"):
        try:
            resp = session.get(f"{base_url}{ep}", timeout=timeout)
            if resp.status_code != 200:
                continue
            data = resp.json()
            if data.get("success"):
                user_info = data.get("resultData", {}).get("userInfo", {}) or {}
                pid = user_info.get("id")
                real_name = user_info.get("realName")
                if pid or real_name:
                    return (str(pid) if pid is not None else None), real_name
        except Exception:
            continue
    return None, None


def auto_login(username: str, password: str, max_retry: int = 3,
               base_url: str = BASE_URL, ocr=None) -> Optional[LoginResult]:
    """OCR 自动登录：循环 抓验证码→识别→登录，最多 max_retry 次。

    成功返回已登录的 LoginResult；全失败返回 None。
    `ocr` 可传入已实例化的 ddddocr.DdddOcr 以复用；默认懒加载实例化。
    """
    import ddddocr  # 懒加载，避免 import 本模块就拖起 ddddocr

    if ocr is None:
        ocr = ddddocr.DdddOcr(show_ad=False)

    session = new_session(base_url)
    for _ in range(max_retry):
        img = fetch_captcha(session, base_url)
        if img is None:
            time.sleep(1)
            continue
        text, _digits, total = solve_captcha(ocr, img)
        if total is None:
            time.sleep(1)
            continue
        ok, msg = login(session, username, password, str(total), base_url)
        if ok:
            pid, real_name = fetch_user_info(session, base_url)
            if not real_name:
                return None
            return LoginResult(
                session=session, base_url=base_url, current_user=username,
                current_pid=pid, current_real_name=real_name,
            )
    return None
