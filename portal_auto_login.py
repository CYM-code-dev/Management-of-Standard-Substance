# -*- coding: utf-8 -*-
"""
Panabit 认证网关自动登录(nssm 长驻服务)。

定时探测外网连通性;断了就连 WiFi + 回放登录请求,让认证无人值守。
配置: 同目录 portal_config.ini
部署: nssm install PortalAutoLogin <.venv python.exe> <本文件绝对路径>
"""
import configparser
import logging
import os
import subprocess
import sys
import tempfile
import time
from os.path import dirname, join
from urllib.parse import parse_qsl, quote, urlparse
from xml.sax.saxutils import escape as xml_escape

import requests
from Crypto.Cipher import AES

BASE_DIR = dirname(__file__)
CONFIG_PATH = join(BASE_DIR, "portal_config.ini")
LOG_PATH = join(BASE_DIR, "portal_auto_login.log")

ONLINE_MARKER = "Microsoft Connect Test"

# Panabit 认证页密码加密:AES-128-ECB + ZeroPadding,密钥写死在前端 crypto.js。
# 逆向自 portal.ck.cirs.group/assert/crypto.js 的 pa_aes_encode();密钥 Panabit@1024_key。
PORTAL_AES_KEY = b"Panabit@1024_key"


def pa_aes_encode(plaintext):
    """复刻前端 pa_aes_encode:明文按 ZeroPadding 补 0 到整块,AES-128-ECB,输出 hex。"""
    data = plaintext.encode("utf-8")
    pad = (-len(data)) % 16 or 16  # CryptoJS ZeroPadding:已对齐则补一整块
    data += b"\x00" * pad
    return AES.new(PORTAL_AES_KEY, AES.MODE_ECB).encrypt(data).hex()

_handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
_console = logging.StreamHandler()
_console.setFormatter(logging.Formatter("%(asctime)s %(message)s"))
logging.basicConfig(level=logging.INFO, handlers=[_handler, _console])
log = logging.getLogger("portal")

# WLAN profile 模板:WPA2/AES,大多数企业/家用 WPA 网络适用。
# ponytail: 若某网络是 WPA/TKIP,改下面两个字段即可,不做配置化。
WLAN_PROFILE = """<?xml version="1.0" encoding="UTF-8"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
  <name>{ssid}</name>
  <SSIDConfig><SSID><name>{ssid}</name></SSID></SSIDConfig>
  <connectionType>ESS</connectionType>
  <connectionMode>auto</connectionMode>
  <MSM>
    <security>
      <authEncryption><authentication>WPA2PSK</authentication><encryption>AES</encryption><useOneX>false</useOneX></authEncryption>
      <sharedKey><keyType>passPhrase</keyType><protected>false</protected><keyMaterial>{pwd}</keyMaterial></sharedKey>
    </security>
  </MSM>
</WLANProfile>
"""


def load_config():
    cfg = configparser.ConfigParser()
    if not cfg.read(CONFIG_PATH, encoding="utf-8"):
        raise RuntimeError(f"找不到配置文件 {CONFIG_PATH}")
    p = cfg["portal"]
    username = p.get("username", "").strip()
    password = p.get("password", "").strip()
    if not username or not password:
        raise RuntimeError("portal_config.ini 缺 username/password")
    return {
        "username": username,
        "password": password,
        "ssid": p.get("ssid", "").strip(),
        "wifi_password": p.get("wifi_password", "").strip(),
        "interval": p.getint("probe_interval", 60),
        "portal_host": p.get("portal_host", "portal.ck.cirs.group").strip(),
        "probe_url": p.get("probe_url", "http://www.msftconnecttest.com/connecttest.txt").strip(),
    }


def is_online(session, probe_url):
    """真连外网 → connecttest.txt 正文含 ONLINE_MARKER;被 portal 拦截/超时 → False。"""
    try:
        r = session.get(probe_url, timeout=8)
        return ONLINE_MARKER in r.text
    except requests.RequestException:
        return False


def gb2312_quote(name):
    """用户名按 GB2312 百分号编码(抓包里 '崔艳梅' → %B4%DE%D1%DE%C3%B7)。"""
    try:
        return quote(name.encode("gb2312"))
    except UnicodeEncodeError:
        return quote(name.encode("gb18030"))  # 超集兜底


def wifi_connected(ssid):
    try:
        r = subprocess.run(
            ["netsh", "wlan", "show", "interfaces"],
            capture_output=True, timeout=10,
        )
        out = (r.stdout or b"").decode("ascii", errors="ignore")
        return ssid in out
    except subprocess.SubprocessError:
        return False


def wifi_in_range(ssid):
    """CIRS-CK 是否在网卡可见范围内(含已连接)。查询失败时保守放行,沿用尽力尝试。"""
    try:
        r = subprocess.run(
            ["netsh", "wlan", "show", "networks"],
            capture_output=True, timeout=10,
        )
        return ssid in (r.stdout or b"").decode("ascii", errors="ignore")
    except subprocess.SubprocessError:
        return True


def connect_wifi(ssid, wifi_password, force=False):
    """导入 WLAN profile 后发起连接;未保存过该 WiFi 也能连。
    force=True 时无视“已连接”状态强制重连:断开/唤醒后网卡常报 connected 但链路已死,
    不重连会导致 portal 登录一直 code:1(探测离线才能确认链路已失效)。"""
    if not force and wifi_connected(ssid):
        log.info("WiFi 已连接 %s", ssid)
        return
    fd, profile_xml = tempfile.mkstemp(suffix=".xml")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(WLAN_PROFILE.format(ssid=xml_escape(ssid), pwd=xml_escape(wifi_password)))
        subprocess.run(
            ["netsh", "wlan", "add", "profile", f"filename={profile_xml}", "user=all"],
            capture_output=True, timeout=15,
        )
        if force:
            # 先断开 stale 关联,否则 connect 对“已连接”状态是空操作
            subprocess.run(
                ["netsh", "wlan", "disconnect"],
                capture_output=True, timeout=10,
            )
            time.sleep(1)
        subprocess.run(
            ["netsh", "wlan", "connect", f"name={ssid}"],
            capture_output=True, timeout=15,
        )
        log.info("已发起 WiFi 连接 %s%s", ssid, "（强制重连）" if force else "")
    finally:
        try:
            os.remove(profile_xml)
        except OSError:
            pass


def login(session, cfg):
    """触发 portal 重定向拿真实 ip/mac,再用同一会话 POST /api 登录。
    关键:capture 与 POST 共用 requests.Session,带上重定向下发的会话 cookie,
    否则 portal 报 code:1 “请求失败,请刷新页面后重新尝试”。
    username GB2312 编码,password/code = AES 加密后明文。"""
    ip = mac = ""
    try:
        r = session.get(cfg["probe_url"], timeout=8)  # 跟随重定向,cookie 自动入 session
    except requests.RequestException:
        r = None
    if r is not None:
        # ip/mac 出现在 302 的 Location 或落地页 URL 里;逐一翻找
        for url in [h.headers.get("Location", "") for h in r.history] + [r.url]:
            if "wlanuserip" in url or "clientmac" in url:
                q = dict(parse_qsl(urlparse(url).query))
                ip = ip or q.get("wlanuserip", "")
                mac = mac or q.get("clientmac", "")
    u = gb2312_quote(cfg["username"])
    p = pa_aes_encode(cfg["password"])
    code = ""  # 无图形验证码:抓包确认网页传空 code;传 AES("") 会被当作非法验证码 → code:1 请求失败
    q = ("route=webauth&action=user_login&auth_type=panabit"
         f"&ip={ip}&mac={mac}&code={code}&username={u}&password={p}&remember_me=1")
    r = session.post(f"http://{cfg['portal_host']}/api?{q}", timeout=10)
    try:
        body = r.json()
        code, msg = body.get("code"), body.get("msg", "")
    except ValueError:
        code, msg = None, r.text[:160]
    if code == 0:
        log.info("认证成功 (code:0)")
    else:
        # code:1 多半是“当前已在线、无可认证会话”,非真实失败;仅当 captive 态仍 code:1 才需排查
        log.warning("认证未通过 code:%s msg:%s —— 若已在线属正常", code, msg)
    return code


RECOVER_TRIES = 8   # 离线后重试连接+登录的次数;每次间隔 ~8s
DNS_WAIT = 5        # WiFi“已连接”后等 DHCP/DNS 起步的秒数
POST_AUTH_WAIT = 75  # 认证成功(code:0)后等网络真正联通的秒数(经验值约1分钟,留余量)
POST_AUTH_POLL = 5  # 认证后轮询在线的间隔


def _wait_online(cfg, session, budget):
    """轮询到在线或超时;认证已通过,期间不再重复登录。"""
    deadline = time.monotonic() + budget
    n = 0
    while time.monotonic() < deadline:
        n += 1
        if is_online(session, cfg["probe_url"]):
            log.info("等待第%d次探测:在线", n)
            return True
        time.sleep(POST_AUTH_POLL)
    return False


def recover(cfg, session):
    """原地重试到认证通过;开机/唤醒时网卡可能还没扫到 SSID,先等它就绪再连。
    认证成功(code:0)后网络真正联通一般还要约1分钟,此时停止重复登录,
    只轮询到在线(POST_AUTH_WAIT),避免在联通前把 8 次重试耗光而误判失败。"""
    log.info("检测到离线,尝试恢复")
    connected = False
    for i in range(1, RECOVER_TRIES + 1):
        if is_online(session, cfg["probe_url"]):
            log.info("第%d次探测:在线", i)
            return
        if cfg["ssid"] and not wifi_in_range(cfg["ssid"]):
            log.info("第%d次:%s 不在范围内,等网卡就绪", i, cfg["ssid"])
            time.sleep(DNS_WAIT)
            continue
        if cfg["ssid"] and not connected:
            connect_wifi(cfg["ssid"], cfg["wifi_password"], force=True)
            connected = True
        time.sleep(DNS_WAIT)
        code = login(session, cfg)
        if code == 0:
            log.info("已认证(code:0),等待网络联通(约1分钟)")
            if _wait_online(cfg, session, POST_AUTH_WAIT):
                return
            log.warning("等待 %ds 后仍离线,等下一轮", POST_AUTH_WAIT)
            return
        time.sleep(3)
    log.warning("重试 %d 次后仍未认证,等下一轮", RECOVER_TRIES)


def main():
    try:
        cfg = load_config()
    except Exception as e:
        log.error("配置加载失败,退出: %s", e)
        sys.exit(1)

    session = requests.Session()
    log.info("启动:每 %ss 探测一次", cfg["interval"])
    while True:
        try:
            if is_online(session, cfg["probe_url"]):
                log.info("在线,跳过")
            else:
                recover(cfg, session)
        except Exception as e:
            log.error("本轮异常: %s", e)  # 吞掉,下一轮继续;nssm 兜底重启
        time.sleep(cfg["interval"])


if __name__ == "__main__":
    main()
