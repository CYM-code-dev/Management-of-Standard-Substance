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
from urllib.parse import quote
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


def connect_wifi(ssid, wifi_password):
    """导入 WLAN profile 后发起连接;未保存过该 WiFi 也能连。"""
    if wifi_connected(ssid):
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
        subprocess.run(
            ["netsh", "wlan", "connect", f"name={ssid}"],
            capture_output=True, timeout=15,
        )
        log.info("已发起 WiFi 连接 %s", ssid)
    finally:
        try:
            os.remove(profile_xml)
        except OSError:
            pass


def login(session, cfg):
    """发登录请求:username GB2312 编码,password = AES 加密后的明文密码。"""
    u = gb2312_quote(cfg["username"])
    p = pa_aes_encode(cfg["password"])
    q = (
        "route=webauth&action=user_login&auth_type=panabit"
        f"&ip=&mac=&code=&username={u}&password={p}&remember_me=1"
    )
    r = session.post(f"http://{cfg['portal_host']}/api?{q}", timeout=10)
    try:
        body = r.json()
        code, msg = body.get("code"), body.get("msg", "")
    except ValueError:
        code, msg = None, r.text[:160]
    if code == 0:
        log.info("认证成功 (code:0)")
    else:
        # code:1 多半是"当前已在线、无可认证会话",非真实失败;仅当 captive 态仍 code:1 才需排查
        log.warning("认证未通过 code:%s msg:%s —— 若已在线属正常", code, msg)


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
                log.info("检测到离线,尝试恢复")
                if cfg["ssid"] and not wifi_in_range(cfg["ssid"]):
                    log.info("%s 不在范围内,等信号恢复", cfg["ssid"])
                else:
                    if cfg["ssid"]:
                        connect_wifi(cfg["ssid"], cfg["wifi_password"])
                        time.sleep(5)
                    login(session, cfg)
                    time.sleep(3)
                    online2 = is_online(session, cfg["probe_url"])
                    log.info("登录后 %s", "在线" if online2 else "仍离线")
        except Exception as e:
            log.error("本轮异常: %s", e)  # 吞掉,下一轮继续;nssm 兜底重启
        time.sleep(cfg["interval"])


if __name__ == "__main__":
    main()
