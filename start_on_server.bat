@echo off
net use \\files.cirs-ck.com /delete >nul 2>&1
net use \\files.cirs-ck.com /user:"轻工公盘管理" "83Zp3Rh9OqNh" /persistent:yes
cd /d "E:\Management-of-Standard-Substance"
.venv\Scripts\python login_html.py
