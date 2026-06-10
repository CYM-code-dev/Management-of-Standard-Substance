@echo off
chcp 65001 >nul
echo 正在初始化 LIMS 接口扫描器...
echo.

python scripts\setup_scanner.py

echo.
echo 完成！git add . 后 commit push 即可。
pause
