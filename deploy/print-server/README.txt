Niimbot 标签打印服务 部署说明
================================

目标电脑：实验室连打印机的电脑（IP 以运行部署脚本时显示的本机 IP 为准）

部署步骤（只做一次）
--------------------
1. 把整个 print-server 文件夹（含 node_modules，约 40MB）拷到该电脑，如 C:\print-server
   （依赖已打好包，拷过去无需联网；若提示缺依赖才需要联网 npm install）
2. 右键 deploy_print_server.bat →「以管理员身份运行」
   脚本会自动：装防火墙规则(5001)、设开机自启、启动服务、本机自测
3. 若提示缺 Node.js：按弹出的网页装 LTS 版（一路下一步），再重新跑一遍脚本
4. 打印机与电脑【蓝牙配对】：Windows 设置 → 蓝牙和其他设备 → 添加设备，
   配对 Niimbot 打印机（如 B1）；配对成功系统会自动生成串口（设备管理器→端口 COMx），
   首次点打印时服务器会自动扫描并连接。注意：打印机是蓝牙连接，不是 USB！

日常使用（后台服务模式）
------------------------
- 打印服务以 Windows 计划任务 NiimbotPrint 在后台运行（SYSTEM 账户，无窗口、无需登录，
  开机自动启动、崩溃自动重启；已取消 Windows 默认的 72 小时运行上限）
- 看门狗任务 NiimbotPrintWatchdog 每 5 分钟检查一次 localhost:5001，服务挂掉/假死时
  自动杀掉并拉起，通常 5 分钟内自愈，无需人工干预
- 电脑上不需要保持任何黑窗口，关掉也不影响打印
- 查看运行状态：schtasks /Query /TN NiimbotPrint
- 手动停止：右键管理员运行 stop_print_server.bat（同时停看门狗，否则 5 分钟后会自动复活；
  下次开机仍会自动启动）
- 手动启动：右键管理员运行 start_print_server.bat（无窗口后台启动，并重新启用看门狗）
- 打印机断电/拔线后再点一次打印即可，服务会自动重扫串口重连

取消开机自启（如需）
--------------------
schtasks /Delete /TN NiimbotPrint /F
schtasks /Delete /TN NiimbotPrintWatchdog /F
netsh advfirewall firewall delete rule name="NiimbotPrint5001"

在公司服务器上验证连通
----------------------
curl http://<打印机电脑当前IP>:5001/connected
（IP 见部署脚本结尾显示的 This PC IP；电脑 IP 变了就填新 IP）
返回 JSON（connected true/false 均可）即通；连不上=服务没起或防火墙未放行
