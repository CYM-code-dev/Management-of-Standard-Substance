# 标准溶液标签打印功能 — WiFi 标签打印机方案

## Context

储备液(B)、应用液(C)、工作液(D)配置完成后，需要打印标签贴在试剂瓶上。
打印机通过 WiFi 连接，与 Flask 服务器在同一局域网。用户点击按钮后打印机自动出纸，无需人工操作。

## 架构

```
浏览器点击"打印标签" → POST /api/lims/print_label → Flask 后端生成 TSPL → TCP Socket → 打印机
```

## 标签内容

溶液编号、溶液名称、介质/浓度、配制人/配制日期、有效期至、储存条件。

### 编号规则

- **储备液(B)、应用液(C)**：直接使用页面 `#bbcdSolutionCode` 的值（自编号）
- **工作液(D)**：每个稀释点一个标签，编号格式 `{前缀}-{单点浓度}-{日期}`，如 `CK-CG-2026068-108.75-20260518`。前缀从 `#bbcdSolutionCode` 中提取（去掉浓度范围和日期部分），浓度从 `.bbcd-d-conc[data-dilution="${i}"]` 读取，日期取配制日期

## 前端（`templates/SolutionConfig.html`）

**新增"打印标签"按钮和 `bbcdPrintLabel()` 函数**

- 按钮放在导出按钮旁边
- 收集页面上的配置数据，POST 到 `/api/lims/print_label`
- 工作液(D)：每个稀释点单独发一次请求

## 后端（`login_html.py`）

**新增 `/api/lims/print_label` 接口**

1. 接收 JSON：`{ solution_name, solution_code, concentration, configure_date, validity_date, storage_condition, medium, creator_name }`
2. 生成 TSPL 指令（Thermal Printer Specification Language）：
   - `SIZE` 设置标签尺寸（默认 40×20mm）
   - `GAP` 设置标签间距
   - `TEXT` 绘制标签内容（名称、编号、浓度、日期等）
   - `PRINT` 触发打印
3. 通过 TCP Socket 连接打印机 `IP:9100`（9100 是标准打印端口）
4. 发送 TSPL 指令字节流
5. 关闭连接

**打印机配置**

- 打印机 IP 地址存放在配置文件或数据库中
- 提供一个配置页面或接口可修改打印机 IP

## TSPL 指令示例（40×20mm 标签）

```
SIZE 40 mm, 20 mm
GAP 2 mm, 0 mm
CLS
TEXT 5, 2, "3", 0, 1, 1, "TSCA111"
TEXT 5, 6, "2", 0, 1, 1, "100.00μg/mL"
TEXT 5, 10, "2", 0, 1, 1, "2026-05-18 ~ 2026-06-17"
TEXT 5, 14, "2", 0, 1, 1, "4°C冷藏"
PRINT 1, 1
```

## 关键文件

- `templates/SolutionConfig.html` — 新增打印按钮和 `bbcdPrintLabel()` 函数
- `login_html.py` — 新增 `/api/lims/print_label` 接口

## 验证

1. 标签打印机连上 WiFi，记录 IP 地址
2. Flask 后端配置打印机 IP
3. 配置一个储备液，点击"打印标签"
4. 打印机自动出纸，标签内容正确
5. 工作液模式：每个稀释点打印一个标签
