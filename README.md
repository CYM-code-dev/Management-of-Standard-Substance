# Management of Standard Substance

有机标准品与标准溶液配置管理工具。项目基于 Flask 提供本地 Web 页面，用于读取和维护标准品 Excel 台账、查询并回写 LIMS、查看标准品证书，以及导出标准溶液配制记录 Word 文档。

## 功能概览

- LIMS 登录：支持验证码登录，登录后缓存会话并自动保持会话有效。
- 有机标准品管理：读取共享盘 Excel 台账，展示、查询、新增和编辑标准品信息。
- Excel 同步：将新增标准品写入配置的 Excel 文件，并更新标准品使用情况。
- 证书查看：根据实验室编号在配置的证书目录中查找并打开 PDF 证书。
- LIMS 查询：按实验室编号、关键字、机构、状态等条件查询 LIMS 耗材记录。
- LIMS 回写：支持更新浓度单位、确认领用、保存标准溶液配置。
- 标准物质配置：支持从标准品配置储备液，也支持从已配置储备液继续配置储备液、应用液、工作液。
- 配制记录导出：基于 Word 模板导出标准溶液配制记录。

## 项目结构

```text
.
├── login_html.py                 # Flask 后端入口，包含 LIMS 对接、Excel 写入、Word 导出等接口
├── config.json                   # Excel 台账路径和证书目录配置
├── requirements.txt              # Python 依赖
├── templates/
│   ├── login.html                # 登录页
│   ├── OrganicStd.html           # 有机标准品管理页
│   └── SolutionConfig.html       # 标准物质配置页
└── word_templates/
    ├── RF10-09 标准溶液配制记录(1).docx
    └── RF10-10 标准溶液配制记录（稀释）(1).docx
```

## 环境要求

- Python 3.10 或更高版本
- 可访问内部 LIMS 服务：`http://192.168.12.234:60015`
- 可访问配置中的共享盘 Excel 文件和证书目录

## 安装依赖

建议在虚拟环境中运行：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 配置

项目使用 `config.json` 保存 Excel 台账路径和证书目录路径：

```json
{
  "excelPath": "\\\\files.cirs-ck.com\\轻工公盘\\有机\\3-有机耗品\\有机标准品.xlsx",
  "certPath": "\\\\files.cirs-ck.com\\公共文件夹\\标品证书"
}
```

也可以在“有机标准品管理”页面右上角点击“设置”，通过页面修改：

- Excel 文件 UNC 路径
- 证书文件夹 UNC 路径

## 启动服务

```powershell
python login_html.py
```

默认访问地址：

- 登录页：http://127.0.0.1:5000/login
- 有机标准品管理：http://127.0.0.1:5000/organic-std
- 标准物质配置：http://127.0.0.1:5000/solution-config

服务默认监听 `0.0.0.0:5000`，同一局域网内其他机器可通过本机 IP 访问。

## 使用流程

1. 启动 Flask 服务。
2. 打开登录页，输入 LIMS 账号、密码和验证码。
3. 进入“有机标准品管理”页面，加载 Excel 台账。
4. 通过实验室编号或关键字查询标准品/LIMS 耗材记录。
5. 新增或编辑标准品后，同步写入 Excel。
6. 进入“标准物质配置”页面，选择标准品或已配置溶液，填写配制参数。
7. 根据需要确认领用、保存配置，并导出 Word 配制记录。

## 主要页面

### 有机标准品管理

路径：`/organic-std`

用于维护有机标准品台账，包含：

- 加载 Excel 标准品数据
- 查询、重置、筛选标准品
- 新增标准品
- 查询 LIMS 耗材记录并回填表单
- 打开标准品证书
- 配置 Excel 和证书目录路径

### 标准物质配置

路径：`/solution-config`

用于标准溶液配置，包含：

- 从标准品配置储备液
- 从已配置储备液配置储备液、应用液、工作液
- 批量填写保存条件、有效期、定容体积等参数
- 确认领用并保存到 LIMS
- 导出标准溶液配制记录 Word 文档

## 后端接口概览

常用接口包括：

- `GET /api/captcha`：获取 LIMS 登录验证码
- `POST /api/login`：登录 LIMS
- `GET /api/status`：获取当前登录状态
- `POST /api/logout`：退出登录
- `GET /api/query`：查询 LIMS 耗材记录
- `GET|POST /api/config`：读取或保存系统配置
- `POST /api/add_to_excel`：新增标准品并写入 Excel
- `POST /api/update_excel_usage`：更新 Excel 中的使用情况
- `POST /api/update_lims_unit`：更新 LIMS 浓度单位
- `POST /api/lims/receive`：确认领用
- `POST /api/lims/save_solution`：保存标准溶液配置
- `GET /api/lims/list_configured_solutions`：查询已配置溶液
- `POST /api/lims/export_docx`：导出标准溶液配制记录
- `POST /api/lims/export_bbcd_docx`：导出储备液/应用液/工作液配制记录

## Word 模板

导出配制记录依赖 `word_templates` 目录中的模板文件。请保持模板文件名不变，否则需要同步修改 `login_html.py` 中的模板名称。

当前使用的模板：

- `RF10-09 标准溶液配制记录(1).docx`
- `RF10-10 标准溶液配制记录（稀释）(1).docx`

## 注意事项

- 项目会读写共享盘 Excel 文件，请确认运行账号具有读写权限。
- `.xls` 和 `.xlsx` 均有部分兼容处理，但建议优先使用 `.xlsx`。
- LIMS 服务地址当前硬编码在 `login_html.py` 中，如服务地址变化需要修改 `base_url`。
- 当前 Flask 以 `debug=True` 启动，适合内网开发和日常工具使用；如需部署到生产环境，应关闭 debug 并使用 WSGI 服务。
- 登录会话会保存为 `session_<用户名>.json`，用于减少重复登录。

