# 工作液（D）合并导出规则

## 一、合并候选识别（前端）

导出工作液时，扫描配置表源溶液行，满足以下全部条件的记录为合并候选：

| 条件 | 说明 |
|------|------|
| `configureOrder` 以 `B-` 或 `C-` 开头 | 排除 A 类（标准品）和 D 类（工作液自身） |
| 浓度单位不是 `%` | 排除百分比浓度的源 |
| 多组分只取第一行 | `isComponent=true` 且 `componentIdx=0` |
| 同一 `configureOrder` 去重 | 避免重复候选 |

前端提取的候选数据：

```javascript
{
    configureOrder: "C-3797",
    solutionName: "标准品/正己烷中邻苯二甲酸二异壬酯，1000μg/mL",
    solutionCode: "CK-CG-2025102-1.00-20260519",
    solutionId: 23564,      // LIMS 记录的数字 ID
    concUnit: "mg/L",
    concValue: 0.01,
}
```

---

## 二、合并弹窗触发条件

有候选后，前端调 `/api/lims/batch_solution_details`，传入：

- `solution_codes`：候选的 solutionCode 列表
- `solution_ids`：候选的 LIMS 数字 ID 列表
- `d_configure_date`：D 的配置日期（来自日期选择器）
- `d_configurator_name`：D 的配置人（来自当前登录用户）

后端返回的每个候选记录包含 `configureDate` 和 `configuratorName`。前端检查：

```javascript
const canMerge = detailData.data.every(s =>
    (s.configureDate || '').startsWith(dDate) &&
    (s.configuratorName || '').trim() === dUser.trim()
);
```

| 条件 | 说明 |
|------|------|
| 所有候选的 `configureDate` 以 D 的日期开头 | 同一天配置 |
| 所有候选的 `configuratorName` 与 D 的配置人一致 | 同一人配置 |
| 两个条件同时满足 | 弹出合并确认弹窗 |

不满足则跳过弹窗，直接导出 D 单独的配制记录。

---

## 三、后端数据流

### 3.1 候选记录获取

1. 前端传入 `solution_ids`（LIMS 数字 ID）
2. 后端直接调 `selectBySCId?id={id}` 获取每个候选的完整记录
3. **不依赖 `getSolutionAdata`**（该接口对 B 类记录返回 500）

### 3.2 父级记录获取

1. 从候选记录的 `originalCode` 字段读取父级编号（如 `B-3377`）
2. 调 `_resolve_order_to_id()` 通过 `getSolutionAdata?type=SOLUTION_TYPE_B` 解析父级编号为数字 ID
3. 用数字 ID 调 `selectBySCId` 获取父级记录

### 3.3 detail_item 重建（`_reconstruct_detail_from_record`）

| 字段 | 有父记录 | 无父记录 |
|------|---------|---------|
| `originalName` | `parent.solutionName` | `concentrationCount` 冒号前部分 |
| `originalConcentration` | `parent.concentration` | `concentrationCount` 冒号后部分 |
| `receivedQuantity` | `result_cv × vol / parent_cv`（C1V1=C2V2） | 同公式，但 parent_cv 可能等于 result_cv |
| `medium` | `record.medium` | `record.medium` |
| `volume` | `record.constantVolume` | `record.constantVolume` |
| `configurationConcentration` | `record.concentration`（如 `0.01(mg/L)`） | 同 |
| `configurationUnit` | 从 concentration 解析（如 `mg/L`） | 同 |
| `originalNo` | `"{originalCode}\n{configureOrder}\n{solutionCode}"` | 同 |
| `customNum` | `record.controlledNo` | `record.controlledNo` |

### 3.4 表头溯源（`_collect_all_top_ancestors`）

沿 `originalCode` 递归向上追溯，**同一配置人 + 同一配置日期**范围内追到最顶层：

```
D → C → B → A（同人同日全链路）
         ↘ 停止（B 不同日/不同人）
```

| 场景 | 溯源链 | 表头取自 |
|------|--------|---------|
| 同人同日全链路 | A→B→C→D | A 的信息 |
| B 不同人/不同日 | B→C→D（B 处截断） | C 的信息 |
| 多源 B4←B1,B2,B3 | B1→A1, B2→A2, B3→A3 | A1/A2/A3 去重合并 |

溯源结果收集为 `headerInfo`：

```python
{
    "names": ["DIDP", "DINP"],           # 溯源顶层 solutionName 去重
    "codes": ["CIRS401001", "CIRS401002"], # 溯源顶层 controlledNo 去重
    "concs": [1000, 500],                 # 溯源顶层 concentration 数值去重
    "concUnit": "μg/mL",                  # 浓度单位（从 concentration 字段解析）
}
```

---

## 四、Word 表格结构

### 4.1 表头（Row 1）

| 位置 | 有合并 + headerInfo | 有合并 + 无 headerInfo | 无合并 |
|------|-------------------|---------------------|--------|
| 标准物质名称 | headerInfo.names 用 `；\n` 连接 | bc_detail_rows 的 originalName 去重 | detail_list 的 originalName |
| 编号 | headerInfo.codes 用 `；\n` 连接 | customNum 去重，回退 originalNo | originalNo 最后一段 |
| 浓度 | 1 个值显示具体数值，多个显示"见下表" | 同 | 单源显示具体值，多源"见下表" |
| 浓度列标题 | `浓度({headerInfo.concUnit})` | `浓度({detail.configurationUnit})` | 同 |

### 4.2 列标题（Row 3）

```
母体标液(浓度单位) | 取量(mL) | 溶剂 | 稀释至,mL | 浓度(浓度单位) | 编号 | 配制日期 | 有效期
```

### 4.3 数据行排列

```
[B/C 行]   is_bc=true   dilutionIdx=-(N)   ← 合并候选的配制记录
────────── BC 与 D 段之间不合并 ──────────
[D 行 1]   is_bc=false  dilutionIdx=0      ← 第 1 个稀释点
[D 行 2]   is_bc=false  dilutionIdx=1      ← 第 2 个稀释点（逐级链）
```

---

## 五、B/C 行字段填充规则

### 5.1 普通单组分

```python
row = {
    'name':          originalName,         # → 母体标液（两行：名称 + 浓度）
    'conc_val':      originalConcentration,# → 母体标液第二行（源浓度）
    'qty':           receivedQuantity,     # → 取量
    'medium':        medium,               # → 溶剂
    'volume':        constantVolume,       # → 稀释至
    'result_conc':   configurationConcentration,  # → 浓度
    'result_code':   solutionCode（浓度修正后）,     # → 编号
    'dilutionIdx':   -(ms_idx + 1),        # 负数标识 B/C 段
    'is_bc':         True,
    'configure_date': 候选的配置日期,
    'validity_date':  候选的有效期,
}
```

### 5.2 多组分（浓度含 `;` 和 `:`）

浓度格式如 `DIDP:447.10;DINP:682.32`，拆分为独立行：

```
DIDP  447.10  ← 第一组分
DINP  682.32  ← 第二组分
```

### 5.3 编号浓度修正

solutionCode 格式为 `CK-CG-{编号}-{浓度}-{日期}`。

B/C 行的 solutionCode 浓度部分需替换为该行实际浓度值：

```python
# CK-CG-2025102-1.00-20260519（原始，浓度 1.00）
# → CK-CG-2025102-0.01-20260519（修正后，浓度 0.01）
ms_code = re.sub(r'(?<=-)(\d+\.\d+)(?=-)', f'{result_cv:.2f}', ms_code, count=1)
```

---

## 六、D 行字段填充规则

```python
row = {
    'name':          originalName,
    'conc_val':      originalConcentration,
    'qty':           receivedQuantity,
    'medium':        medium,
    'volume':        volume,
    'result_conc':   configurationConcentration（空则从 solutionCode 提取）,
    'result_code':   resultCode（空则用 solutionCode）,
    'dilutionIdx':   0, 1, 2...（稀释级数）,
    'is_bc':         False,
    'configure_date': D 的配置日期,
    'validity_date':  D 的有效期,
}
```

---

## 七、垂直单元格合并规则

### 7.1 合并列集

| 场景 | B/C 行 merge_cols | D 行 merge_cols |
|------|-------------------|-----------------|
| 有合并 + 工作液 + 多源 | {2,3,4,5,6,7} | {2,3,4,5,6,7} |
| 有合并 + 工作液 + 单源 | {2,3,4,5,6,7} | {4,5} |
| 有合并 + 非工作液 | {2,3,4,5,6,7} | {2,3,4,5,6,7} |
| 无合并 + 工作液 + 多源 | {2,3,4,5,6,7} | {2,3,4,5,6,7} |
| 无合并 + 其他 | {4,5} | {4,5} 或 {2,3,4,5,6,7} |

列编号：2=溶剂, 3=稀释至, 4=浓度, 5=编号, 6=配制日期, 7=有效期

### 7.2 单元格合并条件（全部满足）

1. 列索引在 `merge_cols` 集合中
2. 当前行的 `is_bc` 与上一行相同（BC 段内或 D 段内）
3. 当前行的 `dilutionIdx` 与上一行相同
4. 当前值 == 上一行的值

满足时：清空单元格内容，设置 `vmerge=continue`。
不满足时：写入值，如果在 merge_cols 中则设置 `vmerge=restart`。

### 7.3 母体标液列（tc[0]）显示

| 情况 | 显示 |
|------|------|
| B/C 行 | 两行文本：源名称 + 源浓度 |
| D 单源 + 第 1 稀释点 | 只显示浓度数值 |
| D 多源 + 第 1 稀释点 | 两行文本：源名称 + 源浓度 |
| D 逐级链（dilutionIdx > 0） | 只显示前一稀释点结果浓度 |

---

## 八、取量计算公式

从 C1V1 = C2V2 稀释公式推导：

```
取量(receivedQuantity) = 结果浓度 × 定容体积 / 源浓度
                        = result_cv × vol / parent_cv
```

| 参数 | 来源 |
|------|------|
| result_cv | `record.concentration` 解析数值（如 `0.01`） |
| vol | `record.constantVolume`（如 `10.0`） |
| parent_cv | `parent.concentration` 解析数值（如 `1.0`） |

计算结果：`0.01 × 10.0 / 1.0 = 0.10 mL`

格式化为保留 2 位小数。当 `calc > vol`（取量超过定容体积）时不显示。

---

## 九、浓度单位解析

LIMS 的 `concentrationUnitName` 字段存储的是体积单位（如 `mL`），不是浓度单位。

浓度单位从 `concentration` 字段解析：

```
"0.01(mg/L)"  → mg/L
"1000μg/mL"   → μg/mL（无括号时取数字后部分）
```

解析函数 `_parse_conc_unit()`：
1. 优先提取括号内文本：`\( ([^)]+) \)`
2. 回退匹配数字后文本：`[\d.]+\s*(.+)`
3. 默认值 `μg/mL`

---

## 十、涉及的关键文件

| 文件 | 关键函数/位置 |
|------|-------------|
| `templates/SolutionConfig.html` | `bbcdGetMergeCandidates()` L3592, `bbcdExport()` L3663, `showMergePickerModal()` L3619 |
| `login_html.py` | `_reconstruct_detail_from_record()` L53, `_trace_to_top()` L162, `_collect_all_top_ancestors()` L197, `lims_batch_solution_details()` L1417, `lims_export_bbcd_docx()` L1481 |

---

## 十一、当前已知限制

1. **溯源依赖 `getSolutionAdata`**：父级查找需要该接口配合正确的 `type` 参数（`SOLUTION_TYPE_B` 等）。如果 LIMS 仍返回 500，溯源链会断在当前层，表头显示当前层而非源头的编号和浓度。
2. **取量计算依赖父记录**：无法获取父级记录时，`concentrationCount` 中的浓度是**结果浓度**而非源浓度，会导致取量计算错误。
3. **solutionCode 浓度修正**：仅替换第一个匹配的 `X.XX` 浓度值，如果编号格式变化可能匹配错误。

---

## 十二、LIMS API 溯源规则（D→C→B→A）

### 12.1 链路结构

- D/C/B 记录都在 `dtSolutionConfigure` 表中，通过 `originalCode` 字段链接到父级
- A 记录在 `consumableReceive` 系统中，是**不同的 API 和字段体系**（`num`, `receiveCode`, `certNo`, `controlledNo`）
- 示例：D 的 `originalCode="C-3797"`，C 的 `originalCode="B-3377"`
- `originalCode` 可逗号分隔表示多源（如 `"B-3405,B-3406,B-3407"`）

### 12.2 LIMS 关键字段

| 字段 | 含义 | 注意事项 |
|------|------|----------|
| `originalCode` | 父级的 `configureOrder` | 可逗号分隔多源 |
| `configureOrder` | 配置编号如 `C-3797` | 前缀表示类型 (B-/C-/D-) |
| `solutionCode` | 溶液编码如 `CK-CG-2025102-0.01-20260519` | 含浓度和日期，**表头编号取此字段** |
| `controlledNo` | 受控编号/证书号如 `CIRS402178-3` | 不是溶液编码，不要用于表头 |
| `concentration` | **配置后的结果浓度**如 `0.01(mg/L)` | 不是源浓度，是配置后该记录自身的浓度 |
| `concentrationCount` | 格式 `父名:结果浓度(单位)` | 冒号后是**结果浓度**，不是源浓度 |
| `concentrationUnitName` | 存的是**体积单位**如 `mL` | 不是浓度单位！需从 `concentration` 字段解析 |
| `constantVolume` | 定容体积如 `10.0` | 单位 mL |
| `medium` | 溶剂如 `甲醇` | |
| `saveDetailList` | **始终为 null** | 不可用，需从其他字段重建 |
| `solutionName` | 溶液名称 | |

### 12.3 LIMS API 接口

| 接口 | 用途 | 关键注意 |
|------|------|----------|
| `selectBySCId?id=数字ID` | 获取单条记录详情 | **只接受数字 ID**，不接受 configureOrder 字符串 |
| `getSolutionAdata?type=SOLUTION_TYPE_B` | 列表查询，建 order→id 映射 | **必须传 type 参数**（B/C/D/E），不传返回 500；type 根据前缀推断：B-→SOLUTION_TYPE_B |
| `detail?id=数字ID` | 另一个详情接口 | 也只接受数字 ID |

### 12.4 configureOrder 前缀与 type 映射

| 前缀 | type 参数 |
|------|----------|
| `B-` | `SOLUTION_TYPE_B` |
| `C-` | `SOLUTION_TYPE_C` |
| `D-` | `SOLUTION_TYPE_D` |
| `E-` | `SOLUTION_TYPE_E` |

### 12.5 solutionCode 命名规则

格式：`CK-CG-{编号}-{浓度}-{日期}`

```
CK-CG-2025102-0.01-20260519
│     │      │     │
前缀  编号   浓度  配制日期
```

- 浓度部分是该记录**配置后的结果浓度**（如 C 配置后浓度为 0.01 mg/L）
- 表格行的编号需将浓度部分替换为该行实际配置浓度
- 表头编号直接使用溯源顶层的 `solutionCode`（不做浓度替换）

### 12.6 编号取值规则

| 位置 | 取值来源 | 说明 |
|------|---------|------|
| 表头编号 | 溯源顶层的 `solutionCode` | 不是 `controlledNo` |
| B/C 表格行编号 | 该记录的 `solutionCode`（浓度部分替换为实际配置浓度） | |
| D 表格行编号 | `resultCode`（空则用 D 的 `solutionCode`） | |

### 12.7 浓度解析

`concentration` 字段示例：

```
"0.01(mg/L)"   → 数值 0.01，单位 mg/L
"1000μg/mL"    → 数值 1000，单位 μg/mL
"13种有机锡混标:0.01(mg/L)" → concentrationCount 格式，冒号前为父名，冒号后为结果浓度
```
