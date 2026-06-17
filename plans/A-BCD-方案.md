# A-BCD 支持：把 A-B 标签页升级为 A-BCD（类型选择 B/C/D）

> **状态**：方案已定稿。应要求**暂不实施代码改动**，仅把本方案归档到项目仓库。后续实现时在分支 `A-BCD` 上进行。

## Context（为什么这么改）
LIMS 支持 A-BCD：从标准品 A 直接配制成 **储备液 B / 应用液 C / 工作液 D**。当前 `SolutionConfig.html` 的 A-B 标签页只能配 B。用户决定：**把 A-B 标签页改名为 A-BCD**，选中标准品后弹出「类型选择 B/C/D」，按类型路由到不同配置弹窗：
- **B** → 现有储备液配置弹窗 `#solConfigModal`（不变）。
- **C** → 复用同一 `#solConfigModal`，按类型参数化（`solutionType=C`），支持多条同时配置（每个标准品一瓶，分开）。
- **D** → 复用 bbcd 工作液弹窗 `#bbcdConfigModal`，把选中的标准品 A 作为 `%` 源（称量）**混到一瓶**，锁单瓶（强制 1 个稀释点）；不支持多条工作液。

**核心复用**（已验证）：
- bbcd 已自带完整 `%` 源（称量）通路——`bbcdPctConcNum`（`:2389`，纯度/100×称量×1e6/体积）、`bbcdRecalcRow` `%` 分支（`:4143`）、`bbcdBuildPayload` 对 `%` 源写 `receivedUint:'g'`（`:4484/4637`）、`bbcdPopulateBCConfigTable` 靠浓度串 `%` 识别（`:3511`）。
- A-B 的 `solCalcRowConc` 称量浓度公式与 bbcd `%` 完全相同；`solSave` 已用 `solType` 作 `solution_type`（`:2271`），故把 `solType` 设成 C 即存应用液。
- B/C 走 solSave（pageType **A**）；D 走 bbcd 工作液（pageType **D**）——不产生 A 与 B 同记录混用，**规避删重建 NPE**。

用户已确认（含两次精修）：① A-BCD = A 直达 C/D；② 多个 A 源可混合、但**禁止混入已配 B**；③ C 复用 `#solConfigModal`；④ D 复用 bbcd 工作液弹窗（接受其 pageType 耦合）；⑤ 范围 = 配制/保存 + 导出 + 查询/列表 + 打印标签 + 审核。

---

## 改动文件
- `templates/SolutionConfig.html`（唯一；后端 `login_html.py` 不改）

## 实施前置
- 从 `main-1` 新建并切换分支：`git checkout -b A-BCD`，所有改动在此分支。

## 约束
- 改动集中在 `#abPane`（A-BCD 标签页）入口 + `#solConfigModal`（C 参数化）+ bbcd 工作液入口（D 复用）。**BC-BCD(bbcd) 标签页既有功能保持不变**，二者并存（A-BCD 从标准品起；BC-BCD 从已配溶液起）。
- 「多个 A 可混、禁止混入 B」：D 入口只读 A-BCD 的选中标准品、**不读** `bbcdSelectedItems`，B 源进不来。

---

## 设计总览（路由）
`#abPane` 标签 "A-B" → "A-BCD"（`:418`）。`solResultAddBtn`（`:7351`）改为：先弹类型选择 B/C/D，再按类型路由：
- **B** → `solType='SOLUTION_TYPE_B'` → 开 `#solConfigModal`（原逻辑）。
- **C** → `solType='SOLUTION_TYPE_C'` → 开 `#solConfigModal`（参数化标题）。
- **D** → 选中标准品 A 映射成 `%` 源 → 开 `#bbcdConfigModal` 工作液模式、锁 1 稀释点。

---

## 1. UI 改动
- **标签文案**（`:418`）：`A-B` → `A-BCD`。
- **类型选择弹窗**：新增轻量 `#abcdTypePickerModal`（3 按钮 储备液 B / 应用液 C / 工作液 D）。**不复用** `#bbcdTypePickerModal`（避免与 bbcd 的 confirm 处理耦合）；confirm 处理按下方 §3/§4 路由。
- **`#solConfigModal` 标题参数化**（`:741`，当前无 id、写死「储备液配置参数」）：给标题元素加 id，开弹窗前按 `solType` 设文案——B→「储备液配置参数」、C→「应用液配置参数」（参考 `:7453`/`:3417` 的 `titleMap` 模式）。
- **来源徽标**：D 进入 `#bbcdConfigModal` 时在标题旁显示 `<span class="badge bg-success">标准品 A</span>`（仅 A-BCD 的 D）。

---

## 2. B（不变）
现有 `#solConfigModal` + `solSave`（`solutionType=B`、`pageType=A`、`/api/lims/save_solution`）。零改动。

---

## 3. C（参数化 `#solConfigModal`）
- 类型选 C → `document.getElementById('solType').value='SOLUTION_TYPE_C'` → 标题设「应用液配置参数」→ 开 `#solConfigModal`。
- `solSave` 已用 `solType` 作 `solution_type`（`:2271`），故自动存 **solutionType=C、pageType=A**。
- **多条同时配置**：沿用 `solItems` 多行（每个标准品 → 一条独立 C，分开成瓶），与 B 一致。
- 有效期默认（`:2027-2032`）：C 沿用 B 的默认（非 D 不触发 30 天）；如应用液有不同效期规则，实现时按需调整。
- 导出/查询/打印/审核：走 B 同一路径（仅 type=C）。导出模板默认同 B（RF10-09/RF10-10），如应用液需不同模板，实现时确认。

---

## 4. D（复用 bbcd 工作液弹窗，混到一瓶）
新增入口 **`bbcdOpenConfigFromA_D()`**（镜像 `bbcdOpenConfig` `:3073`，但只服务 A-BCD 的 D）：
1. 读 A-BCD 选中标准品（`solSelectedMap`/`solItems`，跨标签同页可读），逐个 `mapStandardAToPctSource(item)` → `%` 源；存 `window.bbcdPendingASources`。
2. 进 `#bbcdConfigModal`：`bbcdSolType` 强制 `SOLUTION_TYPE_D`；**只 seed 这些 `%` 源**、**不读** `bbcdSelectedItems`（确保不混入 B）；**强制稀释点=1**（`#bbcdDilutionCount` `:885` 设 1 并禁用、`window.bbcdCurrentDilutionCount=1`，不允许多瓶）；置 `window.bbcdIsABCDFlow=true`；显示来源徽标；末尾 `Modal.show` + `bbcdRecalcAll()`。
3. 用户填共享字段 + 每行称量(g) + 共享定容体积(mL) → 浓度经 `bbcdRecalcRow` `%` 分支即时算出。
4. 保存走 **`bbcdBuildWorkingSolutionPayload`（`:4891`）** → pageType **D**、30 天效期、`[min-max]` 编号、单瓶即 1 标签。

**复用/新增函数**：
- `mapStandardAToPctSource(item)`：标准品 A → `%` 源行（`concValue`=纯度、`concUnit='%'`、`concentrationRaw`含`%`、`isPctSource:true`、`pctConcUnit:'μg/mL'`、`constantVolume:0`、`configureOrder:'A-'+id`、`_fromStandardA:true`）。
- `bbcdSeedPctSourceRow(src)`：按 `bbcdOpenConfig` 单组分分支（`:3251-3285`）生成一行 `<tr>`（含共享 `#bbcdMedium`/`#bbcdVolume`），append 到 `#bbcdConfigTable tbody`；多 A 则多行（混瓶）。

**改动现有**：
- 类型选择 confirm（D 分支）→ 调 `bbcdOpenConfigFromA_D()`。
- `#bbcdConfigModal` 的 `hidden.bs.modal` 清理：追加清 `window.babcdIsABCDFlow`/`bbcdPendingASources`、隐藏来源徽标、复原 `#bbcdDilutionCount` 可用状态。

---

## 5. pageType / solutionType 决策（规避 NPE）
| 类型 | solutionType | pageType | 保存路径 |
|---|---|---|---|
| B | B | **A** | solSave（既有） |
| C | C | **A** | solSave（参数化） |
| D | D | **D** | bbcd 工作液 `bbcdBuildWorkingSolutionPayload` |

- B/C 同为 pageType A（同一路径），互不冲突；D 为 pageType D。**不产生 A 与 B 同记录混用** → 规避「%源 B 老流程 pageType=A vs bbcd pageType=B」删重建 NPE。
- D 的源是标准品 A（合法 pageType A 起点），工作液 D 为 pageType D，链路 A→D 正常（后端 `_is_weighing_top`：`received_unit=='g'` 判定称量顶源，`login_html.py:1432-1434`）。

---

## 6. 各能力影响
| 能力 | B | C | D |
|---|---|---|---|
| 配制/保存 | solSave（既有） | solSave(type=C) | bbcd 工作液 payload |
| 导出 Word | RF10-09/10（既有） | 同 B（模板按需确认） | bbcd 工作液导出 |
| 查询/列表 | 既有 | 同 B（type=C 归应用液） | bbcd 列表（D） |
| 打印标签 | 既有 | 同 B | bbcd 工作液标签（单瓶1标签） |
| 审核 | 既有 | 同 B | bbcd 审核 |

---

## 7. 风险 / 待确认
- **D 后端链路**：D 源为标准品 A（非已配溶液记录），后端工作液对 `%`/g 称量顶源的处理需实测（`login_html.py` `_is_weighing_top` `:1432-1434`、A 前缀跳过 `:1478`）。
- **C 导出模板**：默认同 B，若应用液需不同 Word 模板，实现时确认。
- **锁单瓶**：强制稀释点=1 且禁用增减；确认 UI 不泄露多瓶/逐级稀释入口。
- **列表归类**：C 记录归「应用液」、D 归「工作液」（沿用 `detectType` `:2952` 按 `configureOrder` 前缀）。
- **标题元素 id**：`#solConfigModal` 标题（`:741`）当前无 id，实现时加 id 再参数化。

---

## 8. 验证步骤（端到端）
1. **回归（B 不变）**：A-BCD 选标准品 → 类型 B → 与原 A-B 完全一致（配置/保存/查询/导出/打印/审核）。
2. **C（多条）**：选 2 个标准品 → 类型 C → 弹窗标题「应用液配置参数」、2 行 → 分别称量溶解 → 保存 2 条 solutionType=C/pageType=A → 查询归应用液 → 导出/打印/审核。
3. **D（混到一瓶）**：选 2-3 个标准品 → 类型 D → bbcd 工作液弹窗、多行 `%` 源、稀释点锁 1（不可加）→ 各行称量 + 共享体积 → 保存 **1 条**工作液 pageType=D/30 天 → 查询/导出/打印/审核；确认「混到一瓶、未产生多瓶、未混入 B」。
4. **pageType/NPE**：确认 B/C=A、D=D；删除一条 D 后用同标准品重配 → 无异常。
5. **开关/清理**：类型选择取消、弹窗关闭后状态清空，不影响 BC-BCD 标签页既有 D 配置。
