# 财务 Excel Agent 前端与全流程实施方案

## 1. 目标与交付边界

产品主模型从“创建项目后进入固定 Python 流水线”调整为“选择公司后进入持续的月度 Agent 会话”。大模型负责理解材料、识别文件、编译规则、规划和逐人处理，并通过受控 Excel 工具完成修改与复核；财务用户只负责上传材料、说明要求和确认真正无法自动判断的异常。

本方案分两层交付：

- **明日可验收版**：完成可配置模型 Provider 和 Excel 工具闭环，跑通现有单租户的一家公司、两个真实月度样本、对话式处理、逐人异常确认、结果发布和刷新恢复。
- **生产完整版**：增加租户下多公司、规则文档/录音自动编译、公司权限、长期运行队列、完整审计和规则治理。

不能伪装完成的能力：未配置 API Key 时的真实模型处理、尚未接入的录音自动转写、多公司数据库模型、跨进程可靠任务队列。界面必须明确展示当前能力和阻断原因；没有模型配置时只能预检和人工处理，不能自动发布。

## 2. 核心工程决策

1. 继续使用 Next.js 14、React 18、Tailwind 3、Radix 和现有 shadcn 组件，不迁移框架版本。
2. 参考 `assistant-ui` 的会话、消息、附件和 composer 结构，但不整包引入依赖；财务消息是结构化工作记录，不是通用聊天气泡。
3. 大模型是流程主控：它持续选择和调用有类型 Excel 工具、读取结果并纠错。确定性后端执行器负责受控落盘和验证，不取代模型的业务处理职责。
4. 当前 `Project` 保持“一个月度批次”的语义；生产版新增 `Company`，由 `Company 1-N Project` 组织月度任务。
5. 先建立事件和状态契约，再并行开发 UI 与 API，避免前端继续兼容多套临时返回结构。
6. 旧流程只保留在 `/legacy-payroll`，所有默认入口、待办和结果回跳统一进入 Agent。

### 2.1 模型主控与 Excel 工具闭环

模型在一次 Run 中持续完成：资料理解、规则优先级处理、文件/Sheet/月份识别、人员分类、源到目标映射、逐人 WorkItem 规划、工具调用、结果自检、限次重试和异常提问。它不是只在最后生成一个 `Decision`。

第一期工具固定为 `inspect_workbook`、`classify_file`、`find_table`、`read_range`、`match_person`、`propose_changes`、`apply_cell_changes`、`copy_formula_from_reference`、`validate_workbook`、`rollback_work_item` 和 `publish_workbook`。每个工具均使用 Pydantic/JSON Schema 输入输出，校验租户、公司、运行 revision、目标范围、公式保护和权限；模型不能调用 Python、Shell、宏、自由 SQL 或任意公式。

```text
模型理解任务 -> 调用工具 -> 工具校验并执行 -> 返回结构化 observation
             -> 模型检查/修正/回滚/提问 -> 独立验证 -> 发布
```

模型可自主完成只读检查、active 规则下的唯一映射、可逆低风险写入和可验证重试。身份冲突、目标不唯一、同优先级金额冲突、非法金额、公式覆盖、证据冲突和发布权限不足必须暂停。每步记录模型/提示词/规则版本、工具调用、来源证据、目标单元格、旧值、新值、验证结果和重试次数。

没有 API Key 时，页面允许上传、画像、预检、人工结构化决定和验证，但 Run 停在 `MODEL_CONFIGURATION_REQUIRED`，不得宣称 Agent 已完成业务判断，也不得自动发布。

## 3. 用户主流程

```text
登录
  -> 选择公司
  -> 选择历史月份或新建月度任务
  -> 拖入总表、变更表、手册、录音
  -> 输入本次要求
  -> Agent 识别文件角色、公司和月份
  -> 显示处理计划和阻断项
  -> 自动执行低风险人员
  -> 逐个询问异常人员
  -> 全局验证
  -> 生成草稿/修订稿
  -> 人工发布或按公司策略自动发布
  -> 归档对话、规则版本、审计和文件哈希
```

规则优先级固定为：`当次明确指令 > 生效规则包 > 最新手册 > 录音`。

## 4. 页面与路由

### 4.1 明日版路由

| 路由 | 页面职责 |
|---|---|
| `/login` | 登录和服务状态 |
| `/projects` | 暂时作为月度任务列表，入口统一指向 Agent |
| `/projects/{projectId}/agent` | 主 Agent 对话工作台 |
| `/work-queue` | 跨任务异常待办，点击定位到 Agent 的人员项 |
| `/projects/{projectId}/result` | 兼容历史下载，主操作回 Agent |
| `/projects/{projectId}/legacy-payroll` | 仅兼容旧任务，不作为默认入口 |

### 4.2 生产版路由

| 路由 | 页面职责 |
|---|---|
| `/companies` | 用户有权限处理的公司列表 |
| `/companies/{companyId}/agent` | 公司工作区，默认打开最近月份 |
| `/companies/{companyId}/runs/{runId}` | 可分享/恢复的具体月度会话 |
| `/companies/{companyId}/rules` | 生效、候选、挂起规则和证据 |
| `/companies/{companyId}/files` | 按月份归档原始材料和结果 |
| `/companies/{companyId}/history` | 运行、发布、回滚和审计历史 |
| `/companies/{companyId}/settings` | 自动发布、密码密钥、负责人和脱敏策略 |

旧 `/projects/{id}` 路由服务端重定向到对应 Agent，不能同时维护两套默认流程。

## 5. Agent 工作台信息架构

### 5.1 顶部任务栏

- 公司/项目名称、月份、运行状态、规则版本。
- 已处理人数、待确认人数和阻断数。
- 暂停/继续、重新分析、发布结果。
- 发布按钮始终可见，但不满足条件时显示明确阻断原因。

### 5.2 左侧任务栏

- 当前月度会话与历史月份。
- 本次文件包：总表、来源表、手册、录音、当次附件。
- 人员队列筛选：全部、待确认、失败、已完成。
- 每个人只显示姓名/脱敏标识、问题字段和状态，支持快速搜索。

### 5.3 中间对话时间线

固定消息类型：

1. `file_manifest`：文件角色、月份、Sheet 和识别置信度。
2. `plan`：规则版本、预计影响范围、自动/人工数量。
3. `progress`：批量进度，不逐条刷屏。
4. `person_review`：单个人员的异常问题和操作按钮。
5. `user_instruction`：用户本次要求或补充说明。
6. `memory_candidate`：是否保存为候选经验。
7. `validation`：公式、人员、金额、月份和来源追溯结果。
8. `publication`：正式稿、修订稿、审计报告和文件哈希。

底部 composer 固定在视口内：附件按钮、自然语言输入、运行模式和发送按钮。输入 `Enter` 发送、`Shift+Enter` 换行；人员确认快捷键只在输入框未聚焦时生效。

### 5.4 右侧上下文检查器

- 当前人员和人员类别。
- 来源文件、Sheet、行/列、候选值。
- 总表 Sheet、目标单元格、当前值和拟写入值。
- 命中规则、来源证据、规则优先级和置信度。
- 最近决策和可撤销边界。
- 生效规则和候选 Memory 摘要。

移动端将左栏和右栏改为抽屉，中间对话保持唯一主视图。

## 6. 重度使用交互

- 处理一个异常后自动选择下一个待确认人员。
- `Alt+A` 应用建议、`Alt+K` 保留原值、`Alt+S` 跳过、`Alt+N` 下一项；按钮 tooltip 提示快捷键。
- 已自动处理人员折叠为批次摘要，避免数百条消息拖慢扫描。
- 所有筛选、当前人员和滚动位置写入 URL 或本地非敏感 UI 状态。
- 页面刷新后通过 `runId` 恢复运行、人员队列和当前选择，不用重新开始。
- 失败只影响单个人员，支持单人重试和撤销；全局校验失败才冻结发布。
- 金额统一右对齐并使用 tabular numerals；状态不能只依赖颜色。
- 列表支持键盘上下移动，焦点始终落在下一个可操作项。

## 7. 前端组件拆分

```text
features/agent/
  agent-workbench-page.tsx       # 数据容器，控制在 200 行左右
  use-agent-run.ts               # 加载、轮询、恢复、乐观更新
  agent-shell.tsx                # 三栏响应式框架
  run-header.tsx                 # 状态、进度和主操作
  task-rail.tsx                  # 文件、人员和筛选
  conversation-timeline.tsx      # 虚拟化/折叠后的事件流
  conversation-composer.tsx      # 附件、指令和发送
  messages/
    file-manifest-message.tsx
    plan-message.tsx
    progress-message.tsx
    person-review-message.tsx
    validation-message.tsx
    publication-message.tsx
  context-inspector.tsx          # 证据、规则、审计
  memory-rule-list.tsx
  agent-empty-state.tsx
  agent-error-boundary.tsx
```

`agent-workflow.ts` 只保留纯状态选择器和时间线投影，组件不得自行重复计算“是否可发布”“下一异常”等业务状态。

## 8. 前端状态模型

### Run 状态

```text
draft -> ingesting -> planning
planning -> blocked | executing
executing -> review | validating | failed
review -> executing | validating
validating -> blocked | ready_to_publish
ready_to_publish -> publishing -> published
```

### WorkItem 状态

```text
queued -> auto_applied
queued -> needs_review -> resolved | skipped
queued/needs_review -> failed -> queued
resolved -> needs_review  # 单人撤销后重新确认
```

前端按钮只由状态选择器决定，不能根据文案猜测状态。

## 9. API 契约

### 9.1 明日必须补齐

```text
GET  /api/agent/runs?project_id={id}              # 列表和恢复最新运行
POST /api/agent/runs                              # project_id/instruction/auto_publish
GET  /api/agent/runs/{runId}                      # 状态与摘要
GET  /api/agent/runs/{runId}/events               # 对话时间线
POST /api/agent/runs/{runId}/messages             # 批次级补充指令
POST /api/agent/runs/{runId}/plan
POST /api/agent/runs/{runId}/execute
GET  /api/agent/runs/{runId}/items?status=&query=
POST /api/agent/runs/{runId}/items/{itemId}/message
POST /api/agent/runs/{runId}/items/{itemId}/apply
POST /api/agent/runs/{runId}/items/{itemId}/retry
POST /api/agent/runs/{runId}/pause
POST /api/agent/runs/{runId}/publish
```

所有写接口带 `expected_revision` 或幂等键，避免用户双击和网络重试造成重复写表。

### 9.2 生产版补齐

```text
GET/POST/PATCH /api/companies
GET /api/companies/{companyId}/summary
GET /api/companies/{companyId}/runs
GET /api/companies/{companyId}/rules
POST /api/companies/{companyId}/rules/{ruleId}/activate|suspend
GET /api/companies/{companyId}/files
GET /api/companies/{companyId}/publications
```

统一返回结构：`data`、`revision`、`server_time`、`request_id`；错误返回稳定 `code`，前端不解析中文错误文案决定行为。

## 10. 数据模型与迁移

生产版新增：

- `companies`：租户下业务公司、发布策略、规则配置和凭据引用。
- `projects.company_id`：月度批次归属公司。
- `agent_runs`：状态、instruction、rule_version、revision、auto_publish、时间戳。
- `agent_events`：类型化对话和系统事件。
- `agent_work_items`：人员级状态、目标位置、风险和决策。
- `agent_decisions`：旧值、新值、来源、操作者和确认范围。
- `company_credentials`：仅保存加密后的凭据引用，不返回明文。

当前项目没有迁移框架。生产版先引入 Alembic，再新增 Company 表；禁止只靠 `create_all()` 修改已有数据库。历史 Project 迁移到每个租户的“待归类公司”，不根据项目名称自动猜公司。

## 11. 文件与规则材料

文件角色扩展为：

- `financial_master`
- `financial_source`
- `rule_manual`
- `recording`
- `instruction_attachment`

上传支持 `.xlsx/.xls/.docx/.txt/.m4a/.mp3/.wav`；每类独立大小限制和 MIME/魔数校验。DOCX 提取段落/表格定位，录音保留时间戳转写，规则编译产物必须包含证据定位。首次输入的文件密码使用公司密钥加密保存；日志、API 和 UI 不返回密码。

## 12. 执行与恢复

明日版可以使用单进程后台任务，但必须先保存 Run，再启动处理；前端轮询 `GET run/events`。进程重启后，处于 `executing` 的任务标记为 `interrupted` 并允许恢复，不能默认为成功。

生产版使用持久任务队列；每个 WorkItem 独立幂等执行。模型通过有类型工具推动处理并读取每次 observation；写表工具采用草稿副本、单元格变更日志和原子替换。模型不能执行 Python、Shell、宏或生成代码，但可以选择、组合和重试白名单 Excel 工具。

## 13. 两份真实样本的实操

### 样本 1：2026-07

1. 目标总表：`202608（所属月202607）...xlsx`。
2. 来源：`薪资数据-科园-7月薪资（8.14发薪）.xlsx`。
3. 加载科园 active 规则包；缺失规则才读取最新手册/录音。
4. 生成逐人 WorkItem；唯一非零奖金列自动应用，两列有值进入确认。
5. 通过公式、月份、人员和来源追溯验证后生成草稿和审计。

### 样本 2：2026-06

来源属于 `2026.06`，目前选中的总表文件名包含 `所属月202605`。预检必须阻断并询问目标总表，不能自动发布。用户确认目标文件后才能继续同一运行。

## 14. 实施任务与顺序

### Phase A：契约与恢复能力

1. 固定 Run/Event/WorkItem/Decision/ToolCall/ToolResult TypeScript 与 Pydantic 契约。
2. 增加可配置模型 Provider、模型主控循环、超时/限次重试和无 Key 阻断状态。
3. 增加运行列表、最新运行恢复、事件流和 revision 幂等控制。
4. 为状态机、下一待办、工具审批和发布条件增加单元测试。

### Phase B：核心对话工作区

5. 拆分 Agent 页面为 shell、header、rail、timeline、composer、inspector。
6. 实现文件识别、计划、工具调用、进度、人员确认、验证和发布七种核心消息。
7. 接入真实上传、初始 instruction、模型事件流、运行轮询和刷新恢复。
8. 实现应用/保留/跳过/重试后自动选择下一异常。

### Phase C：全局工作流

9. 项目列表改为财务任务语义；新建后直接进入 Agent composer。
10. 待办全部链接到 `/agent?run={runId}&item={itemId}`，停止进入旧人工页。
11. 结果页改为 Agent 发布归档视图，保留历史下载兼容。
12. 收敛 SaaS Shell 导航和服务状态，删除重复入口。

### Phase D：公司与规则治理

13. 引入 Alembic，建立 Company 与归属迁移。
14. 实现公司列表、公司切换、规则中心、文件中心、历史和设置。
15. 接入 DOCX/录音规则编译、证据审阅和规则激活流程。

### Phase E：生产硬化

16. 持久任务队列、暂停/恢复、进程重启恢复和并发锁。
17. 权限、敏感字段脱敏、凭据加密和下载审计。
18. 桌面/平板/手机浏览器验收、性能和无障碍检查。

## 15. 验收标准

- 用户从登录到发布不进入旧流程页面。
- 刷新 Agent 页面后恢复相同 run 和当前人员。
- 100+ 人任务只显示汇总和异常，不生成 100+ 条展开消息。
- 一个异常确认后 300ms 内切到下一项，且重复点击不重复写表。
- 所有改单元格展示来源、旧值、新值、规则版本和操作者。
- 模型能从资料理解开始连续调用 Excel 工具，读取工具结果后继续或修正；审计可重放完整调用链。
- 月份冲突、公式风险、人员匹配冲突和双候选值必须阻断。
- 无待确认且验证通过时才允许发布；自动发布受公司配置控制。
- 1440px 三栏、1024px 双栏、768px/375px 单栏抽屉均无横向溢出。
- 前端 lint/build、后端 pytest、真实工作簿回归和浏览器 console 均通过。

## 16. 风险与控制

| 风险 | 影响 | 控制 |
|---|---|---|
| 只改 UI 未补恢复 API | 刷新丢任务 | Phase A 必须先完成 |
| 通用聊天组件过重 | 引入无关依赖与交互 | 只借鉴结构，使用现有组件 |
| 多公司直接复用 Tenant | 无法管理多个业务公司 | 生产版增加 Company |
| 同步处理大文件 | 页面超时 | 运行先落库/落盘，再后台执行和轮询 |
| 模型生成任意代码直接写表 | 财务高风险 | 模型只能调用有类型白名单工具，工具独立校验和落盘 |
| 只使用确定性流水线却显示为 Agent | 用户误判处理能力 | 无 Key 时显式阻断模型阶段并禁止自动发布 |
| 历史项目自动归类 | 公司串数据 | 迁移到“待归类公司”，人工确认 |

## 17. 开发检查点

- **检查点 1**：状态契约、恢复和幂等测试通过后，才开始页面重构。
- **检查点 2**：单个真实样本从上传到异常确认跑通后，才改全局待办。
- **检查点 3**：两个真实样本均输出可对比 Excel 后，才开放发布按钮。
- **检查点 4**：多公司权限与迁移测试通过后，才展示公司切换器。

## 18. 当前仍需业务确认

第二份样本应使用哪个目标总表：当前 `202606（所属月202605）...xlsx`，还是 `202607（所属月202606）...xlsx`。这是阻断性选择，系统和实施人员都不能猜测。
