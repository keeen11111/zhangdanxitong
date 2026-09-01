# 文档驱动 Excel Agent 实施计划（待确认）

## 阶段 0：确认规格与规则边界

- 确认 `docs/document-driven-excel-agent-spec.md` 的自动化等级、来源优先级和零接触发布策略。
- 已确认：规则按公司隔离沉淀；奖金按人员类别择一；当次指令优先；允许零接触发布；密码可首次输入并安全保存。
- 待确认：第二个任务的目标总表是否与 2026-06 来源匹配。

## 阶段 1：规则包与编译器

- 定义 Pydantic/JSON Schema：`RulePackage`、`Rule`、`Fact`、`RuleConflict`、`ExecutionPlan`。
- 实现 DOCX/转写稿解析、证据定位、冲突检测和规则包校验；大模型主动理解和编译，校验器拒绝无证据产物。
- 验证：同一份材料重复编译结果稳定；无来源证据或自由代码的规则包被拒绝。

## 阶段 1A：模型主控与工具契约

- 接入可配置模型 Provider；API Key 最后由用户通过环境变量或密钥管理配置，代码和日志不保存明文。
- 定义 `ToolCall`、`ToolResult`、`ModelTrace` 和 `AgentEvent`，让模型在 Run 内持续理解、选择工具、读取 observation、纠错或提问。
- 提供 `inspect_workbook`、`classify_file`、`find_table`、`read_range`、`match_person`、`propose_changes`、`apply_cell_changes`、`copy_formula_from_reference`、`validate_workbook`、`rollback_work_item`、`publish_workbook`。
- 允许模型自主执行只读操作和 active 规则下可逆低风险操作；身份、金额、公式、月份和发布边界仍由工具与独立验证器强制执行。
- 无 API Key 时运行停在 `MODEL_CONFIGURATION_REQUIRED`；只允许预检、人工结构化决定和验证，不得自动发布或显示为模型处理完成。
- 验证：模型调用链可重放；每个修改均关联模型/提示词/规则版本、工具调用、证据、旧值、新值和验证结果；重复失败达到上限后转人工。

## 阶段 2：通用工作簿运行时

- 复用 `sheet_mapper.py`、`excel_engine.py`、`unified_integration.py` 建立画像、映射和无损执行适配层。
- 将工资逻辑保留为 `payroll` 适配器；大模型通过有类型工具调用通用入口，入口只执行规则包声明的操作。
- 验证：不同 Sheet 名、列顺序和表头仍能生成稳定映射；非目标 Sheet、公式和格式不变。

## 阶段 2A：明日交付最小切片

- 只接入 `keyuan-202607-payroll` 和确认后的第二个 Job Manifest。
- 接入真实模型主控循环和可配置 Provider，不把确定性演示队列作为最终 Agent。
- 前端提供新建任务、规则计划、逐人进度、异常中心和结果发布五个页面状态。
- 后端每个人生成独立 Work Item，支持自动连续执行、单人重试和单人回滚。
- 验证：两份素材均能一次上传、逐人完成、输出修改总表和审计报告；无高风险异常时按配置自动发布。

## 阶段 3：模型逐人处理与公司记忆

- 实现 `ModelOrchestrator` 主控 A/B/C/D 全部 WorkItem：A 类直接选工具，B/C 类判断并调用工具，D 类通过“一个人、一个问题、一个决定”的对话处理。
- 两列奖金恰好一列非零时自动取非零列；两列都有数值时通过人员类别规则或对话处理。
- 把用户确认沉淀为租户隔离的规则建议，经过连续月份一致和历史回放后才能激活。
- 验证：身份冲突、金额冲突和公式风险不能被模型绕过；已确认规则可在下月自动复用。

## 阶段 4：零接触策略与验收

- 增加按租户配置的 `strict_review` / `zero_touch` 发布策略。
- 用科园至少 3 个月和另一家公司至少 1 组脱敏样本做 shadow mode 回放。
- 验证：自动率、阻断率、回放匹配率、公式完整率和追溯覆盖率均有报告。

## 明日最小交付顺序

### 已完成切片（当前工作树）

- [x] 公司级 Memory 生命周期与租户隔离 API。
- [x] Agent Run、计划、逐人 WorkItem、对话、应用、重试、发布 API。
- [x] 结构化 `RulePackage` / `ExecutionPlan` / `Decision` 契约与 J/K 唯一非零选择器。
- [x] 登录后项目列表直接进入 Agent 工作台；旧工资流程保留为兼容入口。
- [x] 无模型 API 时的预检降级和演示队列提示；该状态不得自动发布。
- [ ] 真实模型 Provider、主控循环、工具调用审计与限次重试。

1. 只改 Agent 任务入口和前端工作台，不删除旧页面；历史路由继续可访问。
2. 先跑通任务 1：2026-07 来源更新到 202608（所属月 202607）总表。
3. 用相同流程处理任务 2，但在执行前必须确认目标总表月份。
4. 模型逐人处理 A/B/C/D 类 WorkItem：通过白名单工具执行可确定项，通过对话逐人解决异常，读取全局校验后纠错或发布。
5. 用现有 pytest 与两份输出工作簿做回归，不在明天引入 LangGraph 或数据库迁移。

## 前端改造切片（明日）

- `CompanyPickerPage`：登录后首先选择公司，显示最近月份、运行状态、异常数和规则版本。
- `AgentRunPage`：选择公司后创建月度任务并上传总表/来源/手册；使用服务端返回的公司上下文。
- `AgentPlanPanel`：展示规则版本、月份判定、影响 Sheet 和自动化范围。
- `AgentWorkItemQueue`：按人员显示处理状态，支持筛选、暂停、继续、单人重试。
- `AgentConversationPanel`：显示一个人员一个问题，支持按钮和自然语言；提交结构化 Decision。
- `AgentPublishPanel`：校验汇总、自动发布开关、正式版本下载和审计入口。

复用当前 Radix/Lucide/Tailwind 组件，不迁移 Next/Tailwind 大版本。参考 `shadcn-ui/ui` 的 sidebar、command、table、dialog、tabs 结构；参考 Kiranism 的数据密集表格，但不复制其 Next 16/React 19 依赖。

## 企业财务验收

- 切换公司后，页面和 Agent 上下文完全隔离；无法读取另一公司的文件或规则。
- Agent 每次回答都能显示来源文件、Sheet、行号和规则版本。
- Agent 能主动调用 Excel 工具、读取执行结果并在受控范围内修正，不是只输出一次结构化决定。
- 用户不处理普通人员；只在异常人员上对话，决定可复用但默认不自动扩大范围。
- 任何金额覆盖、规则启用、自动发布和下载都写入审计。
- 失败人员可单独重试，不能造成已完成人员重复写入。
- 桌面端支持固定表头、筛选、键盘导航；移动端使用抽屉展示差异，不产生横向页面溢出。

## 依赖与顺序

阶段 0 必须先获得业务口径确认；阶段 1 的接口确定后，阶段 2 和阶段 3 才能并行；数据库 schema、Agent SDK 和新依赖在单独确认后再引入。
