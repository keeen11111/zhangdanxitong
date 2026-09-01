# 文档驱动 Excel Agent 设计规格（草案）

## 1. 目标

用户上传：

- 本公司的 Excel 总表（目标模板）；
- 一个或多个变更/来源 Excel；
- 本公司的操作手册、历史规则文档和会议录音转写稿。

系统先由大模型把手册和录音编译成**公司隔离、可版本化的规则包**，再由大模型持续主控工作簿画像、文件归类、字段映射、逐人处理、工具选择、结果检查、纠错和异常沟通。Excel 的物理读写通过受控工具完成；大模型不能执行任意 Python、Shell、宏或未声明公式，但不是只生成一次结构化决定后退出流程。

目标是让常规月份达到零人工或接近零人工；只有身份冲突、口径冲突、文件损坏/加密、无法验证的高风险金额等事项才中断。

## 1.1 当前可交付实现

当前版本已提供 `/api/agent/runs` 运行编排接口和 `/projects/{projectId}/agent`
工作台。运行会复用现有通用财务工作簿整合器，自动更新安全变更，并把每个异常
拆成独立 WorkItem；用户可以在单人对话中确认、保留、跳过、重试或回滚该人员，
最后显式发布正式版本。当前代码尚未接入真实模型 Provider；配置模型 API 后才具备完整 Agent 能力。没有配置模型 API 时，系统只允许文件检查、确定性预检、人工结构化决定和结果校验，页面必须显示配置阻断，禁止把降级运行描述为“大模型已处理”，并禁止自动发布。

已落地的安全边界包括：Pydantic Decision 白名单、公式单元格拒绝覆盖、租户哈希
隔离运行记录、月份不一致阻断确认、审计修改记录，以及公司级 Memory 的
`candidate -> active -> suspended` 生命周期。

## 2. 关键边界

以下“自动执行”均表示：大模型根据当前任务上下文选择并调用有类型的 Excel 工具，工具在权限、规则、公式保护和写入前置条件全部通过后执行。确定性逻辑是模型可调用的能力和安全护栏，不是取代大模型的另一条主流程。

### 自动执行

- 文件、Sheet、表格区块、字段和记录标识的画像与候选识别；
- 唯一高置信度的字段映射、覆盖、清空、复制、汇总和公式平移；
- 已由客户预处理的考勤结果直接使用；
- 可验证的月份滚动、固定福利、值班次数、高温费参数和人员台账补齐；
- 自动校验公式、格式、数据类型、金额平衡和来源追溯。

### 自动判断后执行

- 同义 Sheet/字段归并；
- 姓名、工号、身份证号的组合匹配；
- 结构化文本中的补发/补扣金额提取；
- 变更表中的人员类别描述差异判断；
- 多来源值的优先级判断。

这类判断必须输出证据、置信度和规则版本，并通过确定性校验；不满足校验时转异常。

### 必须中断或保留人工

- 同名、重复工号、身份证与工号冲突；
- 同优先级来源给出无法解释的不同金额；
- 目标字段或目标 Sheet 不唯一；
- 产假回岗等需要恢复历史状态且无法由规则验证；
- 无法解析的自由文本金额或算式；
- 总表加密、损坏、外链缺失或公式校验失败；
- 文件名、Sheet 日期、用户指定月份互相冲突且无法通过批次上下文消解。

## 3. 运行架构

```text
手册/录音转写稿
        |
        v
模型主控 -> 规则编译工具 -> 规则包校验 -> 公司规则版本库
   ^                                      |
   |                                      v
总表 + 来源表 -> 模型主控 <-> 有类型 Excel 工具层 <-> 工作簿草稿
                    |              |
                    |              +-> 受控读写/公式保护/原子保存/回滚
                    v
             独立验证器 -> 模型检查与限次纠错 -> 异常清单/审计/正式文件
```

组件职责：

1. `ModelOrchestrator`：理解资料和当次指令，维护运行上下文，逐人规划并选择工具，根据工具结果继续、纠错或提问。
2. `InstructionCompiler`：向模型提供 DOCX、转写稿的结构化内容和定位，并校验模型提取的事实、条件、动作、例外和冲突。
3. `WorkbookProfiler`：作为模型工具识别工作簿、Sheet、表格区块、字段类型、主键候选、公式和样式特征。
4. `PlanBuilder`：校验模型把规则包绑定到实际文件后生成的 `ExecutionPlan`，拒绝无证据或越权计划。
5. `RestrictedExecutor`：执行模型发起的白名单操作（copy/set/add/clear/count/formula/insert/validate），不执行模型生成的代码。
6. `Verifier`：独立校验公式、格式、金额、人员集合、月份和来源追溯；结果返回模型，由模型限次修正或升级人工。

### 3.1 模型主控与 Excel 工具闭环

大模型必须实际处理以下工作，而不是只输出最终 `Decision`：

1. 理解手册、录音转写、历史 Memory 和当次指令，按 `当次明确指令 > 生效规则包 > 最新手册 > 录音` 解决优先级。
2. 识别文件角色、工资所属月、发薪日、Sheet 语义、人员类别和源/目标字段关系。
3. 编译或选用公司规则包，生成按人员拆分的 `ExecutionPlan` 和 `WorkItem`。
4. 对每个 WorkItem 读取必要范围、匹配人员、提出变更并调用 Excel 工具应用。
5. 读取工具的结构化结果和验证报告，自检后继续、在同一 WorkItem 内限次纠错，或向用户提出一个明确问题。
6. 所有人员完成后调用独立验证，再根据公司策略申请人工发布或自动发布。

第一期工具契约：

| 工具 | 作用 | 关键约束 |
|---|---|---|
| `inspect_workbook` | 获取 Sheet、公式、命名区域、保护和月份画像 | 只读，返回文件哈希 |
| `classify_file` | 判断总表、来源表、手册、录音等角色 | 低置信度或月份冲突必须阻断 |
| `find_table` | 定位表头、数据区和字段候选 | 返回候选及置信度，不静默选歧义目标 |
| `read_range` | 读取处理所需的最小范围 | 服务端脱敏，禁止读取无关敏感列 |
| `match_person` | 匹配来源人员与目标人员 | 身份冲突不得由模型强行覆盖 |
| `propose_changes` | 生成带来源证据的改单集合 | 每项包含旧值、新值、目标、规则和理由 |
| `apply_cell_changes` | 在草稿副本应用经校验的改单 | revision/幂等键、公式保护、白名单操作 |
| `copy_formula_from_reference` | 从已存在参考单元格平移公式 | 不接受模型自由编写公式 |
| `validate_workbook` | 检查公式、格式、月份、人员、金额和追溯 | 独立于生成该计划的模型步骤 |
| `rollback_work_item` | 撤销单个人员的本轮变更 | 不影响其他已完成人员 |
| `publish_workbook` | 原子发布正式版本 | 必须满足发布策略、权限和零 blocker |

标准循环为：

```text
模型理解当前 WorkItem
  -> 选择有类型工具并提交参数
  -> 工具校验、执行并返回 observation
  -> 模型检查 observation 和证据
  -> 继续 / 修正 / 回滚 / 提问
  -> 独立验证通过后提交该 WorkItem
```

模型可以自主读取画像、选择 active 规则、处理唯一映射、应用可逆低风险变更和修复可验证的工具参数错误。身份冲突、目标不唯一、同优先级金额冲突、非法/负数金额、公式覆盖、规则来源冲突以及不满足公司自动发布策略的情况必须暂停询问或等待有权限的复核人。

每次模型和工具交互必须记录：`tenant_id`、`company_id`、`run_id`、`work_item_id`、模型供应商与模型名、提示词版本、规则包版本、工具名、参数摘要、来源证据、目标位置、旧值、新值、工具结果、验证结果、重试次数、操作者和时间戳。敏感原文和密钥不得进入普通日志。

模型输出必须经过 JSON Schema/Pydantic 校验。单个 WorkItem 的自动纠错次数受限；超过上限、重复产生相同失败或证据不足时转人工，不允许无限循环。没有 API Key 时保留上传、画像、预检、人工决定和验证工具，但停止在 `MODEL_CONFIGURATION_REQUIRED`，不得自动执行模型判断或自动发布。

## 4. 规则包模型

规则包是每家公司独立的，不把任何公司的文件名、Sheet 名、列号或口头习惯写进通用内核。没有随本次任务上传新手册时，系统使用该公司的 `active` 规则包；只有首次接入或规则包不存在时才要求提供规则来源。

```yaml
rule_package:
  id: company-keyuan-payroll
  tenant_id: keyuan
  version: 2026.08.1
  status: candidate       # candidate | active | retired
  effective_period: 2026-08
  source_evidence:
    - file: 0-详细手册-科园-需每月更新.docx
      locator: paragraph-12
    - file: transcript.txt
      locator: 00:02:05-00:02:55
  precedence: [user_explicit, active_rule, latest_manual, recording]

  glossary:
    employee_id: [工号, 员工编号, 员工号]
    id_number: [身份证号, 身份证号码]
    payroll_month: [所属月, 薪资月, 工资月份]

  identity:
    keys:
      - [employee_id]
      - [id_number]
      - [employee_id, normalized_name]
    reject_on_duplicate: true

  rules:
    - id: attendance.copy-preprocessed
      phase: attendance
      trigger: source.topic == attendance
      operation: copy_fields
      source: {fields: [迟到, 早退, 旷工, 次数, 时长]}
      target: {topic: attendance}
      params: {source_is_preprocessed: true}
      automation: A
      on_conflict: stop
      evidence: [transcript.txt@00:06:40-00:07:01]

    - id: duty.count-by-employee
      phase: allowance
      trigger: source.topic == duty
      operation: aggregate_count
      group_by: employee_id
      target_field: 值班次数
      automation: A
      on_conflict: stop

    - id: festival-benefit.month-policy
      phase: allowance
      trigger: payroll_month in [01, 04, 09, 12]
      operation: set_or_clear
      target_topic: payroll
      params:
        amounts: {"01": 1000, "04": 500, "09": 500, "12": 1000}
        only_if_standard_value_nonzero: true
        clear_previous_temporary_value_outside_target_month: true
        do_not_modify_standard_table: true
      automation: A
      on_conflict: stop
      evidence: [transcript.txt@00:00:26-00:02:55]

    - id: personnel.ignore-description-only-change
      phase: personnel
      trigger: company_unchanged and department_unchanged and data_unchanged
      operation: ignore_change
      params: {description_only_difference: true}
      automation: A
      on_conflict: stop
      evidence: [transcript.txt@00:10:14-00:11:46]

    - id: payroll.restore-returning-employee
      phase: personnel
      trigger: change_type == maternity_return
      operation: restore_from_reference_employee
      params:
        restore_formula: true
        restore_high_temperature_fee_rule: true
        reject_negative_net_pay: true
      automation: B
      on_conflict: manual
      evidence: [transcript.txt@00:08:18-00:09:25]

    - id: bonus.select-nonzero-then-category
      phase: bonus
      trigger: source.topic == bonus
      operation: select_value
      params:
        first_policy: exactly_one_nonzero
        fallback_policy: by_employee_category
        category_mapping_ref: keyuan.bonus_category_mapping
        both_zero: no_update
        both_nonzero: conversation_required
        text_or_invalid_number: conversation_required
        negative_value: conversation_required
      automation: A_to_B
      on_conflict: stop
      evidence: [user_instruction@2026-08-27, transcript.txt@00:12:25-00:15:32]

  validations:
    - id: formula.no-error-values
      check: no_formula_error_values
      severity: blocker
    - id: identity.unique
      check: identity_keys_unique
      severity: blocker
    - id: source.traceable
      check: every_change_has_source_locator
      severity: blocker
```

自动化等级：

- `A`：确定性规则，直接执行；
- `B`：模型解释/归类后执行，必须通过校验并写入证据；
- `C`：可逆、低风险的建议，默认执行并产生 warning；
- `D`：不可验证或高风险，停止并生成异常任务。

## 5. 规则编译流程

1. 文档解析：DOCX 提取段落、表格、标题；录音使用已有转写稿，保留时间戳。
2. 事实抽取：模型生成 `Fact/Condition/Action/Exception` 并主动处理冲突；每条都必须带来源定位并通过结构校验。
3. 术语归一：把“工资核算”“工资总表”等映射为公司内的主题和字段别名。
4. 冲突检测：不同版本手册、录音与表内日期不一致时生成 `RuleConflict`，不静默覆盖。
5. 规则生成：输出 JSON Schema 校验过的规则包；禁止自由代码和未声明字段。
6. 历史回放：用至少 3 个月已完成文件运行 shadow mode，与已交付结果逐单元格比较。
7. 激活：只有回放通过且冲突已解决，规则包才从 `candidate` 变成 `active`。

规则包不需要先做模型微调。优先使用“公司规则包 + 规则版本 + 已确认决定记忆”；微调只有在样本量和收益都明确后再评估。

## 6. 月度零接触流程

```text
一次上传全部文件
  -> 按文件夹/任务建立 Job Manifest
  -> 自动识别总表和来源角色
  -> 加载公司 active 规则包（本次新指令优先）
  -> 拆成“每人一个 Work Item”，每项独立规划、执行、校验和提交
  -> A 类自动执行；B/C 类由 Agent 判断后执行并留证据
  -> D 类才进入人工队列，其他人员继续处理
  -> 全局公式/格式/金额/月份/来源校验
  -> 输出草稿、审计报告、异常清单
```

每个人的 `WorkItem` 至少包含：人员身份、来源行、目标 Sheet/单元格、规则列表、旧值、新值、公式影响、置信度、执行状态和校验结果。单个人失败不会丢弃其他人的结果；重试、回滚和再次校验只针对该人员。全局校验失败时才冻结发布。

“人工发布”可以作为部署策略，而不是业务规则：默认保留严格模式；在公司规则包稳定、连续回放通过后，可对无 `D` 类事项的批次开启 `zero_touch` 自动发布。该开关必须按公司隔离、可审计、可回滚。

## 7. 当前科园样本的规则归类

可直接固化为 A 类：OA 月份滚动、考勤字段映射、值班次数、高温费参数、年节福利月份与金额、已预处理的迟到/早退/旷工、新员工台账补齐、描述差异忽略。

可做 B/C 类：补发补扣结构化文本、产假回岗恢复、绩效来源归并、奖金字段写入目标列。

已确认奖金基础规则：J/K 恰好一列为有效非零数值时取该列；两列都为零或空时不更新；两列均非零、包含文本、负数或非法值时按人员类别 active 规则处理，仍不能唯一判断则询问。当前不能猜测：月份证据冲突的最终口径、同优先级手册冲突的权威版本，以及未提供或无法解密的总表密码。

## 8. 验收指标

- `auto_apply_rate`：自动落地的变更占比；
- `hard_block_rate`：D 类阻断占比；
- `rule_replay_match_rate`：与历史正式结果的单元格匹配率；
- `formula_integrity_rate`：公式/格式校验通过率；
- `source_trace_coverage`：变更可追溯覆盖率；
- `rollback_rate`：发布后回滚率。

首期不预设“100% 无人工”，先用历史样本测基线，再按公司规则包版本持续降低 D 类占比。

## 9. 前端 Agent 工作台（替代旧的固定 Python 步骤）

面向重度财务用户，主界面不是聊天框，而是可暂停、可追溯的处理队列：

1. **新建任务**：拖入文件，Agent 自动标记“总表/来源/规则文档”；允许用户一键纠正角色。
2. **规则与计划**：显示本次使用的公司、规则包版本、月份、来源优先级和预计影响 Sheet；只让用户确认真正冲突项。
3. **逐人处理队列**：显示总人数、已完成、异常、跳过、失败；点击某人可看来源证据和单元格前后差异。
4. **异常中心**：只展示 D 类事项；支持“采用来源/保留原值/编辑后写入/忽略/指定其他目标”，决定会沉淀为该公司的规则建议。
5. **结果与发布**：展示公式、金额、格式、月份和追溯校验；符合公司 `zero_touch` 策略时自动发布，否则保留一键发布。

后端按人员拆分处理，但提供“自动连续执行”按钮；财务人员不需要逐人点击，只有需要关注的人才进入人工界面。旧的六步流水线保留为内部兼容适配器，不再作为前端交互模型。

建议接口：

```text
POST /api/agent/runs
GET  /api/agent/runs/{run_id}
POST /api/agent/runs/{run_id}/plan
POST /api/agent/runs/{run_id}/execute
GET  /api/agent/runs/{run_id}/items?status=...
POST /api/agent/runs/{run_id}/items/{item_id}/retry
POST /api/agent/runs/{run_id}/items/{item_id}/resolve
POST /api/agent/runs/{run_id}/pause
POST /api/agent/runs/{run_id}/publish
```

## 10. 当前两个任务的 Job Manifest（草案）

| Job | 目标总表 | 来源文件 | 规则来源 | 备注 |
|---|---|---|---|---|
| `keyuan-202607-payroll` | `1/202608（所属月202607）...xlsx` | `1/薪资数据-科园-7月薪资（8.14发薪）.xlsx` | 根目录两份科园手册；无新文档时使用 active 规则包 | 周期一致，可作为明天首个交付样本 |
| `keyuan-202606-payroll` | `2/202606（所属月202605）...xlsx`（待确认） | `2/薪资数据-科园-6月薪资（7.15发薪）.xlsx`、`鹤安长泰2026.06.xlsx`、`益药科园2026.06.xlsx` | 根目录两份科园手册；无新文档时使用 active 规则包 | 来源表明确属于 2026-06，当前目标总表显示 202605，预检必须先确认 |

没有操作文档的任务按“数据更新模式”执行：不重新编译规则，只加载该公司的 active 规则包，仍执行完整画像、逐人处理和校验。

## 11. 对话式处理异常

异常不再要求用户打开另一张人工 Excel 清单逐项填写。Agent 以“一个人、一个问题、一个决定”的节奏对话：

```text
Agent：李某的奖金在来源列 J 为 2,000，来源列 K 为空。
      按该人员类别“店员”的规则，应采用哪一列？
用户：按店员规则取 J 列。
Agent：已记录“店员 -> J 列”，将应用于李某，并作为科园规则建议。
```

自然语言先由模型结合证据和会话上下文解析，再形成结构化 `Decision` 和有类型工具调用；原始自然语言不能绕过结构校验直接传给 Excel 执行器：

```json
{
  "item_id": "work-item-123",
  "decision": "select_source",
  "value": "J",
  "scope": "tenant_rule_candidate",
  "reason": "用户明确指定店员取 J 列",
  "confidence": 1.0
}
```

前端同时提供按钮选项（采用来源、保留原值、编辑后写入、忽略），自然语言只是快捷入口。Agent 复述将执行的结果，用户确认后才提交该人员；确认结果先写入租户隔离的经验库，经过至少一次历史回放后才升级为 `active` 规则。

为了减少错误和成本，不为每个单元格单独开启一次模型调用。模型仍主控 A/B/C/D 全部 WorkItem：先批量读取画像和规则命中结果，再按人员处理必要上下文；A 类通常直接调用工具，B/C 类判断后调用工具，D 类发起对话。身份证、银行卡等敏感字段只在服务端匹配，不发送给模型；发送给模型的是脱敏后的字段名、候选值、来源定位和必要上下文。

## 12. 经验库晋级规则

经验库不保存姓名、工号、身份证号等个人条件，只保存可复用的公司级语义条件，例如“奖金主题 + 候选列 J/K + 人员类别 店员”。

一条用户决定的生命周期：

```text
首次确认 -> candidate
同一公司连续月份一致 -> active
出现冲突或用户推翻 -> suspended，重新进入对话
```

`candidate` 可以作为 Agent 建议，但不能自动扩大适用范围；`active` 才能在满足条件时自动执行。经验规则必须记录来源、适用范围、版本、最近使用时间和冲突次数。

## 13. 最快交付架构

### 明天先做

- 复用 `financial_workbooks.py` 的上传、总表选择、无损写入、审计和发布；
- 新增 `agent_runs` 内部状态（先复用现有项目/会话存储，避免明天做数据库迁移）；
- 新增一个 Agent 入口：生成规则摘要、执行计划和逐人 `WorkItem`；
- 自动执行 A 类；B/C/D 类通过一个对话接口逐项解决；
- 前端替换 `MasterUpdatePage` 的固定步骤为“任务-计划-逐人进度-异常对话-结果”；
- 规则先写入现有租户隔离经验库，第二阶段再迁移为正式 `rule_packages`、`rule_versions`、`decisions` 数据表。

### 第二阶段再引入

- Pydantic AI：如果需要更快获得类型化 Agent、模型供应商切换和 FastAPI UI 事件流，可在当前 FastAPI 上接入；
- OpenAI Agents SDK：如果需要内置工具审批、`RunState` 暂停/恢复、guardrails 和 tracing，可替换当前轻量编排器；
- LangGraph：当任务需要跨天暂停、队列恢复、复杂分支和长期记忆时再引入，不作为明天交付的前置依赖。

## 14. 调研依据

- OpenAI Agents SDK：支持工具审批、暂停/恢复和可编程审批。<https://openai.github.io/openai-agents-python/human_in_the_loop/>；guardrails：<https://openai.github.io/openai-agents-python/guardrails/>。
- Pydantic AI：类型化 Agent、供应商切换和 FastAPI UI 事件流。<https://pydantic.dev/docs/ai/core-concepts/agents/>；<https://pydantic.dev/docs/ai/integrations/ui/overview/>。
- LangGraph：持久执行、interrupts 和长期状态。<https://docs.langchain.com/oss/python/langgraph/durable-execution>；<https://docs.langchain.com/oss/python/langgraph/interrupts>。
- Microsoft MarkItDown：文档到结构化文本的解析参考。<https://github.com/microsoft/markitdown>。
- Excel 原子工具参考：<https://github.com/SylvianAI/sv-excel-agent>、<https://github.com/Bread-Technologies/Bread_Excel_Agent>。两者只借鉴工具边界，不直接授权模型任意修改工资总表。
- Next.js + LangGraph 工作台参考：<https://github.com/agentailor/fullstack-langgraph-nextjs-agent>。只参考流式状态和审批交互，不整体搬入当前项目。

## 15. 已确认的开放决策

1. 允许无 D 类事项时按公司配置自动发布。
2. 绩效/奖金先自动取唯一非零列；两列都有数值时再按人员类别择一；具体类别到来源列的映射仍需在规则包中配置和回放验证。
3. 文件名、Sheet 日期和发薪日期是每月更新的月份元数据；冲突时先用批次上下文推断，无法唯一确定才阻断。
4. 来源优先级为：`当次明确指令 > 生效规则包 > 最新手册 > 录音`。
5. 总表密码允许首次输入，并以租户密钥安全保存。
6. **待确认**：第二个任务是否应将目标总表改为根目录的 `202607（所属月202606）...xlsx`？
7. **待确认**：经验规则达到几次连续月份一致后，从 `candidate` 自动晋级为 `active`？建议默认 3 次。

## 16. 登录后的公司选择

进入系统后先显示当前用户有权限处理的公司列表，不直接进入某个月的 Excel 页面。公司卡片/表格显示：公司名称、最近处理月份、最近一次运行状态、待处理异常数、生效规则包版本和最近发布人。

选择公司后才加载：

- 该公司的 `active` 规则包；
- 该公司的历史任务和规则建议；
- 该公司的发布策略（`strict_review` / `zero_touch`）；
- 该公司的加密文件凭据引用；
- 该公司的审计范围。

每个 Agent 请求必须携带服务端确定的 `company_id` 和 `run_id`，前端不能只靠提示词指定公司。用户切换公司时清空当前对话上下文，禁止把 A 公司的规则、文件或人员数据带到 B 公司。

当前项目的 `Tenant` 可先作为公司选择的来源；当一个客户账号管理多个法人主体时，再新增 `CompanyProfile`（隶属于 Tenant），不要把公司名称写死在前端。

## 17. 企业财务规范与权限

- **操作员**：上传文件、发起任务、与 Agent 处理当前批次异常。
- **规则管理员**：查看规则证据、启用/暂停规则、配置人员类别映射。
- **复核/发布人**：查看校验报告、发布正式稿、回滚版本。
- **审计只读**：查看规则、对话、改单元格和版本归档，不可修改。

高风险操作必须服务端授权：修改规则、改变发布策略、覆盖金额、发布正式稿、下载含敏感数据的文件。所有操作记录用户、公司、任务、规则版本、模型版本、来源定位、旧值、新值和文件哈希。

## 18. 前端模板复用策略

当前前端是 Next.js 14.2.5 + React 18.3.1 + Tailwind 3.4 + Radix + Lucide。候选开源模板的兼容性结论：

| 参考项目 | 结论 |
|---|---|
| `shadcn-ui/ui` | 适合继续复用可访问组件和后台 blocks，MIT；与现有 Radix 体系方向一致。 |
| `Kiranism/next-shadcn-dashboard-starter` | 结构和数据表值得参考，但当前是 Next 16/React 19/Tailwind 4，不能整包搬入。 |
| `satnaing/shadcn-admin` | 视觉和表格可参考，但它是 Vite + React 19，不能作为当前路由骨架。 |
| `TailAdmin/free-nextjs-admin-dashboard` | 页面较全，但当前依赖 Next 16/React 19，且包含不需要的营销/图表模块。 |
| `abderrahimghazali/shadcn-fintech` | 可参考金融数据密度，但视觉偏展示型，不适合工资核算主工作台。 |

实施选择：沿用当前组件和颜色变量，吸收 `shadcn-ui` 的 sidebar、command、table、dialog、tabs 结构；不引入新模板运行时，不迁移 Next/Tailwind 大版本。页面采用浅色、高密度、低动画的财务工作台风格：固定左侧导航、顶部公司切换器、主体任务表格、右侧 Agent 对话面板。

## 19. 公司工作台线框

```text
┌─────────────────────────────────────────────────────────────────────┐
│ 公司切换器 [北京科园 ▼]   处理月份 [2026-07 ▼]   规则 v2026.08.1  用户 │
├───────────────┬──────────────────────────────────┬──────────────────┤
│ 任务导航       │ 当前任务 / 人员处理队列            │ Agent 对话        │
│                 │ ┌ 总人数 ┐ ┌ 自动 ┐ ┌ 异常 ┐       │ 当前人员：李某     │
│ 公司概览        │ │ 42     │ │ 39   │ │ 3    │       │ 来源证据          │
│ 月度任务        │ └────────┘ └──────┘ └──────┘       │ 修改前/后值       │
│ 规则中心        │ 状态 | 人员 | 规则 | 目标 | 结果     │ [采用J] [采用K]    │
│ 发布归档        │ ...                                │ 输入自然语言      │
│ 审计记录        │ [暂停] [继续自动处理] [导出报告]    │ [发送]            │
└───────────────┴──────────────────────────────────┴──────────────────┘
```

移动端不强行压缩宽表：人员队列切换为卡片，字段差异和来源证据进入抽屉；桌面端保留固定表头、筛选、批量选择和键盘操作。
