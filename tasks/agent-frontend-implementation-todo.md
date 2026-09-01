# 财务 Excel Agent 实施清单

## Checkpoint 1：契约与恢复

- [x] T01 固定 Run/Event/WorkItem/Decision 契约和错误码
  - 验收：前后端不再兼容多种临时结构；状态机测试覆盖全部转换。
  - 文件：`backend/routers/agent.py`、`core/document_agent/contracts.py`、`frontend/src/lib/api.ts`、`frontend/src/lib/agent-workflow.ts`。
- [x] T02 增加运行列表、最新运行恢复和事件查询
  - 验收：刷新 `/agent?run=` 后恢复相同进度和人员项。
  - 文件：Agent router、API client、`use-agent-run.ts`、API tests。
- [x] T03 增加 revision/idempotency 与重复提交测试
  - 验收：双击“应用建议”只产生一次改单元格记录。

## Checkpoint 2：Agent 对话主流程

- [x] T04 建立三栏 `AgentShell`、顶部状态和响应式抽屉
  - 验收：1440/1024/768/375 四种宽度无重叠和横向溢出。
- [x] T05 实现类型化时间线消息与批量折叠
  - 验收：100 人自动处理只显示一条进度摘要，异常逐人展示。
- [x] T06 实现固定 composer、附件、初始 instruction 和运行模式
  - 验收：指令真实传到 `POST /runs`，附件上传状态可见。
- [x] T07 实现人员确认、快捷键和自动下一项
  - 验收：应用/保留/跳过后自动定位下一待确认人员。
- [x] T08 实现右侧证据、规则、审计和 Memory
  - 验收：每项展示来源位置、旧值、新值和规则版本。

## Checkpoint 3：全局页面闭环

- [ ] T09 项目列表改为月度任务语义并统一 Agent 入口
- [ ] T10 待办链接改为精确定位 Agent run/item
- [ ] T11 结果与发布回到 Agent 时间线和归档视图
- [ ] T12 清理重复导航和旧流程默认回跳

## Checkpoint 4：真实样本

- [ ] T13 跑通 2026-07 科园样本并逐单元格对比已完成总表
- [ ] T14 阻断 2026-06 月份冲突，等待目标总表确认
- [ ] T15 验证 J/K 唯一非零、双非零、双零、文本和负数分支
- [ ] T16 验证公式、格式、Sheet、月份、人员集合和来源追溯

## Checkpoint 5：生产版多公司

- [ ] T17 引入 Alembic 并新增 Company/Project 归属迁移
- [ ] T18 公司列表、公司切换和数据隔离
- [ ] T19 规则中心、文件中心、历史和设置页面
- [ ] T20 DOCX/录音规则编译与证据审阅（已具备编译 API，录音仍需外部转写，证据审阅 UI 待完成）
- [ ] T21 持久任务队列、暂停恢复和进程重启恢复
- [ ] T22 权限、凭据加密、脱敏和下载审计

## 最终验证命令

- [ ] `pytest -q`
- [ ] `npm run lint`
- [ ] `npm run build`
- [ ] `git diff --check`
- [ ] 浏览器桌面/移动截图、DOM、console、键盘操作验证
- [ ] 两份真实工作簿的输出、审计和源文件不变性对比
