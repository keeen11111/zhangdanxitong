# Spec: 财务 Agent 可恢复端到端工作流

## Objective

财务人员上传一份原始总表、一份或多份更新表及可选操作文档后，Agent 先解析文档，再规划并在独立副本中执行。页面实时显示读取、工具调用、写入和校验过程；只在真实歧义处暂停提问。完成后进入统一结果视图，显示修改清单并允许下载更新后工作簿。

## Confirmed Assumptions

1. 上传原件永不覆盖，所有写入只发生在服务端独立结果副本中。
2. 总表和至少一份更新表齐备后可自动开始；上传的手册/说明先于计划被解析。
3. 明确、可验证、工具支持的步骤不需要额外“确认计划”；月份冲突、多个同等候选、缺少必要来源或会覆盖公式时必须提问。
4. 网页刷新、SSE 断线、模型暂时失败或单次模型轮次耗尽不能丢失进度；工作流从已持久化检查点续跑。
5. 结果页展示实际工具写入产生的修改清单，不把模型文字当成写入证据。

## Interfaces

- `GET /api/agent/runs/{run_id}/events/stream?after_revision=N`
  - `text/event-stream`，支持事件 `id`、心跳和断线续传。
  - 事件类型为现有 Agent 事件的向后兼容扩展。
- `POST /api/agent/runs/{run_id}/process`
  - 幂等地启动或续跑后台工作流，立即返回当前 run，不让浏览器请求挂住数分钟。
- `GET /api/agent/runs/{run_id}/output`
  - 返回结果文件、校验状态、修改数量和分页修改明细。
- 保留现有 `/plan`、`/execute`、`/events`、`/acceptance/download` 以兼容旧页面和已保存运行。

## Workflow States

`planning -> processing -> awaiting_review | completed | failed`

- `planning`：解析材料、建立只读计划。
- `processing`：复制原表、读取来源、写入、校验。
- `awaiting_review`：存在必须由人确认的问题，已完成的安全写入保留。
- `completed`：已有结果文件且服务端校验通过，可查看与下载。
- `failed`：不可自动恢复的明确失败，必须包含错误码、失败阶段和重试入口。

## Reliability Rules

- 每一个工具调用前后都持久化事件和检查点。
- 工作流不依赖单个 HTTP 请求寿命；后台工作线程从磁盘重新加载 run，不长期共享请求中的可变对象。
- 模型轮次分段有上限时，写入已持久化检查点并自动开启下一段；不再把“超过若干轮”当成整个流程的终点。
- 持久化心跳时间；超过阈值的 `processing` 运行在下次打开或显式续跑时标记为可恢复，而不是永久卡住。
- 后台任务必须在 `finally` 中落盘终态或可恢复状态。

## Result View

- 结果文件名、处理时间、校验状态。
- 按 Sheet 分组的修改明细：单元格/行、旧值、新值、来源文件/依据。
- “查看处理结果”进入页内结果视图；“下载更新后表格”下载经服务端验收的副本。
- 如有未决项，结果视图明确标注“草稿/待确认”，不提供“正式稿”误导性文案。

## Testing Strategy

- Python 单元测试：事件流续传、幂等启动、轮次分段续跑、卡住恢复、结果摘要。
- TypeScript 单元测试：SSE `id/event/data`、心跳、分片/多行数据、断线续传。
- 集成测试：最小工作簿 + 更新表 + 手册，实际产生可重新打开的 `.xlsx`，修改清单与单元格一致。
- 浏览器验收：观察网络流、刷新恢复、进度、提问、结果视图和下载，控制台无错误。

## Commands

- Backend focused: `python -m pytest -q tests/test_agent_events.py tests/test_document_agent_core.py tests/test_agent_api.py`
- Backend full: `python -m pytest -q`
- Frontend tests: `npm run test:agent`
- Frontend type check: `npx tsc --noEmit`
- Frontend build: `npm run build`

## Boundaries

- Always: 校验租户权限、文件摘要、修改证据；保留旧 API；不覆盖原件。
- Ask first: 新增外部依赖、数据库 schema 变更、自动发布未验收结果。
- Never: 用模型文字代替实际工具写入；猜测金额/人员匹配；因超时或轮次用完而静默停止。

## Success Criteria

1. 一次上传后无需再点“确认计划”即可执行无歧义处理。
2. 实际工具调用在执行期间逐条出现，不是结束后一次性显示。
3. 中途刷新或断开流式连接后，从最后 revision 继续，无事件重复或缺失。
4. 人为模拟模型临时失败和分段轮次耗尽后，运行保留已完成写入并可续跑到终态。
5. 结束时页面展示结果视图、修改清单和可用的下载按钮；下载文件可被 openpyxl 重新打开。
