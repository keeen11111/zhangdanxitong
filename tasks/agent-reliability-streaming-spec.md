# 财务 Agent：可靠对话与正式稿交付规范

## 目标

让用户能在同一任务中可靠地发送和接收对话；回复、工具进度和执行状态应实时显示。完成服务器验收后，页面只显示“下载正式稿”按钮。未决事项不得生成或下载正式稿。

## 已确认的产品决定

- 使用现有 OpenAI 兼容模型配置，不替换为 Vercel AI SDK。
- 保留 Python/FastAPI 的 Excel 工具和现有认证、项目数据。
- 不提供草稿下载入口；只有验收通过的正式稿可下载。
- 模型或工具耗时不能被前端 20 秒总超时中断。

## 流式接口契约

新增 `POST /api/agent/runs/{run_id}/message/stream`。请求体为 `{ "message": string }`；保持 Bearer 认证。

响应为 `text/event-stream`。事件具有递增 `revision`、可持久化 `event_id` 和以下类型：

- `message.accepted`：服务端已保存用户消息。
- `message.delta`：Agent 回复的增量文本。
- `tool.started` / `tool.completed`：安全的工具进度摘要，不包含密钥、完整提示词或文件原文。
- `run.status`：计划、执行、待确认、验收和发布状态。
- `message.completed`：回复结束，可用作重连后的完整状态。
- `error`：结构化、可展示且不泄露内部细节的失败原因。

前端使用带 Authorization 请求头的 `fetch` + `ReadableStream` 读取 SSE，而不是原生 `EventSource`；连接失败允许重连，已持久化事件按 revision 去重。取消请求只取消显示，不撤销已经持久化的用户消息。

## 交付状态

`PLANNING → EXECUTING → AWAITING_CONFIRMATION | VALIDATING → READY_FOR_FINAL → PUBLISHED`

- `AWAITING_CONFIRMATION`：显示人工确认卡片，禁用正式稿下载。
- `READY_FOR_FINAL`：显示一个“下载正式稿”按钮。
- `PUBLISHED`：保留同一按钮，并显示已发布版本信息。
- 任何失败均显示“重试”与明确原因；不丢失用户已发送的消息。

## 实施任务

1. 为流事件编码、SSE 格式和安全脱敏建立后端单元测试；新增流式消息端点。
2. 为前端流读取和重连/去重逻辑建立逻辑测试；移除聊天路径的 20 秒总超时。
3. 将工作台改为即时显示用户消息、增量回复、工具状态和可恢复错误状态。
4. 收紧正式稿下载门槛与按钮状态，补齐“未决项 / 验收失败 / 验收通过”回归测试。
5. 运行后端测试、前端逻辑测试、lint、生产构建和本地端到端验证。

## 完成标准

- 后端短暂不可用时，用户消息不丢失且页面可重试。
- 模型回应超过 20 秒时，对话不会被前端主动中止。
- 用户能看到逐段回复和处理状态，而不是等待完整长报告。
- 有未决项时无法下载正式稿；验收通过时能下载正式稿。
