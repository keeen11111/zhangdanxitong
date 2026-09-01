# 规格：安全财务账单批次（第一阶段）

## 目标

让企业财库人员将本期来源文件更新到公司财务总表后，只获得“草稿”；正式账单必须由人工明确发布。系统以阻止不确定结果流出为第一目标。

## 用户与工作流

操作人员上传一份通用财务总表和多份通用来源文件，发起整合，得到草稿、自动更新审计和人工事项。复核人员确认该批次通过后点击“发布正式账单”；在此之前，系统不得将该结果作为正式文件下载。

```text
来源文件 + 总表
  → 安全整合
  → 草稿 / 待人工处理 / 阻止发布
  → 发布前控制校验
  → 人工发布
  → 正式账单下载
```

## 第一阶段范围

- 每次整合结果拥有一个状态：`review_required`、`ready_for_release`、`published` 或 `blocked`。
- 有任何人工事项时为 `review_required`，不得发布。
- 没有可处理的来源 Sheet 时为 `blocked`，不得发布。
- 只有通过控制校验且无人工事项的草稿才能由人工发布。
- 草稿可单独下载；正式下载端点只允许已发布结果。
- 发布操作、发布时间和发布前校验结果记录在该项目当前批次元数据中。
- 发布时生成独立的正式版本档案，记录发布人、发布时间和归档文件 SHA-256；历史版本不受之后上传或重新整合影响。

## 不在本阶段范围

- 多人审批流、审批人身份、双人复核和不可篡改审计表。
- 数据库迁移、多人审批、多实例并发锁，以及数据库级不可篡改审计表。
- 未处理人工事项的在线编辑与重新整合闭环。

## 接口约定

- `POST /api/financial-workbooks/{project_id}/integrations`：生成草稿。
- `POST /api/financial-workbooks/{project_id}/integrations/latest/release`：人工发布，仅接受可发布草稿。
- `GET /api/financial-workbooks/{project_id}/integrations/latest/draft/download`：下载草稿。
- `GET /api/financial-workbooks/{project_id}/integrations/latest/download`：仅下载已发布正式账单。
- `GET /api/financial-workbooks/{project_id}/integrations/published`：查看正式版本历史。
- `GET /api/financial-workbooks/{project_id}/integrations/published/{version_id}/download`：下载指定正式版本。
- `GET /api/financial-workbooks/work-queue`：汇总当前用户需人工复核、需发布或被阻断的真实批次；按风险和最新批次排序。

## 验收标准

1. 有未匹配、公式冲突或其他人工事项时，结果不能发布。
2. 空的或无法识别的来源 Sheet 不能被默默忽略，必须阻止发布。
3. 无人工事项且存在已确认映射的结果可由人工发布；发布后才可下载正式账单。
4. 自动更新审计、发布状态和控制校验结果能由 API 返回。
5. 发布后更改当前上传文件不会影响历史正式版本；归档文件校验不一致时必须拒绝下载。

## 工程约束

- 后端测试：`python -m pytest -q`
- 前端检查：在 `frontend/` 执行 `npm run lint` 与 `npm run build`
- 不新增依赖、不修改数据库 schema；数据库化审批和版本历史必须先取得确认。
