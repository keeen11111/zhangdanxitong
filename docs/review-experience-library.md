# 人工审核经验库

## 目标

经验库让财务人员只处理首次出现或无法确定的差异。它会保存已确认的处理理由，并在下次出现满足相同条件的差异时给出建议；系统不会据此自动确认、忽略或改写任何薪资数据。

## 使用流程

1. 完成模块 A-C，得到待审核差异。
2. 查询 `GET /api/pipeline/{project_id}/experience-suggestions`，查看当前待处理项命中的历史建议。
3. 人工核对后，通过 `POST /api/pipeline/{project_id}/experiences` 确认或忽略当前项，同时选择可复用的匹配字段并写明处理理由。
4. 后续项目再次遇到相同条件时，系统返回建议；仍由人工决定是否采纳。

## 创建经验的请求示例

```json
{
  "diff_type": "salary_change",
  "item_index": 0,
  "match_fields": ["变动项", "source"],
  "decision": "confirmed",
  "note": "已核验正式调薪通知。"
}
```

`match_fields` 应选择业务条件，例如“变动项”“来源”“原部门”。工号、姓名、身份证号等个人标识不会被写入经验条件；如果只选择这些字段，接口会拒绝保存。

## 接口

| 接口 | 作用 |
| --- | --- |
| `GET /api/pipeline/{project_id}/experience-suggestions` | 获取当前待处理差异可参考的历史建议。 |
| `POST /api/pipeline/{project_id}/experiences` | 保存一次人工确认/忽略决定，并把当前差异状态同步为该决定。 |
| `GET /api/pipeline/{project_id}/experiences?page=1&page_size=20` | 分页查看当前租户已沉淀的经验规则。 |

## 安全与边界

- 每个请求都先校验项目归属；经验库按租户隔离，不能跨租户查询。
- 规则仅在本地数据目录的 `experience_rules/` 中保存，不调用模型或外部服务。
- 规则是“建议”，不能触发自动处理。人工修改差异内容仍使用现有的差异审核接口。
- 当前版本采用精确字段匹配。收到更多真实案例后，可在人工确认下逐步扩充可复用的字段组合和规则类型。

## 验证

```powershell
python -m pytest tests/test_experience_store.py tests/test_review_experience_api.py -q
```
