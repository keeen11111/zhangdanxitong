# 任务清单：通用语义 Sheet 归并

- [x] T1 工作簿级主题画像与唯一匹配判定
  - 验收：全部 Sheet 都参与匹配；同主题异名归并；非唯一不自动合并。
  - 验证：`python -m pytest -q tests/test_unified_integration.py`
  - 文件：`core/unified_integration.py`、`tests/test_unified_integration.py`

- [x] T2 同主题自动映射与覆盖
  - 验收：人员或其他业务主键唯一时直接覆盖且不新增 Sheet；公式与样式保留。
  - 验证：聚焦 pytest 与公式签名校验。
  - 文件：`core/unified_integration.py`、`tests/test_unified_integration.py`

- [x] T3 最小人工兜底与交付衔接
  - 验收：仅身份/目标/来源冲突进入独立人工队列，来源可追溯。
  - 验证：`python -m pytest -q tests/test_unified_integration.py`、`python -m pytest -q`
  - 文件：`core/unified_integration.py`、`tests/test_unified_integration.py`、必要时 `backend/routers/pipeline.py`
