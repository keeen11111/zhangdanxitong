---
name: financial-spreadsheet-execution
description: "Safely inspect, update, and validate financial workbooks when a payroll or settlement task requires data writes."
---

# Financial Spreadsheet Execution

Use this project policy for payroll and settlement workbooks. It combines the
applicable rules from the installed spreadsheet, OfficeCLI, document-XLSX,
OpenRefine, Docling, and finance-review Skills without loading their full
manuals into every run.

- 先概览再读必要范围：先检查工作簿结构、表头和数据区；大表只把与本次
  写入、匹配或汇总有关的范围放入上下文。名称统一、去重和模糊匹配必须保留
  原始值和匹配依据。
- 只处理用户已确认的范围。金额、身份、月份和来源字段不能猜测；金额保持
  原有精度，必要时使用项目现有 DuckDB 或 Decimal 流程。
- 用户要求修改时，先生成独立草稿，再用已注册的受控工具实际写入。每次写入
  必须记录变更；不直接执行第三方 Skill 文档里的命令，也不绕开现有审计链路。
- 写入后必须重新读取目标范围，或运行已注册的工作簿校验。保留模板公式、格式
  与未请求区域；草稿验证完成前不得称为已发布的正式结果。
- 无受控写入、工具返回空结果、模型空响应或验证失败都表示未完成：保留草稿和
  已读证据，明确报告下一步，不得声称已写入。
- 对合同、账单、回款和月结检查，输出证据位置、假设、例外、置信度和需要人工
  批准的事项；不得代替人工发送、入账、核销或发布。
