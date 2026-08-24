# CONTEXT.md — 外服账单系统 2.0 业务术语表

> 本文件建立项目的共享语言（ubiquitous language）。AI 和人在沟通时使用这些术语，避免歧义。
> 术语定义的权威来源是 [core/schema_config.py](core/schema_config.py)，本文件是其业务化解读。

## 核心概念

### 工资核算（Payroll Calculation）
每月对一批员工计算应发工资、扣除社保公积金和个税、得出实发金额的完整流程。系统服务的核心业务。

### Master 模板（Master Template）
带公式的 Excel 模板，A-CN 共 93 列。实体数据按列位填入，公式列（如应发合计 BA、实发 BT）不覆盖，由模板自动计算。

### 月度快照（Monthly Snapshot）
某月核算定稿后的数据存档。可被下月克隆为原始库基础，实现月间数据延续。

## 5 个业务实体

系统将源 Excel 数据归一为 5 个实体。每个实体有主键和外键关联。

| 实体 | 标识 | 主键 | 外键 | 说明 |
|---|---|---|---|---|
| 员工档案 | `employee_profile` | 工号 | 无（独立） | 人员基本信息，所有关联的根 |
| 薪资明细 | `salary_detail` | 工号 | 工号 | 薪资结构与各项标准 |
| 考勤记录 | `attendance_record` | 工号 | 工号 | 出勤与扣款 |
| 社保公积金 | `social_security` | 身份证号 | 身份证号 | 社保/公积金基数与明细 |
| 个税记录 | `tax_record` | 工号 | 工号 | 个税扣缴与累计数据 |

**关联方式**：员工档案以工号为根，薪资/考勤/个税通过工号关联，社保通过身份证号关联（因社保表无工号列，需 LEFT JOIN 用身份证号补工号）。

### 实体字段示例（员工档案）

| 字段 | 类型 | 必填 | 别名（源列名可能的形式） |
|---|---|---|---|
| 工号 | text | 是 | 新工号、员工工号、员工编号、工牌号 |
| 姓名 | text | 是 | 人员姓名、员工姓名 |
| 身份证号 | text | 是 | 身份证号码、证件号码 |
| 入职日期 | date | 否 | 入职时间、入职日 |
| 转正日期 | date | 否 | 转正时间、试用期结束时间 |

> 完整字段定义见 [core/schema_config.py](core/schema_config.py) 的 `ALL_ENTITIES`。改实体定义需先问（Ask first 边界）。

## 7 类异动（Personnel Changes）

每月核算时，通过 DuckDB SQL 对比原始库与源头表，算出 7 类差异：

| 异动类型 | 标识 | 含义 |
|---|---|---|
| 入职 | `hire` | 源头表有、原始库无（新员工） |
| 离职 | `depart` | 原始库有、源头表无 |
| 调动 | `transfer` | 部门/岗位/工作地点变化 |
| 薪资变动 | `salary_change` | 基本工资/绩效标准变化 |
| 考勤异动 | `attendance` | 加班/请假等考勤数据变化 |
| 社保异动 | `social` | 社保基数或明细变化 |
| 个税异动 | `tax` | 个税扣缴数据变化 |

每条异动带**审核状态**：
- `pending`（待审核）— 默认状态
- `confirmed`（已确认）— 用户确认后纳入最终库
- `ignored`（已忽略）— 用户判定不处理

## 校验规则

数据补齐阶段，校验引擎对每行做两项检查：

### Blocker（阻断项，红色 🔴）
必须补齐才能导出，否则阻断：
- 工号
- 姓名
- 身份证号
- 基本工资

### Warning（警告项，黄色 🟡）
建议补齐但不阻断导出：
- 转正日期
- 试用期标准
- 车贴（交通通讯补贴）

## Pipeline 6 模块（业务主线）

后端业务以 pipeline 为唯一主线，6 个模块对应月度核算的 6 个阶段：

```
A 原始库 → B 源头表 → C 差异计算 → D 审核 → E 最终库 → F 导出
```

| 模块 | 名称 | 输入 | 输出 | 对应引擎 |
|---|---|---|---|---|
| A | 原始库 | 上月快照或空 | 原始库数据 | entity_engine |
| B | 源头表 | 本月源 Excel | 解析后的实体 | entity_engine + normalizer |
| C | 差异计算 | A + B | 7 类异动清单 | db_engine + diff_engine |
| D | 审核 | C 的异动清单 | 确认后的异动 | （用户交互） |
| E | 最终库 | A + B + 确认异动 | 合并定稿数据 | entity_engine |
| F | 导出 | E + Master 模板 | 填好的 Excel | entity_engine + excel_engine |

## Master 模板列位结构

EXPORT_MAPPING 定义实体字段到模板列的映射。公式列标记为 True，导出时不覆盖。

| 列位 | 内容 | 来源实体 | 公式列 |
|---|---|---|---|
| A-O | 员工基本信息 | employee_profile | 否 |
| P-W | 薪资标准 | salary_detail | 否 |
| X-Y | 月度合计/固浮比 | — | **是** |
| Z-AB | 基本薪资/绩效标准 | salary_detail | 否 |
| AC-AH | 奖金与调差 | salary_detail | 否 |
| AI-AT | 考勤数据 | attendance_record | 否 |
| AU-AZ | 本月各类补贴 | salary_detail | 否 |
| BA | 应发工资合计 | — | **是** |
| BB-BO | 社保公积金 | social_security | 否 |
| BP-BR | 个税扣缴 | tax_record | 否 |
| BS | 子女配偶补充医疗 | tax_record | 否 |
| BT | 工资实发金额 | — | **是** |
| BU | 应收应付工资 | — | **是** |
| BV | 补偿金 | salary_detail | 否 |
| BW | 本次实发金额 | — | **是** |
| BX-BZ | 身份信息 | employee_profile | 否 |
| CA-CN | 累计税务数据 | tax_record | 部分 |

## 源文件类型

系统通过文件名嗅探文件类型，决定如何解析：

| 类型标识 | 文件名特征 | 映射实体 | 解析规则 |
|---|---|---|---|
| `salary` | 薪资/工资/salary | 员工档案+薪资明细+考勤 | 按 Sheet 名分区：入离职/考勤/奖金 |
| `social_hean` | 鹤安/hean | 社保公积金 | Sheet1，第2行起为表头 |
| `social_yiyao` | 益药/yiyao | 社保公积金 | Sheet1，第2行起为表头 |
| `tax` | 个税/tax | 个税记录 | — |

> 嗅探逻辑见 [core/schema_config.py](core/schema_config.py) 的 `detect_file_type()`。

## 数据归一化规则

源 Excel 数据需归一化后才能进实体（[core/normalizer.py](core/normalizer.py)）：

- **文本**：去首尾空格，工号保留前导零
- **金额**：去 ¥/逗号，统一为 float
- **日期**：统一 YYYY-MM-DD，兼容 Excel 序列号
- **整数**：转 int

## 无损 Excel 操作

导出时需对模板做行级操作（[core/excel_engine.py](core/excel_engine.py)）：

- **公式平移**：插行时，正则改 A1 引用的行号，绝对引用 `$N` 不动
- **插行**：复制上一行样式 + 公式平移 1 行
- **删行**：从下往上删，避免行号偏移
- **公式列保留**：标记为公式的列位不覆盖原值

## 多租户概念

| 术语 | 含义 |
|---|---|
| 租户（Tenant） | 一个客户组织，数据隔离的边界 |
| 项目（Project） | 租户下的一个月度核算单元，关联文件和快照 |
| 上传文件（UploadFile） | 项目下的源 Excel 文件 |
| 实体快照（EntitySnapshot） | 某月定稿的实体数据存档 |

## 缩写对照

| 缩写 | 全称 | 含义 |
|---|---|---|
| DTP | Desktop Publishing | 排版计件绩效（特定岗位） |
| RLS | Row Level Security | 行级安全（数据库多租户隔离） |
| ADR | Architecture Decision Record | 架构决策记录 |
| LCP | Largest Contentful Paint | 前端性能指标 |
