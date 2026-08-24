# 技术实现计划：外服账单系统 2.0 企业级改造

> 本文件是改造的总体技术计划，配合 SPEC.md 使用。任务清单见 todo.md。

## 一、改造性质

**Brownfield 增量改造**（非重写）。核心引擎 `core/` 保留增强，外层规范化。

依据：`addyosmani/agent-skills` 的 adoption-guide Path B（已装在 `.skills/`）。

## 二、关键决策（已确认）

| 决策项 | 选择 | 理由 |
|---|---|---|
| 目标形态 | 企业级内部 Web，架构支持后续对外开放 | 多租户保留 |
| 业务主线 | pipeline 唯一主线，删 modules/payroll/entities | 收敛 4 套并行路由 |
| 数据库 | SQLite → PostgreSQL（Docker 部署） | 生产可用 |
| 文件存储 | 本地磁盘 + 抽象存储接口 | 内部够用，预留对象存储扩展 |
| 认证 | 保留 JWT + 修复硬编码 + refresh token | 适合多端 |
| 前端 UI | Ant Design（替换 Radix UI） | 企业级中后台标准 |
| 中间态存储 | session JSON → 数据库表 | 支持并发与审计 |

## 三、组件与依赖关系

```
┌─────────────────────────────────────────────────────┐
│ 前端 (Next.js + Ant Design)                          │
│   登录 / 项目列表 / 工作台(pipeline 6模块)            │
└────────────────────┬────────────────────────────────┘
                     │ REST API
┌────────────────────▼────────────────────────────────┐
│ 后端 (FastAPI)                                       │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌────────┐ │
│  │ auth.py  │ │projects  │ │pipeline  │ │ health │ │
│  │(JWT+刷新)│ │(文件管理)│ │(6模块主线)│ │(监控)  │ │
│  └────┬─────┘ └────┬─────┘ └────┬─────┘ └────────┘ │
│       │            │            │                    │
│  ┌────▼────────────▼────────────▼────┐              │
│  │  数据层 (SQLAlchemy + Alembic)      │              │
│  │  PostgreSQL (Docker)                │              │
│  └─────────────────────────────────────┘              │
└────────────────────┬────────────────────────────────┘
                     │
┌────────────────────▼────────────────────────────────┐
│ 核心引擎 core/ (保留增强)                            │
│  normalizer / schema_config / entity_engine /        │
│  db_engine / diff_engine / validator / excel_engine  │
└─────────────────────────────────────────────────────┘
```

**依赖方向**：前端 → 后端路由 → 数据层 + 核心引擎。核心引擎不依赖后端。

## 四、实现顺序与验证检查点

### 阶段 0：守护与上下文（不改业务代码）
- **产出**：AGENTS.md、CONTEXT.md、.cursor/rules/、pre-commit
- **验证检查点**：AI 能正确引用项目约定，pre-commit 能拦截
- **下一阶段前置**：必须完成，否则后续无安全网

### 阶段 1：架构收敛与数据层重构
- **依赖**：阶段 0 完成
- **产出**：单套路由、PostgreSQL、Alembic、配置分层
- **验证检查点**：`alembic upgrade head` 建库成功，无硬编码密钥，无 JSON 文件存储
- **风险**：路由合并可能丢失 entities 的快照逻辑 → 缓解：先写特征化测试固化行为

### 阶段 2：后端规范化与安全加固
- **依赖**：阶段 1 完成
- **产出**：统一响应/异常、审计日志、安全加固、核心测试
- **验证检查点**：security-auditor 审查无 Critical，pytest 覆盖率 ≥80%
- **并行可能**：安全加固与测试可部分并行

### 阶段 3：前端重写
- **依赖**：阶段 2 的 API 规范化完成（需要统一响应格式）
- **产出**：完整 Next.js + Ant Design 前端
- **验证检查点**：E2E 跑通月度核算全流程
- **风险**：前端工作量大 → 缓解：按 pipeline 6 模块分片交付，每模块可独立验证

### 阶段 4：部署与运维
- **依赖**：阶段 3 完成
- **产出**：Dockerfile、docker-compose、CI/CD、监控
- **验证检查点**：`docker-compose up` 一键启动，CI 通过所有质量门

## 五、风险登记册

| 风险 | 概率 | 影响 | 缓解措施 |
|---|---|---|---|
| 路由合并丢失 entities 快照逻辑 | 中 | 高 | 先写特征化测试固化现有行为 |
| DuckDB SQL 拼接注入风险 | 中 | 高 | 阶段 2 参数化 |
| 前端重写工期长 | 高 | 中 | 按 pipeline 模块分片，每片可验证 |
| PostgreSQL 与 SQLite SQL 差异 | 低 | 中 | 用 SQLAlchemy ORM，避免原生 SQL |
| 多租户隔离不彻底 | 中 | 高 | 阶段 2 安全审查重点查租户隔离 |

## 六、质量门（每阶段必须通过）

每个阶段交付前，必须通过对应 skill 的验证：

- **阶段 0**：`documentation-and-adrs` 验证清单
- **阶段 1**：`incremental-implementation` + `code-review-and-quality`
- **阶段 2**：`security-and-hardening` + `test-driven-development`
- **阶段 3**：`frontend-ui-engineering` + `browser-testing-with-devtools`
- **阶段 4**：`ci-cd-and-automation` + `shipping-and-launch`

## 七、不变量（全程遵守）

1. `core/` 引擎的公开接口不破坏性变更（内部可重构）
2. 每次提交原子化（~100 行），可 bisect
3. 不提交密钥、不编辑 .next 构建产物
4. 改老代码前先写特征化测试
5. 数据库变更走 Alembic 迁移，不手动改库
