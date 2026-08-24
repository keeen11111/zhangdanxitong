# AGENTS.md — 外服账单系统 2.0

> 本文件给 AI coding agent 提供项目规则。配合 CONTEXT.md（业务术语）和 tasks/（改造计划）使用。
> 改造性质：Brownfield 增量改造，依据 addyosmani/agent-skills adoption-guide Path B。

## 项目目标

企业级薪资数据处理 SaaS 系统（AutoPayroll-Pro）。当前为内部 Web 系统，架构须支持后续多租户对外开放。

**核心业务**：每月将多张来源不同的 Excel（上月模板、本月异动、社保表、个税表）合并、清洗、补齐，回填到带公式的 Master 模板导出。

**用户**：财务人员（内部）→ 多租户客户（对外）。

## 技术栈

| 层 | 技术 | 说明 |
|---|---|---|
| 后端 | FastAPI + Python 3.11+ | 异步框架，type hints |
| 核心引擎 | Pandas + DuckDB + openpyxl | 数据处理中台 |
| 数据库 | PostgreSQL（Docker） | 现状 SQLite，阶段 1 迁移 |
| ORM | SQLAlchemy + Alembic | 阶段 1 引入 Alembic |
| 认证 | JWT + refresh token | 阶段 2 加固 |
| 前端 | Next.js 14 + Ant Design + TypeScript | 阶段 3 重写 |
| 部署 | Docker + Nginx + PM2 | 阶段 4 |

## 命令

### 后端（当前）
```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --reload --port 8000    # 开发
uvicorn main:app --host 0.0.0.0 --port 8000  # 生产
```

### 后端（改造后追加）
```bash
pytest                                   # 跑测试
pytest --cov backend --cov core --cov-report=html  # 覆盖率
alembic upgrade head                     # 数据库迁移
alembic revision -m "description"        # 新建迁移
```

### 前端（阶段 3 重写后）
```bash
cd frontend
npm install
npm run dev                              # 开发
npm run build                            # 构建
npm run lint                             # 检查
```

### Docker（阶段 4）
```bash
docker-compose up -d                     # 启动全栈
docker-compose down                      # 停止
```

## 项目结构

```
外服账单系统2.0/
├── AGENTS.md                 # 本文件（项目规则）
├── CONTEXT.md                # 业务术语表
├── tasks/                    # 改造计划与任务
├── docs/adr/                 # 架构决策记录
│
├── app.py                    # Streamlit 单文件原型（仅参考，不发布）
│
├── core/                     # 核心业务引擎（项目核心资产，保留增强）
│   ├── normalizer.py         # 数据归一化（文本/金额/日期）
│   ├── schema_config.py      # ★ 实体定义 + EXPORT_MAPPING 列位映射（最关键）
│   ├── entity_engine.py      # 实体引擎：解析/合并/导出
│   ├── db_engine.py          # DuckDB 中台：异动计算/主表富化
│   ├── diff_engine.py        # 差异引擎：7类差异+审核状态
│   ├── validator.py          # 校验引擎：Blocker/Warning
│   └── excel_engine.py       # 无损Excel：公式平移/插删行
│
├── backend/
│   ├── main.py               # FastAPI 入口，注册路由
│   ├── database.py           # DB engine/session（阶段1改PostgreSQL）
│   ├── models.py             # SQLAlchemy 模型：Tenant/User/Project/UploadFile/EntitySnapshot
│   ├── schemas.py            # Pydantic 请求/响应模型
│   ├── auth.py               # JWT 认证（阶段2加固）
│   ├── config/              # 配置分层（阶段1新增）
│   ├── routers/
│   │   ├── auth.py           # /api/auth 注册/登录/me
│   │   ├── projects.py       # /api/projects 项目+文件+数据画像
│   │   ├── pipeline.py       # /api/pipeline ★ 业务主线（6模块+快照）
│   │   ├── modules.py        # 待删除（阶段1）
│   │   ├── payroll.py        # 待删除（阶段1）
│   │   ├── entities.py       # 待合并到 pipeline（阶段1）
│   │   └── health.py         # 待新增（阶段2）
│   ├── tests/                # 测试（阶段1起补充）
│   └── alembic/              # 数据库迁移（阶段1新增）
│
└── frontend/                 # 阶段3重写（当前只有 .next 构建产物）
    ├── app/                  # Next.js App Router
    ├── components/           # Ant Design 组件
    └── ...
```

## 代码风格

### Python（后端 + core）
- 使用 type hints，async 路由用 `async def`
- Pydantic 做输入/输出校验，不裸用 dict
- 数据库操作用 SQLAlchemy ORM，避免原生 SQL（DuckDB 例外，但需参数化）

```python
# 好：有类型、有校验、async
async def get_project(project_id: int, db: Session = Depends(get_db)) -> Project:
    project = db.query(Project).filter(Project.id == project_id).first()
    if not project:
        raise HTTPException(status_code=404, detail="Project not found")
    return project

# 不好：无类型、无校验
def get_project(id, db):
    return db.query(Project).filter(Project.id == id).first()
```

### 前端（阶段 3）
- TypeScript strict 模式
- Ant Design 组件为主，不混用其他 UI 库
- Server Components 优先，交互部分用 'use client'

## 测试策略

| 层 | 框架 | 覆盖率目标 | 策略 |
|---|---|---|---|
| core/ 引擎 | pytest | ≥80% | 阶段1先写特征化测试固化现有行为 |
| 后端路由 | pytest + httpx | ≥60% | API 集成测试 |
| 前端 | Jest + Testing Library | 关键流程 | E2E 用 Playwright |
| E2E | Playwright | 核心流程 | 月度核算全流程跑通 |

**测试位置**：`backend/tests/` 镜像源码结构。

**改老代码前先写特征化测试**（brownfield 铁律）。

## 边界（Boundaries）

### Always（始终遵守）
- 提交前跑测试
- 遵循命名约定（实体名、字段别名见 CONTEXT.md）
- 输入校验（系统边界：用户输入、外部 API）
- 数据库变更走 Alembic 迁移
- 操作日志写入审计表

### Ask first（先问再做）
- 数据库 schema 变更
- 新增依赖（requirements.txt / package.json）
- 改 CI 配置
- 改 core/ 的公开接口（schema_config 的实体定义、EXPORT_MAPPING）
- 删除任何文件（除了 todo.md 明确标记删除的）

### Never（绝不）
- 提交密钥、token、.env 文件
- 编辑 frontend/.next/ 构建产物
- 删除失败的测试（除非经批准）
- 在 core/ 用 f-string 拼 SQL（用参数化）
- 硬编码 JWT 密钥、DB 连接串
- 跳过多租户隔离校验

## 已知雷区

1. **frontend/ 源码缺失** — 只有 .next 构建产物，阶段 3 整体重写
2. **4 套并行路由** — modules/payroll/entities/pipeline，阶段 1 收敛为 pipeline 唯一主线
3. **session JSON 文件存储** — 不支持并发，阶段 1 迁移到数据库
4. **JWT 密钥硬编码**（auth.py:16）— 阶段 2 修环境变量
5. **DuckDB SQL f-string 拼接**（db_engine.py:116）— 阶段 2 参数化
6. **SQLite** — 生产不可用，阶段 1 迁 PostgreSQL
7. **app.py Streamlit 版** — 原型仅参考，不发布，改动优先级最低

## 多租户隔离规则

- 所有业务查询必须带 `tenant_id` 过滤
- `Project.tenant_id` 关联 `Tenant.id`
- 文件存储按 tenant 隔离目录
- JWT payload 含 tenant_id，路由层自动注入

## Skill 使用映射

本项目按 addyosmani/agent-skills 的 brownfield 路径使用 skill：

| 场景 | 使用的 skill | 位置 |
|---|---|---|
| 写规范/改架构前 | spec-driven-development | .skills/addyosmani-agent-skills/ |
| 改老代码前 | 先写特征化测试（test-driven-development） | 同上 |
| 安全加固 | security-and-hardening | 同上 |
| API 设计 | api-and-interface-design | 同上 |
| 代码审查 | code-review-and-quality | 同上 |
| 可观测性 | observability-and-instrumentation | 同上 |
| 需求对齐 | grill-with-docs | .skills/mattpocock-skills/ |
| 代码库设计 | codebase-design | 同上 |

## 参考文档
- [tasks/plan.md](tasks/plan.md) — 技术实现计划
- [tasks/todo.md](tasks/todo.md) — 任务清单
- [CONTEXT.md](CONTEXT.md) — 业务术语表
- [docs/adr/](docs/adr/) — 架构决策记录
