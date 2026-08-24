# 任务清单：外服账单系统 2.0 企业级改造

> 配合 plan.md 使用。按依赖顺序排列，每任务可单次专注完成。

## 阶段 0：守护与上下文（进行中）

- [x] T0.1 确认改造方案与关键决策
  - Acceptance: 方案文档 SPEC.md 完成，6 个假设与 4 个 Open Questions 已确认
  - Verify: 用户确认"继续"
  - Files: tasks/plan.md

- [ ] T0.2 创建项目规则文件 AGENTS.md
  - Acceptance: 记录真实代码约定、构建命令、目录含义、已知雷区、不变量
  - Verify: 内容覆盖 addyosmani context-engineering 要求的 6 项
  - Files: AGENTS.md

- [ ] T0.3 创建业务术语表 CONTEXT.md
  - Acceptance: 记录 5 个实体、异动类型、Master 模板列位等业务语言
  - Verify: 团队新成员能据此理解业务
  - Files: CONTEXT.md

- [ ] T0.4 抽取关键规范到 .cursor/rules/
  - Acceptance: 安全、API 设计、TDD、Git 工作流 4 条规则就位
  - Verify: .cursor/rules/ 下有 4 个 .mdc 文件
  - Files: .cursor/rules/*.mdc

- [ ] T0.5 配置 pre-commit 钩子
  - Acceptance: 提交前自动检查（格式、密钥、大文件）
  - Verify: git commit 触发钩子
  - Files: .pre-commit-config.yaml

- [ ] T0.6 记录阶段 0 架构决策（ADR）
  - Acceptance: 记录"为何 brownfield 增量改造""为何选 Ant Design"等决策
  - Verify: docs/adr/ 下有 ADR 文件
  - Files: docs/adr/0001-brownfield-incremental.md, docs/adr/0002-antd-frontend.md

## 阶段 1：架构收敛与数据层重构

- [ ] T1.1 给 core/ 写特征化测试（固化现有行为）
  - Acceptance: 7 个引擎模块各有测试，覆盖关键路径
  - Verify: pytest --cov core/ ≥ 60%
  - Files: backend/tests/core/

- [ ] T1.2 给 routers/pipeline.py 写特征化测试
  - Acceptance: 6 模块流程有测试覆盖
  - Verify: pytest backend/tests/routers/test_pipeline.py 通过
  - Files: backend/tests/routers/test_pipeline.py

- [ ] T1.3 删除 modules.py 和 payroll.py 路由
  - Acceptance: main.py 不再注册这两套路由
  - Verify: uvicorn 启动无报错，/docs 不含这两套
  - Files: backend/routers/modules.py(删除), backend/routers/payroll.py(删除), backend/main.py

- [ ] T1.4 合并 entities 快照能力到 pipeline
  - Acceptance: pipeline 提供快照 CRUD，entities 路由删除
  - Verify: 快照保存/加载/克隆接口在 /api/pipeline 下可用
  - Files: backend/routers/pipeline.py, backend/routers/entities.py(删除)

- [ ] T1.5 SQLite → PostgreSQL
  - Acceptance: database.py 连 PostgreSQL，连接池配置
  - Verify: 连接 PG 成功，建表正常
  - Files: backend/database.py, backend/models.py

- [ ] T1.6 引入 Alembic 数据库迁移
  - Acceptance: alembic init + 首个 migration 可从零建库
  - Verify: alembic upgrade head 成功
  - Files: backend/alembic/

- [ ] T1.7 session JSON → PipelineSession 数据表
  - Acceptance: 中间态存数据库，不再写 JSON 文件
  - Verify: 运行流程后无 .json 文件产生
  - Files: backend/models.py, backend/routers/pipeline.py

- [ ] T1.8 配置分层 + 环境变量
  - Acceptance: config/dev.py, config/prod.py, .env.example，JWT/DB 走环境变量
  - Verify: 无硬编码密钥，grep 查不到
  - Files: backend/config/, .env.example, backend/auth.py

## 阶段 2：后端规范化与安全加固

- [ ] T2.1 安全加固（security-and-hardening skill）
  - Acceptance: CORS、文件上传校验、路径穿越防护、密码策略
  - Verify: security-auditor 审查无 Critical
  - Files: backend/main.py, backend/routers/projects.py, backend/auth.py

- [ ] T2.2 API 规范化（api-and-interface-design skill）
  - Acceptance: 统一响应格式、统一异常处理、输入校验补全
  - Verify: 所有 API 返回 {code, data, message} 结构
  - Files: backend/schemas.py, backend/main.py, backend/routers/

- [ ] T2.3 可观测性（observability-and-instrumentation skill）
  - Acceptance: structlog 结构化日志、审计日志表、健康检查
  - Verify: /api/health 返回 200，操作有审计记录
  - Files: backend/main.py, backend/models.py, backend/routers/health.py(新增)

- [ ] T2.4 DuckDB SQL 参数化
  - Acceptance: db_engine.py 不再用 f-string 拼 SQL
  - Verify: 代码审查无注入风险
  - Files: core/db_engine.py

- [ ] T2.5 核心引擎测试补全
  - Acceptance: pytest --cov core/ ≥ 80%
  - Verify: 覆盖率报告
  - Files: backend/tests/core/

- [ ] T2.6 代码审查质量门（code-review-and-quality skill）
  - Acceptance: 五轴审查通过
  - Verify: code-reviewer agent 报告无 Required 级别问题
  - Files: 全后端

## 阶段 3：前端重写

- [ ] T3.1 前端项目初始化
  - Acceptance: Next.js + Ant Design + TypeScript 脚手架，清理旧 .next
  - Verify: npm run dev 启动
  - Files: frontend/

- [ ] T3.2 登录页
  - Acceptance: JWT 登录，token 存储
  - Verify: 登录成功跳转
  - Files: frontend/app/login/

- [ ] T3.3 项目列表页
  - Acceptance: 项目 CRUD，文件上传
  - Verify: 能创建项目、上传 Excel
  - Files: frontend/app/projects/

- [ ] T3.4 工作台-模块A原始库
  - Acceptance: 显示原始库数据，可编辑
  - Verify: 数据网格可编辑保存
  - Files: frontend/app/projects/[id]/module-a/

- [ ] T3.5 工作台-模块B源头表
  - Acceptance: 显示源头表数据，可编辑
  - Verify: 数据网格可编辑保存
  - Files: frontend/app/projects/[id]/module-b/

- [ ] T3.6 工作台-模块C差异计算
  - Acceptance: 显示新增/离职/异动差异
  - Verify: 差异列表正确
  - Files: frontend/app/projects/[id]/module-c/

- [ ] T3.7 工作台-模块D审核
  - Acceptance: 待审核/确认/忽略状态切换
  - Verify: 状态切换保存
  - Files: frontend/app/projects/[id]/module-d/

- [ ] T3.8 工作台-模块E最终库
  - Acceptance: 显示最终合并数据
  - Verify: 数据正确合并
  - Files: frontend/app/projects/[id]/module-e/

- [ ] T3.9 工作台-模块F导出
  - Acceptance: 按模板导出 Excel，公式保留
  - Verify: 导出文件可打开，公式正确
  - Files: frontend/app/projects/[id]/module-f/

- [ ] T3.10 E2E 测试（browser-testing-with-devtools skill）
  - Acceptance: 完整流程跑通
  - Verify: E2E 测试通过
  - Files: frontend/e2e/

## 阶段 4：部署与运维

- [ ] T4.1 后端 Dockerfile
  - Acceptance: 可构建镜像
  - Verify: docker build 成功
  - Files: backend/Dockerfile

- [ ] T4.2 前端 Dockerfile
  - Acceptance: 可构建镜像
  - Verify: docker build 成功
  - Files: frontend/Dockerfile

- [ ] T4.3 docker-compose 编排
  - Acceptance: 前后端 + PostgreSQL 一键启动
  - Verify: docker-compose up 成功
  - Files: docker-compose.yml

- [ ] T4.4 Nginx 反向代理
  - Acceptance: 前后端统一入口
  - Verify: 访问 nginx 能到前端和 API
  - Files: nginx/

- [ ] T4.5 CI/CD pipeline
  - Acceptance: 自动测试、构建、部署
  - Verify: CI 通过
  - Files: .github/workflows/

- [ ] T4.6 健康检查与监控
  - Acceptance: 健康检查、错误告警
  - Verify: 模拟故障能告警
  - Files: backend/routers/health.py

- [ ] T4.7 上线检查（shipping-and-launch skill）
  - Acceptance: 通过上线质量门
  - Verify: go/no-go 决策通过
  - Files: docs/release-checklist.md
