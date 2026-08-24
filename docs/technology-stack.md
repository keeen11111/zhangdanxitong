# 技术栈与架构

## 技术栈（当前实现）

| 层级 | 技术 | 职责 |
| --- | --- | --- |
| 前端 | Next.js 14、React 18、TypeScript 5 | 登录、项目管理、上传、工作台与成果下载页面。 |
| 前端 UI | Tailwind CSS、Radix UI、Lucide | 当前的界面样式、基础交互组件和图标。 |
| 后端 | Python 3.11+、FastAPI、Uvicorn | REST API、身份校验、项目与流水线编排。 |
| 后端数据模型 | SQLAlchemy、Pydantic | ORM 数据模型、请求和响应校验。 |
| 核心数据处理 | Pandas、DuckDB | Excel 表格解析、归一化、差异计算与临时数据查询。 |
| Excel 引擎 | openpyxl | 公式/格式保留、插删行、输出工作簿和人工处理表。 |
| 身份认证 | JWT、Passlib/bcrypt、python-jose | 密码校验、令牌签发和当前用户识别。 |
| 当前存储 | SQLite、本地文件、session JSON | 开发期项目数据、上传文件和流水线中间态。 |
| 桌面封装 | Electron、electron-builder | 可选的 Windows 桌面打包能力。 |
| 验证 | pytest、ESLint、Next.js build | Python 回归测试、前端静态检查和生产构建。 |

## 当前架构

```text
Next.js 前端
    │ HTTP / JWT
    ▼
FastAPI 路由层
    ├── auth / projects：身份、租户、项目、文件
    └── pipeline：月度整合主线、人工处理和导出
             │
             ▼
        core/ 业务引擎
        ├── entity_engine：解析与实体合并
        ├── diff_engine + db_engine：差异计算
        ├── validator：Blocker / Warning 校验
        ├── unified_integration：总表、来源 Sheet、人工事项
        └── excel_engine：公式和格式安全写入
             │
             ▼
SQLite + 本地上传/导出文件 + session JSON
```

## 目标架构与现状差异

| 主题 | 当前 | 已记录目标 | 状态 |
| --- | --- | --- | --- |
| 数据库 | SQLite | PostgreSQL + 连接池 | 待实施 |
| 迁移 | `create_all` / 现有库 | Alembic 版本化迁移 | 待实施 |
| 会话与审计 | 本地 JSON 为主 | PipelineSession、审核和审计表 | 待实施 |
| 路由 | 旧路由与 pipeline 并存 | `pipeline` 为唯一业务主线 | 待实施 |
| UI 标准 | Tailwind + Radix | 历史计划指定 Ant Design | 待产品决策 |
| 部署 | 本地脚本 / 可选桌面包 | Docker、Nginx、内网 PostgreSQL | 待实施 |
| 自动化交付 | 手工验证 | CI/CD 和发布检查 | 待实施 |

## 关键工程原则

- 工资计算、字段匹配、差异识别和写入使用确定性 Python 规则；不使用大模型作账务判断。
- 工资主模板是公式、格式和 Sheet 结构的权威来源；导出不得破坏模板公式。
- 自动更新必须可唯一确认主题、人员、字段与来源优先级；否则生成可追溯人工事项。
- 所有项目和文件访问必须受租户边界约束。
- 真实薪资数据不能进入 Git、测试断言、公开日志或外部服务。

## 常用命令

```powershell
# 全量 Python 测试
python -m pytest -q

# 前端检查与生产构建
cd frontend
npm run lint
npm run build

# 本地开发（根目录）
uvicorn backend.main:app --reload --port 8000
```

进一步的业务术语和字段定义见 [CONTEXT.md](../CONTEXT.md)，架构决策见 [ADR 目录](adr/)。
