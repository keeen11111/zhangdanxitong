# AutoPayroll-Pro

外服账单系统 2.0 是一个面向财务人员的薪资数据处理 Web 系统。当前基础框架由 FastAPI 后端、Next.js 前端和 `core/` Excel 处理引擎组成。

## 当前可用框架

- 登录与租户身份
- 项目创建和月份管理
- Excel 上传与文件列表
- 六阶段薪资处理工作台
- 后端服务连接状态
- 模板导出入口

## 本地启动

推荐 Python 3.11 和 Node.js 20。

### 一键启动（推荐）

双击项目根目录的 `start-project.cmd` 即可启动前后端。脚本会复用已运行的服务，不会重复安装依赖；服务日志保存在 `logs/`。

前端地址固定为：`http://127.0.0.1:3002`

### 1. 后端

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r backend\requirements.txt -r requirements.txt
uvicorn backend.main:app --reload --port 8000
```

后端健康检查地址：`http://127.0.0.1:8000/api/health`

### 2. 前端

复制 `frontend/.env.example` 为 `frontend/.env.local`，然后执行：

```powershell
cd frontend
npm install
npm run dev
```

前端地址：`http://127.0.0.1:3000`

## 验证命令

```powershell
python -m pytest -q
cd frontend
npm run lint
npm run build
```

## 数据安全

- `.env`、SQLite 数据库、上传文件、session JSON、导出文件和真实 Excel 均不得提交。
- 真实薪资数据仅在本地或批准的内网环境中处理。
- PostgreSQL、Alembic 和历史数据迁移按 `tasks/week-mvp-plan.md` 后续实施。
