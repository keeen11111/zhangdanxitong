# ADR-0003: PostgreSQL Docker 部署 + 本地文件存储预留接口

- 状态: Accepted
- 日期: 2026-08-06
- 决策者: 项目团队

## 背景

初版用 SQLite + 本地文件存储 + session JSON。生产环境需解决：
- SQLite 不支持并发（多用户同时核算）
- JSON 文件不支持并发和审计
- 文件存储需考虑多租户隔离和后续对象存储迁移

## 决策

### 数据库
采用 **PostgreSQL**，通过 **Docker** 部署（docker-compose）。

### 中间态
session JSON 迁移到 **PipelineSession 数据表**（PostgreSQL）。

### 文件存储
当前用**本地磁盘**，但代码抽象出**存储接口**（StorageInterface），后续可无缝切换对象存储。

## 理由

**PostgreSQL + Docker**：
- 当前为内部系统，Docker 部署成本最低，运维简单
- PostgreSQL 支持并发、行级安全（RLS）、JSON 字段
- 上线时若需云数据库，切换连接串即可
- Alembic 管理迁移，环境一致

**文件存储抽象接口**：
- 内部阶段本地磁盘够用
- 对外开放时需对象存储（OSS/S3）做多租户隔离
- 抽象接口避免届时大改：

```python
class StorageInterface(ABC):
    @abstractmethod
    async def save(self, tenant_id: int, project_id: int, file: UploadFile) -> str: ...
    @abstractmethod
    async def load(self, path: str) -> bytes: ...
    @abstractmethod
    async def delete(self, path: str) -> None: ...

class LocalStorage(StorageInterface): ...  # 当前实现
class OssStorage(StorageInterface): ...   # 未来实现
```

## 后果

**好处**：
- 开发环境一键启动（docker-compose up）
- 数据库并发问题解决
- 文件存储可扩展

**代价**：
- 团队需熟悉 Docker
- 本地开发需装 Docker Desktop
- PostgreSQL 与 SQLite 有 SQL 差异（用 ORM 规避）

**部署方案**：
```yaml
# docker-compose.yml（阶段4）
services:
  postgres:
    image: postgres:16
    environment:
      POSTGRES_DB: payroll
      POSTGRES_USER: payroll
      POSTGRES_PASSWORD: ${DB_PASSWORD}
    volumes:
      - pgdata:/var/lib/postgresql/data
  backend:
    build: ./backend
    depends_on: [postgres]
  frontend:
    build: ./frontend
```
