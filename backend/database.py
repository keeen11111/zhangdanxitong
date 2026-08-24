"""数据库连接与会话管理。

开发环境默认使用 SQLite；部署环境通过 PAYROLL_DATABASE_URL 切换数据库。
"""
import os
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base

from backend.config import data_dir_from_mapping, settings

# 数据目录
DATA_DIR = str(data_dir_from_mapping(os.environ))
os.makedirs(DATA_DIR, exist_ok=True)

# 上传文件存储目录
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# 导出文件存储目录
EXPORT_DIR = os.path.join(DATA_DIR, "exports")
os.makedirs(EXPORT_DIR, exist_ok=True)

DATABASE_URL = settings.database_url


def engine_options(database_url: str) -> dict:
    """返回与数据库驱动匹配的 SQLAlchemy engine 参数。"""
    options: dict = {"echo": False}
    if database_url.startswith("sqlite"):
        # SQLite 需要关闭线程检查（FastAPI 多线程访问）
        options["connect_args"] = {"check_same_thread": False}
    return options


engine = create_engine(DATABASE_URL, **engine_options(DATABASE_URL))

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


def get_db():
    """FastAPI 依赖注入：每请求一个 session。"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db():
    """建表（幂等）。"""
    # 必须先 import models 让 ORM 类注册到 Base.metadata
    from backend import models  # noqa: F401
    Base.metadata.create_all(bind=engine)
