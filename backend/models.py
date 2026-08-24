"""ORM 模型：多租户 + 项目空间 + 上传文件元数据 + 实体快照。

关系：
    Tenant 1 ── n User
    User  1 ── n Project
    Project 1 ── n UploadFile
    Project 1 ── n EntitySnapshot
"""
import uuid
from datetime import datetime

from sqlalchemy import Column, String, Integer, DateTime, ForeignKey, Text, Boolean
from sqlalchemy.orm import relationship

from backend.database import Base


def _uuid() -> str:
    return uuid.uuid4().hex


class Tenant(Base):
    """租户（企业/组织），数据隔离的最高层级。"""
    __tablename__ = "tenants"
    id = Column(String, primary_key=True, default=_uuid)
    name = Column(String(100), nullable=False, comment="租户名称")
    created_at = Column(DateTime, default=datetime.utcnow)
    users = relationship("User", back_populates="tenant")


class User(Base):
    """用户。属于某个租户。"""
    __tablename__ = "users"
    id = Column(String, primary_key=True, default=_uuid)
    tenant_id = Column(String, ForeignKey("tenants.id"), nullable=False, index=True)
    email = Column(String(200), unique=True, nullable=False, index=True)
    name = Column(String(100), nullable=False)
    hashed_password = Column(String(255), nullable=False)
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    tenant = relationship("Tenant", back_populates="users")
    projects = relationship("Project", back_populates="owner", cascade="all, delete-orphan")


class Project(Base):
    """薪资项目空间。每个用户可有多个项目（对应不同月份/批次）。"""
    __tablename__ = "projects"
    id = Column(String, primary_key=True, default=_uuid)
    owner_id = Column(String, ForeignKey("users.id"), nullable=False, index=True)
    name = Column(String(200), nullable=False, comment="项目名称，如 2026.06 薪资批次")
    salary_month = Column(String(20), nullable=False, default="2026.06")
    status = Column(String(50), default="import", comment="import/config/work/export")
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    owner = relationship("User", back_populates="projects")
    files = relationship("UploadFile", back_populates="project", cascade="all, delete-orphan")
    snapshots = relationship("EntitySnapshot", back_populates="project", cascade="all, delete-orphan")


class UploadFile(Base):
    """上传文件元数据。物理文件落在 UPLOAD_DIR。"""
    __tablename__ = "upload_files"
    id = Column(String, primary_key=True, default=_uuid)
    project_id = Column(String, ForeignKey("projects.id"), nullable=False, index=True)
    file_type = Column(String(50), nullable=False,
                       comment="last_month / current / social / template")
    original_name = Column(String(500), nullable=False)
    stored_path = Column(String(500), nullable=False, comment="相对 UPLOAD_DIR 的路径")
    sheet_name = Column(String(200), nullable=True,
                        comment="社保表场景下用户指定的关联表名")
    row_count = Column(Integer, default=0)
    col_count = Column(Integer, default=0)
    created_at = Column(DateTime, default=datetime.utcnow)

    project = relationship("Project", back_populates="files")


class EntitySnapshot(Base):
    """实体数据快照：按项目+月份存储完整的实体数据JSON。

    每月一个快照，支持历史查阅和跨月克隆。
    """
    __tablename__ = "entity_snapshots"
    id = Column(String, primary_key=True, default=_uuid)
    project_id = Column(String, ForeignKey("projects.id"), nullable=False, index=True)
    salary_month = Column(String(20), nullable=False, index=True,
                          comment="薪资月份，如 2026.06")
    label = Column(String(200), nullable=True, comment="快照标签，如'首次导入'、'6月定稿'")
    data_json = Column(Text, nullable=False, comment="完整实体数据JSON")
    row_count = Column(Integer, default=0, comment="总记录数")
    created_at = Column(DateTime, default=datetime.utcnow)

    project = relationship("Project", back_populates="snapshots")
