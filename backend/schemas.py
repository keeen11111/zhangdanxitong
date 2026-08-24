"""Pydantic 请求/响应模型。"""
from datetime import datetime
from typing import Optional

from pydantic import BaseModel, ConfigDict, EmailStr, Field


# ---------- Auth ----------
class RegisterIn(BaseModel):
    email: EmailStr
    password: str = Field(min_length=6, max_length=100)
    name: str = Field(min_length=1, max_length=100)
    tenant_name: Optional[str] = Field(None, max_length=100,
                                       description="租户名，留空则用 email 域名")


class LoginIn(BaseModel):
    email: EmailStr
    password: str


class UserOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: str
    email: EmailStr
    name: str
    tenant_id: str
    tenant_name: str
    is_active: bool
    created_at: datetime

class TokenOut(BaseModel):
    access_token: str
    token_type: str = "bearer"
    user: UserOut


# ---------- Project ----------
class ProjectIn(BaseModel):
    name: str = Field(min_length=1, max_length=200)
    salary_month: str = Field(default="2026.06", max_length=20)


class ProjectOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: str
    name: str
    salary_month: str
    status: str
    created_at: datetime
    updated_at: datetime
    file_count: int = 0

# ---------- File ----------
class FileOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: str
    file_type: str
    original_name: str
    sheet_name: Optional[str]
    row_count: int
    col_count: int
    created_at: datetime
    processing_seconds: Optional[float] = None
    normalization_note: Optional[str] = None

class FileTypeUpdateIn(BaseModel):
    file_type: str


class FilePreview(BaseModel):
    """文件预览：前 N 行 + 列名。"""
    file_id: str
    columns: list[str]
    rows: list[dict]
    row_count: int
    col_count: int


# ---------- 业务：异动配置 ----------
class DiffConfigIn(BaseModel):
    last_month_file_id: str
    current_file_id: str
    id_col_lm: str
    id_col_cs: str
    compare_cols: list[str] = []
    social_file_ids: list[str] = []


class DiffResult(BaseModel):
    new_hires: list[dict]
    departed: list[dict]
    transfers: list[dict]


# ---------- 业务：诊断 ----------
class AuditResult(BaseModel):
    columns: list[str]
    rows: list[dict]
    status_col: list[str]            # 每行诊断状态
    missing_col: list[list[str]]     # 每行缺失字段
    metrics: dict


class AuditSaveIn(BaseModel):
    """用户在线补齐后回传的整表。"""
    rows: list[dict]


# ---------- 业务：导出 ----------
class ExportResult(BaseModel):
    file_id: str
    filename: str
    log: list[str]
