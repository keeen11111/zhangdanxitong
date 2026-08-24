"""板块化数据工作台路由。

核心理念：上传 Excel 后，按 Sheet 拆分成多个"板块"，每个板块是独立的可编辑数据网格。
用户在各板块里改单元格 / 增删行，全部改完后一键导出合并后的 Excel。

板块数据全量存在 session JSON（backend/data/sessions/），编辑时整表回传，简单可靠。
"""
import io
import json
import os
import uuid
from datetime import datetime

import openpyxl
import pandas as pd
from fastapi import APIRouter, Depends, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from sqlalchemy.orm import Session

from backend.auth import get_current_user
from backend.database import DATA_DIR, EXPORT_DIR, UPLOAD_DIR, get_db
from backend.models import User, Project, UploadFile

router = APIRouter(prefix="/api/modules", tags=["modules"])

SESSION_DIR = os.path.join(DATA_DIR, "sessions")
os.makedirs(SESSION_DIR, exist_ok=True)


# ============================================================
# 数据模型
# ============================================================
class ModuleInfo(BaseModel):
    id: str
    name: str
    source_file: str
    source_sheet: str
    row_count: int
    col_count: int


class ModuleData(BaseModel):
    id: str
    name: str
    source_file: str
    source_sheet: str
    columns: list[str]
    rows: list[dict]


class ModuleSaveIn(BaseModel):
    """整表回传保存。"""
    columns: list[str]
    rows: list[dict]


class RowIn(BaseModel):
    """新增行。"""
    row: dict


class ExportLog(BaseModel):
    filename: str
    log: list[str]


# ============================================================
# 工具
# ============================================================
def _load_project_or_404(project_id: str, user: User, db: Session) -> Project:
    p = db.query(Project).filter(Project.id == project_id).first()
    if not p:
        raise HTTPException(status_code=404, detail="项目不存在")
    if p.owner.tenant_id != user.tenant_id:
        raise HTTPException(status_code=403, detail="无权访问")
    return p


def _session_path(project_id: str) -> str:
    return os.path.join(SESSION_DIR, f"{project_id}_modules.json")


def _load_modules(project_id: str) -> dict | None:
    p = _session_path(project_id)
    if not os.path.exists(p):
        return None
    with open(p, "r", encoding="utf-8") as fp:
        return json.load(fp)


def _save_modules(project_id: str, data: dict):
    with open(_session_path(project_id), "w", encoding="utf-8") as fp:
        json.dump(data, fp, ensure_ascii=False, default=str)


def _df_to_records(df: pd.DataFrame) -> list[dict]:
    out = df.where(df.notna(), None)
    return [{str(k): v for k, v in r.items()} for r in out.to_dict(orient="records")]


# ============================================================
# 板块加载：读取项目所有上传文件的所有 Sheet，拆成板块
# ============================================================
@router.post("/{project_id}/load", response_model=list[ModuleInfo])
def load_modules(project_id: str, user: User = Depends(get_current_user),
                 db: Session = Depends(get_db)):
    """扫描项目所有已上传 Excel，按 Sheet 拆分成板块。"""
    p = _load_project_or_404(project_id, user, db)
    files = db.query(UploadFile).filter(UploadFile.project_id == p.id).all()
    if not files:
        raise HTTPException(status_code=400, detail="请先上传至少一个 Excel 文件")

    modules: list[dict] = []
    for f in files:
        fp = os.path.join(UPLOAD_DIR, f.stored_path)
        if not os.path.exists(fp):
            continue
        try:
            # 读取所有 sheet
            xls = pd.ExcelFile(fp)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(fp, sheet_name=sheet_name)
                if df.empty:
                    continue
                # 列名转字符串
                df.columns = [str(c) for c in df.columns]
                mod_id = uuid.uuid4().hex[:12]
                modules.append({
                    "id": mod_id,
                    "name": sheet_name,
                    "source_file": f.original_name,
                    "source_sheet": sheet_name,
                    "columns": list(df.columns),
                    "rows": _df_to_records(df),
                })
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"解析 {f.original_name} 失败：{e}")

    _save_modules(project_id, {"modules": modules, "loaded_at": datetime.now().isoformat()})

    return [
        ModuleInfo(
            id=m["id"], name=m["name"], source_file=m["source_file"],
            source_sheet=m["source_sheet"],
            row_count=len(m["rows"]), col_count=len(m["columns"]),
        )
        for m in modules
    ]


# ============================================================
# 板块列表
# ============================================================
@router.get("/{project_id}", response_model=list[ModuleInfo])
def list_modules(project_id: str, user: User = Depends(get_current_user),
                 db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_modules(p.id)
    if not data:
        return []
    return [
        ModuleInfo(
            id=m["id"], name=m["name"], source_file=m["source_file"],
            source_sheet=m["source_sheet"],
            row_count=len(m["rows"]), col_count=len(m["columns"]),
        )
        for m in data["modules"]
    ]


# ============================================================
# 获取单个板块数据
# ============================================================
@router.get("/{project_id}/{module_id}", response_model=ModuleData)
def get_module(project_id: str, module_id: str,
               user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_modules(p.id)
    if not data:
        raise HTTPException(status_code=400, detail="尚未加载板块，请先点击导入")
    for m in data["modules"]:
        if m["id"] == module_id:
            return ModuleData(
                id=m["id"], name=m["name"], source_file=m["source_file"],
                source_sheet=m["source_sheet"],
                columns=m["columns"], rows=m["rows"],
            )
    raise HTTPException(status_code=404, detail="板块不存在")


# ============================================================
# 整表保存（用户编辑后回传）
# ============================================================
@router.put("/{project_id}/{module_id}", response_model=ModuleData)
def save_module(project_id: str, module_id: str, payload: ModuleSaveIn,
                user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_modules(p.id)
    if not data:
        raise HTTPException(status_code=400, detail="尚未加载板块")
    for m in data["modules"]:
        if m["id"] == module_id:
            m["columns"] = payload.columns
            m["rows"] = payload.rows
            _save_modules(p.id, data)
            return ModuleData(
                id=m["id"], name=m["name"], source_file=m["source_file"],
                source_sheet=m["source_sheet"],
                columns=m["columns"], rows=m["rows"],
            )
    raise HTTPException(status_code=404, detail="板块不存在")


# ============================================================
# 新增行
# ============================================================
@router.post("/{project_id}/{module_id}/rows", response_model=ModuleData)
def add_row(project_id: str, module_id: str, payload: RowIn,
            user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_modules(p.id)
    if not data:
        raise HTTPException(status_code=400, detail="尚未加载板块")
    for m in data["modules"]:
        if m["id"] == module_id:
            m["rows"].append(payload.row)
            _save_modules(p.id, data)
            return ModuleData(
                id=m["id"], name=m["name"], source_file=m["source_file"],
                source_sheet=m["source_sheet"],
                columns=m["columns"], rows=m["rows"],
            )
    raise HTTPException(status_code=404, detail="板块不存在")


# ============================================================
# 删除行（按索引，从 0 开始）
# ============================================================
@router.delete("/{project_id}/{module_id}/rows/{row_idx}", response_model=ModuleData)
def delete_row(project_id: str, module_id: str, row_idx: int,
               user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_modules(p.id)
    if not data:
        raise HTTPException(status_code=400, detail="尚未加载板块")
    for m in data["modules"]:
        if m["id"] == module_id:
            if row_idx < 0 or row_idx >= len(m["rows"]):
                raise HTTPException(status_code=400, detail="行索引越界")
            m["rows"].pop(row_idx)
            _save_modules(p.id, data)
            return ModuleData(
                id=m["id"], name=m["name"], source_file=m["source_file"],
                source_sheet=m["source_sheet"],
                columns=m["columns"], rows=m["rows"],
            )
    raise HTTPException(status_code=404, detail="板块不存在")


# ============================================================
# 重命名板块
# ============================================================
class RenameIn(BaseModel):
    name: str


@router.patch("/{project_id}/{module_id}/rename", response_model=ModuleInfo)
def rename_module(project_id: str, module_id: str, payload: RenameIn,
                  user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_modules(p.id)
    if not data:
        raise HTTPException(status_code=400, detail="尚未加载板块")
    for m in data["modules"]:
        if m["id"] == module_id:
            m["name"] = payload.name
            _save_modules(p.id, data)
            return ModuleInfo(
                id=m["id"], name=m["name"], source_file=m["source_file"],
                source_sheet=m["source_sheet"],
                row_count=len(m["rows"]), col_count=len(m["columns"]),
            )
    raise HTTPException(status_code=404, detail="板块不存在")


# ============================================================
# 导出：把所有板块写回 Excel（每板块一个 Sheet）
# ============================================================
@router.post("/{project_id}/export", response_model=ExportLog)
def export_modules(project_id: str, user: User = Depends(get_current_user),
                   db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_modules(p.id)
    if not data or not data["modules"]:
        raise HTTPException(status_code=400, detail="没有可导出的板块")

    log: list[str] = []
    wb = openpyxl.Workbook()
    # 删除默认 sheet
    default_ws = wb.active
    wb.remove(default_ws)

    for m in data["modules"]:
        # Sheet 名不能重复且 ≤31 字符
        sheet_name = m["name"][:31] or "Sheet"
        ws = wb.create_sheet(title=sheet_name)
        # 写表头
        ws.append(m["columns"])
        # 写数据
        for row in m["rows"]:
            ws.append([row.get(c) for c in m["columns"]])
        log.append(f"板块 [{m['name']}]：{len(m['rows'])} 行 × {len(m['columns'])} 列")

    os.makedirs(os.path.join(EXPORT_DIR, p.id), exist_ok=True)
    out_name = f"薪资数据_{p.salary_month}_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    out_path = os.path.join(EXPORT_DIR, p.id, out_name)
    wb.save(out_path)
    log.append(f"导出完成：{out_name}")

    # 记录下载信息
    export_meta = {
        "path": os.path.join(p.id, out_name),
        "filename": out_name,
    }
    with open(os.path.join(SESSION_DIR, f"{p.id}_export.json"), "w", encoding="utf-8") as fp:
        json.dump(export_meta, fp, ensure_ascii=False)

    return ExportLog(filename=out_name, log=log)


@router.get("/{project_id}/export/download")
def download_export(project_id: str, user: User = Depends(get_current_user),
                    db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    meta_path = os.path.join(SESSION_DIR, f"{p.id}_export.json")
    if not os.path.exists(meta_path):
        raise HTTPException(status_code=400, detail="尚未生成导出文件")
    with open(meta_path, "r", encoding="utf-8") as fp:
        meta = json.load(fp)
    fp_full = os.path.join(EXPORT_DIR, meta["path"])
    if not os.path.exists(fp_full):
        raise HTTPException(status_code=404, detail="导出文件已丢失")
    with open(fp_full, "rb") as f:
        data = f.read()
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{meta["filename"]}"'},
    )
