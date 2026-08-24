"""实体化数据工作台路由。

核心理念：源数据 → 5个业务实体（结构化存储）→ 按模板导出。

流程：
    1. 用户上传源文件（薪资/社保）+ 导出模板
    2. POST /import 触发导入：源数据解析为5个实体
    3. 用户在各实体板块中编辑数据
    4. POST /export 按模板导出：实体数据 → 模板「工资核算」Sheet
"""
import io
import json
import os
import uuid
from datetime import datetime

import openpyxl
import pandas as pd
from fastapi import APIRouter, Depends, HTTPException, UploadFile as FastUploadFile, File, Form
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, ConfigDict
from sqlalchemy.orm import Session

from backend.auth import get_current_user
from backend.database import DATA_DIR, EXPORT_DIR, UPLOAD_DIR, get_db
from backend.models import User, Project, UploadFile, EntitySnapshot
from core.entity_engine import (
    import_file_to_entities,
    import_template_to_entities,
    merge_entities,
    export_to_template,
    get_entity_overview,
)
from core.schema_config import ALL_ENTITIES, ENTITY_BY_ID

router = APIRouter(prefix="/api/entities", tags=["entities"])

SESSION_DIR = os.path.join(DATA_DIR, "sessions")
os.makedirs(SESSION_DIR, exist_ok=True)


# ============================================================
# 数据模型
# ============================================================
class EntityOverview(BaseModel):
    id: str
    name: str
    icon: str
    primary_key: str
    row_count: int
    field_count: int
    completeness_pct: float


class EntityData(BaseModel):
    id: str
    name: str
    icon: str
    primary_key: str
    fields: list[dict]  # [{name, dtype, required}]
    columns: list[str]  # 字段名列表
    rows: list[dict]    # 数据行


class EntitySaveIn(BaseModel):
    columns: list[str]
    rows: list[dict]


class ExportResult(BaseModel):
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
    return os.path.join(SESSION_DIR, f"{project_id}_entities.json")


def _load_entities(project_id: str) -> dict | None:
    p = _session_path(project_id)
    if not os.path.exists(p):
        return None
    with open(p, "r", encoding="utf-8") as fp:
        return json.load(fp)


def _save_entities(project_id: str, data: dict):
    with open(_session_path(project_id), "w", encoding="utf-8") as fp:
        json.dump(data, fp, ensure_ascii=False, default=str)


def _find_template_files(project_id: str, db: Session) -> list[tuple[UploadFile, str]]:
    """查找项目的全部有效模板，按上传时间从早到晚排列。"""
    files = db.query(UploadFile).filter(
        UploadFile.project_id == project_id,
        UploadFile.file_type == "template",
    ).order_by(UploadFile.created_at.asc()).all()
    return [
        (file, path)
        for file in files
        if os.path.exists(path := os.path.join(UPLOAD_DIR, file.stored_path))
    ]


def _find_template_file(project_id: str, db: Session) -> str | None:
    """返回最新上传的模板，作为最终导出的主模板。"""
    templates = _find_template_files(project_id, db)
    return templates[-1][1] if templates else None


# ============================================================
# 上传导出模板
# ============================================================
@router.post("/{project_id}/template", response_model=dict, status_code=201)
def upload_template(project_id: str,
                    file: FastUploadFile = File(...),
                    user: User = Depends(get_current_user),
                    db: Session = Depends(get_db)):
    """上传导出模板 Excel 文件。"""
    p = _load_project_or_404(project_id, user, db)

    content = file.file.read()
    # 验证是有效的 Excel
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content))
        if "工资核算" not in wb.sheetnames:
            raise HTTPException(status_code=400, detail="模板中缺少「工资核算」Sheet")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Excel 解析失败：{e}")

    # 保存模板；项目允许同时保留多份模板
    ext = os.path.splitext(file.filename or "template.xlsx")[1] or ".xlsx"
    stored_name = f"{uuid.uuid4().hex}{ext}"
    stored_path = os.path.join(p.id, stored_name)
    abs_path = os.path.join(UPLOAD_DIR, stored_path)
    os.makedirs(os.path.dirname(abs_path), exist_ok=True)
    with open(abs_path, "wb") as f:
        f.write(content)

    rec = UploadFile(
        project_id=p.id,
        file_type="template",
        original_name=file.filename or "template.xlsx",
        stored_path=stored_path,
        row_count=0,
        col_count=0,
    )
    db.add(rec)
    db.commit()
    db.refresh(rec)

    return {"id": rec.id, "filename": rec.original_name, "message": "模板上传成功"}


# ============================================================
# 导入：将所有上传的源文件解析为实体
# ============================================================
@router.post("/{project_id}/import", response_model=list[EntityOverview])
def import_entities(project_id: str, user: User = Depends(get_current_user),
                    db: Session = Depends(get_db)):
    """导入流程：

    - 首次使用（无历史快照）：导入模板初始数据 + 源数据
    - 后续月份（有历史快照）：克隆最近快照 + 导入新源数据
    """
    p = _load_project_or_404(project_id, user, db)

    all_logs: list[str] = []
    collected: dict[str, list[dict]] = {}

    # ---- 检查是否有历史快照 ----
    latest_snap = db.query(EntitySnapshot).filter(
        EntitySnapshot.project_id == p.id
    ).order_by(EntitySnapshot.salary_month.desc()).first()

    if latest_snap:
        # 有历史快照：克隆上月数据作为基础
        all_logs.append(f"克隆上月快照: {latest_snap.salary_month} ({latest_snap.label})")
        snap_data = json.loads(latest_snap.data_json)
        snap_entities = snap_data.get("entities", {})
        # 清空月度动态字段
        for eid in ["attendance_record", "tax_record"]:
            if eid in snap_entities:
                snap_entities[eid] = []
        for rec in snap_entities.get("employee_profile", []):
            rec["本月状态"] = "在职"
            rec.pop("离职日期", None)
        for eid, records in snap_entities.items():
            collected.setdefault(eid, []).extend(records)
    else:
        # 无历史快照：首次使用，导入模板初始数据
        templates = _find_template_files(p.id, db)
        if templates:
            for template, template_path in templates:
                all_logs.append(f"首次导入模板初始数据: {template.original_name}")
                try:
                    tpl_result = import_template_to_entities(template_path)
                    for eid, records in tpl_result.items():
                        if eid == "_log":
                            all_logs.extend(records)
                            continue
                        collected.setdefault(eid, []).extend(records)
                except Exception as e:
                    all_logs.append(f"模板解析失败: {e}")
        else:
            all_logs.append("未上传导出模板，跳过初始数据导入")

    # ---- 导入源数据文件 ----
    files = db.query(UploadFile).filter(
        UploadFile.project_id == p.id,
        UploadFile.file_type != "template",
    ).all()

    if not files and not collected:
        raise HTTPException(status_code=400, detail="请先上传导出模板或源数据文件")

    for f in files:
        fp = os.path.join(UPLOAD_DIR, f.stored_path)
        if not os.path.exists(fp):
            all_logs.append(f"文件丢失: {f.original_name}")
            continue

        all_logs.append(f"处理源文件: {f.original_name} (类型: {f.file_type})")

        try:
            result = import_file_to_entities(fp, f.original_name)
            for eid, records in result.items():
                if eid == "_log":
                    all_logs.extend(records)
                    continue
                collected.setdefault(eid, []).extend(records)
        except Exception as e:
            all_logs.append(f"解析失败: {e}")

    # 按主键去重合并
    merged = merge_entities(collected)

    # 保存到 session
    session_data = {
        "entities": merged,
        "imported_at": datetime.now().isoformat(),
        "logs": all_logs,
    }
    _save_entities(p.id, session_data)

    return get_entity_overview(merged)


# ============================================================
# 实体概览列表
# ============================================================
@router.get("/{project_id}", response_model=list[EntityOverview])
def list_entities(project_id: str, user: User = Depends(get_current_user),
                  db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_entities(p.id)
    if not data:
        return []
    return get_entity_overview(data.get("entities", {}))


# ============================================================
# 获取单个实体数据（用于编辑）
# ============================================================
@router.get("/{project_id}/{entity_id}", response_model=EntityData)
def get_entity(project_id: str, entity_id: str,
               user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_entities(p.id)
    if not data:
        raise HTTPException(status_code=400, detail="尚未导入数据，请先点击导入")

    entity_def = ENTITY_BY_ID.get(entity_id)
    if entity_def is None:
        raise HTTPException(status_code=404, detail="实体不存在")

    records = data["entities"].get(entity_id, [])
    fields = [{"name": f.name, "dtype": f.dtype, "required": f.required} for f in entity_def.fields]
    columns = [f.name for f in entity_def.fields]

    # 过滤掉溯源字段，只返回定义的字段
    clean_rows = []
    for rec in records:
        clean = {k: v for k, v in rec.items() if not k.startswith("_")}
        clean_rows.append(clean)

    return EntityData(
        id=entity_def.id,
        name=entity_def.name,
        icon=entity_def.icon,
        primary_key=entity_def.primary_key,
        fields=fields,
        columns=columns,
        rows=clean_rows,
    )


# ============================================================
# 保存实体数据（用户编辑后回传）
# ============================================================
@router.put("/{project_id}/{entity_id}", response_model=EntityData)
def save_entity(project_id: str, entity_id: str, payload: EntitySaveIn,
                user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_entities(p.id)
    if not data:
        raise HTTPException(status_code=400, detail="尚未导入数据")

    entity_def = ENTITY_BY_ID.get(entity_id)
    if entity_def is None:
        raise HTTPException(status_code=404, detail="实体不存在")

    # 更新实体数据
    data["entities"][entity_id] = payload.rows
    _save_entities(p.id, data)

    fields = [{"name": f.name, "dtype": f.dtype, "required": f.required} for f in entity_def.fields]
    return EntityData(
        id=entity_def.id,
        name=entity_def.name,
        icon=entity_def.icon,
        primary_key=entity_def.primary_key,
        fields=fields,
        columns=payload.columns,
        rows=payload.rows,
    )


# ============================================================
# 新增行
# ============================================================
class RowIn(BaseModel):
    row: dict


@router.post("/{project_id}/{entity_id}/rows", response_model=EntityData)
def add_row(project_id: str, entity_id: str, payload: RowIn,
            user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_entities(p.id)
    if not data:
        raise HTTPException(status_code=400, detail="尚未导入数据")

    entity_def = ENTITY_BY_ID.get(entity_id)
    if entity_def is None:
        raise HTTPException(status_code=404, detail="实体不存在")

    data["entities"].setdefault(entity_id, []).append(payload.row)
    _save_entities(p.id, data)

    fields = [{"name": f.name, "dtype": f.dtype, "required": f.required} for f in entity_def.fields]
    columns = [f.name for f in entity_def.fields]
    return EntityData(
        id=entity_def.id,
        name=entity_def.name,
        icon=entity_def.icon,
        primary_key=entity_def.primary_key,
        fields=fields,
        columns=columns,
        rows=data["entities"][entity_id],
    )


# ============================================================
# 删除行
# ============================================================
@router.delete("/{project_id}/{entity_id}/rows/{row_idx}")
def delete_row(project_id: str, entity_id: str, row_idx: int,
               user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_entities(p.id)
    if not data:
        raise HTTPException(status_code=400, detail="尚未导入数据")

    records = data["entities"].get(entity_id, [])
    if row_idx < 0 or row_idx >= len(records):
        raise HTTPException(status_code=400, detail="行索引越界")
    records.pop(row_idx)
    _save_entities(p.id, data)
    return {"ok": True, "remaining": len(records)}


# ============================================================
# 导入日志
# ============================================================
@router.get("/{project_id}/logs")
def get_logs(project_id: str, user: User = Depends(get_current_user),
             db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_entities(p.id)
    if not data:
        return {"logs": []}
    return {"logs": data.get("logs", []), "imported_at": data.get("imported_at")}


# ============================================================
# 导出：实体 → 模板
# ============================================================
@router.post("/{project_id}/export", response_model=ExportResult)
def export_entities(project_id: str, user: User = Depends(get_current_user),
                    db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    data = _load_entities(p.id)
    if not data or not data.get("entities"):
        raise HTTPException(status_code=400, detail="没有可导出的实体数据")

    # 查找模板文件
    template_path = _find_template_file(p.id, db)
    if not template_path:
        raise HTTPException(status_code=400, detail="请先上传导出模板文件")

    # 输出路径
    os.makedirs(os.path.join(EXPORT_DIR, p.id), exist_ok=True)
    out_name = f"工资核算_{p.salary_month}_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    out_path = os.path.join(EXPORT_DIR, p.id, out_name)

    # 执行导出
    try:
        log = export_to_template(
            template_path=template_path,
            entities=data["entities"],
            output_path=out_path,
            freeze_formulas=True,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"导出失败: {e}")

    # 记录下载信息
    export_meta = {
        "path": os.path.join(p.id, out_name),
        "filename": out_name,
    }
    with open(os.path.join(SESSION_DIR, f"{p.id}_entity_export.json"), "w", encoding="utf-8") as fp:
        json.dump(export_meta, fp, ensure_ascii=False)

    return ExportResult(filename=out_name, log=log)


# ============================================================
# 下载导出文件
# ============================================================
@router.get("/{project_id}/export/download")
def download_export(project_id: str, user: User = Depends(get_current_user),
                    db: Session = Depends(get_db)):
    p = _load_project_or_404(project_id, user, db)
    meta_path = os.path.join(SESSION_DIR, f"{p.id}_entity_export.json")
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


# ============================================================
# 快照管理：按月保存/加载/克隆实体数据
# ============================================================

class SnapshotInfo(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: str
    salary_month: str
    label: str | None
    row_count: int
    created_at: datetime

class SnapshotSaveIn(BaseModel):
    label: str | None = None


@router.get("/{project_id}/snapshots", response_model=list[SnapshotInfo])
def list_snapshots(project_id: str, user: User = Depends(get_current_user),
                   db: Session = Depends(get_db)):
    """列出项目所有已保存的实体数据快照。"""
    p = _load_project_or_404(project_id, user, db)
    snaps = db.query(EntitySnapshot).filter(
        EntitySnapshot.project_id == p.id
    ).order_by(EntitySnapshot.salary_month.desc()).all()
    return snaps


@router.post("/{project_id}/snapshots", response_model=SnapshotInfo, status_code=201)
def save_snapshot(project_id: str, payload: SnapshotSaveIn,
                  user: User = Depends(get_current_user),
                  db: Session = Depends(get_db)):
    """将当前 session 中的实体数据保存为月度快照。

    如果同月份已有快照则覆盖更新。
    """
    p = _load_project_or_404(project_id, user, db)
    data = _load_entities(p.id)
    if not data or not data.get("entities"):
        raise HTTPException(status_code=400, detail="当前没有可保存的实体数据")

    entities = data["entities"]
    total_rows = sum(len(v) for v in entities.values() if isinstance(v, list))
    data_json = json.dumps(data, ensure_ascii=False, default=str)

    # 查找是否已有同月快照
    existing = db.query(EntitySnapshot).filter(
        EntitySnapshot.project_id == p.id,
        EntitySnapshot.salary_month == p.salary_month,
    ).first()

    if existing:
        existing.data_json = data_json
        existing.row_count = total_rows
        existing.label = payload.label or existing.label
        db.commit()
        db.refresh(existing)
        return existing

    snap = EntitySnapshot(
        project_id=p.id,
        salary_month=p.salary_month,
        label=payload.label or f"{p.salary_month} 数据",
        data_json=data_json,
        row_count=total_rows,
    )
    db.add(snap)
    db.commit()
    db.refresh(snap)
    return snap


@router.post("/{project_id}/snapshots/{snapshot_id}/load", response_model=list[EntityOverview])
def load_snapshot(project_id: str, snapshot_id: str,
                  user: User = Depends(get_current_user),
                  db: Session = Depends(get_db)):
    """加载指定快照的数据到当前 session（用于查阅历史或作为新月份基础）。"""
    p = _load_project_or_404(project_id, user, db)
    snap = db.query(EntitySnapshot).filter(
        EntitySnapshot.id == snapshot_id,
        EntitySnapshot.project_id == p.id,
    ).first()
    if not snap:
        raise HTTPException(status_code=404, detail="快照不存在")

    data = json.loads(snap.data_json)
    _save_entities(p.id, data)
    return get_entity_overview(data.get("entities", {}))


@router.post("/{project_id}/snapshots/{snapshot_id}/clone", response_model=list[EntityOverview])
def clone_snapshot(project_id: str, snapshot_id: str,
                   user: User = Depends(get_current_user),
                   db: Session = Depends(get_db)):
    """克隆指定快照数据到当前月份（用于新月份基于上月数据修改）。

    克隆后清空本月状态字段，保留所有员工档案和薪资标准。
    """
    p = _load_project_or_404(project_id, user, db)
    snap = db.query(EntitySnapshot).filter(
        EntitySnapshot.id == snapshot_id,
        EntitySnapshot.project_id == p.id,
    ).first()
    if not snap:
        raise HTTPException(status_code=404, detail="快照不存在")

    data = json.loads(snap.data_json)
    entities = data.get("entities", {})

    # 克隆时清空本月动态字段（考勤、个税等月度数据），保留档案和薪资标准
    for eid in ["attendance_record", "tax_record"]:
        if eid in entities:
            entities[eid] = []

    # 清空员工档案中的本月状态（设为在职）
    for rec in entities.get("employee_profile", []):
        rec["本月状态"] = "在职"
        # 清空离职日期等月度字段
        rec.pop("离职日期", None)

    cloned_data = {
        "entities": entities,
        "imported_at": datetime.now().isoformat(),
        "logs": [f"克自快照 {snap.salary_month} ({snap.label})，已清空月度动态数据"],
    }
    _save_entities(p.id, cloned_data)
    return get_entity_overview(entities)


@router.delete("/{project_id}/snapshots/{snapshot_id}")
def delete_snapshot(project_id: str, snapshot_id: str,
                    user: User = Depends(get_current_user),
                    db: Session = Depends(get_db)):
    """删除指定快照。"""
    p = _load_project_or_404(project_id, user, db)
    snap = db.query(EntitySnapshot).filter(
        EntitySnapshot.id == snapshot_id,
        EntitySnapshot.project_id == p.id,
    ).first()
    if not snap:
        raise HTTPException(status_code=404, detail="快照不存在")
    db.delete(snap)
    db.commit()
    return {"ok": True}
